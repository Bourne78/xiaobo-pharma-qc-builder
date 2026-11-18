# -*- coding: utf-8 -*-
"""
通过本会话的 GitHub MCP 代理端点**全量同步**技能目录。
令牌取自 CODEBUDDY_MCP_CONFIG（仅进程内使用，不落盘、不打印）。
文件内容由本脚本自己读取，不经对话搬运。

技能根目录默认自动探测为“本脚本所在 tools/ 的上级”（clone 后开箱即用）；
若把脚本挪出技能包，用 --skill-root 显式指定。

用法（在技能根执行）：
  python tools/push_skill_via_mcp.py --dry-run      # 只列清单/分批，不连 MCP、不推送
  python tools/push_skill_via_mcp.py                # 全量分批 push_files，随后逐文件 sha 校验
  python tools/push_skill_via_mcp.py --verify-only  # 只逐文件比对远端 blob sha
可选覆盖：--skill-root / --owner / --repo / --branch / --message
（fork 到自己账号时改 --owner/--repo；注意本仓库含内部 QC 信息，慎推公开 fork）
"""
import os, re, sys, json, hashlib, argparse, urllib.request, urllib.error
from pathlib import Path

DEFAULT_SKILL_ROOT = Path(__file__).resolve().parent.parent  # tools/ 的上级 = 技能根
OWNER, REPO, BRANCH = "Bourne78", "xiaobo-pharma-qc-builder", "main"
DEFAULT_MESSAGE = "chore(skill): full skill dir sync via MCP push_files"

ALLOW_EXT = {".py", ".ps1", ".md", ".json", ".txt"}
EXCLUDE_DIRS = {"__pycache__", ".git", ".workbuddy", ".idea", ".vscode"}
EXCLUDE_NAMES = {"thumbs.db", "desktop.ini", ".ds_store"}
EXCLUDE_SUFFIX = (".pyc", ".pyo", ".tmp", ".bak", ".zip", ".swp")
MAX_BYTES = 1024 * 1024          # GitHub Contents API 单文件上限
BATCH_MAX_FILES = 8
BATCH_MAX_CHARS = 200000


# ---------- 文件收集 ----------
def collect_files(root):
    """返回 [(rel_path, content)]，rel 用 '/' 分隔，按 rel 排序。"""
    out = []
    for dirpath, dirs, fs in os.walk(root):
        dirs[:] = [d for d in dirs if d not in EXCLUDE_DIRS and not d.startswith(".")]
        for name in fs:
            full = os.path.join(dirpath, name)
            rel = os.path.relpath(full, root).replace(os.sep, "/")
            low = name.lower()
            if name.startswith(".") or low in EXCLUDE_NAMES or low.endswith("~"):
                continue
            if low.endswith(EXCLUDE_SUFFIX):
                continue
            if os.path.splitext(name)[1].lower() not in ALLOW_EXT:
                continue
            size = os.path.getsize(full)
            if size > MAX_BYTES:
                raise RuntimeError("文件超过 1MB 上限：%s (%d B)" % (rel, size))
            with open(full, "r", encoding="utf-8", newline="") as f:
                content = f.read().replace("\r\n", "\n")
            out.append((rel, content))
    out.sort(key=lambda x: x[0])
    return out


def make_batches(files):
    batches, cur, chars = [], [], 0
    for rel, content in files:
        n = len(content.encode("utf-8"))
        if cur and (len(cur) >= BATCH_MAX_FILES or chars + n > BATCH_MAX_CHARS):
            batches.append(cur)
            cur, chars = [], 0
        cur.append((rel, content))
        chars += n
    if cur:
        batches.append(cur)
    return batches


def blob_sha(content):
    raw = content.encode("utf-8")
    return hashlib.sha1(b"blob " + str(len(raw)).encode() + b"\0" + raw).hexdigest()


# ---------- MCP RPC ----------
def rpc(url, headers, payload, session=None, timeout=180):
    h = dict(headers)
    h.update({"Content-Type": "application/json",
              "Accept": "application/json, text/event-stream"})
    if session:
        h["Mcp-Session-Id"] = session
    data = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    req = urllib.request.Request(url, data=data, headers=h, method="POST")
    with urllib.request.urlopen(req, timeout=timeout) as r:
        sid = r.headers.get("Mcp-Session-Id") or session
        raw = r.read().decode("utf-8", "ignore")
        return sid, raw


def parse_sse(raw):
    for line in raw.splitlines():
        if line.startswith("data:"):
            s = line[5:].strip()
            if s and s != "[DONE]":
                try:
                    return json.loads(s)
                except json.JSONDecodeError:
                    continue
    try:
        return json.loads(raw)
    except Exception:
        return None


def call_tool(url, headers, sid, rid, name, args):
    _, raw = rpc(url, headers, {
        "jsonrpc": "2.0", "id": rid, "method": "tools/call",
        "params": {"name": name, "arguments": args},
    }, session=sid)
    res = parse_sse(raw)
    if res is None:
        raise RuntimeError("无有效响应: %r" % (raw[:300],))
    if "error" in res:
        raise RuntimeError("RPC error: %s" % json.dumps(res["error"], ensure_ascii=False)[:300])
    result = res.get("result", {})
    if result.get("isError"):
        raise RuntimeError("tool error: %s" % json.dumps(result, ensure_ascii=False)[:500])
    return result


def commit_sha(result):
    # GitHub MCP 工具常返回 ref 对象（content[].text 为转义 JSON 字符串），
    # commit sha 在 object.sha；也可能直接给 commit 对象。两种都兼容。
    for c in result.get("content", []):
        if c.get("type") != "text":
            continue
        try:
            obj = json.loads(c.get("text", ""))
        except json.JSONDecodeError:
            continue
        if isinstance(obj, dict):
            o = obj.get("object")
            if isinstance(o, dict) and re.fullmatch(r"[0-9a-f]{40}", str(o.get("sha", ""))):
                return o["sha"]
            if re.fullmatch(r"[0-9a-f]{40}", str(obj.get("sha", ""))):
                return obj["sha"]
    # 兜底：正则兼容带反斜杠转义的 "sha" 键
    m = re.search(r'\\?"sha\\?"\s*:\s*\\?"([0-9a-f]{40})\\?"',
                  json.dumps(result, ensure_ascii=False))
    return m.group(1) if m else None


# ---------- 远端读取（校验用） ----------
def extract_remote_sha(result):
    """get_file_contents 的 sha 有两种来源：
    1) type=text 提示 'successfully downloaded text file (SHA: <40hex>)'；
    2) type=resource 里的 JSON 对象含 sha。两者都兼容。
    """
    for c in result.get("content", []):
        if c.get("type") == "text":
            m = re.search(r"SHA:\s*([0-9a-f]{40})", c.get("text", ""), re.I)
            if m:
                return m.group(1)
    for c in result.get("content", []):
        if c.get("type") == "resource":
            txt = c.get("resource", {}).get("text", "")
        elif c.get("type") == "text":
            txt = c.get("text", "")
        else:
            continue
        try:
            obj = json.loads(txt)
        except json.JSONDecodeError:
            continue
        if isinstance(obj, dict) and re.fullmatch(r"[0-9a-f]{40}", str(obj.get("sha", ""))):
            return obj["sha"]
    return None


# ---------- 主流程 ----------
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--dry-run", action="store_true", help="只列清单与分批，不连 MCP")
    ap.add_argument("--verify-only", action="store_true", help="只校验远端，不推送")
    ap.add_argument("--skill-root", default=str(DEFAULT_SKILL_ROOT),
                    help="技能根目录，默认=本脚本所在 tools/ 的上级")
    ap.add_argument("--owner", default=OWNER)
    ap.add_argument("--repo", default=REPO)
    ap.add_argument("--branch", default=BRANCH)
    ap.add_argument("--message", default=DEFAULT_MESSAGE)
    args = ap.parse_args()

    skill_root = Path(args.skill_root).resolve()
    if not (skill_root / "SKILL.md").is_file():
        print("[FAIL] 技能根无效（缺 SKILL.md）：%s；请用 --skill-root 指定" % skill_root)
        return 2
    files = collect_files(str(skill_root))
    batches = make_batches(files)
    total = sum(len(c.encode("utf-8")) for _, c in files)
    print("[plan] 文件数=%d 总字节=%d 批次=%d -> %s/%s@%s" %
          (len(files), total, len(batches), args.owner, args.repo, args.branch))
    for i, b in enumerate(batches, 1):
        print("  batch %d/%d: %d files, %s ... %s" %
              (i, len(batches), len(b), b[0][0], b[-1][0]))
    if args.dry_run:
        print("[dry-run] 只预演，未连接 MCP、未推送。")
        return 0

    cfg = json.loads(os.environ["CODEBUDDY_MCP_CONFIG"])
    g = cfg["mcpServers"]["github"]
    url, headers = g["url"], dict(g.get("headers", {}))

    sid, raw = rpc(url, headers, {
        "jsonrpc": "2.0", "id": 1, "method": "initialize",
        "params": {"protocolVersion": "2025-03-26", "capabilities": {},
                   "clientInfo": {"name": "wb-local-fullpush", "version": "3.0"}},
    })
    init = parse_sse(raw)
    assert init and "result" in init, "initialize 失败: %r" % (raw[:200],)
    print("[OK] initialize, session=%s" % (sid or "")[:16])
    try:
        rpc(url, headers, {"jsonrpc": "2.0", "method": "notifications/initialized"},
            session=sid)
    except Exception as e:
        print("[warn] initialized 通知:", type(e).__name__)

    shas = []
    if not args.verify_only:
        for i, batch in enumerate(batches, 1):
            payload_files = [{"path": rel, "content": c} for rel, c in batch]
            msg = args.message + (" (batch %d/%d)" % (i, len(batches)) if len(batches) > 1 else "")
            res = call_tool(url, headers, sid, 100 + i, "push_files", {
                "owner": args.owner, "repo": args.repo, "branch": args.branch,
                "message": msg, "files": payload_files,
            })
            sha = commit_sha(res)
            if not sha:
                print("[FAIL] batch %d 未拿到 commit sha: %s" %
                      (i, json.dumps(res, ensure_ascii=False)[:300]))
                return 1
            shas.append(sha)
            print("[OK] batch %d/%d commit=%s" % (i, len(batches), sha))
    else:
        print("[verify-only] 跳过推送，只校验远端。")

    # ---- 逐文件校验 ----
    print("[verify] 开始逐文件比对远端 blob sha ...")
    bad = []
    for rid, (rel, content) in enumerate(files, 200):
        res = call_tool(url, headers, sid, rid, "get_file_contents", {
            "owner": args.owner, "repo": args.repo, "branch": args.branch, "path": rel,
        })
        remote = extract_remote_sha(res)
        local = blob_sha(content)
        if remote != local:
            bad.append((rel, local, remote))
            print("  [MISMATCH] %s local=%s remote=%s" % (rel, local[:10], str(remote)[:10]))
    if bad:
        print("[FAIL] %d/%d 文件 sha 不一致" % (len(bad), len(files)))
        return 1
    print("[VERIFY OK] %d/%d 文件远端 blob sha 一致%s" %
          (len(files), len(files),
           ("；%d 个 commit 均有效" % len(shas)) if shas else ""))
    return 0


if __name__ == "__main__":
    sys.exit(main())
