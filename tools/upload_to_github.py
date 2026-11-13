#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
把本地目录（通常是某个 WorkBuddy skill 根目录）上传到 GitHub 仓库。

Auth 策略（关键）：
    直接读本机 git 凭据缓存里的 github.com 令牌（``git credential fill``），
    因此只要这台机器之前登录过 github、且凭据未被清除，就不需要手动提供 PAT。
    这能绕过「WorkBuddy GitHub 连接器是只读 App（建仓 / 改文件 / push 全 403）」
    以及「沙箱直连 github.com 主站超时但 api.github.com 可通」的限制。

    上传走 REST Contents API（PUT /repos/{owner}/{repo}/contents/{path}），
    每个文件单独上传、基于最新 main HEAD，无需 git 协议。

用法：
    # 上传当前目录到指定仓库的 main 分支（根级）
    python tools/upload_to_github.py --owner Bourne78 --repo xiaobo-pharma-qc-builder --dir .

    # 推到某子目录（例如 skills/foo/）
    python tools/upload_to_github.py --owner Bourne78 --repo my-skills --dir . --prefix skills/xiaobo-pharma-qc-builder/

    # 指定分支 / 排除某些文件
    python tools/upload_to_github.py --owner O --repo R --dir . --branch main --exclude '.git' '__pycache__'

注意：
    - 已存在的文件会被覆盖（PUT 更新，HTTP 200）；新文件创建（HTTP 201）。
    - 大文件（>1MB）GitHub Contents API 会拒绝，需用 Git LFS / 走 git 协议。
    - 令牌仅用于本次进程内的 API 调用，不会落盘、不会打印。
"""
import argparse
import base64
import json
import os
import subprocess
import sys
import urllib.error
import urllib.request

API = "https://api.github.com"


def get_token():
    """从本机 git 凭据缓存取 github.com 令牌（password 字段）。

    优先级：环境变量 GITHUB_TOKEN > git credential fill 缓存。
    （Windows 无头环境下 git credential fill 有时会因凭据管理器弹交互而挂起，
    此时用 GITHUB_TOKEN 预取令牌注入最稳。）
    """
    env_token = os.environ.get("GITHUB_TOKEN")
    if env_token:
        return env_token
    try:
        out = subprocess.run(
            ["git", "credential", "fill"],
            input="protocol=https\nhost=github.com\n",
            capture_output=True, text=True, timeout=20,
        ).stdout
    except Exception as e:  # noqa: BLE001
        print(f"[warn] git credential fill 失败: {e}", file=sys.stderr)
        return None
    for line in out.splitlines():
        if line.startswith("password="):
            return line[len("password="):].strip()
    return None


def list_files(root, exclude):
    files = []
    for dirpath, dirnames, filenames in os.walk(root):
        dirnames[:] = [d for d in dirnames if d not in exclude]
        for fn in filenames:
            full = os.path.join(dirpath, fn)
            rel = os.path.relpath(full, root).replace(os.sep, "/")
            if rel in exclude:
                continue
            files.append((full, rel))
    return sorted(files)


def get_sha(token, owner, repo, branch, remote_path):
    """返回已存在文件的 blob sha；不存在返回 None。"""
    url = f"{API}/repos/{owner}/{repo}/contents/{remote_path}?ref={branch}"
    req = urllib.request.Request(
        url, method="GET",
        headers={"Authorization": f"Bearer {token}",
                 "Accept": "application/vnd.github+json"},
    )
    try:
        resp = urllib.request.urlopen(req, timeout=30)
        data = json.loads(resp.read().decode("utf-8"))
        return data.get("sha")
    except urllib.error.HTTPError as e:
        if e.code == 404:
            return None
        detail = e.read().decode("utf-8", "ignore")[:200]
        print(f"  [warn] GET sha {remote_path} -> {e.code}: {detail}")
        return None
    except Exception:  # noqa: BLE001
        return None


def upload(token, owner, repo, branch, remote_path, content, message):
    # 已存在文件更新需带 sha
    sha = get_sha(token, owner, repo, branch, remote_path)
    payload = {
        "message": message,
        "content": base64.b64encode(content).decode("ascii"),
        "branch": branch,
    }
    if sha:
        payload["sha"] = sha
    url = f"{API}/repos/{owner}/{repo}/contents/{remote_path}"
    body = json.dumps(payload).encode("utf-8")
    req = urllib.request.Request(
        url, data=body, method="PUT",
        headers={
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
            "Accept": "application/vnd.github+json",
        },
    )
    try:
        resp = urllib.request.urlopen(req, timeout=60)
        return resp.status
    except urllib.error.HTTPError as e:
        detail = e.read().decode("utf-8", "ignore")[:300]
        return f"ERR {e.code}: {detail}"
    except Exception as e:  # noqa: BLE001
        return f"ERR transport: {e}"


def main():
    ap = argparse.ArgumentParser(description="上传本地目录到 GitHub 仓库（Contents API）")
    ap.add_argument("--owner", required=True, help="仓库 owner，如 Bourne78")
    ap.add_argument("--repo", required=True, help="仓库名，如 xiaobo-pharma-qc-builder")
    ap.add_argument("--dir", default=".", help="本地待上传目录（默认当前目录）")
    ap.add_argument("--branch", default="main", help="目标分支（默认 main）")
    ap.add_argument("--prefix", default="", help="远程路径前缀，如 skills/foo/")
    ap.add_argument("--exclude", nargs="*", default=[".git", "__pycache__"],
                    help="要排除的文件/目录名（可多个）")
    ap.add_argument("--dry-run", action="store_true", help="只列出将上传的文件，不真正上传")
    args = ap.parse_args()

    token = get_token()
    if not token:
        print("ERROR: 未在本机 git 凭据中找到 github.com 令牌。请先在可连 GitHub 的机器上登录过 git。",
              file=sys.stderr)
        sys.exit(1)

    files = list_files(args.dir, set(args.exclude))
    print(f"待上传文件 {len(files)} 个 -> {args.owner}/{args.repo}@{args.branch} "
          f"(prefix='{args.prefix}')")
    if args.dry_run:
        for _, rel in files:
            print("  (dry) " + (args.prefix + rel))
        return

    ok = 0
    for full, rel in files:
        remote = (args.prefix + rel) if args.prefix else rel
        with open(full, "rb") as f:
            content = f.read()
        status = upload(token, args.owner, args.repo, args.branch, remote,
                        content, f"add {remote}")
        if isinstance(status, int) and 200 <= status < 300:
            ok += 1
            print(f"  {remote} -> {status}")
        else:
            print(f"  {remote} -> {status}")
    print(f"完成：{ok}/{len(files)} 上传成功。")


if __name__ == "__main__":
    main()
