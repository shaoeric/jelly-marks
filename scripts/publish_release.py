#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""把打好的 exe 发布到 GitHub Release。

CI 里由 .github/workflows/release-windows.yml 调用;也可以本地手动执行:
    GH_TOKEN=xxx TAG=v0.1.0 uv run python scripts/publish_release.py

为什么要在 Windows 上用 Python 去调用 gh,而不是直接在 bash 里调用:
中文文件名作为 argv 从 bash(msys)传给原生程序 gh 时会被破坏,gh 拿不到合法的
资源名,GitHub 于是把它退化成 default.exe(v0.1.0 那次就是这样发出去的)。
Python 用 UTF-16 传 argv,不经过那个边界;同时这里用 路径#资源名 显式指定名字,
并在上传后回查确认,名字不对就直接让任务失败,不会静默产出错名资源。
"""

import glob
import json
import os
import subprocess
import sys


def gh(*args):
    """调用 gh CLI,返回 CompletedProcess(输出按 UTF-8 解码)。"""
    return subprocess.run(["gh", *args], capture_output=True, text=True,
                          encoding="utf-8", errors="replace")


def find_exe(dist_dir="dist"):
    """取打包目录里的 exe;找不到就报错退出。"""
    paths = sorted(glob.glob(os.path.join(dist_dir, "*.exe")))
    if not paths:
        raise SystemExit("::error::%s 下没有找到 .exe 产物" % dist_dir)
    return paths[0]


def asset_names(tag):
    """Release 上现有的资源名;Release 不存在时返回 None。"""
    r = gh("release", "view", tag, "--json", "assets")
    if r.returncode != 0:
        return None
    return [a["name"] for a in json.loads(r.stdout)["assets"]]


def publish(tag, exe, prerelease=False):
    """创建或更新 Release 并上传 exe;返回最终资源名列表。"""
    name = os.path.basename(exe)
    labeled = "%s#%s" % (exe, name)

    if asset_names(tag) is None:
        args = ["release", "create", tag, labeled,
                "--title", tag, "--generate-notes"]
        if prerelease:
            args.append("--prerelease")
        action = "创建 Release 并上传"
    else:
        args = ["release", "upload", tag, labeled, "--clobber"]
        action = "覆盖上传"

    r = gh(*args)
    if r.returncode != 0:
        raise SystemExit("::error::%s 失败:\n%s\n%s"
                         % (action, r.stdout, r.stderr))
    print("%s完成: %s" % (action, name))

    # 先确认期望的资源名确实在 Release 上,再清理历史错名资源。
    # 顺序很关键:万一名字又被改写,这里必须先失败退出,不能顺手把刚上传的
    # 那个 exe 也删掉——那样 Release 上一个 exe 都不剩,比留着一个名字错、
    # 内容正确的文件更糟(用户至少还能下载下来改名)。
    final = asset_names(tag) or []
    if name not in final:
        raise SystemExit("::error::资源名应为 %r,实际为 %s。"
                         "未删除任何资源,可先手动下载确认内容是否正常。"
                         % (name, final))

    for old in final:
        if old != name and old.lower().endswith(".exe"):
            print("删除历史遗留的错名资源: %r" % old)
            gh("release", "delete-asset", tag, old, "--yes")

    final = asset_names(tag) or []
    print("Release 上的资源: %s" % final)
    print("资源名校验通过: %r" % name)
    return final


def main():
    tag = os.environ.get("TAG", "").strip()
    if not tag:
        raise SystemExit("::error::缺少环境变量 TAG")
    if not os.environ.get("GH_TOKEN"):
        raise SystemExit("::error::缺少环境变量 GH_TOKEN")
    prerelease = os.environ.get("PRERELEASE", "").strip().lower() == "true"
    exe = find_exe()
    print("本地产物: %r" % os.path.basename(exe))
    publish(tag, exe, prerelease)


if __name__ == "__main__":
    main()
