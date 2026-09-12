#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""把打好的 exe 发布到 GitHub Release。

CI 里由 .github/workflows/release-windows.yml 调用;本地手动执行:
    GH_TOKEN=xxx GITHUB_REPOSITORY=owner/repo TAG=v0.1.0 \
        uv run python scripts/publish_release.py

为什么不用 gh CLI,而是直接调 GitHub REST API:
资源名叫「小小工具.exe」,是非 ASCII。用 gh 时这个名字会丢,服务端退化成
default.exe(v0.1.0 那两次都是这样)。注意 gh 手册里 `路径#文字` 设置的其实是
display label(显示标签),并不是资源名,所以它管不到这件事。
REST API 的上传接口把资源名放在 `?name=` 查询参数里(URL 编码的 UTF-8),
这是文档规定的机制,不依赖命令行工具怎么传 argv,也不依赖 HTTP 头能否承载
非 ASCII。

上传完成后会回查资源名:名字不对就让任务失败,并且不删任何东西——
那样用户至少还能把文件下载下来改名,比 Release 上一个 exe 都不剩要好。
"""

import glob
import json
import os
import urllib.error
import urllib.parse
import urllib.request

API = "https://api.github.com"
UPLOADS = "https://uploads.github.com"
TIMEOUT = 180

# 只允许访问这两个固定域名,防止请求被引向别处
ALLOWED_HOSTS = frozenset({"api.github.com", "uploads.github.com"})


class _NoRedirect(urllib.request.HTTPRedirectHandler):
    """这几个接口不需要跳转;一律不跟随,避免被重定向到非白名单主机。"""

    def redirect_request(self, req, fp, code, msg, headers, newurl):
        return None


_OPENER = urllib.request.build_opener(_NoRedirect)


def check_url(url):
    """校验目标地址:必须 https 且在白名单域名内。"""
    parts = urllib.parse.urlsplit(url)
    if parts.scheme != "https" or parts.hostname not in ALLOWED_HOSTS:
        raise SystemExit("::error::拒绝请求白名单之外的地址: %s" % url)


def token():
    value = os.environ.get("GH_TOKEN") or os.environ.get("GITHUB_TOKEN")
    if not value:
        raise SystemExit("::error::缺少环境变量 GH_TOKEN")
    return value


def repository():
    """返回 owner/name,并约束格式(只允许一段 owner 和一段仓库名)。"""
    value = os.environ.get("GITHUB_REPOSITORY", "").strip()
    owner, sep, name = value.partition("/")
    if not sep or not owner or not name or "/" in name:
        raise SystemExit("::error::缺少或非法的 GITHUB_REPOSITORY: %r" % value)
    return "%s/%s" % (owner, name)


def request(method, url, data=None, content_type=None, want_json=True):
    """发一个带鉴权的请求;404 返回 None,交给调用方判断。"""
    check_url(url)
    req = urllib.request.Request(url, data=data, method=method)
    req.add_header("Authorization", "Bearer %s" % token())
    req.add_header("Accept", "application/vnd.github+json")
    req.add_header("X-GitHub-Api-Version", "2022-11-28")
    if content_type:
        req.add_header("Content-Type", content_type)
    try:
        with _OPENER.open(req, timeout=TIMEOUT) as resp:
            body = resp.read()
    except urllib.error.HTTPError as exc:
        if exc.code == 404:
            return None
        detail = exc.read().decode("utf-8", "replace")
        # 上面禁用了跳转。万一 GitHub 真返回跳转,把目标地址打出来,
        # 免得只看到一个光秃秃的状态码
        location = exc.headers.get("Location") if exc.headers else None
        if location:
            detail += "\n(服务端要求跳转到 %s,已被拒绝)" % location
        raise SystemExit("::error::%s %s 失败: HTTP %d\n%s"
                         % (method, url, exc.code, detail))
    return json.loads(body) if want_json else body


def find_exe(dist_dir="dist"):
    paths = sorted(glob.glob(os.path.join(dist_dir, "*.exe")))
    if not paths:
        raise SystemExit("::error::%s 下没有找到 .exe 产物" % dist_dir)
    return paths[0]


def get_release(repo, tag):
    url = "%s/repos/%s/releases/tags/%s" % (API, repo, urllib.parse.quote(tag))
    return request("GET", url)


def create_release(repo, tag, prerelease):
    payload = json.dumps({
        "tag_name": tag,
        "name": tag,
        "generate_release_notes": True,
        "prerelease": bool(prerelease),
    }).encode("utf-8")
    return request("POST", "%s/repos/%s/releases" % (API, repo),
                   data=payload, content_type="application/json")


def list_assets(repo, release_id):
    url = "%s/repos/%s/releases/%d/assets" % (API, repo, release_id)
    return request("GET", url) or []


def delete_asset(repo, asset_id):
    url = "%s/repos/%s/releases/assets/%d" % (API, repo, asset_id)
    request("DELETE", url, want_json=False)


def upload_asset(repo, release_id, path, name):
    """上传文件,资源名由 ?name= 指定(URL 编码的 UTF-8)。"""
    with open(path, "rb") as fh:
        blob = fh.read()
    url = "%s/repos/%s/releases/%d/assets?name=%s" % (
        UPLOADS, repo, release_id, urllib.parse.quote(name, safe=""))
    return request("POST", url, data=blob,
                   content_type="application/octet-stream")


def publish(repo, tag, exe, prerelease=False):
    """创建或更新 Release 并上传 exe;返回最终资源名列表。"""
    name = os.path.basename(exe)

    release = get_release(repo, tag)
    if release is None:
        print("Release %s 不存在,创建" % tag)
        release = create_release(repo, tag, prerelease)
    release_id = release["id"]

    # 先删掉同名旧资源,等价于 gh 的 --clobber
    for asset in release.get("assets", []):
        if asset["name"] == name:
            print("删除同名旧资源: %r" % asset["name"])
            delete_asset(repo, asset["id"])

    uploaded = upload_asset(repo, release_id, exe, name)
    print("上传完成: %r (asset id=%s)" % (uploaded["name"], uploaded["id"]))

    # 先确认期望的资源名真的在 Release 上,再清理历史错名资源。
    # 顺序很关键:万一名字又被改写,这里必须先失败退出,不能顺手把刚上传的
    # 那个 exe 也删掉——那样 Release 上一个 exe 都不剩,比留着一个名字错、
    # 内容正确的文件更糟(用户至少还能下载下来改名)。
    assets = list_assets(repo, release_id)
    names = [a["name"] for a in assets]
    if name not in names:
        raise SystemExit("::error::资源名应为 %r,实际为 %s。"
                         "未删除任何资源,可先手动下载确认内容是否正常。"
                         % (name, names))

    for asset in assets:
        other = asset["name"]
        if other != name and other.lower().endswith(".exe"):
            print("删除历史遗留的错名资源: %r" % other)
            delete_asset(repo, asset["id"])

    final = [a["name"] for a in list_assets(repo, release_id)]
    print("Release 上的资源: %s" % final)
    print("资源名校验通过: %r" % name)
    return final


def main():
    tag = os.environ.get("TAG", "").strip()
    if not tag:
        raise SystemExit("::error::缺少环境变量 TAG")
    repo = repository()
    prerelease = os.environ.get("PRERELEASE", "").strip().lower() == "true"
    exe = find_exe()
    print("仓库: %s | 产物: %r" % (repo, os.path.basename(exe)))
    publish(repo, tag, exe, prerelease)


if __name__ == "__main__":
    main()
