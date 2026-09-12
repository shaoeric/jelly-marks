#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""把打好的 exe 发布到 GitHub Release。

CI 里由 .github/workflows/release-windows.yml 调用;本地手动执行:
    GH_TOKEN=xxx GITHUB_REPOSITORY=owner/repo TAG=v0.1.0 \
        uv run python scripts/publish_release.py

为什么 exe 文件名用 ASCII、而不用中文:
GitHub 的 Release 上传接口会重命名文件名——官方文档原话是 "GitHub renames asset
filenames that have special characters, non-alphanumeric characters, and leading
or trailing periods"。「小小工具」是非字母数字字符,被剥离后主干清空,GitHub 就
填成了 default,于是资源变成 default.exe(gh 和本脚本两条独立路径都得到同样结果,
确认是服务端的既定行为,不是工具链问题)。
中文改由资源 label 呈现:label 在发布页会替代文件名显示,由 ASSET_LABEL 传入。
所以 exe 名必须是 ASCII(见 mark-tool.spec 的 EXE_NAME)。

为什么这里直接调 REST API 而不是 gh CLI:
原先用 gh 时也需要靠这个接口的 ?name= 参数指定资源名(gh 手册里的 `路径#文字`
设置的是 display label,不是资源名)。直接用 urllib 可以少一层依赖,也便于把
地址白名单、重试与回查都写在一起。

遇到 422 already_exists 的处理:
实测 Release 上只列出一个 default.exe,但目标名字却被认为已占用——资源名的唯一性
记录与列表显示的名字对不上,按显示名去找同名资源找不到冲突源,所以这种情况直接
清掉 Release 上所有 exe 资源再重试一次。

上传完成后会回查资源名:名字不对就让任务失败,并且不删任何东西——
那样用户至少还能把文件下载下来改名,比 Release 上一个 exe 都不剩要好。
"""

import glob
import ipaddress
import json
import os
import socket
import urllib.error
import urllib.parse
import urllib.request

API = "https://api.github.com"
UPLOADS = "https://uploads.github.com"
TIMEOUT = 180

# 只允许访问这两个固定域名,防止请求被引向别处
ALLOWED_HOSTS = frozenset({"api.github.com", "uploads.github.com"})


class AssetConflict(Exception):
    """目标资源名已被占用(HTTP 422 already_exists)。"""


class _NoRedirect(urllib.request.HTTPRedirectHandler):
    """这几个接口不需要跳转;一律不跟随,避免被重定向到非白名单主机。"""

    def redirect_request(self, req, fp, code, msg, headers, newurl):
        return None


_OPENER = urllib.request.build_opener(_NoRedirect)


def check_host_is_public(host):
    """解析主机名,拒绝本机 / 环回 / 私有 / 链路本地 / 保留 / 组播地址。"""
    try:
        infos = socket.getaddrinfo(host, None)
    except socket.gaierror as exc:
        raise SystemExit("::error::无法解析主机 %s: %s" % (host, exc))
    for info in infos:
        addr = ipaddress.ip_address(info[4][0])
        if (addr.is_loopback or addr.is_private or addr.is_link_local
                or addr.is_reserved or addr.is_multicast
                or addr.is_unspecified):
            raise SystemExit("::error::拒绝访问非公网地址: %s (%s)"
                             % (host, addr))


def check_url(url):
    """发请求前的地址校验:必须 https、域名在白名单内、且解析到公网地址。"""
    parts = urllib.parse.urlsplit(url)
    if parts.scheme != "https":
        raise SystemExit("::error::只允许 https 请求: %s" % url)
    host = parts.hostname or ""
    if host not in ALLOWED_HOSTS:
        raise SystemExit("::error::拒绝请求白名单之外的地址: %s" % url)
    check_host_is_public(host)


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


def request(method, url, data=None, content_type=None, want_json=True,
            fatal=True):
    """发一个带鉴权的请求。

    404 返回 None,交给调用方判断。fatal=False 时其余错误只打印告警并返回
    None,用于 label 这类"失败也不该影响发布"的附加操作。
    """
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
        # 资源名冲突单独抛出,交给上层做清理重试
        if exc.code == 422 and "already_exists" in detail:
            raise AssetConflict(detail)
        # 上面禁用了跳转。万一 GitHub 真返回跳转,把目标地址打出来,
        # 免得只看到一个光秃秃的状态码
        location = exc.headers.get("Location") if exc.headers else None
        if location:
            detail += "\n(服务端要求跳转到 %s,已被拒绝)" % location
        message = "::error::%s %s 失败: HTTP %d\n%s" % (
            method, url, exc.code, detail)
        if not fatal:
            print("警告:" + message.replace("::error::", ""))
            return None
        raise SystemExit(message)
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
    """列出 Release 上的全部资源(显式分页,避免默认只看最近一批)。"""
    url = "%s/repos/%s/releases/%d/assets?per_page=100" % (
        API, repo, release_id)
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


def drop_exe_assets(repo, release_id, keep_id=None):
    """删掉 Release 上现有的 exe 资源(只可能是本工具历史上传的)。"""
    for asset in list_assets(repo, release_id):
        if asset["id"] == keep_id:
            continue
        if asset["name"].lower().endswith(".exe"):
            print("  删除资源: %r (id=%s)" % (asset["name"], asset["id"]))
            delete_asset(repo, asset["id"])


def set_asset_label(repo, asset_id, label):
    """给资源设一个展示用 label(发布页用它替代文件名显示)。

    best-effort:失败只告警,不影响发布本身。
    """
    url = "%s/repos/%s/releases/assets/%d" % (API, repo, asset_id)
    payload = json.dumps({"label": label}).encode("utf-8")
    result = request("PATCH", url, data=payload,
                     content_type="application/json", fatal=False)
    if result is not None:
        print("已设置展示标签: %r" % result.get("label"))
    return result


def upload_resolving_conflict(repo, release_id, exe, name):
    """上传;若名字已被占用则清掉现有 exe 资源后重试一次。"""
    try:
        return upload_asset(repo, release_id, exe, name)
    except AssetConflict:
        # 名字的唯一性记录可能和列表里显示的名字对不上,光按同名去找是找不到的,
        # 所以这里直接清掉所有 exe 资源再来一次。
        print("资源名 %r 被占用,清理 Release 上的 exe 资源后重试" % name)
        drop_exe_assets(repo, release_id)
        return upload_asset(repo, release_id, exe, name)


def publish(repo, tag, exe, prerelease=False):
    """创建或更新 Release 并上传 exe;返回最终资源名列表。"""
    name = os.path.basename(exe)

    release = get_release(repo, tag)
    if release is None:
        print("Release %s 不存在,创建" % tag)
        release = create_release(repo, tag, prerelease)
    release_id = release["id"]

    existing = list_assets(repo, release_id)
    print("Release 现有资源: %s" % [a["name"] for a in existing])

    # 先删掉同名旧资源,等价于 gh 的 --clobber
    for asset in existing:
        if asset["name"] == name:
            print("删除同名旧资源: %r" % asset["name"])
            delete_asset(repo, asset["id"])

    uploaded = upload_resolving_conflict(repo, release_id, exe, name)
    print("上传完成: %r (asset id=%s)" % (uploaded["name"], uploaded["id"]))

    # 发布页用它替代文件名显示,这样文件名是 ASCII、界面上仍是中文
    label = os.environ.get("ASSET_LABEL", "").strip()
    if label:
        set_asset_label(repo, uploaded["id"], label)

    # 先确认期望的资源名真的在 Release 上,再清理历史错名资源。
    # 顺序很关键:万一名字又被改写,这里必须先失败退出,不能顺手把刚上传的
    # 那个 exe 也删掉——那样 Release 上一个 exe 都不剩,比留着一个名字错、
    # 内容正确的文件更糟(用户至少还能下载下来改名)。
    assets = list_assets(repo, release_id)
    names = [a["name"] for a in assets]
    if name not in names:
        raise SystemExit(
            "::error::资源名应为 %r,实际为 %s。GitHub 会重命名含非字母数字字符的"
            "资源名(中文会被剥离,主干清空后退化成 default),所以 exe 文件名必须"
            "是 ASCII——请检查 mark-tool.spec 里的 EXE_NAME。"
            "未删除任何资源,可先手动下载确认内容是否正常。" % (name, names))

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
