#!/usr/bin/env python3
"""
LIMS API 接口扫描器
扫描代码中对 http://192.168.12.234:60015 的 HTTP 请求，
对比 Apifox 中已记录的端点，发现新接口。

用法:
  python scripts/lims_api_scanner.py          # 扫描并输出新端点
  python scripts/lims_api_scanner.py --json    # JSON 格式输出（供 CI 使用）
"""

import json
import os
import re
import sys
import io
import requests
from collections import defaultdict

if sys.platform == "win32":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_DIR = os.path.dirname(SCRIPT_DIR)
OPENAPI_FILE = os.path.join(PROJECT_DIR, "LIMS_API_openapi.json")

LIMS_HOST = "192.168.12.234"
LIMS_PORT = "60015"

REQUEST_PATTERN = re.compile(
    r'(?:session|requests|self\.\w*session\w*)\.'
    r'(get|post|put|delete|patch)\s*\('
    r'[^"\']*["\']'
    r'([^"\']*)',
    re.IGNORECASE
)

URL_ASSIGN_PATTERN = re.compile(
    r'(?:url|endpoint|api_url|req_url)\s*=\s*["\']([^"\']*/detectionManager/[^"\']*)["\']',
    re.IGNORECASE
)


def get_known_endpoints_from_apifox():
    """从 Apifox API 获取已知端点（优先），失败则读本地文件"""
    token = os.environ.get("APIFOX_API_KEY", "")
    project_id = os.environ.get("PROJECT_ID", "8238274")

    if token:
        try:
            headers = {
                "Authorization": f"Bearer {token}",
                "X-Apifox-Api-Version": "2024-03-28",
                "Content-Type": "application/json",
            }
            resp = requests.post(
                f"https://api.apifox.com/v1/projects/{project_id}/export-openapi?locale=zh-CN",
                headers=headers,
                json={"version": "3.0", "format": "json"},
                timeout=15,
            )
            if resp.status_code == 200:
                spec = resp.json()
                endpoints = set()
                for path, methods in spec.get("paths", {}).items():
                    for method in methods:
                        if method.lower() in ("get", "post", "put", "delete", "patch"):
                            endpoints.add((method.upper(), path))
                print(f"  从 Apifox 获取到 {len(endpoints)} 个已知端点", file=sys.stderr)
                return endpoints
        except Exception as e:
            print(f"  Apifox API 获取失败: {e}，回退到本地文件", file=sys.stderr)

    # 回退到本地文件
    if os.path.exists(OPENAPI_FILE):
        with open(OPENAPI_FILE, "r", encoding="utf-8") as f:
            spec = json.load(f)
        endpoints = set()
        for path, methods in spec.get("paths", {}).items():
            for method in methods:
                if method.lower() in ("get", "post", "put", "delete", "patch"):
                    endpoints.add((method.upper(), path))
        print(f"  从本地文件获取到 {len(endpoints)} 个已知端点", file=sys.stderr)
        return endpoints

    print(f"  警告: 无法获取已知端点", file=sys.stderr)
    return set()


def extract_path(full_url):
    match = re.search(r'(/detectionManager/[^\s"\'`,)}\]]+)', full_url)
    if match:
        path = match.group(1).split("?")[0].rstrip("/")
        return path
    return None


def normalize_path(path):
    return re.sub(r'/\d+', '/{id}', path)


def scan_py_files():
    found = []
    for root, dirs, files in os.walk(PROJECT_DIR):
        dirs[:] = [d for d in dirs if d not in (".venv", "__pycache__", "node_modules", ".git")]
        for fname in files:
            if not fname.endswith(".py"):
                continue
            filepath = os.path.join(root, fname)
            rel_path = os.path.relpath(filepath, PROJECT_DIR)
            with open(filepath, "r", encoding="utf-8", errors="replace") as f:
                for line_num, line in enumerate(f, 1):
                    stripped = line.strip()
                    if stripped.startswith("#"):
                        continue
                    for match in REQUEST_PATTERN.finditer(line):
                        method = match.group(1).upper()
                        url_str = match.group(2)
                        path = extract_path(url_str)
                        if path:
                            found.append((method, path, rel_path, line_num))
                    for match in URL_ASSIGN_PATTERN.finditer(line):
                        url_str = match.group(1)
                        path = extract_path(url_str)
                        if path:
                            found.append(("UNKNOWN", path, rel_path, line_num))
    return found


def run_scan(output_json=False):
    known = get_known_endpoints_from_apifox()
    found = scan_py_files()

    seen = {}
    new_endpoints = []

    for method, path, rel_file, line_num in found:
        norm_path = normalize_path(path)
        is_known = False
        for known_method, known_path in known:
            if known_method == method and known_path == norm_path:
                is_known = True
                break
            if method != "UNKNOWN":
                known_pattern = re.sub(r'\{[^}]+\}', '{id}', known_path)
                if known_pattern == norm_path:
                    is_known = True
                    break

        key = (method, norm_path)
        if key not in seen:
            seen[key] = (rel_file, line_num)
            if not is_known:
                new_endpoints.append({
                    "method": method,
                    "path": norm_path,
                    "file": rel_file,
                    "line": line_num,
                })

    if output_json:
        return new_endpoints

    print(f"{'='*50}")
    print(f"  LIMS 接口扫描报告")
    print(f"{'='*50}")
    print(f"  已知端点: {len(known)} 个")
    print(f"  代码中发现: {len(seen)} 个")
    print(f"  新端点: {len(new_endpoints)} 个")

    if new_endpoints:
        print(f"\n  发现新端点:")
        for ep in new_endpoints:
            print(f"    - {ep['method']} {ep['path']}")
            print(f"      ({ep['file']}:{ep['line']})")
    else:
        print(f"\n  没有发现新端点")

    return new_endpoints


def main():
    output_json = "--json" in sys.argv
    result = run_scan(output_json=output_json)
    if output_json:
        print(json.dumps(result, ensure_ascii=False))


if __name__ == "__main__":
    main()
