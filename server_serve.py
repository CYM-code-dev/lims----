"""http.server 无窗口封装（计划任务用 pythonw 运行本脚本）。
pythonw 下 sys.stderr/stdout 为 None，http.server 每个请求写日志即崩（表现为 curl 空回复），
先堵住日志，再起 ThreadingHTTPServer 服务本脚本所在目录。"""
import os
import sys

sys.stdout = sys.stderr = open(os.devnull, "w")

from functools import partial
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer

DIR = os.path.dirname(os.path.abspath(__file__))
Handler = partial(SimpleHTTPRequestHandler, directory=DIR)
ThreadingHTTPServer(("0.0.0.0", 8010), Handler).serve_forever()
