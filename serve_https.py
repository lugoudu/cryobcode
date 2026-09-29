# -*- coding: utf-8 -*-
# 手机访问用 HTTPS 静态服务器（摄像头必须 HTTPS 才能调用）
# 用法：python3 serve_https.py
# 首次运行自动生成自签名证书（certs/ 目录），并在终端打印二维码，
# 手机扫码或输入 https://<本机局域网IP>:8443/冻存条码核对.html 即可打开。
import os
import socket
import ssl
import subprocess
from functools import partial
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import quote

PORT = 8443
BASE = os.path.dirname(os.path.abspath(__file__))
CERT_DIR = os.path.join(BASE, 'certs')
CERT = os.path.join(CERT_DIR, 'cert.pem')
KEY = os.path.join(CERT_DIR, 'key.pem')


def make_cert():
    os.makedirs(CERT_DIR, exist_ok=True)
    if os.path.exists(CERT) and os.path.exists(KEY):
        return
    print('正在生成自签名证书（仅首次）…')
    subprocess.run([
        'openssl', 'req', '-x509', '-newkey', 'rsa:2048', '-keyout', KEY, '-out', CERT,
        '-days', '3650', '-nodes', '-subj', '/CN=cryobcode-local',
    ], check=True, capture_output=True)


def lan_ip():
    s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    try:
        s.connect(('8.8.8.8', 80))
        return s.getsockname()[0]
    except Exception:
        return '127.0.0.1'
    finally:
        s.close()


def print_qr(url):
    try:
        import qrcode
        qr = qrcode.QRCode(border=1)
        qr.add_data(url)
        qr.print_ascii(invert=True)
    except Exception as e:
        print('（未安装 qrcode 库，跳过二维码：%s）' % e)


def main():
    make_cert()
    ip = lan_ip()
    url = 'https://%s:%d/%s' % (ip, PORT, quote('冻存条码核对.html'))

    handler = partial(SimpleHTTPRequestHandler, directory=BASE)
    httpd = ThreadingHTTPServer(('0.0.0.0', PORT), handler)
    ctx = ssl.SSLContext(ssl.PROTOCOL_TLS_SERVER)
    ctx.load_cert_chain(CERT, KEY)
    httpd.socket = ctx.wrap_socket(httpd.socket, server_side=True)

    print('已启动：https://%s:%d/  （目录：%s）' % (ip, PORT, BASE))
    print('手机与电脑需在同一 Wi-Fi / 局域网。')
    print('手机浏览器打开（或扫描下方二维码）：')
    print('  ' + url)
    print('首次打开出现安全警告时：Safari 点「显示详细信息」→「访问此网站」；Chrome 点「高级」→「继续前往」。')
    print('macOS 若弹出防火墙询问，请选择「允许」。Ctrl+C 停止服务。')
    print()
    print_qr(url)
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        print('\n服务已停止')


if __name__ == '__main__':
    main()
