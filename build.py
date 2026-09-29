# -*- coding: utf-8 -*-
"""
冻存条码核对器 · 构建脚本

用法：
  python3 build.py [xlsx路径] [选项]

选项：
  --name 名称        编码库名称（默认取 Excel 文件名）
  --encrypt[=密码]   将编码库 AES 加密后内嵌（公网分发推荐）；
                     未指定密码时读取 password.txt，不存在则生成新密码
  --save-password    配合 --encrypt：把密码写入 password.txt（默认只打印）
  --pwa              输出 PWA 部署目录（index.html/sw.js/manifest/图标/setup.html）
  --out 路径         输出文件（默认模式）或目录（--pwa 模式，默认 deploy/）
  --grid RxC         冻存盒孔位规格，默认 6x8
  --code-col 关键词  核对码来源列（默认「暂存空间」）
  --code-regex 正则  编码提取规则（默认取来源列末尾编码段，不限开头字母；
                     特殊值 full = 整列文本作为编码）

示例：
  python3 build.py demo/demo.xlsx                      # 演示数据明文单文件
  python3 build.py 名单.xlsx --pwa --out site          # 自己的名单，明文 PWA
  python3 build.py 名单.xlsx --encrypt --pwa           # 加密 PWA（密码打印/存 password.txt）
"""
import argparse
import base64
import json
import os
import re
import secrets
import sys

import openpyxl

BASE = os.path.dirname(os.path.abspath(__file__))
DEMO_XLSX = os.path.join(BASE, 'demo', 'demo.xlsx')


# ---------------------------------------------------------------- 读取编码库
# 核对用条码默认 = 「暂存空间」末尾编码段（如 …第1盒E7000001 → E7000001）；
# 开头字母不限（那只是某批次的巧合），可用 --code-col / --code-regex 自定义来源与规则。
# 行结构：[核对码, 样本号, 冻存条码号, 暂存空间, 行索引, 列索引, 手工标记]
DEFAULT_CODE_RE = r'([A-Za-z0-9][A-Za-z0-9\-]{1,})$'


def load_rows(xlsx_path, code_col_kw='暂存空间', code_regex=None):
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    ws = wb[wb.sheetnames[0]]
    grid = list(ws.iter_rows(values_only=True))
    header = next((i for i, r in enumerate(grid[:5]) if r and any(str(c or '').strip() for c in r)), 0)
    cols = [str(c or '').strip() for c in (grid[header] if grid else [])]

    def find_col(kws):
        for j, name in enumerate(cols):
            if any(k in name for k in kws):
                return j
        return -1

    code_col = find_col([code_col_kw]) if code_col_kw else -1
    if code_col < 0:   # 退回：任何名称含「条码」的列，整列取值
        code_col = find_col(['条码'])
        whole_col = True
    else:
        whole_col = (code_regex == 'full')
    rx = None if whole_col else re.compile(code_regex or DEFAULT_CODE_RE)

    s_col, t_col, b_col = find_col(['样本', '编号', 'yangben']), find_col(['条码号', '冻存条码']), find_col(['暂存', '空间', '位置'])
    r_col, k_col = find_col(['行']), find_col(['列'])
    t_col = t_col if t_col != code_col else -1

    rows, skipped = [], 0
    for r in grid[header + 1:]:
        raw = r[code_col] if code_col < len(r) else None
        if raw is None or not str(raw).strip():
            continue
        text = str(raw).strip()
        if whole_col:
            code = re.sub(r'\s+', '', text)
        else:
            m = rx.search(text)
            code = m.group(1) if m else ''
        if not code:
            skipped += 1
            continue
        get = lambda j: (str(r[j]).strip() if j >= 0 and j < len(r) and r[j] is not None else '')
        rows.append([code, get(s_col), get(t_col), get(b_col) if b_col != code_col else text,
                     get(r_col), get(k_col), ''])
    uniq = len(set(x[0] for x in rows))
    print('读取', xlsx_path, '->', len(rows), '行，唯一码', uniq, '个（来源列:',
          cols[code_col] if code_col >= 0 else '?', '| 规则:', '整列' if whole_col else (code_regex or '默认末尾编码段'), '）')
    if rows:
        print('  码样例:', [x[0] for x in rows[:5]])
    if skipped:
        print('  跳过未提取到编码的行:', skipped)
    return rows


def load_template():
    tpl = open(os.path.join(BASE, 'app_template.html'), encoding='utf-8').read()
    lib_js = open(os.path.join(BASE, 'html5-qrcode.min.js'), encoding='utf-8').read()
    lib_js = re.sub(r'</script', r'<\\/script', lib_js, flags=re.I)
    return tpl, lib_js


def assemble(tpl, lib_js, data_js, pwa_head='', pwa_reg='', grid=None):
    html = tpl
    assert html.count('/*__H5Q__*/') == 1 and html.count('/*__DATA__*/') == 1
    html = html.replace('/*__H5Q__*/', lib_js)
    html = html.replace('/*__DATA__*/', data_js)
    html = html.replace('<!--__PWAHEAD__-->', pwa_head)
    html = html.replace('/*__PWAREG__*/', pwa_reg)
    if grid and grid != (6, 8):
        html = html.replace('const R = 6, C = 8; /*__GRID__*/', 'const R = %d, C = %d;' % grid)
    assert '/*__' not in html.split('</head>')[0] or True
    return html


# ---------------------------------------------------------------- 加密
def load_or_create_password(explicit=None, save=False):
    pw_file = os.path.join(BASE, 'password.txt')
    if explicit:
        return explicit
    if os.path.exists(pw_file):
        return open(pw_file).read().strip()
    pw = 'CRYO-' + ''.join(secrets.choice('23456789ABCDEFGHJKMNPQRSTUVWXYZ') for _ in range(8))
    if save:
        open(pw_file, 'w').write(pw + '\n')
        os.chmod(pw_file, 0o600)
        print('已生成新密码并保存到 password.txt')
    return pw


def encrypt_rows(pw, rows):
    from cryptography.hazmat.primitives import hashes
    from cryptography.hazmat.primitives.ciphers.aead import AESGCM
    from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC
    salt, iv = os.urandom(16), os.urandom(12)
    key = PBKDF2HMAC(algorithm=hashes.SHA256(), length=32, salt=salt, iterations=150000).derive(pw.encode())
    plain = json.dumps(rows, ensure_ascii=False, separators=(',', ':')).encode()
    ct = AESGCM(key).encrypt(iv, plain, None)
    b64 = lambda b: base64.b64encode(b).decode()
    return {'salt': b64(salt), 'iv': b64(iv), 'iter': 150000, 'ct': b64(ct)}


# ---------------------------------------------------------------- PWA 资产
def make_icons(d):
    from PIL import Image, ImageDraw
    def icon(size, path):
        img = Image.new('RGBA', (size, size), (0, 0, 0, 0))
        dr = ImageDraw.Draw(img)
        dr.rounded_rectangle([0, 0, size - 1, size - 1], radius=size // 5, fill=(15, 23, 42, 255))
        bar_w, gap = size // 14, size // 22
        x = size // 4
        heights = [size // 2.6, size // 3.4, size // 2.2, size // 3, size // 2.8]
        pattern = [3, 1, 2, 1, 4, 1, 1, 2, 3, 1, 2]
        for i, w in enumerate(pattern):
            h = heights[i % len(heights)]
            y = (size - h) / 2
            dr.rectangle([x, y, x + bar_w * w - gap, y + h], fill=(238, 242, 251, 255))
            x += bar_w * w
        img.save(path)
    icon(192, os.path.join(d, 'icon-192.png'))
    icon(512, os.path.join(d, 'icon-512.png'))
    icon(180, os.path.join(d, 'apple-touch-icon.png'))


SW_JS = '''const CACHE='cbpwa-__BUILD__';
const ASSETS=['./','./index.html','./manifest.webmanifest','./icon-192.png','./icon-512.png','./apple-touch-icon.png'];
self.addEventListener('install',e=>{e.waitUntil((async()=>{const c=await caches.open(CACHE);await c.addAll(ASSETS);await self.skipWaiting()})())});
self.addEventListener('activate',e=>{e.waitUntil((async()=>{const ks=await caches.keys();await Promise.all(ks.filter(k=>k!==CACHE).map(k=>caches.delete(k)));await self.clients.claim()})())});
// 缓存优先（打开永远秒开、断网可用）+ 后台静默更新（联网时拉新版，下次打开生效）
self.addEventListener('fetch',e=>{
  if(e.request.method!=='GET')return;
  e.respondWith((async()=>{
    const c=await caches.open(CACHE);
    const cached=await c.match(e.request,{ignoreSearch:e.request.mode==='navigate'});
    const refresh=fetch(e.request).then(res=>{
      if(res&&(res.status===200||res.type==='opaque')) c.put(e.request,res.clone());
      return res;
    }).catch(()=>null);
    if(cached) return cached;
    const res=await refresh;
    if(res) return res;
    return c.match('./index.html');
  })());
});'''

MANIFEST = {
    'name': '冻存条码核对', 'short_name': '条码核对',
    'start_url': './', 'scope': './', 'display': 'standalone',
    'background_color': '#0f172a', 'theme_color': '#0f172a',
    'icons': [
        {'src': './icon-192.png', 'sizes': '192x192', 'type': 'image/png'},
        {'src': './icon-512.png', 'sizes': '512x512', 'type': 'image/png'},
    ],
}

SETUP_HTML = os.path.join(BASE, 'deploy', 'setup.html')   # 若已存在则复用


def pwa_assets(build_id):
    return {
        'sw.js': SW_JS.replace('__BUILD__', build_id),
        'manifest.webmanifest': json.dumps(MANIFEST, ensure_ascii=False, indent=2),
    }


# ---------------------------------------------------------------- 主流程
def main():
    ap = argparse.ArgumentParser(description='冻存条码核对器构建脚本')
    ap.add_argument('xlsx', nargs='?', default=DEMO_XLSX, help='Excel 名单路径（默认 demo/demo.xlsx）')
    ap.add_argument('--name', help='编码库名称（默认取文件名）')
    ap.add_argument('--encrypt', nargs='?', const='', default=None,
                    help='AES 加密内嵌；可跟 =密码，否则读/生成 password.txt')
    ap.add_argument('--save-password', action='store_true', help='把生成的密码写入 password.txt')
    ap.add_argument('--pwa', action='store_true', help='输出 PWA 部署目录')
    ap.add_argument('--out', help='输出路径（文件或目录）')
    ap.add_argument('--grid', default='6x8', help='冻存盒孔位规格，如 6x8')
    ap.add_argument('--code-col', default='暂存空间',
                    help='核对码来源列的关键词（默认「暂存空间」；找不到时退回任何含「条码」的列）')
    ap.add_argument('--code-regex', default=None,
                    help='从来源列文本提取编码的正则（默认取末尾编码段，不限开头字母）；'
                         '特殊值 full = 整列文本作为编码')
    args = ap.parse_args()

    if not os.path.exists(args.xlsx):
        sys.exit('找不到 Excel：%s' % args.xlsx)
    m = re.fullmatch(r'(\d+)\s*[xX×]\s*(\d+)', args.grid.strip())
    if not m:
        sys.exit('--grid 格式应为 RxC，如 6x8')
    grid = (int(m.group(1)), int(m.group(2)))
    if not (1 <= grid[0] <= 16 and 1 <= grid[1] <= 16):
        sys.exit('--grid 行列须在 1–16 之间')

    rows = load_rows(args.xlsx, args.code_col, args.code_regex)
    tpl, lib_js = load_template()
    name = args.name or re.sub(r'\.(xlsx|xls)$', '', os.path.basename(args.xlsx), flags=re.I)

    if args.encrypt is not None:
        pw = load_or_create_password(args.encrypt or None, args.save_password)
        enc = encrypt_rows(pw, rows)
        data_js = json.dumps({'name': name, 'enc': enc}, ensure_ascii=False, separators=(',', ':'))
        print('访问密码：%s' % pw)
    else:
        data_js = json.dumps({'name': name, 'src': 'default', 'rows': rows},
                             ensure_ascii=False, separators=(',', ':'))

    if args.pwa:
        import time
        build_id = format(int(time.time()), 'x')
        pwa_head = ('<link rel="manifest" href="./manifest.webmanifest">\n'
                    '<link rel="apple-touch-icon" href="./apple-touch-icon.png">\n'
                    '<meta name="description" content="offline barcode check">')
        pwa_reg = (
            "if ('serviceWorker' in navigator) { window.addEventListener('load', function () {\n"
            "  navigator.serviceWorker.register('./sw.js').then(function (reg) {\n"
            "    reg.addEventListener('updatefound', function () {\n"
            "      var nw = reg.installing; if (!nw) return;\n"
            "      nw.addEventListener('statechange', function () {\n"
            "        if (nw.state === 'installed' && navigator.serviceWorker.controller) {\n"
            "          var t = document.getElementById('toast');\n"
            "          if (t) {\n"
            "            t.textContent = '新版本已就绪：关闭页面后重新打开即更新'; t.classList.add('show');\n"
            "            setTimeout(function () { t.classList.remove('show'); }, 3200);\n"
            "          }\n"
            "        }\n"
            "      });\n"
            "    });\n"
            "  }).catch(function () {});\n"
            "}); }\n"
        )
        html = assemble(tpl, lib_js, data_js, pwa_head, pwa_reg, grid)
        out_dir = args.out or os.path.join(BASE, 'deploy')
        os.makedirs(out_dir, exist_ok=True)
        open(os.path.join(out_dir, 'index.html'), 'w', encoding='utf-8').write(html)
        for fn, content in pwa_assets(build_id).items():
            open(os.path.join(out_dir, fn), 'w', encoding='utf-8').write(content)
        make_icons(out_dir)
        # setup.html（自签服务器信任引导）与 README 由仓库提供，存在则不覆盖
        setup_src = os.path.join(BASE, 'setup.html')
        if os.path.exists(setup_src):
            import shutil
            shutil.copy(setup_src, os.path.join(out_dir, 'setup.html'))
        print('PWA 已输出到', out_dir)
    else:
        html = assemble(tpl, lib_js, data_js, grid=grid)
        out = args.out or os.path.join(BASE, '冻存条码核对.html')
        open(out, 'w', encoding='utf-8').write(html)
        print('单文件已输出到', out)


if __name__ == '__main__':
    main()
