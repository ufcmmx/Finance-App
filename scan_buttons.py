"""扫描所有 setFixedWidth/setFixedSize 的按钮，估算文字宽度，对比是否够装"""
import re
import os
from pathlib import Path

# 字符宽度估算（Microsoft YaHei 13px）
def estimate_text_width(text: str, bold: bool = False) -> int:
    total = 0
    for ch in text:
        if '一' <= ch <= '鿿':  # 中文
            total += 16 if bold else 14
        elif ch == ' ':
            total += 5
        elif ch in '+－－＋✕✖✗':
            total += 8
        elif ch in '0123456789':
            total += 8
        elif ord(ch) < 128:  # ASCII
            total += 8 if bold else 7
        elif ord(ch) > 0x2000:  # emoji 等宽字符
            total += 18
        else:
            total += 14
    return total

# 全局 padding
PADDING_H = 12 * 2  # padding-left + padding-right (current: 6 12)

# 扫描所有 py 文件
ROOT = Path("/sessions/gifted-inspiring-franklin/mnt/Finance")
button_pattern = re.compile(
    r'(\w+)\s*=\s*QPushButton\(\s*["\']([^"\']+)["\']\s*\)'
)
fixedwidth_pattern = re.compile(r'(\w+)\.setFixedWidth\(\s*(\d+)\s*\)')
fixedsize_pattern  = re.compile(r'(\w+)\.setFixedSize\(\s*(\d+)\s*,\s*(\d+)\s*\)')
objectname_pattern = re.compile(r'(\w+)\.setObjectName\(\s*["\'](btn_\w+)["\']\s*\)')

records = []  # (file, line, var, text, width, objectname, bold)

for py in ROOT.rglob("*.py"):
    if "/__pycache__/" in str(py): continue
    text = py.read_text(encoding='utf-8', errors='ignore')
    lines = text.split('\n')
    
    # 收集每个文件里的按钮信息（按变量名）
    btns = {}  # var -> {text, line, objname, width}
    
    for i, line in enumerate(lines, 1):
        # 同一行里可能多个匹配
        for m in button_pattern.finditer(line):
            var, btxt = m.group(1), m.group(2)
            btns[var] = {'text': btxt, 'line': i, 'objname': None, 'width': None, 'fixedsize': None}
        for m in objectname_pattern.finditer(line):
            var, on = m.group(1), m.group(2)
            if var in btns:
                btns[var]['objname'] = on
        for m in fixedwidth_pattern.finditer(line):
            var, w = m.group(1), int(m.group(2))
            if var in btns:
                btns[var]['width'] = w
        for m in fixedsize_pattern.finditer(line):
            var, w, h = m.group(1), int(m.group(2)), int(m.group(3))
            if var in btns:
                btns[var]['fixedsize'] = (w, h)
                btns[var]['width'] = w
    
    for var, info in btns.items():
        if info['width'] is None: continue  # 没设固定宽度，跳过
        records.append({
            'file': str(py.relative_to(ROOT)),
            'line': info['line'],
            'var': var,
            'text': info['text'],
            'width': info['width'],
            'objname': info['objname'] or '?',
            'fixedsize': info['fixedsize'],
        })

# 分析
print(f"{'文件':<35} {'行':>4} {'文字':<22} {'宽度':>5} {'文字+padding':>14} {'objname':<12} {'状态':<6}")
print("-" * 105)

bold_objs = {'btn_primary'}
risky = []
ok = []

for r in records:
    is_bold = r['objname'] in bold_objs
    txt_w = estimate_text_width(r['text'], bold=is_bold)
    need = txt_w + PADDING_H
    diff = r['width'] - need
    status = '❌截断' if diff < 0 else ('⚠️紧贴' if diff < 6 else '✅OK')
    line = f"{r['file']:<35} {r['line']:>4} {r['text']:<22} {r['width']:>5} {need:>14} {r['objname']:<12} {status}"
    print(line)
    if diff < 6:
        risky.append((r, txt_w, need, diff))

print(f"\n小结：{len(records)} 个按钮设了固定宽度，其中 {len(risky)} 个有问题")
print("\n问题清单（建议宽度 = 文字+padding+8px 余量）：")
for r, txt_w, need, diff in risky:
    suggest = need + 8
    print(f"  {r['file']}:{r['line']}  {r['text']:<22}  现在 {r['width']:>3} → 建议 {suggest:>3}")
