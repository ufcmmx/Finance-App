#!/usr/bin/env python3
"""
gen_license.py — 调用 Workers 后台生成激活码

用法：
    python gen_license.py --email customer@example.com --price 199 --type annual
    python gen_license.py --email user@test.com --price 499 --type permanent

需要先在 .env 里配置：
    ADMIN_KEY=...
    WORKERS_URL=https://wiseledger-license.xxxx.workers.dev
"""
import argparse
import json
import os
import sys
import urllib.request
from pathlib import Path


def load_env():
    """读取 .env 文件到 os.environ"""
    env_path = Path(__file__).parent / ".env"
    if not env_path.exists():
        print("❌ .env 文件不存在", file=sys.stderr)
        sys.exit(1)
    with open(env_path) as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            k, v = line.split("=", 1)
            os.environ.setdefault(k.strip(), v.strip())


def main():
    load_env()

    parser = argparse.ArgumentParser(description="生成 WiseLedger 激活码")
    parser.add_argument("--email", required=True, help="客户邮箱")
    parser.add_argument("--price", type=float, default=0, help="售价")
    parser.add_argument("--note", default="", help="备注（如渠道、套餐）")
    parser.add_argument(
        "--type", choices=["annual", "permanent"], default="permanent",
        help="激活码类型：annual (¥199/年 订阅，原价 ¥398) | permanent (¥499 永久，原价 ¥998，默认)",
    )
    args = parser.parse_args()

    admin_key = os.environ.get("ADMIN_KEY")
    workers_url = os.environ.get("WORKERS_URL", "").rstrip("/")
    if not admin_key:
        print("❌ .env 里缺 ADMIN_KEY", file=sys.stderr)
        sys.exit(1)
    if not workers_url:
        print("❌ .env 里缺 WORKERS_URL（部署完 Workers 后填上）", file=sys.stderr)
        sys.exit(1)

    payload = json.dumps({
        "customer_email": args.email,
        "price": args.price,
        "note": args.note,
        "type": args.type,
    }).encode("utf-8")

    req = urllib.request.Request(
        f"{workers_url}/admin/generate",
        data=payload,
        method="POST",
        headers={
            "Content-Type": "application/json",
            "X-Admin-Key": admin_key,
            # 避免 Cloudflare 把默认 Python UA 当机器人拦截
            "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) WiseLedger-Admin/1.0",
        },
    )

    try:
        with urllib.request.urlopen(req, timeout=10) as resp:
            result = json.loads(resp.read().decode("utf-8"))
    except urllib.error.HTTPError as e:
        body = e.read().decode("utf-8")
        print(f"❌ HTTP {e.code}: {body}", file=sys.stderr)
        sys.exit(1)
    except Exception as e:
        print(f"❌ 请求失败: {e}", file=sys.stderr)
        sys.exit(1)

    code = result["license_code"]
    lic_type = result.get("type", args.type)
    type_zh = {
        "annual":    "订阅版（¥199/年 活动价，原 ¥398）",
        "permanent": "永久版（¥499 买断 活动价，原 ¥998）",
    }[lic_type]
    print()
    print(f"✓ 激活码已生成：{code}")
    print(f"  类型：{type_zh}")
    print(f"  客户邮箱：{result['customer_email']}")
    print(f"  生成时间：{result['sold_at']}")
    print()
    print("─" * 60)
    print(f"复制以下文字发给客户：")
    print("─" * 60)
    print()
    print(f"您的智一盈小账 (WiseLedger) 激活码：")
    print(f"  {code}")
    print(f"  类型：{type_zh}")
    print()
    print(f"使用说明：")
    print(f"  1. 下载软件：https://download.wisdompluscn.com/wiseledger/latest.zip")
    print(f"  2. 打开软件 → 弹出激活窗 → 勾选同意协议")
    print(f"  3. 输入上述激活码 → 联网验证成功即可使用")
    if lic_type == "annual":
        print(f"  4. 订阅版有效期 1 年，到期前 30 天软件会开始提示续费")
        print(f"     · 到期后 7 天宽限期照常使用")
        print(f"     · 超过宽限期进入只读模式（仍可查看历史数据）")
    else:
        print(f"  4. 永久版无到期时间，激活一次终身使用")
    print(f"  5. 激活码绑定首次使用的电脑，每年可换机 2 次")
    print()


if __name__ == "__main__":
    main()
