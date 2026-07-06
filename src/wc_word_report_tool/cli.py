from __future__ import annotations

import argparse

from .numbers import number_to_chinese_upper


def main() -> None:
    parser = argparse.ArgumentParser(description="数字转中文金额大写")
    parser.add_argument("value", help="要转换的数字，例如 12345.67")
    args = parser.parse_args()
    print(number_to_chinese_upper(args.value))
