#!/usr/bin/env python3
from __future__ import annotations

import argparse
import os

from openai import OpenAI


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--file-id", required=True)
    parser.add_argument("--model", default="qwen-long")
    parser.add_argument("--question", default="请总结这个 docx 文档的主要内容。")
    parser.add_argument("--base-url", default=os.environ.get("DASHSCOPE_BASE_URL", "https://dashscope.aliyuncs.com/compatible-mode/v1"))
    args = parser.parse_args()
    api_key = os.environ.get("DASHSCOPE_API_KEY", "")
    if not api_key:
        raise RuntimeError("DASHSCOPE_API_KEY is missing.")
    client = OpenAI(api_key=api_key, base_url=args.base_url)
    messages = [
        {"role": "system", "content": "你是文档总结助手。"},
        {"role": "system", "content": f"fileid://{args.file_id}"},
        {"role": "user", "content": args.question},
    ]
    resp = client.chat.completions.create(model=args.model, messages=messages, stream=False)
    choices = getattr(resp, "choices", None) or []
    if not choices:
        print("")
        return 0
    print(getattr(getattr(choices[0], "message", None), "content", "") or "")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
