#!/usr/bin/env python3
from __future__ import annotations

import argparse

from dashscope_doc_review import get_default_client_from_env


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--limit", type=int, default=20)
    parser.add_argument("--base-url", default="https://dashscope.aliyuncs.com/compatible-mode/v1")
    args = parser.parse_args()
    client = get_default_client_from_env(base_url=args.base_url)
    items = client.files.list(limit=max(1, args.limit))
    data = getattr(items, "data", []) or []
    for f in data:
        print(
            f"id={getattr(f, 'id', '')} "
            f"filename={getattr(f, 'filename', '')} "
            f"purpose={getattr(f, 'purpose', '')} "
            f"status={getattr(f, 'status', '')} "
            f"created_at={getattr(f, 'created_at', '')}"
        )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
