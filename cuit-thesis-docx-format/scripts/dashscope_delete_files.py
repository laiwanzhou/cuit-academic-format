#!/usr/bin/env python3
from __future__ import annotations

import argparse

from dashscope_doc_review import get_default_client_from_env


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("file_ids", nargs="+")
    parser.add_argument("--yes", action="store_true")
    parser.add_argument("--base-url", default="https://dashscope.aliyuncs.com/compatible-mode/v1")
    args = parser.parse_args()
    client = get_default_client_from_env(base_url=args.base_url)
    for file_id in args.file_ids:
        if not args.yes:
            print(f"dry-run delete file_id={file_id}")
            continue
        client.files.delete(file_id)
        print(f"deleted file_id={file_id}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
