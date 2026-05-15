#!/usr/bin/env python3
from __future__ import annotations

import argparse
from pathlib import Path
import sys

SCRIPT_DIR = Path(__file__).resolve().parents[2] / "scripts"
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

from dashscope_doc_review import (
    get_default_client_from_env,
    resolve_spec_file_id,
    run_qwen_long_docx_review,
    upload_target_docx,
)


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--spec-file-id")
    parser.add_argument("--spec-docx")
    parser.add_argument("--target-docx", required=True)
    parser.add_argument("--model", default="qwen-long")
    parser.add_argument("--output")
    parser.add_argument("--base-url", default="https://dashscope.aliyuncs.com/compatible-mode/v1")
    args = parser.parse_args()

    client = get_default_client_from_env(base_url=args.base_url)
    spec_path = Path(args.spec_docx) if args.spec_docx else None
    spec_file_id = resolve_spec_file_id(client, args.spec_file_id, spec_path)
    target_file_id = upload_target_docx(client, Path(args.target_docx))
    text = run_qwen_long_docx_review(
        client=client,
        spec_file_id=spec_file_id,
        target_file_id=target_file_id,
        model=args.model,
        user_prompt="请根据规范文件检查论文并给出修改建议与 safe_edit_plan、manual_review_items。",
    )
    if args.output:
        Path(args.output).write_text(text, encoding="utf-8", newline="\n")
    else:
        print(text)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
