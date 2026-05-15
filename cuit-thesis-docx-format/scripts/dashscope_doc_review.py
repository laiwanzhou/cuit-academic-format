#!/usr/bin/env python3
from __future__ import annotations

import os
import time
from pathlib import Path

try:
    from openai import OpenAI
except ModuleNotFoundError as exc:  # pragma: no cover
    OpenAI = None
    OPENAI_IMPORT_ERROR = exc
else:
    OPENAI_IMPORT_ERROR = None


DEFAULT_BASE_URL = "https://dashscope.aliyuncs.com/compatible-mode/v1"


def get_dashscope_client(api_key: str, base_url: str = DEFAULT_BASE_URL) -> OpenAI:
    if OpenAI is None:
        raise RuntimeError("openai package is required for DashScope document review.") from OPENAI_IMPORT_ERROR
    return OpenAI(api_key=api_key, base_url=base_url)


def upload_file_for_extract(client: OpenAI, path: Path) -> str:
    with path.open("rb") as fp:
        uploaded = client.files.create(file=fp, purpose="file-extract")
    file_id = getattr(uploaded, "id", None)
    if not file_id:
        raise RuntimeError("DashScope file upload succeeded but no file id returned.")
    return str(file_id)


def wait_file_processed(client: OpenAI, file_id: str, timeout_seconds: int = 600, interval_seconds: int = 3):
    deadline = time.time() + timeout_seconds
    while time.time() < deadline:
        file_obj = client.files.retrieve(file_id)
        status = str(getattr(file_obj, "status", "") or "")
        if status == "processed":
            return file_obj
        if status in {"failed", "error"}:
            raise RuntimeError(f"DashScope file processing failed: file_id={file_id}, status={status}")
        time.sleep(interval_seconds)
    raise TimeoutError(f"Timed out waiting for file processed: file_id={file_id}")


def resolve_spec_file_id(
    client: OpenAI,
    spec_file_id: str | None,
    spec_docx_path: Path | None,
    timeout_seconds: int = 600,
) -> str:
    if spec_file_id:
        return spec_file_id
    if spec_docx_path is None:
        raise ValueError("spec_file_id and spec_docx_path are both missing.")
    if not spec_docx_path.exists():
        raise FileNotFoundError(spec_docx_path)
    uploaded_id = upload_file_for_extract(client, spec_docx_path)
    wait_file_processed(client, uploaded_id, timeout_seconds=timeout_seconds)
    return uploaded_id


def upload_target_docx(client: OpenAI, target_docx_path: Path, timeout_seconds: int = 600) -> str:
    if not target_docx_path.exists():
        raise FileNotFoundError(target_docx_path)
    uploaded_id = upload_file_for_extract(client, target_docx_path)
    wait_file_processed(client, uploaded_id, timeout_seconds=timeout_seconds)
    return uploaded_id


def build_qwen_long_docx_review_messages(spec_file_id: str, target_file_id: str, user_prompt: str) -> list[dict]:
    return [
        {
            "role": "system",
            "content": (
                "你是成都信息工程大学本科论文格式审查助手。"
                "接下来会提供两份文档：第一份是学校论文格式规范，第二份是待检查论文。"
                "请严格依据第一份规范检查第二份论文。"
            ),
        },
        {"role": "system", "content": f"fileid://{spec_file_id}"},
        {"role": "system", "content": f"fileid://{target_file_id}"},
        {"role": "user", "content": user_prompt},
    ]


def run_qwen_long_docx_review(
    *,
    client: OpenAI,
    spec_file_id: str,
    target_file_id: str,
    model: str = "qwen-long",
    user_prompt: str | None = None,
    timeout_seconds: int = 600,
) -> str:
    prompt = user_prompt or (
        "请输出：\n"
        "一、总体判断\n"
        "二、格式问题清单\n"
        "三、需要修改的位置与修改方式\n"
        "四、可直接替换的修改后文本\n"
        "五、可安全自动修改项 safe_edit_plan\n"
        "六、必须人工复核项 manual_review_items"
    )
    messages = build_qwen_long_docx_review_messages(spec_file_id, target_file_id, prompt)
    stream = client.chat.completions.create(
        model=model,
        messages=messages,
        stream=True,
        stream_options={"include_usage": True},
        timeout=timeout_seconds,
    )
    parts: list[str] = []
    for chunk in stream:
        choices = getattr(chunk, "choices", None) or []
        if not choices:
            continue
        delta = getattr(choices[0], "delta", None)
        content = getattr(delta, "content", "") if delta else ""
        if content:
            parts.append(str(content))
    return "".join(parts)


def delete_uploaded_file(client: OpenAI, file_id: str):
    return client.files.delete(file_id)


def get_default_client_from_env(base_url: str = DEFAULT_BASE_URL) -> OpenAI:
    api_key = os.environ.get("DASHSCOPE_API_KEY", "")
    if not api_key:
        raise RuntimeError("DASHSCOPE_API_KEY is missing.")
    return get_dashscope_client(api_key=api_key, base_url=base_url)
