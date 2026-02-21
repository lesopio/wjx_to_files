#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
问卷星导出工具 Gradio 前端

运行:
    python app_gradio.py
"""

from __future__ import annotations

import csv
import re
import socket
import sys
import tempfile
import zipfile
from datetime import datetime
from pathlib import Path
from typing import Iterable, List, Optional, Tuple

import gradio as gr

from wjx_to_docx import BatchResult, ExportResult, export_one_url, validate_url


CSS = """
:root {
  --bg-main: #EDE3D4;
  --panel: #F7EFE6;
  --accent-soft: #F3C9D8;
  --accent-strong: #B43A4B;
  --accent-strong-hover: #962b3a;
  --text-main: #3F2E2E;
  --border: #D8C5B6;
}

.gradio-container {
  font-family: "Noto Serif SC", "Microsoft YaHei", "Segoe UI", sans-serif !important;
  color: var(--text-main);
  background: linear-gradient(135deg, var(--bg-main) 0%, #f8e7ec 100%);
}

.app-card {
  background: var(--panel) !important;
  border: 1px solid var(--border) !important;
  border-radius: 16px !important;
  box-shadow: 0 6px 18px rgba(86, 44, 44, 0.08);
}

.main-btn button {
  background: var(--accent-strong) !important;
  color: #fff !important;
  border: none !important;
}

.main-btn button:hover {
  background: var(--accent-strong-hover) !important;
}

.soft-btn button {
  background: var(--accent-soft) !important;
  color: var(--text-main) !important;
  border: 1px solid #e6a8bf !important;
}

.log-box textarea {
  background: #f9f4ec !important;
  border: 1px solid var(--border) !important;
  color: var(--text-main) !important;
  font-family: Consolas, "Courier New", monospace !important;
}

.section-title h2, .section-title h3 {
  color: var(--accent-strong) !important;
}
"""


def now_str() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def split_url_tokens(raw: str) -> List[str]:
    if not raw:
        return []
    parts = re.split(r"[,\n，;；\s]+", raw.strip())
    urls = [p.strip() for p in parts if p.strip()]
    return urls


def read_text_file_auto(path: Path) -> str:
    encodings = ["utf-8-sig", "utf-8", "gbk", "gb18030"]
    for enc in encodings:
        try:
            return path.read_text(encoding=enc)
        except UnicodeDecodeError:
            continue
    return path.read_text(encoding="utf-8", errors="ignore")


def extract_urls_from_file(path: Optional[str]) -> List[str]:
    if not path:
        return []
    p = Path(path)
    if not p.exists():
        return []

    content = read_text_file_auto(p)
    if p.suffix.lower() == ".csv":
        rows: List[str] = []
        reader = csv.reader(content.splitlines())
        for row in reader:
            rows.extend(row)
        return split_url_tokens("\n".join(rows))
    return split_url_tokens(content)


def dedupe_keep_order(items: Iterable[str]) -> List[str]:
    seen = set()
    out: List[str] = []
    for item in items:
        if item not in seen:
            seen.add(item)
            out.append(item)
    return out


def find_free_port(start: int = 7860, end: int = 7999) -> int:
    for port in range(start, end + 1):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
            sock.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
            try:
                sock.bind(("127.0.0.1", port))
                return port
            except OSError:
                continue
    raise RuntimeError("未找到可用端口，请关闭占用端口后重试。")


def ensure_writable_dir(path: Path) -> bool:
    try:
        path.mkdir(parents=True, exist_ok=True)
        probe = path / ".write_test"
        probe.write_text("ok", encoding="utf-8")
        probe.unlink(missing_ok=True)
        return True
    except OSError:
        return False


def create_run_output_dir(run_id: str) -> Path:
    local = Path.cwd() / "outputs" / run_id
    if ensure_writable_dir(local):
        return local

    temp_dir = Path(tempfile.gettempdir()) / "wjx_export" / run_id
    if ensure_writable_dir(temp_dir):
        return temp_dir
    raise RuntimeError("无法创建输出目录，请检查目录权限。")


def write_results_csv(path: Path, results: List[ExportResult]) -> Path:
    csv_path = path / "results.csv"
    with csv_path.open("w", encoding="utf-8-sig", newline="") as f:
        writer = csv.writer(f)
        writer.writerow(
            ["url", "status", "title", "message", "docx_path", "json_path", "md_path", "error_code"]
        )
        for item in results:
            writer.writerow(
                [
                    item.url,
                    item.status,
                    item.title,
                    item.message,
                    item.docx_path or "",
                    item.json_path or "",
                    item.md_path or "",
                    item.error_code,
                ]
            )
    return csv_path


def write_failed_urls(path: Path, results: List[ExportResult]) -> Path:
    failed_path = path / "failed_urls.txt"
    failed_lines = [f"{item.url}\t{item.message}" for item in results if item.status != "success"]
    failed_path.write_text("\n".join(failed_lines), encoding="utf-8")
    return failed_path


def write_log_file(path: Path, run_id: str, logs: str) -> Path:
    log_path = path / f"run_{run_id}.log"
    log_path.write_text(logs, encoding="utf-8")
    return log_path


def build_zip_bundle(path: Path, run_id: str, files: List[str], extras: List[Path]) -> Path:
    zip_path = path / f"bundle_{run_id}.zip"
    with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for file_path in files:
            p = Path(file_path)
            if p.exists():
                zf.write(p, arcname=p.name)
        for extra in extras:
            if extra.exists():
                zf.write(extra, arcname=extra.name)
    return zip_path


def parse_input_urls(url_text: str, upload_file: Optional[str]) -> Tuple[List[str], List[str]]:
    from_text = split_url_tokens(url_text)
    from_file = extract_urls_from_file(upload_file)
    merged = dedupe_keep_order(from_text + from_file)

    valid: List[str] = []
    invalid: List[str] = []
    for u in merged:
        if validate_url(u):
            valid.append(u)
        else:
            invalid.append(u)
    return valid, invalid


def results_to_rows(results: List[ExportResult]) -> List[List[str]]:
    rows: List[List[str]] = []
    for idx, item in enumerate(results, start=1):
        rows.append(
            [
                str(idx),
                item.url,
                "成功" if item.status == "success" else "失败",
                item.title or "-",
                item.message,
            ]
        )
    return rows


def run_batch_export(url_text: str, upload_file: Optional[str]):
    logs: List[str] = []
    results: List[ExportResult] = []
    download_files: List[str] = []
    rows: List[List[str]] = []

    def logger(message: str) -> None:
        line = f"[{now_str()}] {message}"
        print(line)
        logs.append(line)

    def log_text() -> str:
        return "\n".join(logs)

    logger("开始收集输入链接...")
    yield log_text(), rows, download_files, None, None

    valid_urls, invalid_urls = parse_input_urls(url_text, upload_file)
    if invalid_urls:
        for bad in invalid_urls:
            logger(f"[跳过] 非法或不支持域名链接: {bad}")

    if not valid_urls:
        logger("没有检测到可处理的有效链接，任务结束。")
        yield log_text(), rows, download_files, None, None
        return

    run_id = datetime.now().strftime("%Y%m%d_%H%M%S")
    out_dir = create_run_output_dir(run_id)
    logger(f"任务目录: {out_dir}")
    logger(f"待处理链接数: {len(valid_urls)}")
    yield log_text(), rows, download_files, None, None

    success_count = 0
    failed_count = 0
    for idx, url in enumerate(valid_urls, start=1):
        logger(f"======== 处理进度 {idx}/{len(valid_urls)} ========")
        item = export_one_url(url, out_dir=out_dir, logger=logger)
        results.append(item)
        if item.status == "success":
            success_count += 1
            for f in (item.docx_path, item.json_path, item.md_path):
                if f:
                    download_files.append(f)
        else:
            failed_count += 1

        rows = results_to_rows(results)
        yield log_text(), rows, download_files, None, None

    batch = BatchResult(
        success_count=success_count,
        failed_count=failed_count,
        results=results,
        all_files=download_files.copy(),
    )
    logger(f"批量处理结束: 成功 {batch.success_count}，失败 {batch.failed_count}")

    results_csv = write_results_csv(out_dir, batch.results)
    failed_txt = write_failed_urls(out_dir, batch.results)
    logger(f"已写入结果清单: {results_csv}")
    logger(f"已写入失败清单: {failed_txt}")

    log_path = write_log_file(out_dir, run_id, log_text())
    logger(f"已写入日志文件: {log_path}")

    zip_path = build_zip_bundle(
        out_dir, run_id, batch.all_files, extras=[results_csv, failed_txt, log_path]
    )
    logger(f"已生成打包文件: {zip_path}")

    batch.log_path = str(log_path)
    batch.zip_path = str(zip_path)
    final_files = dedupe_keep_order(download_files)
    yield log_text(), rows, final_files, batch.zip_path, batch.log_path


def build_ui() -> gr.Blocks:
    with gr.Blocks(
        title="问卷星导出工具",
        fill_height=True,
    ) as demo:
        gr.Markdown(
            """
            # 问卷星导出工具
            输入一个或多个问卷链接（支持中英文逗号/分号/换行分隔），可选上传 `txt/csv` URL 列表。
            输出 `docx + json + md`，并提供独立下载、总 ZIP、日志文件。
            """,
            elem_classes=["section-title"],
        )

        with gr.Row():
            with gr.Column(scale=2, elem_classes=["app-card"]):
                url_input = gr.Textbox(
                    label="问卷链接输入",
                    placeholder=(
                        "可输入多个链接，例如:\n"
                        "https://v.wjx.cn/vm/xxx.aspx，https://v.wjx.cn/vm/yyy.aspx"
                    ),
                    lines=8,
                )
                upload_file = gr.File(
                    label="上传 URL 列表文件（可选）",
                    file_types=[".txt", ".csv"],
                    type="filepath",
                )
                run_btn = gr.Button("开始导出", elem_classes=["main-btn"])
                reset_btn = gr.Button("清空输出", elem_classes=["soft-btn"])

            with gr.Column(scale=3, elem_classes=["app-card"]):
                log_box = gr.Textbox(
                    label="终端日志（实时）",
                    lines=16,
                    interactive=False,
                    elem_classes=["log-box"],
                )

        result_table = gr.Dataframe(
            headers=["序号", "URL", "状态", "标题", "消息"],
            datatype=["str", "str", "str", "str", "str"],
            row_count=(0, "dynamic"),
            column_count=(5, "fixed"),
            label="处理结果",
            interactive=False,
            wrap=True,
        )

        with gr.Row():
            files_output = gr.Files(label="独立文件下载（docx/json/md）")
            zip_output = gr.File(label="总ZIP下载")
            log_output = gr.File(label="日志下载")

        run_btn.click(
            fn=run_batch_export,
            inputs=[url_input, upload_file],
            outputs=[log_box, result_table, files_output, zip_output, log_output],
        )

        reset_btn.click(
            fn=lambda: ("", [], [], None, None),
            inputs=None,
            outputs=[log_box, result_table, files_output, zip_output, log_output],
        )

    return demo


def main() -> None:
    demo = build_ui()
    port = find_free_port()
    print(f"[{now_str()}] 启动 Gradio 服务: http://127.0.0.1:{port}")
    demo.queue(default_concurrency_limit=1).launch(
        server_name="127.0.0.1",
        server_port=port,
        inbrowser=True,
        share=False,
        css=CSS,
    )


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print(f"[{now_str()}] 已收到中断信号，退出程序。")
        sys.exit(0)
