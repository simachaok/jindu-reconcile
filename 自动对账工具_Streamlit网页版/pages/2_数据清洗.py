from __future__ import annotations

import tempfile
from datetime import datetime
from pathlib import Path

import streamlit as st

from data_cleaning_core import clean_files


BASE_DIR = Path(__file__).resolve().parents[1]
TEMPLATE_DIR = BASE_DIR / "数据清洗模板"

st.set_page_config(page_title="数据清洗", page_icon="🧹", layout="wide")

st.markdown(
    """
    <style>
    .block-container {padding-top: 1.4rem; max-width: 1280px;}
    .clean-hero {
        border: 1px solid #dbe4ef;
        background: linear-gradient(135deg, #ffffff 0%, #eef8f4 58%, #f7fbff 100%);
        border-radius: 8px;
        padding: 20px 24px;
        margin-bottom: 18px;
    }
    .clean-title {font-size: 30px; font-weight: 760; color: #102033; margin: 0;}
    .clean-sub {font-size: 15px; color: #5b6675; margin: 8px 0 0;}
    .clean-pill {
        display: inline-block;
        margin: 14px 8px 0 0;
        padding: 6px 10px;
        border: 1px solid #dbe4ef;
        border-radius: 6px;
        background: #ffffff;
        color: #223047;
        font-size: 13px;
    }
    </style>
    <div class="clean-hero">
        <p class="clean-title">数据清洗</p>
        <p class="clean-sub">上传供应商出库 Excel，按模板表头映射、药品资料、供货商编码和供货单位编码生成标准化结果。</p>
        <span class="clean-pill">支持 XLS / XLSX</span>
        <span class="clean-pill">自动识别供货商</span>
        <span class="clean-pill">输出标准化结果与失败原因</span>
    </div>
    """,
    unsafe_allow_html=True,
)

left, right = st.columns([1.2, 0.8], gap="large")

with left:
    st.subheader("上传待清洗文件")
    uploads = st.file_uploader(
        "供应商出库单",
        type=["xls", "xlsx"],
        accept_multiple_files=True,
        help="可以一次上传多个供应商文件。PDF/OCR 版后续单独接入。",
    )
    run = st.button("开始数据清洗", type="primary", disabled=not uploads, use_container_width=True)

with right:
    st.subheader("规则模板")
    template_files = sorted([p.name for p in TEMPLATE_DIR.glob("*")]) if TEMPLATE_DIR.exists() else []
    st.metric("模板文件", len(template_files))
    with st.expander("查看当前模板", expanded=False):
        for name in template_files:
            st.write(name)
    st.info("当前网页第一版支持 Excel 出库单标准化。PDF/OCR/AI 解析涉及云端密钥，后续可用 Streamlit Secrets 单独配置。")

if not uploads:
    st.warning("请先上传一个或多个供应商出库 Excel 文件。")
    st.stop()

if not run:
    st.stop()

with st.spinner("正在清洗数据并生成结果，请稍等..."):
    try:
        with tempfile.TemporaryDirectory(prefix="jd_clean_") as temp_name:
            temp_dir = Path(temp_name)
            paths = []
            for item in uploads:
                target = temp_dir / item.name
                target.write_bytes(item.getbuffer())
                paths.append(target)
            zip_bytes, summaries = clean_files(paths, TEMPLATE_DIR)

        st.success("数据清洗完成。")
        c1, c2, c3 = st.columns(3)
        c1.metric("文件数", len(summaries))
        c2.metric("成功行数", sum(item.success_rows for item in summaries))
        c3.metric("失败行数", sum(item.failed_rows for item in summaries))

        st.download_button(
            "下载清洗结果压缩包",
            data=zip_bytes,
            file_name=f"数据清洗结果_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip",
            mime="application/zip",
            type="primary",
            use_container_width=True,
        )

        st.subheader("清洗摘要")
        st.dataframe(
            [
                {
                    "文件名": item.file_name,
                    "识别供货商": item.supplier,
                    "总行数": item.total_rows,
                    "成功行数": item.success_rows,
                    "失败行数": item.failed_rows,
                }
                for item in summaries
            ],
            use_container_width=True,
            hide_index=True,
        )
    except Exception as exc:
        st.error(f"数据清洗失败：{exc}")
        st.exception(exc)
