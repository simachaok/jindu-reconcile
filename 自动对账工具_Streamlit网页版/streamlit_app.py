from __future__ import annotations

import tempfile
from pathlib import Path
from typing import Optional

import streamlit as st

from auto_reconcile_tool import FileSelection, UserVisibleError, run_reconcile


APP_NAME = "泾都药业自动对账"
BASE_DIR = Path(__file__).resolve().parent
DEFAULT_CONFIG = BASE_DIR / "模板文件" / "业务规则配置总表.xlsx"
DEFAULT_HIGH_INVOICE = BASE_DIR / "模板文件" / "高开票对账模板.xlsx"


st.set_page_config(
    page_title=APP_NAME,
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)


def save_upload(uploaded_file, folder: Path, default_name: str) -> Optional[Path]:
    if uploaded_file is None:
        return None
    suffix = Path(uploaded_file.name or default_name).suffix
    name = Path(default_name).with_suffix(suffix).name
    target = folder / name
    target.write_bytes(uploaded_file.getbuffer())
    return target


def read_file_bytes(path: Path) -> bytes:
    return path.read_bytes()


def render_header() -> None:
    st.markdown(
        """
        <style>
        .block-container {padding-top: 1.4rem; padding-bottom: 2rem; max-width: 1280px;}
        [data-testid="stSidebar"] {background: #f7f9fc;}
        .jd-hero {
            border: 1px solid #dbe4ef;
            background: linear-gradient(135deg, #ffffff 0%, #eef6ff 58%, #f7fbf3 100%);
            border-radius: 8px;
            padding: 22px 24px;
            margin-bottom: 18px;
        }
        .jd-title {font-size: 30px; font-weight: 760; color: #102033; margin: 0;}
        .jd-sub {font-size: 15px; color: #5b6675; margin: 8px 0 0;}
        .jd-rule {
            display: inline-block;
            margin: 14px 8px 0 0;
            padding: 6px 10px;
            border: 1px solid #dbe4ef;
            border-radius: 6px;
            background: #ffffff;
            color: #223047;
            font-size: 13px;
        }
        .stButton > button, .stDownloadButton > button {border-radius: 6px; min-height: 42px;}
        [data-testid="stMetricValue"] {font-size: 24px;}
        </style>
        <div class="jd-hero">
            <p class="jd-title">泾都药业自动对账</p>
            <p class="jd-sub">上传本次药店、银行、申报表数据，网页端生成带筛选和摘要的 Excel 对账结果。</p>
            <span class="jd-rule">银行1笔付款公司：按公司月汇总</span>
            <span class="jd-rule">陕西海通药业有限责任公司：银行金额 × 1.22</span>
            <span class="jd-rule">高开票：供货商、分部、日期/顺序匹配</span>
        </div>
        """,
        unsafe_allow_html=True,
    )


def main() -> None:
    render_header()

    with st.sidebar:
        st.subheader("规则文件")
        config_upload = st.file_uploader("业务规则配置总表", type=["xlsx"], key="config")
        high_invoice_upload = st.file_uploader("高开票对账表", type=["xlsx"], key="high_invoice")
        use_default_high_invoice = st.checkbox("未上传时使用模板文件中的高开票表", value=False)

        st.divider()
        st.subheader("部署提示")
        st.caption("Streamlit Cloud 只会保存代码仓库里的模板，不会保存你在页面上传的业务数据。")

    left, right = st.columns([1.15, 0.85], gap="large")

    with left:
        st.subheader("本次对账数据")
        drug_upload = st.file_uploader("药店数据表", type=["xls", "xlsx"], key="drug")
        bank_upload = st.file_uploader("银行数据表", type=["xls", "xlsx"], key="bank")
        declare_upload = st.file_uploader("申报表", type=["xls", "xlsx"], key="declare")

        ready = bool(drug_upload and bank_upload and declare_upload)
        run_clicked = st.button("执行自动对账", type="primary", disabled=not ready, use_container_width=True)

    with right:
        st.subheader("当前状态")
        c1, c2 = st.columns(2)
        c1.metric("必传文件", f"{sum(bool(x) for x in [drug_upload, bank_upload, declare_upload])}/3")
        c2.metric("规则文件", "默认" if config_upload is None else "已上传")
        st.info("如果高开票表里有多个分部，程序会按当前银行账户/申报单位自动过滤对应分部。")

        with st.expander("示例与模板", expanded=False):
            st.write("仓库里包含 `模板文件`、`示例文件`、`说明资料` 三个文件夹。")
            if DEFAULT_CONFIG.exists():
                st.download_button(
                    "下载默认业务规则配置总表",
                    data=read_file_bytes(DEFAULT_CONFIG),
                    file_name="业务规则配置总表.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
            if DEFAULT_HIGH_INVOICE.exists():
                st.download_button(
                    "下载高开票对账模板",
                    data=read_file_bytes(DEFAULT_HIGH_INVOICE),
                    file_name="高开票对账模板.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )

    if not ready:
        st.warning("请先上传药店数据表、银行数据表、申报表三个文件。")
        return

    if not run_clicked:
        return

    logs: list[str] = []

    def log(message: str) -> None:
        logs.append(message)

    with st.spinner("正在对账并生成 Excel，请稍等..."):
        try:
            with tempfile.TemporaryDirectory(prefix="jd_reconcile_") as temp_name:
                temp_dir = Path(temp_name)
                input_dir = temp_dir / "input"
                output_dir = temp_dir / "output"
                input_dir.mkdir(parents=True, exist_ok=True)
                output_dir.mkdir(parents=True, exist_ok=True)

                drug_file = save_upload(drug_upload, input_dir, "药店数据表.xls")
                bank_file = save_upload(bank_upload, input_dir, "银行数据表.xls")
                declare_file = save_upload(declare_upload, input_dir, "申报表.xlsx")
                config_file = save_upload(config_upload, input_dir, "业务规则配置总表.xlsx") or DEFAULT_CONFIG
                high_invoice_file = save_upload(high_invoice_upload, input_dir, "高开票对账表.xlsx")
                if high_invoice_file is None and use_default_high_invoice and DEFAULT_HIGH_INVOICE.exists():
                    high_invoice_file = DEFAULT_HIGH_INVOICE

                selection = FileSelection(
                    drug_file=drug_file,
                    bank_file=bank_file,
                    declare_file=declare_file,
                    output_folder=output_dir,
                    config_file=config_file if config_file.exists() else None,
                    high_invoice_file=high_invoice_file,
                )
                output_path = run_reconcile(selection, log)
                result_bytes = output_path.read_bytes()
                result_name = output_path.name

            st.success("对账完成，结果文件已经生成。")
            st.download_button(
                "下载对账结果汇总表",
                data=result_bytes,
                file_name=result_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True,
            )
        except UserVisibleError as exc:
            st.error(str(exc))
        except Exception as exc:
            st.error(f"对账失败：{exc}")
            st.exception(exc)

    with st.expander("运行日志", expanded=True):
        if logs:
            st.code("\n".join(logs), language="text")
        else:
            st.write("本次没有生成日志。")


if __name__ == "__main__":
    main()
