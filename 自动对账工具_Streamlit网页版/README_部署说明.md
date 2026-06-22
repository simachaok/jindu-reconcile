# 泾都药业自动对账工具 Streamlit 网页版

## Streamlit Cloud 免费部署

1. 把本文件夹上传到一个 GitHub 仓库。
2. 打开 Streamlit Community Cloud，选择 `Create app`。
3. Repository 选择刚才的仓库。
4. Branch 一般选择 `main`。
5. Main file path 填：

```text
streamlit_app.py
```

6. 点击 Deploy。

## 仓库里需要保留的文件

- `streamlit_app.py`
- `auto_reconcile_tool.py`
- `requirements.txt`
- `模板文件/业务规则配置总表.xlsx`
- `模板文件/高开票对账模板.xlsx`
- `示例文件/`
- `说明资料/`

## 使用方式

网页打开后上传：

- 药店数据表
- 银行数据表
- 申报表

可选上传：

- 业务规则配置总表
- 高开票对账表

执行后页面会提供 `对账结果汇总表` 下载。

## 隐私提醒

Streamlit Community Cloud 是免费云平台。业务 Excel 会上传到云端服务器参与本次计算，虽然本程序不会主动保存上传文件，但如果数据特别敏感，仍建议优先使用局域网网页版或桌面版。
