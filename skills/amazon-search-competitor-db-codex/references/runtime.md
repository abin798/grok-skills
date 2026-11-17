# Codex运行时和WPS兼容

Python助手使用3.9+标准库。基表由Codex Spreadsheets工作流提供的bundled Node和artifact-tool创建，不修改系统或依赖目录。用目标环境load_workspace_dependencies查路径；默认当前用户 `~/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/`，不匹配时设置工具返回的CODEX_NODE_BIN和CODEX_NODE_MODULES。未提供表格运行时时说明缺少依赖，不能声称安装Skill就包含全部数据服务。

```text
python3 <skill>/scripts/render_workbook.py <input.json> --output <output.xlsx> --preview-dir <previews>
python3 <skill>/scripts/validate_workbook.py <output.xlsx> --input <input.json>
```

默认WPS兼容；--no-wps-click-layer关闭透明层；--allow-missing-images只用于明确部分交付。以--help为准。

## 图片和链接

E列用t=e/vm、XLRICHVALUE、richData和内置PNG/JPEG保留原生Place in Cell图片。D列ASIN、E列主图和AG列卖家页有标准Hyperlink样式及链接，D列文字链接指向 `https://www.amazon.com/dp/{ASIN}`。WPS可能读取E链接却拦截图片点击，所以默认在每个图片单元格上方加透明DrawingML矩形链接热区（xdr:sp+a:hlinkClick），不是浮动替代图片。twoCellAnchor随单元格移动/缩放；D文字链接提供另一入口。

该点击方案在一次真实Mac WPS任务中观测成功，不保证所有版本排序/筛选行为。目标版本可用时核查排序/筛选后的ASIN对应；结构验证不能声称已经原生验证。透明形状是sp，替代图片是pic，不要把所有drawing都误判为浮动图片。

_localImage的Text是替代文字，不能编造成URL字段。最终文件不得经不保留richData的库再次保存。已有richData/metadata的任意历史工作簿不是生成器基表输入，避免覆盖旧图片索引。

## 日期与测试

WPS可能把ISO t=d直接显示成时刻字符串。K列上架时间和AA列卖家精灵数据更新时间都写Excel整数日期，格式 `yyyy-mm-dd`，不含时间；AA列先把真实更新时间换到北京时间再取日期，不能用查询时间代填。这两列列宽为12，避免WPS把日期显示成井号。整数列用 `#,##0`，价格和FBA配送费用 `$#,##0.00`。表头行高40并换行。浅橙FCE4D6只出现在近一月销量。

安装测试用合成数据覆盖1条、超过143条、缺图和零行，核对图字节、链接、日期、底色及原子交付，不访问Amazon或付费API。测试文件留临时工作目录，不进入发布Skill。
