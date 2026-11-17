# amazon-search-competitor-db

见 SKILL.md（33 列最终版流程）。

- `scripts/embed_v2.py`、`scripts/add_links_v2.py`：中间文件生成嵌入单元格主图（Place in Cell）和主图/卖家主页超链接；行数和列位置按一次具体调研写死，使用前按商品数量调整。
- `scripts/add_asin_links.py` → `drop_cols_IJ.py` → `compact_format.py` → `add_fba_fee_col.py` → `add_pkg_cols.py` → `reorder_cols.py`：纯 XML 后处理，用法统一为 `python3 脚本 输入.xlsx 输出.xlsx`，不经 openpyxl，嵌入图片不受影响。
- `add_fba_fee_col.py`、`add_pkg_cols.py` 从 `/workspace/sorftime_cache.json`（Sorftime product_detail 结果，键为 ASIN）读 FBA 配送费和包装尺寸重量。
- `ref/xl/richData/`：嵌入图片所需的 Excel 富数据模板。
- 卖家精灵密钥从环境变量 `SELLERSPRITE_MCP_KEY` 读取，仓库里不含任何密钥。
