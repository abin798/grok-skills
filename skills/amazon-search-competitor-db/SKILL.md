---
name: 亚马逊搜索词竞品表
description: >-
  用户只发来亚马逊搜索词时用：抓 amazon.com 前 3 页自然排名竞品，生成带嵌入主图、FBA 配送费、包装尺寸重量和卖家主体信息的 33 列标准
  Excel 竞品表。
---
# 亚马逊搜索词竞品表

用户只发搜索词（例如 "napkin rings"，可以一次发多个，每行一个或用逗号隔开）时，就是要这张完整的表。直接开工，不用确认；回复全部用中文。工具包在 `/home/box/agent-data/amazon_competitor_kit/`：
- 中间流程脚本：`embed_v2.py`、`add_links_v2.py`、`restructure8.py`；富数据模板 `ref/xl/richData`
- 后处理脚本（都是纯 XML，用法统一为 `python3 脚本 输入.xlsx 输出.xlsx`）：`add_asin_links.py`、`drop_cols_IJ.py`、`compact_format.py`、`add_fba_fee_col.py`、`add_pkg_cols.py`、`reorder_cols.py`
- 示例：`template_final_example.xlsx`（旧版 32 列，只看样式参考）、`template_stage_example.xlsx`

## 1. 抓取（amazon.com 美国站，配送邮编 60463）
- 用 curl_cffi 抓搜索结果第 1–3 页，只要自然排名，剔除所有 Sponsored 和广告组件。
- 按 ASIN 去重，排名在 3 页之间连续编号，同时记录页码和页内位置。
- 进每个商品的详情页，取标题、品牌、价格、评分、评论数、BSR（大类和小类分开）、库存状态、变体（另做「变体明细」表）、Buy Box 卖家 ID，并下载主图。
- 被亚马逊 WAF 拦（首页 202 挑战页、搜索 503）时立即停手，不要连续重试设邮编；间隔 20–30 分钟轻量探活。用户同意后可改走第 1A 节的备用方案。绝不编数据。

## 1A. 备用方案：亚马逊被拦时用 Sorftime 替代前台数据
只在亚马逊前台被拦、且用户同意时使用；站点仍是美国站，只要自然排名，33 列格式不变。
- Sorftime MCP `keyword_search_results`（美国站）取第 1–3 页，共约 60 个自然排名结果，按原样保留，不剔除和关键词不太相关的商品，也不擅自换搜索词。
- `product_detail` 取标题、品牌、价格、评分、评论数、小类 BSR、父 ASIN、变体数、上架时间、卖家、配送方式、FBA 费、包装尺寸和重量；详情和搜索结果不一致时用详情的值。上架日期为 null 时用 listing 日期；明显的占位值（如 1970-01-01）留空。
- `product_content` 取主图（直接返回，不用 write_to_file），`seller_detail` 取卖家名称、国家、公司主体和地址。
- 这种模式下「大类BSR排名」「库存状态」留空；亚马逊自营行的公司地址和卖家主页留空；Sorftime 卖家为 N/A 的按无购买框处理，价格和卖家相关列都清空。
- 所有结果按 ASIN 或卖家 ID 缓存，跨搜索词复用。

## 2. 近一月销量：SellerDex amzUnit
- 「近一月销量」只填亚马逊页面上「过去一个月卖出 X+」的取整值，来源是 SellerDex MCP `competitor_lookup` 返回的 `amzUnit` 字段（参数 variation=Y）。
- 不用 `units`、Sorftime 或卖家精灵的估算值；`amzUnit` 为空或 0 就留空。格式是纯整数 `#,##0`，不加底色。
- SellerDex 地址 `https://mcp.sellerdexmcp.com/mcp`（POST，curl 加 User-Agent），请求头 `secret-key` 取环境变量 `SELLERDEX_MCP_KEY`，绝不打印密钥。每次最多 40 个 ASIN，扣 1 点；失败的调用不扣点。
- 先查缓存，只查新 ASIN。开工前和结束时都查一下余额；余额为 0 时停下来告诉用户，不要用别的数据顶替。

## 3. 卖家精灵 MCP 补数据（可用时）
- 地址 `https://mcp.sellersprite.com/v2/mcp`，请求头 `secret-key` 取环境变量 `SELLERSPRITE_MCP_KEY`，绝不打印密钥。成功调用一次扣 1 点，尽量批量；额度用完时跳过这一步。
- `asin.detail`：每批 20 个，取上架时间 availableDate、卖家、配送方式、父 ASIN、记录更新时间。
- 上架时间仍为空的，用 `keepa.trend` 的最早记录近似补上。
- 已查过的 ASIN 结果缓存下来，同一批多个搜索词之间复用，不重复扣点。ISBN 形态的图书 ASIN 卖家精灵不收录，留空即可。

## 4. Sorftime 补 FBA 配送费和包装尺寸重量
- 用 Sorftime MCP 的 `product_detail`（每 ASIN 扣 1 点）取 `fba_fee`、包装尺寸（厘米）、包装重量（克），原样取值不换算；返回 -1 或 null 时留空。
- 结果写进共享缓存 `/workspace/sorftime_cache.json`（键为 ASIN），`add_fba_fee_col.py` 和 `add_pkg_cols.py` 从这里读；先查缓存，只补新 ASIN。

## 5. 卖家主体
- 正常模式：卖家名称和国家取自卖家精灵；公司全称和地址并发抓取亚马逊卖家资料页 `https://www.amazon.com/sp?seller=ID`。备用模式下用 Sorftime `seller_detail`。亚马逊自营的地址留空。
- 第三方卖家必须写实际店名，不能误写成 Amazon。
- 不做任何商标（USPTO）调研，表里也不设商标列。

## 6. 「竞品明细」表结构（最终 33 列，顺序固定）
A 自然排名、B 页码、C 页内位置、D ASIN、E 主图、F 标题、G 品牌、H 价格、I 近一月销量、J 变体数量、K 上架时间、L FBA配送费、M 包装尺寸(cm)、N 包装重量(g)、O 评分、P 评论数、Q 大类BSR排名、R 大类、S 小类BSR排名、T 小类名称、U 父ASIN、V 库存状态、W 卖家、X 配送方式、Y 主体所在国家/地区、Z 公司地址、AA 卖家精灵数据更新时间、AB 卖家ID、AC 店铺名称、AD 公司主体、AE 卖家ASIN数量、AF 卖家最佳自然排名、AG 卖家主页。

不要「原价/参考价」「折扣」和任何商标列。

格式规则：
- 表头只写中文（包装尺寸、重量两列带单位括号），表头行高 40，窄列表头能完整换行。
- 价格和 FBA 配送费格式 `$#,##0.00`；整数列（近一月销量、评论数、BSR、变体数量、排名等）用千位逗号 `#,##0`。
- 库存状态为「无购买框」的行，价格一律清空。
- 上架时间和卖家精灵数据更新时间都只留 `yyyy-mm-dd`，这两列列宽 12（WPS 下不显示井号）。
- 包装尺寸写成「长 × 宽 × 高」。
- 库存状态统一写成「有货」「库存紧张（剩N）」「无购买框」。
- D 列 ASIN 是文本，带指向 `https://www.amazon.com/dp/{ASIN}` 的蓝色下划线超链接。E 列主图是嵌入单元格的图片（Place in Cell），**不加超链接**，也不能有浮在单元格上的图片。卖家主页是可点击超链接。用户用 WPS，靠 ASIN 列打开商品页。
- 冻结窗格、全表筛选开启。

其他 sheet：「变体明细」表头只写中文；「说明」中文写清搜索词、抓取时间（北京时间）、数据来源模式（亚马逊前台或 Sorftime 备用）、广告剔除规则和各列数据来源，不写商标和折扣。

## 7. 出表流水线（关键：嵌入图片）
1. openpyxl 只能改中间文件；最终文件一旦被 openpyxl 保存，嵌入图片就毁了，后面所有步骤只改 zip 内 XML。
2. 中间文件改好后跑 `embed_v2.py`（主图占位单元格换成 `t="e" vm=` 图片单元格），再跑 `add_links_v2.py`，但只加卖家主页链接，不给 E 列主图加链接（worksheet 根节点声明 xmlns:r）。这两个脚本写死了行数和列，要按商品数调整。
3. 依次跑：`add_asin_links.py` → `drop_cols_IJ.py`（中间表仍是旧 32 列含原价/折扣时才需要）→ `compact_format.py`（千位逗号、日期只留年月日、行高列宽）→ `add_fba_fee_col.py` → `add_pkg_cols.py` → `reorder_cols.py`（按表头名称重排成第 6 节的 33 列顺序）。
4. 已知问题：`add_pkg_cols.py` 有时会把 N 列（包装重量）误涂成 `FCE4D6` 橙色，每份都要检查，有就把这些格子改指向一个无填充的数字样式。I 列也不能有任何底色。
5. 校验每份：33 个表头与第 6 节逐一一致；`xl/media` 图片数等于商品行数，没有 drawing/vml 浮层；D 列链接数等于商品数，E 列链接数为 0；I 列、N 列橙色数为 0；记录 I 列空格数，以及 FBA、尺寸、重量的覆盖数。
6. 每份出表后算 sha256，追加到当天的哈希记录文件，并用 `sha256sum -c` 全部校验通过。

## 8. 提速（多个搜索词时尤其重要）
- 详情、卖家、主图、SellerDex 结果都做全局缓存，跨搜索词共享，只查新的，临时文件加 rename 原子写回。
- 亚马逊直抓时并发适度，持续监控验证码和 503 比例，失败率上升就降回；一旦整体被 WAF 封锁就按第 1 节停手。
- 多个搜索词时每路用独立中间目录；每个阶段落盘，中断后从已有文件续跑，不重抓。

## 9. 交付
- 文件名固定为 `当季热销品竞品数据库--<搜索词><YYYYMMDD>.xlsx`：搜索词原样保留空格，日期取北京时间当天，紧跟在搜索词后面，不加分隔符，例如 `当季热销品竞品数据库--napkin rings20261114.xlsx`。保存到 `/workspace/`，作为附件发出。
- 多个搜索词时每个词一份，做好一份发一份。
- 附几句中文说明：商品数量、近一月销量空了几行、亚马逊自营、FBM 和无购买框行数、没补上的数据（FBA、尺寸、重量、上架时间缺几个）、SellerDex 用了多少点和剩余多少点；余额出现不是自己用掉的变化时要提醒用户。
- 用户要求下载到本地时，直接放进其电脑的「下载」文件夹，不建子文件夹，每次只并行复制 2–3 个，最后核对文件大小。

## 10. 同步
这个技能每次修改后，都要同步到用户的 GitHub 技能仓库。
