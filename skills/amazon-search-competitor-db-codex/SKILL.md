---
name: amazon-search-competitor-db-codex
description: 在Codex中按亚马逊美国站搜索词制作前3页自然排名竞品Excel，固定33列，含FBA配送费、包装尺寸重量、原生单元格内主图、WPS商品链接、真实更新时间和卖家主体。用于明确的竞品表制作请求或该工作流中的搜索词续用。
---

# 亚马逊搜索词竞品表（Codex）

以用户原始搜索词生成 `当季热销品竞品数据库--<搜索词><YYYYMMDD>.xlsx`，日期取交付时北京时间，保留搜索词空格；非法路径分隔符安全替换并说明。中文回复。主表「竞品明细」固定33列及顺序，另保留「变体明细」「说明」，不得为标注来源增删主表列。主表以AG列「卖家主页」结束。不设原价、折扣或商标列，不再进行商标查询、补查或商标持有人背景调研，也不计算折扣。字段见 [references/schema.json](references/schema.json)。这是独立Codex Skill，既有Grok版不属于本目录。未建立该工作流时，不把无关的孤立词语解释为付费调研请求。

33列顺序（A–AG）：自然排名、页码、页内位置、ASIN、主图、标题、品牌、价格、近一月销量、变体数量、上架时间、FBA配送费、包装尺寸(cm)、包装重量(g)、评分、评论数、大类BSR排名、大类、小类BSR排名、小类名称、父ASIN、库存状态、卖家、配送方式、主体所在国家/地区、公司地址、卖家精灵数据更新时间、卖家ID、店铺名称、公司主体、卖家ASIN数量、卖家最佳自然排名、卖家主页。

## 采集

在当前任务独立 `work/` 子目录工作，脚本相对本Skill定位。采集amazon.com前3页自然商品，排除Sponsored、品牌及视频广告组件。保存HTML、URL、采集起止时间及可核实邮编。按ASIN去重，保留首次页码和自然卡片页内位置，连续编号；不能拿数据库关键词排序冒充本次自然排名。

```text
python3 <skill>/scripts/collect_amazon.py --keyword '<搜索词>' --work-dir '<本次目录>' --details
```

助手只抓页面证据并下载真实主图，不调用付费API。遇验证码、访问错误、重复分页或详情ASIN不匹配时保留证据并注明覆盖，不反复请求受阻页面或绕过验证码。邮编未核实不能声称已设置。离线格式见 [references/collector.md](references/collector.md)。

## 补齐资料

执行 [references/enrichment.md](references/enrichment.md) 的补齐规则；工具和限流处理见 [references/sellersprite.md](references/sellersprite.md)。优先用目标Codex已有的卖家精灵、Sorftime连接，不重复索取密钥，不将其写入Skill、表格或回复。安装Skill不授予外部数据权限。

- 按确切ASIN合并。批量接口可能扩展同父子体并分页，完整读取后精确匹配；额外变体只进副表。BSR采用对应标签或可追溯映射，不取祖先类目代替。
- AA列缺更新时间时继续查详情的真实 `updatedTime`，保存原始UTC毫秒 `_sellersprite_updated_at_ms`，由生成器按北京时间取日期。禁止查询时刻、createdTime或销量快照时间代填。分钟流控缓存后等待窗口恢复，不当作永久缺失。
- 公司地址（Z）和公司主体（AD）优先Amazon正式公司及地址；页面不可得时，可用按卖家ID精确查询的第三方明确字段，标来源及未直接核实，不把已有资料全部留空。品牌和店铺不能推断为公司；Amazon自营无正式主体/地址证据则留空。
- `1K+ bought in past month` 仅取下限，保留 `_sales_display` 和 `_sales_lower_bound: true`。第三方估算只有周期/粒度明确匹配才能补主表并标橙。未确认子体粒度的同父units保留副表/证据，不作子体销量、不相加。
- FBA配送费、包装尺寸(cm)、包装重量(g)只来自 Sorftime `product_detail`（每个ASIN 1点）：`fba_fee` 为美元，包装尺寸和重量按接口的厘米、克原样取值，不换算。尺寸写成「长 × 宽 × 高」，重量写成整数。先查任务目录里按ASIN键控的 `sorftime_cache.json`，已有记录不重复扣点；成功后原子写回。缺失、空值和负数保持 null，不猜测、不用相似商品或父体填补。

## 输入与生成

JSON记录键采用原始表头，缺失为null。`_image` 为本地PNG/JPEG（相对JSON目录或绝对路径），逐字段来源用 `_sources`，来源性质用 `_field_provenance`。第三方公司资料浅黄FFF2CC，估算销量浅橙FCE4D6，且浅橙只涂在近一月销量。事实、状态前缀和来源由agent据证据填，生成器不猜。详细属性见 [references/schema.json](references/schema.json)。

复用旧输入时，移除旧商标字段、来源、详情及相关说明，不将其转移到副表或「说明」页；保留品牌和卖家公司资料。同时移除原价/参考价和折扣，不把它们改写进别的列。

用目标Codex的Spreadsheets工作流及 `load_workspace_dependencies` 找bundled Node/Python与artifact-tool，再运行：

```text
python3 <skill>/scripts/render_workbook.py <数据.json> --output <交付.xlsx> --preview-dir <预览目录>
```

运行时见 [references/runtime.md](references/runtime.md)。D列ASIN为文本，并带指向 `https://www.amazon.com/dp/{ASIN}` 的超链接。价格和FBA配送费为 `$#,##0.00`。整数列（自然排名、页码、页内位置、近一月销量、变体数量、评论数、两个BSR、包装重量、卖家ASIN数量、卖家最佳自然排名）为 `#,##0`。库存状态为「无购买框」的行，价格留空。K列上架时间和AA列卖家精灵数据更新时间都是Excel整数日期、格式 `yyyy-mm-dd`、不含时间；更新时间取北京时间的日期。这两列列宽为12，避免WPS显示井号。不能以t=d的ISO时刻串交付。表头行高40并换行。冻结F2，全表筛选。

主图必须是原生Place in Cell内嵌图，不能替换成浮动图片或远程IMAGE公式。默认WPS兼容在图片上方加透明链接热区，D列ASIN也有文字备用链接；E原生图片和单元格链接保留。纯Excel环境可用 `--no-wps-click-layer`。最终文件不能经不保留richData的库再次保存。数量动态，不保留示例行数限制。

缺图默认报错，先从真实来源补齐；确实无法取得时可用 `--allow-missing-images` 明确部分交付，不能宣称图已全。零商品报告未取得结果，用户需要时才交付空表。

## 验收与交付

运行 `validate_workbook.py <文件> --input <数据.json>` 核对33列表头及顺序、没有原价/折扣/商标列、新列数字格式、浅橙只出现在近一月销量、两个日期列只有日期、逐行ASIN/图像、源字节、richData/metadata、D/E/卖家链接、透明热区、F2冻结和筛选。抽查真实更新时间及来源、日期、底色和长文本，查看所有表预览。有目标WPS/Excel时实际打开并点击两个不同ASIN的E图；XML链接存在不等于WPS可点击。透明层随单元格移动/缩放，不同版本排序/筛选行为需实测，未测不声称已测，D链接可核对目标。

交付到当前聊天约定的outputs。简短报告商品数、主要卖家国家、证据支持的头部主体及缺口（含FBA、尺寸、重量、销量缺多少），区分直接资料与第三方候选。计费只报实证，未返回点数就说明，不把请求次数当费用。卖家精灵和Sorftime分开计数。删除/改写过时中间说明。真实HTML、图片、API响应和表格留任务目录，不进入公开Skill。
