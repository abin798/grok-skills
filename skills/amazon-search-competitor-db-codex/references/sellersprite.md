# 工具、分页与恢复

以当次可调用工具定义及响应为准，工具可见不代表权限/余额已核实。不要将已有连接密钥复制进Skill。

| 目的 | 工具示例 | 要点 |
|---|---|---|
| 详情/真实更新时间 | mcp__ss_mcp__asin_detail | 单ASIN，marketplace=US；只用真实updatedTime |
| 批量商品快照 | mcp__ss_mcp__competitor_lookup | asins最多40个；variation=N可能扩展同父系列并分页 |
| 历史观测 | mcp__ss_mcp__keepa_info | 不含销量；返回窗口起点不证明上架日 |
| 商标 | trademark_country_list/list/detail | 先确认局编码，再按实际品牌归一化精确匹配 |
| 卖家公司/地址 | mcp__sorftime__seller_detail | 美国站，query_type=2，name=卖家ID；明确第三方快照 |

## 批量结果

按响应pages/分页标记读取到结束，不能因输入40项就认为结果最多40项。完整缓存后对输入ASIN精确join，保持搜索页顺序和排名；额外同父变体仅放副表。字段随接口不同，list可能用bsr，详情用bsrRank/bsrLabel；nodeLabelPath首段不一定是当前BSR类目。

units在同父子体中可能完全相同，粒度未确认时不填主表子体销量，也不把同父子体相加。

## X列

批量未返回更新时间时，用详情接口继续补查。把响应ASIN及原始updatedTime UTC毫秒存证，记录 `_sellersprite_updated_at_ms`；由生成器转北京时间。createdTime、查询时刻、amzUnitDate均不能代替商品更新时间。接口确实不返回才留缺口。

## 分钟流控

缓存每个成功ASIN/品牌/卖家响应。限制以服务实际返回为准；默认可从最多4并发、每恢复窗不超过30–35次开始。遇ERROR_MAXIMUM_ACCESS_PER_MINUTE停当前批次，等待窗口后只重试失败/未查项；可用65秒间隔，拆成不超过60秒的等待，继续报告进度。

同一失败项每窗口最多一次；连续两个窗口仍受同一流控则降批量，后续仍无法取得时交付具体缺口，不无限轮询。权限无效、总额度耗尽或不可恢复参数错误应停相应路径，不当作分钟限流重试。请求次数不等于点数。

品牌组合筛选可能大小写敏感。空结果有依据存疑时可做一次简化检索（如text+已核实商标局），本地精确匹配；不把一次空结果当成没有商标。Sorftime共享余额的前后变化不证明永久费率，与卖家精灵点数分别记录。
