# 工具、分页与恢复

以当次可调用工具定义及响应为准，工具可见不代表权限/余额已核实。不要将已有连接密钥复制进Skill。

| 目的 | 工具示例 | 要点 |
|---|---|---|
| 详情/真实更新时间 | mcp__ss_mcp__asin_detail | 单ASIN，marketplace=US；只用真实updatedTime |
| 批量商品快照 | mcp__ss_mcp__competitor_lookup | asins最多40个；variation=N可能扩展同父系列并分页 |
| 历史观测 | mcp__ss_mcp__keepa_info | 不含销量；返回窗口起点不证明上架日 |
| 卖家公司/地址 | mcp__sorftime__seller_detail | 美国站，query_type=2，name=卖家ID；明确第三方快照 |
| FBA配送费和包装尺寸重量 | Sorftime product_detail | 每个确切ASIN 1点；fba_fee、厘米尺寸、克重量原样取值 |

## 批量结果

按响应pages/分页标记读取到结束，不能因输入40项就认为结果最多40项。完整缓存后对输入ASIN精确join，保持搜索页顺序和排名；额外同父变体仅放副表。字段随接口不同，list可能用bsr，详情用bsrRank/bsrLabel；nodeLabelPath首段不一定是当前BSR类目。

units在同父子体中可能完全相同，粒度未确认时不填主表子体销量，也不把同父子体相加。

## AA列

批量未返回更新时间时，用详情接口继续补查。把响应ASIN及原始updatedTime UTC毫秒存证，记录 `_sellersprite_updated_at_ms`；由生成器转成北京时间的日期，主表只写 `yyyy-mm-dd`。createdTime、查询时刻、amzUnitDate均不能代替商品更新时间。接口确实不返回才留缺口。

## FBA与包装缓存

`product_detail` 的结果按ASIN放进本次任务目录的 `sorftime_cache.json`。键是ASIN，值至少包含 `fba_fee`、`package_size_cm`、`weight_g`；没有的字段写 null，不要省略成另一个商品的值。读取时先命中缓存，只为缓存里没有的ASIN发起请求，避免同一ASIN扣两次点。写回用临时文件加 rename。多个搜索词共用这一份任务缓存，不把缓存路径写成某一环境的固定绝对路径。

合并时只认确切ASIN。尺寸保持厘米，重量保持克，FBA配送费保持美元。空值、负数和无法解析的值留在缺口里，不估算。

## 分钟流控

缓存每个成功ASIN/卖家响应。限制以服务实际返回为准；默认可从最多4并发、每恢复窗不超过30–35次开始。遇ERROR_MAXIMUM_ACCESS_PER_MINUTE停当前批次，等待窗口后只重试失败/未查项；可用65秒间隔，拆成不超过60秒的等待，继续报告进度。

同一失败项每窗口最多一次；连续两个窗口仍受同一流控则降批量，后续仍无法取得时交付具体缺口，不无限轮询。权限无效、总额度耗尽或不可恢复参数错误应停相应路径，不当作分钟限流重试。请求次数不等于点数。

Sorftime共享余额的前后变化不证明永久费率，与卖家精灵点数分别记录。product_detail 每成功一个新ASIN记1点；缓存命中不计点。
