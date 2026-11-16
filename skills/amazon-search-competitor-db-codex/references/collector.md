# 搜索结果采集助手

`scripts/collect_amazon.py` 仅采集有页面证据的字段，不调用卖家精灵、不消耗其点数。它使用 Python 标准库；若环境已经安装 `curl_cffi` 或 `requests`，自动优先使用。不要求安装新依赖。

```bash
python3 scripts/collect_amazon.py --keyword 'napkin rings' --work-dir /绝对路径/work/napkin-rings --pages 3
```

输出 `competitors.json`、`images/` 主图和 `evidence/` HTML。JSON 可继续由 agent 补充卖家精灵、变体、卖家主体和商标证据，再传给工作簿生成器。`status: search_complete` 只表示本次搜索页采集结束；`search_coverage_complete` 标记搜索覆盖，`details_complete` 和 `images_complete` 分别说明详情、主图覆盖。`enrichment_complete` 默认 `false`，不得据此宣称35列已全部补齐。抓取受阻返回退出码2，但仍保存已取得的记录、覆盖状态与原因。

默认只抓搜索页；加 `--details` 按ASIN顺序请求详情，`--max-details N` 限制数量（默认200）。详情遇验证码、重定向、HTTP错误或ASIN不一致即停止后续详情补充，已有搜索页数据保留。缺失字段留空，不把推荐商品的值冒充目标商品。

`--postcode 90001` 只核验页面顶部实际显示的邮编；脚本不修改会话地址。未核实、与请求邮编不符或不同页面邮编不一致时记录缺口。需要固定邮编时，可在正常浏览器设置后导入保存的页面，不声称脚本已自动设置。

`bought in past month` 的数字写入原有近一月销量列，原文另存 `_sales_display`；带加号的展示值用 `_sales_lower_bound: true` 标明是下限。它不是精确销量，也不是卖家精灵估算。后续若用第三方估算替换，设置 `_sales_estimated: true`、记录来源，并由生成器以原规则的浅橙色标记。

离线输入支持以下结构，整个离线模式不会请求网络：

```text
fixtures/
  page1.html
  page2.html
  page3.html
  details/B012345678.html      # 仅 --details 时需要
  images/网页图片URL中的文件名.png
```

```bash
python3 scripts/collect_amazon.py --keyword 'offline sample' --work-dir /绝对路径/work/check --offline-html-dir /绝对路径/fixtures
```

HTML来自离线导入时，`captured_at` 是本次处理时间；原始页面抓取时间须另行提供，不能把测试处理时间当作真实调研快照时间。广告过滤、字段解析依赖当前页面结构，运行后须抽样核对保留的HTML与自然排名，特别是布局变化、缺页或无结果情况。
