# 亚马逊竞品表（独立Codex版）

面向Codex的可安装Skill：根据美国站搜索词整理前3页自然竞品，固定32列Excel、原生单元格主图和WPS商品链接。相邻amazon-search-competitor-db为既有Grok版，独立维护。

## 安装

把下列地址发给目标Codex，并说“使用skill-installer安装这个Skill”：

https://github.com/abin798/grok-skills/tree/main/skills/amazon-search-competitor-db-codex

或用Codex自带安装器：

```text
python <skill-installer目录>/scripts/install-skill-from-github.py --repo abin798/grok-skills --path skills/amazon-search-competitor-db-codex
```

安装后下一条消息使用：

```text
使用 $amazon-search-competitor-db-codex，分析美国站搜索词 <搜索词>，生成32列竞品Excel。
```

后续不进行商标调研，已删除原末尾三列：商标持有人、商标持有人国家/地区、商标信息。主表共32列，截止AF「卖家主页」。

## 运行条件

目标Codex需有Spreadsheets工作流的bundled Node、Python、artifact-tool，路径从目标环境查询；卖家精灵提供商品/真实更新时间，Sorftime可补公司地址。工具名称和权限以目标环境为准。Skill不含密钥、账号或真实调研文件，安装不会复制另一电脑的授权。

缺失资料应继续合理补查；不能核实的明确留缺口。第三方公司资料保留来源、底色和核实状态。完整流程见SKILL.md。

## 离线验收

在已具备上述表格运行时的目标Codex中，可运行以下测试。它只生成合成图片与数据，不访问Amazon或付费接口：

```text
python <本Skill目录>/tests/smoke_test.py --work-dir <本次测试目录>
```
