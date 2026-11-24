# grok-skills

William 团队共用的 Grok Bot 技能库。

## 安装方法（同事）
1. 在自己的 Grok Bot 里连接 GitHub（公开库不连也可以读取）。
2. 把技能目录地址发给自己的 Grok Bot，例如：
   `https://github.com/abin798/grok-skills/tree/main/skills/amazon-search-competitor-db`
   并说：「读取这个目录下的 SKILL.md，按原样保存成我的技能，scripts 和 ref 文件夹也下载到你电脑上的 /home/box/agent-data/amazon_competitor_kit/」。
3. 卖家精灵 MCP 密钥需要自己配置：让 Grok Bot 用安全输入框保存为环境变量 `SELLERSPRITE_MCP_KEY`。密钥绝不要写进本仓库。

安装「团队调度」时，把目录换成：
`https://github.com/abin798/grok-skills/tree/main/skills/team-dispatch`
并说：「读取这个目录下的 SKILL.md，按原样保存成我的技能」。该技能是调度手册，导入后不能自动创建 Bot、侧边栏分组或例行任务。

## 技能列表
| 技能 | 用途 |
|---|---|
| [amazon-search-competitor-db](skills/amazon-search-competitor-db/) | 只发一个亚马逊搜索词，自动生成「当季热销品竞品数据库」Excel：前 3 页自然排名竞品、嵌入主图、卖家主体、商标信息 |
| [团队调度](skills/team-dispatch/) | 收到任务时先判断该由哪个部门、哪个 Bot 处理；需要调整分工、新建 Bot，或安排重要结论复核时使用 |

> 注意：本仓库为公开仓库，请勿提交任何密钥、账号或客户数据。
