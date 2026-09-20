# 来源入口

01的数据获取指定调用**Google Ads Keywords to Sheets or CSV**（`google-ads-keyword-to-sheets`），逐用户基础词运行详细CSV模式，每词一个Excel子表；见[keyword-planning.md](keyword-planning.md)。帮助链接是规则来源，不能替代真实CSV。

以下不是已取得的客户/市场数据。每次重新打开相关来源，核对标题、适用范围和访问状态；失效时查同机构现行文档，保留替换依据。

| 步骤 | 官方来源 | 用途 |
|---|---|---|
| 01 | https://support.google.com/google-ads/answer/7337243 | 关键词规划师，使用时核验可访问状态 |
| 02 | https://support.google.com/google-ads/answer/2497941 | 受众群体与系列适用范围 |
| 03 | https://support.google.com/google-ads/answer/7684791 | RSA描述≤90、双宽计数、单条广告资源数 |
| 04 | https://support.google.com/google-ads/answer/2375416 | 站内链接，链接文字≤25 |
| 04 | https://developers.google.com/google-ads/api/reference/rpc/v22/SitelinkAsset | 两行说明各≤35；使用时核对现行版本 |
| 05 | https://support.google.com/google-ads/answer/6079510 | Callout≤25及双宽规则 |
| 06 | https://support.google.com/google-ads/answer/9470665 | 原流程表单入口，使用时核验 |
| 06 | https://support.google.com/google-ads/answer/16726130 | 创建表单的官方步骤 |
| 06 | https://developers.google.com/google-ads/api/reference/rpc/v22/LeadFormAsset | 表单字段、隐私政策及CTA；使用时核对现行版本 |
| 07 | https://support.google.com/google-ads/answer/6280012 | 允许的结构化摘要标头 |
| 07 | https://support.google.com/adspolicy/answer/6283300 | Models/Service catalog/Types语义要求 |

2026-09-20建Skill时核验：原 `answer/6324971` 为“Create a campaign”，不是RSA规则；原 `answer/6077124` 为“Introduction”，故采用上表对应主题页。`9470665` 未成功抓取；`16726130` 检索可见，但正文遇429。原流程Headline≤30、Description≤200保留为制作上限；Business Name、CTA及提交反馈限制执行时继续核验。

## 原流程公开工具与市场入口

- Google Trends：https://trends.google.com/trends/ （另有 https://trends.google.com/ ）
- AnswerThePublic：https://answerthepublic.com/
- LetterCount：https://www.lettercount.com/ （辅助计数，最终用本地脚本）
- WordStream预览：https://www.wordstream.com/google-ads-preview-tool
- Think with Google：https://thinkwithgoogle.com/
- WITS：https://wits.worldbank.org/
- ITC Market Access Map：https://www.macmap.org/
- Google Market Finder：https://marketfinder.thinkwithgoogle.com/ （建立时跳转 `/intl/en_us/`，详细功能与数据仍须实际查询）
- 可选贸易替代：UN Comtrade https://comtradeplus.un.org/ 及目标国海关/统计机构；使用前实际访问。

这些入口不保证全部功能免登录/免费。遇到访问限制记录原因，不把工具名当证据，不声称未运行的工具已使用。Excel参考链接含适用官方/公开工具来源，并补充每行具体事实来源；动态数据记录查询条件。
