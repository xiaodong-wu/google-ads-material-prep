# Google Ads Material Prep

根据目标网站和多个基础关键词，生成8份Excel及1份HTML谷歌广告准备资料。文案默认英文，研究说明默认中文；仅准备本地资料，不创建或发布广告。

## 安装

将本仓库放到 `~/.codex/skills/google-ads-material-prep/`。关键词步骤还需安装 [Google Ads Keywords to Sheets or CSV](https://github.com/xiaodong-wu/google-ads-keyword-to-sheets)。

需要可控制已登录Chrome的工具（例如Codex的`mcp__cua_repl`）、可使用Keyword Planner并下载CSV的Google Ads账号，以及带`openpyxl`的Python环境。`chrome:control-chrome`是可选方案，不是硬性依赖。

## 调用

在Codex中输入：

> 使用 $google-ads-material-prep，为 https://example.com/ 生成谷歌广告准备资料。基础词：product one、product two；语言English，地区全球。每个基础词一个子表。

01逐词查询真实Google Ads数据、导出完整指标CSV，再合并去重与分组。网址仅用于业务分析及落地页匹配，不作为Keyword Planner的网站筛选。

## 输出

默认写入当前目录下的`谷歌广告建组与市场调研资料/`，不同品牌使用独立子目录：

1. 关键词规划表（每基础词独立子表＋汇总）
2. 目标受众画像分析（五维报告＋角色明细，3–4句总结、2–3条建议）
3. 谷歌广告描述语（全站层级＋每栏目5条≤90字符描述，约1/3强CTA，每栏目至少2条有证据的数字卖点）
4. 谷歌广告站内链接
5. 谷歌广告宣传信息
6. 潜在客户表单设置
7. 结构化摘要扩展
8. 海外市场与海关宏观报告（单文件HTML，严格五个核心维度）
9. 谷歌广告投放国家推荐

来源、原始关键词CSV、逐词调用状态与检查结果保存在输出目录的`_核验/`。无法取得的数据明确标注待验证或受阻，不用模型生成数字补齐。静态校验不等于Google Ads审核通过。

完整操作要求见[SKILL.md](SKILL.md)，数据格式与检查命令见[导出说明](references/export.md)。

03不再固定20条或要求至少70字符；数量由真实栏目决定。数字/认证/交期不足时报告证据缺口，不补造承诺。更新前的02/03需按新数据契约补充附表后重新生成；仅升级技能不会改写已有客户文件。
