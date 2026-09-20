# 导出与校验脚本

运行环境：Python 3.9+、openpyxl。Codex桌面可通过工作区依赖工具定位带openpyxl的Python；否则选已有环境，缺依赖时在任务虚拟环境安装openpyxl，不修改全局Python。

## 数据契约

先运行schema取得完整表头和JSON骨架。顶层sheets按01、02、03、04、05、06、07、09编号；每项rows是二维数组，字段顺序与columns一致。filename/columns仅帮助编辑，实际输出使用脚本固定定义。01另有seeds二维数组，4列，按用户基础词去重后填写，至少1行，不设5–10个限制。

01先按[keyword-planning.md](keyword-planning.md)完成逐词调用和来源映射；rows填写汇总规划，seed_tables填写各词CSV数据。脚本不调用Google Ads技能，静态检查通过不代表真实查询完成。

每个seed_tables元素对应一个基础词，结构如下（复制真实CSV数据，不能用示例数字填充）：

~~~json
{
  "seed": "用户的基础词原文",
  "status": "completed",
  "context": "实际语言、地区、期间、网络、币种、导出时间",
  "source_file": "_核验/关键词/批次ID/关键词明细.csv",
  "metadata": [["CSV表头前的原始元数据行"]],
  "keyword_column": 0,
  "columns": ["完整原始CSV表头，保持顺序"],
  "rows": [["每行全部原始字符串单元格"]]
}
~~~

metadata可为空数组；CSV数据使用csv.reader读取，空值保留空字符串、指标不转数值。keyword_column为已核实的关键词列索引（从0开始），不能按未知本地化表头猜测。status为completed、no_results或blocked；no_results保留表头、rows为空；blocked的columns/rows为空，在context写原因，source_file可为空。每个seeds基础词对应一个seed_tables元素，不能因受阻省略。脚本生成合法且唯一的子表名，把运行信息及CSV元数据保存到A1批注，并记录指标摘要以在重读时发现改单元格或换列；这不能替代原CSV逐格核对。

- 03/05/06/07的计数字段可填null，build自动计算，成品重读时复核。
- 参考链接为一个或多个HTTP(S) URL，用JSON换行符分隔；不要混入解释文字。
- notes为字符串数组。数量/字段确实不足时逐项写目标、实际、缺口、原因，导出为“说明”工作表。notes只允许带说明的部分交付，不能豁免超长、重复、错误URL和禁用词。
- 06字符限制填正整数，非文本字段填“不适用”，未知填“平台复核”；内容缺失用null，notes及行备注解释，结果保留警告。
- 普通英文精确计数；W/F双宽文字按2作辅助计数，非ASCII广告文本提示人工复核，不宣称覆盖Google所有字符规则。
- 脚本不联网，不生成研究结论，不判断来源是否支持业务声明。

## 命令

下面使用默认安装路径；若CODEX_HOME不同或在工作区副本运行，修改ads_skill_dir为实际路径。ads_python需指向包含openpyxl的Python。

~~~bash
ads_skill_dir="$HOME/.codex/skills/google-ads-material-prep"
ads_python=python3
mkdir -p '谷歌广告建组与市场调研资料/_核验'
"$ads_python" "$ads_skill_dir/scripts/materials.py" schema > '谷歌广告建组与市场调研资料/_核验/materials.json'
~~~

根据已核验资料填好JSON，再生成八份表格：

~~~bash
"$ads_python" "$ads_skill_dir/scripts/materials.py" build --data '谷歌广告建组与市场调研资料/_核验/materials.json' --out '谷歌广告建组与市场调研资料'
~~~

完成08的HTML后，重读成品并保存检查结果：

~~~bash
"$ads_python" "$ads_skill_dir/scripts/materials.py" check --out '谷歌广告建组与市场调研资料' --require-html > '谷歌广告建组与市场调研资料/_核验/表格校验.json'
~~~

部分Excel更新可用 --steps 09；确认覆盖选定文件时加 --replace，不覆盖其他步骤。HTML由执行者按[research.md](research.md)的五章结构单独编写；仅检查08时运行：

~~~bash
"$ads_python" "$ads_skill_dir/scripts/materials.py" check-html --out '谷歌广告建组与市场调研资料'
~~~

## 检查结果

- invalid / 退出码1：结构、字符、重复、URL或未说明数量问题；修正后再导出。
- needs_review / 退出码0：已说明缺口、待核验字段或CTA识别等警告；逐条处理或披露，不能称完整。
- checks_passed / 退出码0：仅脚本检查通过；还须人工核验事实、落地页、研究及排版。HTML检查要求五个指定section及各自h2、内联CSS、viewport、UTF-8和footer内四个可点击入口；不证明研究质量，也不能代替人工检查是否夹带额外内容。

notes须真实逐项说明；不要为了消除错误随意填notes、将“平台复核”改为“不适用”，或用空表冒充完成版。
