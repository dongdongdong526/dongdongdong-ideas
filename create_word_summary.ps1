$ErrorActionPreference = 'Stop'

$out = 'C:\Users\User\Desktop\清北目标与赚钱思考.docx'
$tmp = Join-Path $env:TEMP ('claude_docx_' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $tmp | Out-Null
New-Item -ItemType Directory -Path (Join-Path $tmp '_rels') | Out-Null
New-Item -ItemType Directory -Path (Join-Path $tmp 'word') | Out-Null
New-Item -ItemType Directory -Path (Join-Path $tmp 'word\_rels') | Out-Null

$paragraphs = @(
    @{ Kind='title'; Text='高中生身份下关于赚钱与清北目标的对话整理' },
    @{ Kind='normal'; Text='本文档整理了刚才的聊天内容，重点保留与你当前“高中生身份”、赚钱想法、清华北大目标相关的核心讨论。' },
    @{ Kind='heading'; Text='1. 身份切换' },
    @{ Kind='quote'; Text='用户：我现在作为高中生的身份再和你聊天' },
    @{ Kind='normal'; Text='回复要点：后续交流按高中生身份来处理，复杂概念会拆成简单步骤；涉及论文、专利、通信科研时，会区分“高中生学习理解”和“科研/专利写作”两个层级。' },
    @{ Kind='heading'; Text='2. 关于“怎么赚大钱”的真实判断' },
    @{ Kind='quote'; Text='用户：你觉得咋能赚大钱' },
    @{ Kind='normal'; Text='核心结论：高中阶段想赚大钱，最现实的路线不是马上找暴利项目，而是尽早积累可以复利的能力：技术、表达、商业理解和信用。' },
    @{ Kind='normal'; Text='不建议走的方向包括：炒币、彩票、刷单、灰产、平台漏洞、所谓稳赚项目、流量课和过早重资产生意。这些更可能让你被别人赚钱，而不是你赚钱。' },
    @{ Kind='normal'; Text='建议押注的能力组合：AI 工具能力 + 编程能力 + 某个行业知识。结合你关注通信、6G、低空通信、无线网络数字孪生等方向，未来可以考虑科研资料自动整理工具、专利检索和分析工具、论文图表生成工具、通信仿真可视化工具、科研人员 AI 助手等。' },
    @{ Kind='normal'; Text='关键商业思维：不要只做一次性接单，要做可以复用和规模化的小产品。一次投入，多次销售，才有放大空间。' },
    @{ Kind='heading'; Text='3. 推荐的长期路线' },
    @{ Kind='normal'; Text='结合“高中生 + 通信科研相关背景”，建议路线是：' },
    @{ Kind='bullet'; Text='高中阶段：AI + 编程 + 科研写作基础。' },
    @{ Kind='bullet'; Text='大学前后：做科研、专利、论文辅助工具。' },
    @{ Kind='bullet'; Text='长期：做面向科研人员、学生、工程师的产品或服务。' },
    @{ Kind='normal'; Text='阶段性目标：30 天内做出一个能帮别人解决问题的小工具或资料产品，例如高中物理错题整理模板、AI 论文阅读笔记模板、通信基础知识小网站、专利交底书生成模板、PDF 文献整理成表格的小脚本等。' },
    @{ Kind='heading'; Text='4. 清华北大目标优先级' },
    @{ Kind='quote'; Text='用户：可是我现在的目标是考上清华北大' },
    @{ Kind='normal'; Text='核心结论：如果目标是清华北大，现在“赚钱”应降级为副线，主线必须是高考、竞赛、强基或相关升学路径。高中阶段真正能让未来赚钱能力放大的最大杠杆，是先把清北这个入口拿到。' },
    @{ Kind='normal'; Text='真实判断：现在不应该把主要精力放在赚钱上，因为高中时间窗口极短，清北录取机会的机会成本极高。当前最值钱的是分数、学科能力、竞赛/科研履历和稳定心态。' },
    @{ Kind='heading'; Text='5. 清北路径需要先判断' },
    @{ Kind='normal'; Text='清北路线大致包括：' },
    @{ Kind='bullet'; Text='裸分高考。' },
    @{ Kind='bullet'; Text='强基计划。' },
    @{ Kind='bullet'; Text='竞赛保送、破格或相关降分路径。' },
    @{ Kind='bullet'; Text='专项计划或综合评价类路径，具体取决于省份和政策。' },
    @{ Kind='bullet'; Text='港校或海外作为备选，如果你愿意。' },
    @{ Kind='normal'; Text='关键问题不是“怎么努力”，而是：我离清北差多少？差在哪一科？靠什么路径补？' },
    @{ Kind='heading'; Text='6. 学习策略要点' },
    @{ Kind='bullet'; Text='分数优先级通常是：数学 > 物理/化学/理综方向 > 英语 > 语文。不是语文不重要，而是数学和理科更容易拉开顶尖差距。' },
    @{ Kind='bullet'; Text='不要平均用力。要找出拖后腿科目、能冲满分科目、最容易提分模块和最容易失误题型，然后集中爆破。' },
    @{ Kind='bullet'; Text='建立错题系统，不是简单抄错题，而是记录错误原因、下次第一步怎么做、同类题模型是什么。' },
    @{ Kind='bullet'; Text='真题和高质量模拟题优先。一套卷子真正吃透，价值大于刷三套不复盘。' },
    @{ Kind='heading'; Text='7. 赚钱副线如何处理' },
    @{ Kind='normal'; Text='如果保留赚钱副线，每周最多 2 小时，并且必须服务于主线学习。可以做高中物理错题模板、数学模型总结、英语作文素材库、高考复盘系统、学习计划自动生成表等。这些既能提高自己，也可能未来变现。' },
    @{ Kind='heading'; Text='8. 建议建立“清北目标作战文档”' },
    @{ Kind='normal'; Text='建议新建一个清北目标作战文档，先定规则再执行。示例结构如下：' },
    @{ Kind='code'; Text="# 清北目标作战规则`n`n## 总目标`n目标：考入清华/北大。`n时间节点：____年高考。`n目标路径：裸分 / 强基 / 竞赛 / 其他。`n`n## 当前基线`n省份：`n年级：`n选科：`n最近排名：`n当前总分：`n目标总分：`n差距：`n`n## 科目策略`n数学：`n物理：`n化学/生物/政治/历史/地理：`n语文：`n英语：`n`n## 每周规则`n- 每周复盘一次成绩和错题`n- 每天固定数学训练`n- 每天固定英语输入`n- 每周至少一次完整限时训练`n- 错题必须分类，不只抄题`n`n## 红线`n- 不为了副业牺牲主线学习`n- 不沉迷低质量资料`n- 不用计划代替执行`n- 不刷题不复盘" },
    @{ Kind='heading'; Text='9. 下一步需要补充的信息' },
    @{ Kind='normal'; Text='为了制定具体清北备考方案，需要填写：' },
    @{ Kind='code'; Text="年级：`n省份：`n选科：`n最近一次大考总分：`n年级排名/全市或全省排名：`n数学分数：`n语文分数：`n英语分数：`n物理/化学/其他分数：`n最强科目：`n最弱科目：`n是否学竞赛：`n目标专业：`n每天可自主学习时间：" },
    @{ Kind='heading'; Text='10. 总结' },
    @{ Kind='normal'; Text='当前主线：清华北大。赚钱不是取消，而是降级为服务学习的副线。最优策略是先用分数和能力拿到更高平台，再把 AI、编程、通信科研理解和产品化能力放大成未来的收入能力。' }
)

$styleMap = @{
    title = '<w:pStyle w:val="Title"/>'
    heading = '<w:pStyle w:val="Heading1"/>'
    quote = '<w:pStyle w:val="Quote"/>'
    bullet = '<w:pStyle w:val="ListBullet"/>'
    code = '<w:pStyle w:val="Code"/>'
}

function XmlEscape([string]$s) {
    return [System.Security.SecurityElement]::Escape($s)
}

function MakeParagraph($kind, $text) {
    $ppr = ''
    if ($styleMap.ContainsKey($kind)) { $ppr = '<w:pPr>' + $styleMap[$kind] + '</w:pPr>' }
    $runs = New-Object System.Collections.Generic.List[string]
    $lines = $text -split "`n", -1
    for ($i = 0; $i -lt $lines.Count; $i++) {
        if ($i -gt 0) { $runs.Add('<w:r><w:br/></w:r>') }
        $line = XmlEscape $lines[$i]
        $space = ''
        if ($kind -eq 'code' -or $lines[$i].StartsWith(' ') -or $lines[$i].EndsWith(' ')) { $space = ' xml:space="preserve"' }
        $runs.Add('<w:r><w:t' + $space + '>' + $line + '</w:t></w:r>')
    }
    return '<w:p>' + $ppr + ($runs -join '') + '</w:p>'
}

$body = ($paragraphs | ForEach-Object { MakeParagraph $_.Kind $_.Text }) -join ''

$documentXml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    $body
    <w:sectPr>
      <w:pgSz w:w="11906" w:h="16838"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/>
    </w:sectPr>
  </w:body>
</w:document>
"@

$stylesXml = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:rPr><w:rFonts w:ascii="Microsoft YaHei" w:eastAsia="Microsoft YaHei"/><w:sz w:val="22"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:basedOn w:val="Normal"/><w:rPr><w:b/><w:rFonts w:ascii="Microsoft YaHei" w:eastAsia="Microsoft YaHei"/><w:sz w:val="36"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:rPr><w:b/><w:rFonts w:ascii="Microsoft YaHei" w:eastAsia="Microsoft YaHei"/><w:sz w:val="28"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="Quote"><w:name w:val="Quote"/><w:basedOn w:val="Normal"/><w:pPr><w:ind w:left="420"/></w:pPr><w:rPr><w:i/><w:color w:val="666666"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="ListBullet"><w:name w:val="List Bullet"/><w:basedOn w:val="Normal"/><w:pPr><w:ind w:left="420" w:hanging="240"/></w:pPr></w:style>
  <w:style w:type="paragraph" w:styleId="Code"><w:name w:val="Code"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:before="120" w:after="120"/></w:pPr><w:rPr><w:rFonts w:ascii="Consolas" w:eastAsia="Microsoft YaHei"/><w:sz w:val="20"/></w:rPr></w:style>
</w:styles>
'@

$contentTypes = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>
'@

$rels = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>
'@

$docRels = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>
'@

[System.IO.File]::WriteAllText((Join-Path $tmp '[Content_Types].xml'), $contentTypes, [System.Text.UTF8Encoding]::new($false))
[System.IO.File]::WriteAllText((Join-Path $tmp '_rels\.rels'), $rels, [System.Text.UTF8Encoding]::new($false))
[System.IO.File]::WriteAllText((Join-Path $tmp 'word\document.xml'), $documentXml, [System.Text.UTF8Encoding]::new($false))
[System.IO.File]::WriteAllText((Join-Path $tmp 'word\styles.xml'), $stylesXml, [System.Text.UTF8Encoding]::new($false))
[System.IO.File]::WriteAllText((Join-Path $tmp 'word\_rels\document.xml.rels'), $docRels, [System.Text.UTF8Encoding]::new($false))

if (Test-Path $out) { Remove-Item $out -Force }
Add-Type -AssemblyName System.IO.Compression.FileSystem
[System.IO.Compression.ZipFile]::CreateFromDirectory($tmp, $out)
Remove-Item $tmp -Recurse -Force
Write-Output $out
