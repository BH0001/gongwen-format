# 结构映射与命令接口

## 命令

`python scripts/gongwen.py prepare INPUT --output PREPARED.docx`

`python scripts/gongwen.py inspect PREPARED.docx --output inspect.json`

`python scripts/gongwen.py format PREPARED.docx --output OUTPUT.docx --profile gbt9704 --map map.json`

`python scripts/gongwen.py check OUTPUT.docx --against PREPARED.docx --output check.json`

可选参数：`--profile sample-notice` 或自定义JSON路径；`--report 新的报告.json`；`--numbering preserve|official`；可重复的 `--font body=仿宋`。默认不覆盖任何已有输出、报告或准备文件。

## 映射内容

`inspect` 输出完整原文、逐段ID、建议角色、原样式及文字片段；它不是经过语义核对的最终结论。技能应核对后输出映射。可以只覆盖部分段落，未覆盖者使用检查阶段的规则建议。含歧义的重点段落应明确指定，尤其是署名、日期、附件和段内题头。

```json
{
  "version": 1,
  "sourceSha256": "从inspect原样复制的SHA-256",
  "paragraphs": [
    {
      "id": "p000006",
      "expectedText": "（一）总体要求。坚持依法依规。",
      "role": "body",
      "spans": [
        {"start": 0, "length": 8, "role": "heading2", "expectedText": "（一）总体要求。"}
      ]
    }
  ]
}
```

上例 `length` 必须按实际字符串重新计算（示例题头“（一）总体要求。”为8个Unicode码点）。**起点与长度按Unicode码点计数，不按UTF-16单元或UTF-8字节数。** Python可用 `len(text)` 和切片计算。每个范围必须同时给出完全匹配的 `expectedText`；范围不能相交，也不能跨字段、制表符、换行、图片等复杂对象。同一run中的普通文字、制表符和换行会先拆分定位；Word的 `lastRenderedPageBreak` 分页缓存不应阻止普通题头排版。原超链接、插入修订等外层容器保留。

角色：

| 角色 | 用途 |
|---|---|
| title / body | 主标题 / 正文 |
| heading1—heading4 | 独立的四级标题 |
| addressee | 主送机关 |
| signature / date | 原有署名 / 原有日期（包括未填完整的占位日期） |
| attachment | “附件：……”说明；续行悬挂对齐 |
| attachment-label | 附件本体起始标识，另面排版 |
| attachment-title | 附件标题，字号同主标题 |
| note | 附注或已有的“此页无正文”等文字 |
| table / blank | 表格段落 / 原有空段或分页符段落 |
| preserve | 保留原有有效段落格式与段内图片大小；必要时复制样式依赖，用于特殊结构 |

段内范围只接受 `body`、`heading1`—`heading4`。保留范围的自定义字符样式使重复排版仍保持题头字体；用户显式把 `spans` 设为空会清除这些题头标记。

主标题及附件标题可额外设置 `"lineBreaks": [12]`，表示在原文第12个Unicode码点之后添加排版换行；必须根据全文语义选择实际位置，不能照抄示例数字。它只插入带有 `GW_LayoutBreak` 样式标记的Word换行元素，不增删字词；内容校验仅忽略这种工具生成的排版换行，仍检查原有制表符、换行与分页。重复排版时重新 `inspect` 并绑定当前文件，不复用旧映射。需要修改断行点时从原始准备文件重新排版。同一段不能同时指定非空 `lineBreaks` 与 `spans`，避免坐标含义不明；断行点不能位于字段、图像或其他复杂对象内部。

## 检查报告与失败

格式化失败返回非零退出码，不发布半成品输出；引擎先在临时文件校验，再创建最终文件。映射不匹配、范围不合法、输出已存在、缺字体、数字签名或宏文档会明确报错。

排版报告包含字体、`layout` 行距与网格参数、应用角色、对象保留结果、原有和新增Open XML错误、格式错误及需要视觉复核的警告。`visualReview` 初始为pending，仅结构检查不能证明Word分页和字体显示正确。原稿已有的结构问题会记录；新增问题阻止交付。`check` 遇到原有结构问题也返回非零，但合法的检查JSON仍写入 `--output`，需结合报告判断；输入不可读取等执行错误不会伪造检查报告。

`check` 的 `formatStatus: not-applicable` 表示没有受管理的公文正文段落，例如未排版的原稿或全保留/全表格文档，不能据此宣称公文格式通过。`checked` 只表示检查了受管理段落与其样式的一致性；仍需按排版报告确认配置并完成视觉检查。`--against` 会核对原有编号、控制符、字段、修订及实际引用的页眉页脚内容，不仅比较可见文字或对象数量。

`--numbering official` 只作用于独立标题，按“一、”“（一）”“1.”“（1）”分别计数，下级随上级重置；中文序号支持1—99。不改段内题头；包含超链接、批注锚点、修订或非文字对象的标题拒绝自动编号。此模式的文字变化应通过标题范围人工核对，不使用全文一致作为通过条件。
