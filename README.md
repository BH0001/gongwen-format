# gongwen-format · 公文格式排版

将 Markdown、TXT、DOCX 排版为可编辑的中文公文 Word。适合作为 Agent Skill 使用，也可以直接运行命令行工具。

默认按 GB/T 9704—2012 的正文版式规则设置页面、字体、标题层级、段内题头、署名日期、附件和页码。保留原文、编号以及文档中的图片、表格、公式、批注和修订，并生成内容与格式检查报告。

## 功能

- A4 页面，上 / 下 / 左 / 右页边距为 37 / 35 / 28 / 26 mm。
- 标题二号小标宋，正文三号仿宋，一级标题黑体，二级标题楷体。
- 默认固定值 28.9 磅行距，普通正文勾选“如果定义了文档网格，则对齐到网格”。28.9 磅是本项目的默认参数，不是国标规定的唯一行距。
- 支持段内题头、长标题语义断行、奇偶页外侧页码和两套可调整配置。
- 结构映射绑定源文件 SHA-256；默认保留编号，可显式选择规范编号。
- 不覆盖源文件或已有输出，检查内容保留、Open XML 结构和排版参数。

本工具处理正文排版，不生成红头、印章或版记，也不代表对完整公文的国标合规认证。复杂表格、浮动对象及分页仍需渲染检查。

## 安装为 Skill

下载 [Releases](https://github.com/BH0001/gongwen-format/releases) 中的 Skill ZIP，将 `gongwen-format` 文件夹放入所用 Agent 支持的技能目录；文件夹内的 `SKILL.md` 是技能入口。也可以克隆本仓库后放入相应目录。

调用示例：

```text
使用 $gongwen-format 将这份 Word 排版为公文格式，保留原文和编号。
```

## 运行环境

| 依赖 | 用途 |
| --- | --- |
| Python 3.10+ | 命令行入口，仅依赖标准库 |
| .NET SDK 8+ | 构建和运行 Open XML 引擎 |
| Pandoc | 仅 Markdown / TXT 输入需要 |
| 公文字体 | 小标宋、仿宋、黑体、楷体、宋体以及 Times New Roman |
| Microsoft Word 或其他 DOCX 渲染器 | 版面核验 |

首次运行需要从 NuGet 获取 `DocumentFormat.OpenXml 3.5.1`，或使用已有本地缓存。构建产物写入工作目录的 `.gongwen-cache`，也可通过 `GONGWEN_CACHE_DIR` 指定缓存目录。

Windows 从字体注册表检查字体；macOS / Linux 需要 `fc-list`。字体不随仓库分发，请自行安装具有使用授权的字体。字体缺失时程序会停止；确认替代方案后可用 `--font title=字体名` 等参数显式指定。更换字体会影响版式，应重新检查。

## 命令行使用

以下示例在仓库根目录执行；输入文件和每次输出路径需按实际情况调整。

```bash
# Markdown / TXT 先转换；已有 DOCX 可跳过此步
python scripts/gongwen.py prepare examples/notice.md --output work/prepared.docx

# 读取全文、原始格式及角色建议
python scripts/gongwen.py inspect work/prepared.docx --output work/inspect.json

# 核对全文后，将角色映射保存为 work/map.json，再执行排版
python scripts/gongwen.py format work/prepared.docx --output work/formatted.docx --profile gbt9704 --map work/map.json

# 检查结构、格式及内容一致性
python scripts/gongwen.py check work/formatted.docx --against work/prepared.docx --output work/check.json
```

`inspect.json` 包含源文件摘要和逐段角色建议，可复制为 `map.json` 后按全文语义修改 `role`、`spans` 和 `lineBreaks`。不要直接把自动建议当作人工确认的结构。映射必须绑定同一份源文件，语法见 [结构映射说明](references/structure-map.md)。

两套配置：

- `gbt9704`：默认正文公文格式。
- `sample-notice`：示例通知样式，主送机关使用楷体，署名日期使用示例定位方式。

详细参数及实现边界见 [格式规则](references/format-rules.md)。示例内容位于 [examples/notice.md](examples/notice.md)，为公开的虚构材料。

## 版面检查

Windows 安装了 Word 时，可只读导出 PDF：

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File scripts/render_word.ps1 -InputFile work/formatted.docx -OutputDirectory work/render -TimeoutSeconds 120
```

脚本只更新页码相关字段，不保存源 DOCX。Word 自动化不可用或导出超时时，使用 LibreOffice 等渲染器核验，并记录实际使用的渲染器；备用渲染结果不能等同于 Word 验证。结构检查通过也不能代替逐页检查。

## 开发与验证

入口回归测试只需要 Python 标准库：

```bash
python scripts/test_workflow.py
```

完整回归测试使用虚构文档，需安装开发依赖、.NET SDK、Pandoc 和配置中的字体。每次使用新的输出目录：

```bash
python -m pip install -r requirements-dev.txt
python scripts/test_gongwen.py --work .test-output/content
python scripts/test_layout.py --work .test-output/layout
python scripts/test_content_regressions.py --work .test-output/regressions
```

完整测试覆盖图片、公式、表格、超链接、批注、修订、脚注尾注、编号、页眉页脚、Unicode 范围映射及失败时不覆盖文件等行为。自动测试不含视觉验收。

生成可安装包：

```bash
python scripts/pack_skill.py --output dist/gongwen-format.zip
```

发布包只包含技能源码、配置、公开示例和说明，不包含原始参照 Word、业务文档、字体、凭据及构建产物。工作文档请放在仓库外或 `work/` 目录。

## 许可证

本仓库代码、说明和虚构示例采用 [MIT License](LICENSE)。第三方依赖和自行安装的字体遵循各自许可证。
