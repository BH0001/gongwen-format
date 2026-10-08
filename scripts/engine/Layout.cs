using System.Xml.Linq;
using System.Text.RegularExpressions;
using System.Text;
using static Program;

internal static class LayoutEngine
{
    static readonly XNamespace WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
    static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";

    internal static void Apply(Package doc, Profile cfg, List<string> warnings)
    {
        var settings = doc.GetXml("word/settings.xml") ?? new XDocument(new XElement(W + "settings", new XAttribute(XNamespace.Xmlns + "w", W)));
        bool hadEven = settings.Root!.Element(W + "evenAndOddHeaders") is XElement evenSetting && Val(evenSetting) is not ("0" or "false" or "off");
        var sdkSettings = new DocumentFormat.OpenXml.Wordprocessing.Settings(settings.Root.ToString(SaveOptions.DisableFormatting));
        sdkSettings.RemoveAllChildren<DocumentFormat.OpenXml.Wordprocessing.EvenAndOddHeaders>();
        sdkSettings.RemoveAllChildren<DocumentFormat.OpenXml.Wordprocessing.UpdateFieldsOnOpen>();
        sdkSettings.AddChild(new DocumentFormat.OpenXml.Wordprocessing.EvenAndOddHeaders { Val = true }, true);
        sdkSettings.AddChild(new DocumentFormat.OpenXml.Wordprocessing.UpdateFieldsOnOpen { Val = true }, true);
        settings = XDocument.Parse(sdkSettings.OuterXml, LoadOptions.PreserveWhitespace);
        doc.PutXml("word/settings.xml", settings);
        doc.EnsureMainRelationship("settings", "settings.xml");
        doc.EnsureContentType("/word/settings.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml");

        var sections = doc.Document.Descendants(W + "sectPr").Where(x => !IsIn(x, "sectPrChange")).ToList();
        if (sections.Count == 0) { var sec = E("sectPr"); doc.Document.Root!.Element(W + "body")!.Add(sec); sections.Add(sec); }
        var inheritedFooters = new Dictionary<string, string>();
        var inheritedHeaders = new Dictionary<string, string>();
        string? defaultHeader = null;
        for (int i = 0; i < sections.Count; i++)
        {
            var sec = sections[i];
            bool landscape = Val(sec.Element(W + "pgSz"), "orient") == "landscape";
            if (landscape) warnings.Add($"第{i + 1}节原有横向页面已保留，属于国标正文版式例外；请检查宽表。");
            Set(sec, new XElement(W + "pgSz", new XAttribute(W + "w", Mm(landscape ? cfg.PageHeightMm : cfg.PageWidthMm)), new XAttribute(W + "h", Mm(landscape ? cfg.PageWidthMm : cfg.PageHeightMm)), new XAttribute(W + "orient", landscape ? "landscape" : "portrait")));
            Set(sec, new XElement(W + "pgMar", new XAttribute(W + "top", Mm(cfg.MarginTopMm)), new XAttribute(W + "right", Mm(cfg.MarginRightMm)), new XAttribute(W + "bottom", Mm(cfg.MarginBottomMm)), new XAttribute(W + "left", Mm(cfg.MarginLeftMm)), new XAttribute(W + "header", Mm(cfg.HeaderMm)), new XAttribute(W + "footer", Mm(cfg.FooterMm)), new XAttribute(W + "gutter", 0)));
            var cols = sec.Element(W + "cols");
            if (cols != null && Val(cols, "num") is not ("" or "1")) warnings.Add($"第{i + 1}节保留原有多栏布局，需人工检查。");
            int linePitch = cfg.LineMode == "grid" ? (int)Math.Round(cfg.ContentHeight / (double)cfg.Rows) : (int)Math.Round(cfg.FixedLinePt * 20);
            // OOXML charSpace adjusts the active run's character advance. Using the
            // source Normal style's smaller font would spread 16pt body text apart.
            int charSpace = (int)Math.Round((cfg.ContentWidth / 20.0 / cfg.Columns - cfg.BodyPt) * 4096);
            Set(sec, new XElement(W + "docGrid", new XAttribute(W + "type", cfg.LineMode == "grid" ? "linesAndChars" : "lines"), new XAttribute(W + "linePitch", linePitch), new XAttribute(W + "charSpace", cfg.LineMode == "grid" ? charSpace : 0)));
            var pn = new XElement(W + "pgNumType", new XAttribute(W + "fmt", "decimal"));
            if (i == 0) pn.SetAttributeValue(W + "start", 1);
            Set(sec, pn);

            // Enabling alternating footers must not make an existing default header disappear on even pages.
            var hr = sec.Elements(W + "headerReference").FirstOrDefault(x => Val(x, "type") == "default");
            defaultHeader = (string?)hr?.Attribute(R + "id") ?? defaultHeader;
            if (!hadEven)
            {
                // Inactive even headers were not visible in the source. Do not
                // activate stale text when enabling alternating page numbers.
                sec.Elements(W + "headerReference").Where(x => Val(x, "type") == "even").Remove();
                if (defaultHeader != null) sec.Add(new XElement(W + "headerReference", new XAttribute(W + "type", "even"), new XAttribute(R + "id", defaultHeader)));
            }
            foreach (var h in sec.Elements(W + "headerReference")) inheritedHeaders[Val(h, "type")] = (string)h.Attribute(R + "id")!;
            foreach (string type in new[] { "default", "even", "first" })
            {
                if (inheritedHeaders.ContainsKey(type)) continue;
                // Word can materialize an empty Header-style paragraph with its
                // built-in bottom rule. Explicit empty headers avoid that artifact.
                string headerName = $"gongwen-header-{i + 1}-{type}.xml";
                var hp = new XElement(W + "pPr", new XElement(W + "pBdr", new XElement(W + "bottom", new XAttribute(W + "val", "nil"))), new XElement(W + "spacing", new XAttribute(W + "before", 0), new XAttribute(W + "after", 0)));
                var header = new XDocument(new XElement(W + "hdr", new XAttribute(XNamespace.Xmlns + "w", W), new XElement(W + "p", hp)));
                doc.PutXml("word/" + headerName, header);
                doc.EnsureContentType("/word/" + headerName, "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml");
                string id = doc.EnsureMainRelationship("header", headerName, $"rIdGongwenHeader{i + 1}{type}");
                sec.Add(new XElement(W + "headerReference", new XAttribute(W + "type", type), new XAttribute(R + "id", id)));
                inheritedHeaders[type] = id;
            }

            foreach (var f in sec.Elements(W + "footerReference")) inheritedFooters[Val(f, "type")] = doc.RelationshipTarget((string)f.Attribute(R + "id")!)!;
            var sources = new Dictionary<string, string>(inheritedFooters);
            sec.Elements(W + "footerReference").Remove();
            foreach (string type in new[] { "default", "even", "first" })
            {
                string? oldPath = sources.GetValueOrDefault(type == "even" && !hadEven ? "default" : type);
                var footer = oldPath == null ? null : doc.GetXml(oldPath);
                footer ??= new XDocument(new XElement(W + "ftr", new XAttribute(XNamespace.Xmlns + "w", W), new XAttribute(XNamespace.Xmlns + "r", R)));
                foreach (var p in footer.Root!.Elements(W + "p").ToList())
                {
                    if (IsPurePageParagraph(p)) p.Remove();
                }
                // Tables, content controls and lower-case/split field codes are
                // legitimate footer content. Preserve them without a second PAGE.
                bool mixedPageField = footer.Descendants(W + "p").Any(HasPageField);
                if (!mixedPageField) footer.Root.Add(PageParagraph(type == "even", cfg));
                else warnings.Add($"第{i + 1}节 {type} 页脚含页码与其他文字混排，原样保留；未额外添加页码，需人工核验。");
                string name = $"gongwen-footer-{i + 1}-{type}.xml";
                string path = "word/" + name;
                // Resolve targets against the original part before moving a footer.
                string? oldRel = oldPath == null ? null : Path.GetDirectoryName(oldPath)!.Replace('\\', '/') + "/_rels/" + Path.GetFileName(oldPath) + ".rels";
                string newRel = "word/_rels/" + name + ".rels";
                if (oldRel != null && doc.GetXml(oldRel) is XDocument footerRels)
                {
                    foreach (var rel in footerRels.Root!.Elements())
                    {
                        if ((string?)rel.Attribute("TargetMode") == "External") continue;
                        var target = (string?)rel.Attribute("Target");
                        if (target == null) continue;
                        var resolved = new Uri(new Uri("http://package/" + oldPath), target);
                        rel.SetAttributeValue("Target", new Uri("http://package/" + path).MakeRelativeUri(resolved).ToString());
                    }
                    doc.PutXml(newRel, footerRels);
                }
                doc.PutXml(path, footer);
                doc.EnsureContentType("/" + path, "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml");
                string id = doc.EnsureMainRelationship("footer", name, $"rIdGongwenFooter{i + 1}{type}");
                sec.Add(new XElement(W + "footerReference", new XAttribute(W + "type", type), new XAttribute(R + "id", id)));
            }
            Order(sec);
        }
        FitTables(doc, cfg, warnings);
        FitImages(doc, cfg, warnings);
        if (doc.Document.Descendants(W + "txbxContent").Any()) warnings.Add("文本框内容原样保留，未应用正文规则；需要逐页检查其位置。");
        if (doc.Document.Descendants(W + "altChunk").Any()) warnings.Add("含外部导入内容 altChunk，未展开；需要在 Word 中检查。");
    }

    static bool IsPageInstruction(string instruction) => Regex.IsMatch(instruction, @"^\s*PAGE(?:\s|$)", RegexOptions.IgnoreCase);

    // Original footer parts may remain in the ZIP after their references change.
    // Compare the content Word actually follows, including section inheritance.
    internal static bool HeaderFooterContentEqual(Package before, Package after)
    {
        try { return ActiveHeaderFooterContent(before).SequenceEqual(ActiveHeaderFooterContent(after)); }
        catch (Exception) { return false; } // Unresolvable content is never a pass.
    }

    static bool IsOn(XElement? setting) => setting != null && Val(setting) is not ("0" or "false" or "off");

    static List<string> ActiveHeaderFooterContent(Package doc)
    {
        bool even = IsOn(doc.GetXml("word/settings.xml")?.Root?.Element(W + "evenAndOddHeaders"));
        var sections = doc.Document.Descendants(W + "sectPr").Where(s => !IsIn(s, "sectPrChange")).ToList();
        if (sections.Count == 0) sections.Add(E("sectPr"));
        var inherited = new Dictionary<string, string>();
        var cached = new Dictionary<string, string>();
        var result = new List<string>();
        foreach (var section in sections)
        {
            foreach (string kind in new[] { "header", "footer" })
            {
                foreach (var reference in section.Elements(W + (kind + "Reference")))
                {
                    string id = (string?)reference.Attribute(R + "id") ?? throw new InvalidOperationException("缺少页眉页脚引用。");
                    inherited[kind + ":" + Val(reference, "type")] = doc.RelationshipTarget(id) ?? throw new InvalidOperationException("页眉页脚关系不可解析。");
                }
                foreach (string page in new[] { "default", "even", "first" })
                {
                    string active = page == "even" && !even || page == "first" && !IsOn(section.Element(W + "titlePg")) ? "default" : page;
                    string? part = inherited.GetValueOrDefault(kind + ":" + active);
                    if (part == null) { result.Add(""); continue; }
                    if (!cached.TryGetValue(part, out var signature))
                    {
                        signature = HeaderFooterSignature(doc, part);
                        cached[part] = signature;
                    }
                    result.Add(signature);
                }
            }
        }
        return result;
    }

    static string HeaderFooterSignature(Package doc, string part)
    {
        var root = doc.GetXml(part)?.Root ?? throw new InvalidOperationException("页眉页脚部件不可读取。");
        // Only top-level pure page-number paragraphs are replaceable. Tables,
        // mixed fields, dates, notes and all other containers stay comparable.
        root.Elements(W + "p").Where(IsPurePageParagraph).Remove();
        string relPath = Path.GetDirectoryName(part)!.Replace('\\', '/') + "/_rels/" + Path.GetFileName(part) + ".rels";
        var relationships = doc.GetXml(relPath)?.Root?.Elements().ToDictionary(x => (string)x.Attribute("Id")!, x => x) ?? new();

        XElement? Canonical(XElement node)
        {
            if (node.Name == W + "pPr" || node.Name == W + "rPr") return null;
            if (node.Name == W + "t" && string.IsNullOrWhiteSpace(node.Value)) return null;
            var copy = new XElement(node.Name);
            foreach (var attribute in node.Attributes().Where(a => !a.IsNamespaceDeclaration).OrderBy(a => a.Name.ToString(), StringComparer.Ordinal))
            {
                string value = attribute.Value;
                if (attribute.Name.Namespace == R)
                {
                    if (!relationships.TryGetValue(value, out var relationship)) throw new InvalidOperationException("页眉页脚对象关系不可解析。");
                    string target = (string?)relationship.Attribute("Target") ?? throw new InvalidOperationException("页眉页脚对象缺少目标。");
                    bool external = (string?)relationship.Attribute("TargetMode") == "External";
                    if (!external) target = new Uri(new Uri("http://package/" + part), target).AbsoluteUri;
                    value = ((string?)relationship.Attribute("Type") ?? "") + "|" + (external ? "external|" : "internal|") + target;
                }
                copy.SetAttributeValue(attribute.Name, value);
            }
            foreach (var child in node.Nodes())
            {
                if (child is XElement element)
                {
                    var canonical = Canonical(element);
                    if (canonical != null) copy.Add(canonical);
                }
                else if (child is XText text && (!string.IsNullOrWhiteSpace(text.Value) || node.Name == W + "t" || node.Name == W + "instrText" || node.Name == W + "delText")) copy.Add(new XText(text.Value));
            }
            if ((node.Name == W + "p" || node.Name == W + "r") && !copy.HasElements && !copy.Nodes().OfType<XText>().Any()) return null;
            return copy;
        }
        return string.Concat(root.Elements().Select(Canonical).Where(x => x != null).Select(x => x!.ToString(SaveOptions.DisableFormatting)));
    }

    static IEnumerable<string> FieldInstructions(XElement p)
    {
        foreach (var field in p.Descendants(W + "fldSimple")) yield return Val(field, "instr");
        var stack = new Stack<StringBuilder>();
        foreach (var node in p.Descendants())
        {
            if (node.Ancestors(W + "fldSimple").Any()) continue;
            if (node.Name == W + "fldChar" && Val(node, "fldCharType") == "begin") stack.Push(new StringBuilder());
            else if (node.Name == W + "instrText" && stack.Count > 0) stack.Peek().Append(node.Value);
            else if (node.Name == W + "fldChar" && Val(node, "fldCharType") == "end" && stack.Count > 0) yield return stack.Pop().ToString();
        }
        // An incomplete field may span paragraphs. Keep it conservatively.
        foreach (var field in stack) yield return field.ToString();
    }

    static bool HasPageField(XElement p) => FieldInstructions(p).Any(IsPageInstruction)
        || p.Descendants(W + "instrText").Any(x => IsPageInstruction(x.Value));

    static bool IsPurePageParagraph(XElement p)
    {
        var instructions = FieldInstructions(p).ToArray();
        if (instructions.Length != 1 || !IsPageInstruction(instructions[0])) return false;
        if (p.Elements().Any(x => x.Name != W + "pPr" && x.Name != W + "r" && x.Name != W + "fldSimple")) return false;
        if (OwnRuns(p).Any(r => r.Elements().Any(x => x.Name != W + "rPr" && x.Name != W + "t" && x.Name != W + "instrText" && x.Name != W + "fldChar" && x.Name != W + "tab"))) return false;
        if (p.Descendants().Any(x => x.Name == W + "drawing" || x.Name == W + "pict" || x.Name == W + "object" || x.Name == W + "sdt" || x.Name == W + "ins" || x.Name == W + "del" || x.Name == W + "hyperlink" || x.Name.LocalName == "oMath" || x.Name.LocalName.StartsWith("comment") || x.Name.LocalName.StartsWith("bookmark"))) return false;
        var outside = new StringBuilder();
        int depth = 0;
        foreach (var node in p.Descendants())
        {
            if (node.Ancestors(W + "fldSimple").Any()) continue;
            if (node.Name == W + "fldChar")
            {
                if (Val(node, "fldCharType") == "begin") depth++;
                else if (Val(node, "fldCharType") == "end" && --depth < 0) return false;
            }
            else if (node.Name == W + "t" && depth == 0) outside.Append(node.Value);
        }
        // Digits outside the field can be dates/identifiers and must survive.
        return depth == 0 && Regex.IsMatch(outside.ToString(), @"^[\s—–\-/]*$");
    }

    static (int Width, int Height) SectionFrame(XElement element, Package doc, Profile cfg)
    {
        var section = doc.Document.Descendants(W + "sectPr").Where(x => !IsIn(x, "sectPrChange")).FirstOrDefault(s =>
        {
            var boundary = s.Parent?.Name == W + "pPr" ? s.Parent.Parent! : s;
            return element.AncestorsAndSelf().Contains(boundary) || XNode.DocumentOrderComparer.Compare(element, boundary) < 0;
        });
        bool landscape = Val(section?.Element(W + "pgSz"), "orient") == "landscape";
        return (Mm((landscape ? cfg.PageHeightMm : cfg.PageWidthMm) - cfg.MarginLeftMm - cfg.MarginRightMm),
            Mm((landscape ? cfg.PageWidthMm : cfg.PageHeightMm) - cfg.MarginTopMm - cfg.MarginBottomMm));
    }

    static XElement PageParagraph(bool even, Profile cfg)
    {
        var pp = new XElement(W + "pPr", E("pStyle", "GW_PageNumber"), E("snapToGrid", 0), new XElement(W + "spacing", new XAttribute(W + "before", 0), new XAttribute(W + "after", 0), new XAttribute(W + "line", 280), new XAttribute(W + "lineRule", "exact")), new XElement(W + "ind", new XAttribute(W + (even ? "left" : "right"), cfg.PageNumberPt * 20)), E("jc", even ? "left" : "right"));
        Order(pp);
        XElement Run(XElement child) => new(W + "r", FontProps("page", cfg), child);
        var p = new XElement(W + "p", pp,
            Run(new XElement(W + "t", "—")),
            Run(new XElement(W + "fldChar", new XAttribute(W + "fldCharType", "begin"), new XAttribute(W + "dirty", 1))),
            Run(new XElement(W + "instrText", new XAttribute(XNamespace.Xml + "space", "preserve"), " PAGE ")),
            Run(new XElement(W + "fldChar", new XAttribute(W + "fldCharType", "separate"))),
            Run(new XElement(W + "t", "1")),
            Run(new XElement(W + "fldChar", new XAttribute(W + "fldCharType", "end"))),
            Run(new XElement(W + "t", "—")));
        return p;
    }

    static void FitTables(Package doc, Profile cfg, List<string> warnings)
    {
        foreach (var table in doc.Document.Descendants(W + "tbl"))
        {
            if (IsIn(table, "tc")) { warnings.Add("嵌套表格保留原结构，需检查是否超出单元格。"); continue; }
            var cols = table.Element(W + "tblGrid")?.Elements(W + "gridCol").ToList() ?? new();
            int contentWidth = SectionFrame(table, doc, cfg).Width;
            double total = cols.Sum(c => double.TryParse(Val(c, "w"), out var n) ? n : 0);
            if (total <= contentWidth || total == 0) continue;
            double scale = contentWidth / total;
            int accumulated = 0;
            for (int i = 0; i < cols.Count; i++)
            {
                int width = i == cols.Count - 1 ? contentWidth - accumulated : (int)Math.Round(int.Parse(Val(cols[i], "w")) * scale);
                cols[i].SetAttributeValue(W + "w", width); accumulated += width;
            }
            var pr = table.Element(W + "tblPr");
            var tw = pr?.Element(W + "tblW");
            if (tw != null && Val(tw, "type") == "dxa") tw.SetAttributeValue(W + "w", contentWidth);
            foreach (var tc in table.Descendants(W + "tc").Where(x => x.Ancestors(W + "tbl").First() == table))
            {
                var cw = tc.Element(W + "tcPr")?.Element(W + "tcW");
                if (cw != null && Val(cw, "type") == "dxa" && double.TryParse(Val(cw, "w"), out var size)) cw.SetAttributeValue(W + "w", (int)Math.Round(size * scale));
            }
            foreach (var row in table.Elements(W + "tr"))
            {
                var height = row.Element(W + "trPr")?.Element(W + "trHeight");
                if (height != null && Val(height, "hRule") == "exact") height.SetAttributeValue(W + "hRule", "atLeast");
            }
            warnings.Add("宽表已按列宽比例缩至版心；单元格内容和合并关系保留。");
        }
    }

    static void FitImages(Package doc, Profile cfg, List<string> warnings)
    {
        foreach (var drawing in doc.Document.Descendants(W + "drawing"))
        {
            if (drawing.Ancestors(W + "p").Any(p => (string?)p.Attribute("data-gw-preserve") == "1")) continue;
            var frame = SectionFrame(drawing, doc, cfg);
            long maxW = (long)frame.Width * 635, maxH = (long)frame.Height * 635;
            var shape = drawing.Element(WP + "inline") ?? drawing.Element(WP + "anchor");
            var extent = shape?.Element(WP + "extent");
            if (extent == null || !long.TryParse((string?)extent.Attribute("cx"), out var cx) || !long.TryParse((string?)extent.Attribute("cy"), out var cy) || cx <= 0 || cy <= 0) continue;
            double scale = Math.Min(1, Math.Min(maxW / (double)cx, maxH / (double)cy));
            if (scale < 1)
            {
                extent.SetAttributeValue("cx", (long)Math.Round(cx * scale)); extent.SetAttributeValue("cy", (long)Math.Round(cy * scale));
                foreach (var ext in drawing.Descendants(A + "xfrm").Elements(A + "ext"))
                {
                    if (long.TryParse((string?)ext.Attribute("cx"), out long x)) ext.SetAttributeValue("cx", (long)Math.Round(x * scale));
                    if (long.TryParse((string?)ext.Attribute("cy"), out long y)) ext.SetAttributeValue("cy", (long)Math.Round(y * scale));
                }
                warnings.Add("超出版心的图片已等比例缩小，原始图片文件保留。");
            }
            if (shape!.Name == WP + "anchor") warnings.Add("浮动图片保留锚点，需检查换页后的环绕位置。");
        }
    }
}
