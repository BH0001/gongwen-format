using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;

internal static class Program
{
    internal static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    internal static readonly XNamespace R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    internal static readonly XNamespace Pkg = "http://schemas.openxmlformats.org/package/2006/relationships";
    internal static readonly XNamespace CT = "http://schemas.openxmlformats.org/package/2006/content-types";
    internal static readonly JsonSerializerOptions Json = new() { PropertyNamingPolicy = JsonNamingPolicy.CamelCase, PropertyNameCaseInsensitive = true, WriteIndented = true, Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping };
    internal static readonly HashSet<string> Roles = new("title body heading1 heading2 heading3 heading4 addressee signature date attachment attachment-label attachment-title note table blank preserve".Split(' '));
    static readonly Regex H1 = new(@"^[一二三四五六七八九十百零〇]+、");
    static readonly Regex H2 = new(@"^[（(][一二三四五六七八九十百零〇]+[）)]");
    static readonly Regex H3 = new(@"^\d+[.．、](?!\d)");
    static readonly Regex H4 = new(@"^[（(]\d+[）)]");
    static readonly Regex DatePattern = new(@"^\s*\d{4}年\s*\d{1,2}月\s*\d{1,2}日\s*$");

    public static int Main(string[] args)
    {
        try
        {
            if (args.Length == 0) throw new Exception("需要 inspect、format 或 check 子命令。");
            var opt = new Dictionary<string, string>();
            for (int i = 1; i < args.Length; i += 2)
            {
                if (!args[i].StartsWith("--") || i + 1 >= args.Length) throw new Exception("参数必须使用 --key value。");
                opt.Add(args[i][2..], args[i + 1]);
            }
            string input = Path.GetFullPath(opt["input"]);
            if (!input.EndsWith(".docx", StringComparison.OrdinalIgnoreCase)) throw new Exception("引擎仅接收 DOCX；MD/TXT 请使用 Python 入口转换。");
            var doc = new Package(input);
            var items = Inspect(doc);
            object result;
            switch (args[0])
            {
                case "inspect":
                    result = new { version = 1, sourceSha256 = HashFile(input), offsetUnit = "Unicode code points", paragraphs = items, sections = doc.Document.Descendants(W + "sectPr").Select(x => x.ToString(SaveOptions.DisableFormatting)).ToArray(), objects = Inventory(doc), schemaErrors = Validate(input) };
                    break;
                case "check":
                    var errors = Validate(input);
                    var formatErrors = FormatCheck(doc);
                    bool? preserved = opt.TryGetValue("against", out var before) ? ContentEqual(new Package(before), doc) : null;
                    bool managed = Paras(doc).Any(p => Val(p.Element(W + "pPr")?.Element(W + "pStyle")).StartsWith("GW_"));
                    result = new { schemaErrors = errors, formatErrors, formatStatus = managed ? "checked" : "not-applicable", formatNotice = managed ? null : "未发现受管理的公文正文段落；未验证正文公文格式（全保留或全表格文档也可能如此）。", contentPreserved = preserved, objects = Inventory(doc) };
                    Console.WriteLine(JsonSerializer.Serialize(result, Json));
                    return errors.Count != 0 || formatErrors.Count != 0 || preserved == false ? 2 : 0;
                case "format":
                    result = Format(doc, input, opt, items);
                    break;
                default: throw new Exception("未知子命令。");
            }
            Console.WriteLine(JsonSerializer.Serialize(result, Json));
            return 0;
        }
        catch (Exception e)
        {
            Console.Error.WriteLine(e.Message);
            return 2;
        }
    }

    internal static string HashFile(string path) => Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path))).ToLowerInvariant();
    internal static int Len(string text) => text.EnumerateRunes().Count();
    internal static string Slice(string text, int start, int length) => string.Concat(text.EnumerateRunes().Skip(start).Take(length).Select(x => x.ToString()));
    internal static XElement E(string name, object? value = null) => value == null ? new(W + name) : new(W + name, new XAttribute(W + "val", value));
    internal static XElement Ensure(XElement parent, string name) { var x = parent.Element(W + name); if (x == null) { x = E(name); parent.AddFirst(x); } return x; }
    internal static void Set(XElement parent, XElement value) { parent.Elements(value.Name).Remove(); parent.Add(value); }
    internal static string Val(XElement? x, string name = "val") => (string?)x?.Attribute(W + name) ?? "";
    internal static int Mm(double x) => (int)Math.Round(x * 1440 / 25.4);
    internal static bool IsIn(XElement e, string name) => e.Ancestors(W + name).Any();
    internal static IEnumerable<XElement> Paras(Package p) => p.Document.Root!.Element(W + "body")!.Descendants(W + "p").Where(x => !IsIn(x, "txbxContent"));
    internal static IEnumerable<XElement> OwnRuns(XElement p) => p.Descendants(W + "r").Where(r => r.Ancestors(W + "p").FirstOrDefault() == p);
    internal static string ElementText(XElement x) => x.Name == W + "t" ? x.Value : x.Name == W + "tab" ? "\t" : x.Name == W + "cr" || x.Name == W + "br" && Val(x, "type") is not ("page" or "column") ? "\n" : "";
    internal static string Text(XElement p) => string.Concat(p.Descendants().Where(x => x.Ancestors(W + "p").FirstOrDefault() == p).Select(ElementText));

    static List<ParagraphInfo> Inspect(Package doc)
    {
        var list = Paras(doc).ToList();
        var styles = doc.GetXml("word/styles.xml")?.Root?.Elements(W + "style").ToDictionary(s => Val(s, "styleId"), s => s) ?? new();
        var result = new List<ParagraphInfo>();
        int first = list.FindIndex(p => !IsIn(p, "tc") && !string.IsNullOrWhiteSpace(Text(p)));
        bool mdTitle = first >= 0 && Val(list[first].Element(W + "pPr")?.Element(W + "pStyle")) == "Heading1";
        for (int i = 0; i < list.Count; i++)
        {
            var p = list[i]; var text = Text(p); var t = text.Trim();
            string styleId = Val(p.Element(W + "pPr")?.Element(W + "pStyle"));
            styles.TryGetValue(styleId, out var style);
            string styleName = Val(style?.Element(W + "name"));
            string role = "body", reason = "普通正文";
            var spans = new List<SpanInfo>();
            bool table = IsIn(p, "tc");
            if (table) { role = "table"; reason = "表格内容，保留局部结构和字形"; }
            else if (t.Length == 0) { role = "blank"; reason = "保留原有空段或分页符"; }
            else if (styleId.StartsWith("GW_") && Roles.Contains(styleId[3..])) { role = styleId[3..]; reason = "已存在的公文角色"; }
            else if (styleId == "Title" || styleName.Equals("title", StringComparison.OrdinalIgnoreCase)) { role = "title"; reason = "标题样式"; }
            else if (i == first && t.Length <= 150 && !styleId.StartsWith("List") && p.Element(W + "pPr")?.Element(W + "numPr") == null && !H1.IsMatch(t) && !H2.IsMatch(t) && !H3.IsMatch(t) && !H4.IsMatch(t) && !t.EndsWith('。') && !t.EndsWith('：')) { role = "title"; reason = "首个短段落；需由技能核对标题"; }
            else if (DatePattern.IsMatch(t)) { role = "date"; reason = "独立成文日期"; }
            else if (Regex.IsMatch(t, @"^附件\s*[：:]")) { role = "attachment"; reason = "附件说明"; }
            else if (Regex.IsMatch(t, @"^附件\s*\d*\s*$")) { role = "attachment-label"; reason = "附件起始标识"; }
            else if (i > first && result.Count(x => x.Role != "blank") == 1 && t.Length < 150 && Regex.IsMatch(t, "[：:]$")) { role = "addressee"; reason = "标题后的主送机关"; }
            else
            {
                var hm = Regex.Match(styleName.Length > 0 ? styleName : styleId, @"^(?:heading\s*|标题\s*)([1-9])$", RegexOptions.IgnoreCase);
                if (hm.Success) { role = "heading" + Math.Clamp(int.Parse(hm.Groups[1].Value) - (mdTitle ? 1 : 0), 1, 4); reason = "原有标题层级"; }
                else if (H1.IsMatch(t) || H2.IsMatch(t) || H3.IsMatch(t) || H4.IsMatch(t))
                {
                    string h = H1.IsMatch(t) ? "heading1" : H2.IsMatch(t) ? "heading2" : H3.IsMatch(t) ? "heading3" : "heading4";
                    int stop = t.IndexOf('。');
                    if (stop > 0 && stop < 70 && stop < t.Length - 1)
                    {
                        int start = Len(text) - Len(text.TrimStart());
                        spans.Add(new() { Start = start, Length = Len(t[..(stop + 1)]), Role = h, ExpectedText = t[..(stop + 1)] });
                        reason = "结构序号开头的段内题头；需核对句界";
                    }
                    else if (t.Length <= 100) { role = h; reason = "结构序号"; }
                    else reason = "序号长段落，无法可靠区分题头；按正文保留并待核对";
                }
            }
            var runs = new List<object>();
            int offset = 0;
            foreach (var r in OwnRuns(p))
            {
                var rp = r.Element(W + "rPr");
                foreach (var child in r.Elements())
                {
                    string rt = ElementText(child);
                    if (child.Name == W + "t" && rt.Length > 0) runs.Add(new { start = offset, length = Len(rt), text = rt, eastAsia = Val(rp?.Element(W + "rFonts"), "eastAsia"), sizeHalfPoints = Val(rp?.Element(W + "sz")), characterStyle = Val(rp?.Element(W + "rStyle")) });
                    offset += Len(rt);
                }
            }
            if (role == "body" && spans.Count == 0)
            {
                int pos = 0;
                foreach (var r in OwnRuns(p))
                {
                    string cs = Val(r.Element(W + "rPr")?.Element(W + "rStyle"));
                    foreach (var child in r.Elements())
                    {
                        string rt = ElementText(child);
                        if (cs.StartsWith("GW_char_") && child.Name == W + "t" && rt.Length > 0)
                        {
                            string cr = cs[8..];
                            var prev = spans.LastOrDefault();
                            if (prev != null && prev.Role == cr && prev.Start + prev.Length == pos) { prev.Length += Len(rt); prev.ExpectedText += rt; }
                            else spans.Add(new() { Start = pos, Length = Len(rt), Role = cr, ExpectedText = rt });
                        }
                        pos += Len(rt);
                    }
                }
            }
            result.Add(new() { Id = $"p{i + 1:D6}", ExpectedText = text, Role = role, Reason = reason, Spans = spans, InTable = table, Style = styleId, Runs = runs });
        }
        for (int i = 0; i < result.Count; i++)
            if (result[i].Role == "date")
            {
                int j = i - 1; while (j >= 0 && result[j].Role == "blank") j--;
                if (j >= 0 && result[j].Role == "body" && result[j].ExpectedText.Trim().Length is > 1 and < 65 && !Regex.IsMatch(result[j].ExpectedText, "[。；：]"))
                { result[j].Role = "signature"; result[j].Reason = "日期前的短独立段落；需核对署名"; }
            }
        return result;
    }

    static object Format(Package doc, string input, Dictionary<string, string> opt, List<ParagraphInfo> items)
    {
        string output = Path.GetFullPath(opt["output"]), reportPath = Path.GetFullPath(opt["report"]);
        if (File.Exists(output) || File.Exists(reportPath) || input == output || output == reportPath) throw new Exception("输出或报告已存在，拒绝覆盖。");
        var cfg = JsonSerializer.Deserialize<Profile>(File.ReadAllText(opt["profile"]), Json)!;
        ValidateProfile(cfg);
        var warnings = new List<string>();
        if (doc.Files.Keys.Any(x => x.StartsWith("_xmlsignatures/"))) throw new Exception("输入含数字签名，排版会使签名失效；请提供未签名副本。");
        if (doc.Files.Keys.Any(x => x.EndsWith("vbaProject.bin"))) throw new Exception("首版不处理含宏的文档。");
        if (opt.TryGetValue("map", out var mapFile))
        {
            var map = JsonSerializer.Deserialize<StructureMap>(File.ReadAllText(mapFile), Json)!;
            if (map.Version != 1 || !map.SourceSha256.Equals(HashFile(input), StringComparison.OrdinalIgnoreCase)) throw new Exception("结构映射与输入文件不一致，请对当前文件重新 inspect。");
            var seen = new HashSet<string>();
            foreach (var change in map.Paragraphs)
            {
                if (!seen.Add(change.Id)) throw new Exception("结构映射包含重复段落 ID。");
                var p = items.FirstOrDefault(x => x.Id == change.Id) ?? throw new Exception("结构映射中有未知段落 ID。");
                if (p.ExpectedText != change.ExpectedText) throw new Exception($"{p.Id} 原文校验失败。");
                p.Role = change.Role; p.Spans = change.Spans; p.LineBreaks = change.LineBreaks; p.Reason = "显式结构映射";
            }
        }
        foreach (var p in items)
        {
            if (!Roles.Contains(p.Role)) throw new Exception($"未知角色：{p.Role}");
            if (p.LineBreaks.Count != 0 && (p.Role is not ("title" or "attachment-title") || p.LineBreaks.Distinct().Count() != p.LineBreaks.Count || p.LineBreaks.Any(x => x <= 0 || x >= Len(p.ExpectedText)))) throw new Exception($"{p.Id} 标题断行位置无效。");
            int end = 0;
            foreach (var s in p.Spans.OrderBy(s => s.Start))
            {
                if (!new[] { "body", "heading1", "heading2", "heading3", "heading4" }.Contains(s.Role) || s.Start < end || s.Length <= 0 || s.Start + s.Length > Len(p.ExpectedText)) throw new Exception($"{p.Id} 文字范围无效或重叠。");
                if (s.ExpectedText == null || Slice(p.ExpectedText, s.Start, s.Length) != s.ExpectedText) throw new Exception($"{p.Id} 段内题头原文校验失败。");
                end = s.Start + s.Length;
            }
            if (p.InTable && p.Role != "table" && p.Role != "preserve") throw new Exception($"{p.Id} 位于表格内，不能映射为普通正文角色。");
            if (p.Reason.Contains("需核对") || p.Reason.Contains("需由") || p.Reason.Contains("无法可靠")) warnings.Add(p.Id + ": " + p.Reason);
        }
        var beforeErrors = Validate(input);
        var expected = items.Select(x => x.ExpectedText).ToArray();
        var paragraphs = Paras(doc).ToList();
        ProtectPreservedStyles(doc, paragraphs.Where((p, i) => items[i].Role is "preserve" or "blank" or "table"));
        for (int i = 0; i < paragraphs.Count; i++) if (items[i].Role == "preserve") paragraphs[i].SetAttributeValue("data-gw-preserve", "1");
        var counters = new int[4];
        for (int i = 0; i < paragraphs.Count; i++)
        {
            var p = paragraphs[i]; var item = items[i];
            if (item.Role is "preserve" or "blank" or "table") { ApplyParagraph(p, item, cfg, items, i); continue; }
            PreserveInheritedNumbering(doc, p);
            if (item.LineBreaks.Count != 0 && item.Spans.Count != 0) throw new Exception($"{item.Id} 不能同时设置标题断行和段内范围；请拆为两次明确映射，避免坐标偏移。");
            if (item.LineBreaks.Count != 0) { InsertTitleBreaks(p, item.LineBreaks); expected[i] = Text(p); }
            if (opt.GetValueOrDefault("numbering", "preserve") == "official" && Regex.IsMatch(item.Role, @"^heading[1-4]$"))
            {
                if (item.Spans.Count != 0) throw new Exception("含段内映射的标题不能自动重编号。");
                Renumber(p, item, counters); expected[i] = Text(p);
            }
            ApplyParagraph(p, item, cfg, items, i);
            if (item.Role is not ("preserve" or "table" or "blank"))
            {
                ApplySpans(p, item.Spans);
                foreach (var run in OwnRuns(p))
                {
                    string? spanRole = (string?)run.Attribute("data-gw-role");
                    string role = spanRole ?? item.Role;
                    run.Attribute("data-gw-role")?.Remove();
                    var rp = run.Element(W + "rPr");
                    if (rp != null && Val(rp.Element(W + "rStyle")).StartsWith("GW_char_")) rp.Elements(W + "rStyle").Remove();
                    if (spanRole != null) Set(Ensure(run, "rPr"), E("rStyle", "GW_char_" + spanRole));
                    ApplyRun(run, role, cfg);
                }
            }
        }
        for (int i = 1; i < paragraphs.Count; i++)
            if (items[i].Role == "signature" && items[i - 1].Role is "body" or "attachment")
            {
                // Keep the final content line with an immediately following
                // signature. Deliberate blank/signature pages remain untouched.
                var pp = Ensure(paragraphs[i - 1], "pPr"); Set(pp, E("keepNext", 1)); Order(pp);
            }
        var styles = doc.GetXml("word/styles.xml") ?? new XDocument(new XElement(W + "styles", new XAttribute(XNamespace.Xmlns + "w", W)));
        AddStyles(styles, cfg);
        doc.PutXml("word/styles.xml", styles);
        doc.EnsureMainRelationship("styles", "styles.xml");
        doc.EnsureContentType("/word/styles.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml");
        Layout(doc, cfg, warnings);
        foreach (var p in paragraphs) p.Attribute("data-gw-preserve")?.Remove();
        doc.PutXml("word/document.xml", doc.Document);
        var afterText = Paras(doc).Select(Text).ToArray();
        if (!expected.SequenceEqual(afterText)) throw new Exception("内容保护检查失败：段落文字发生了非预期变化。");
        string temp = Path.Combine(Path.GetDirectoryName(output)!, ".gongwen-" + Guid.NewGuid().ToString("N") + ".docx");
        Directory.CreateDirectory(Path.GetDirectoryName(output)!);
        try
        {
            doc.Save(temp);
            var errors = Validate(temp);
            var oldKeys = beforeErrors.Select(x => x.Key).ToHashSet();
            var added = errors.Where(x => !oldKeys.Contains(x.Key)).ToList();
            if (added.Count > 0) throw new Exception("Open XML 校验未通过：" + JsonSerializer.Serialize(added.Take(8), Json));
            var produced = new Package(temp);
            var violations = FormatCheck(produced, cfg);
            if (violations.Count > 0) throw new Exception("格式检查未通过：" + string.Join("；", violations));
            if (!ObjectsEqual(new Package(input), produced)) throw new Exception("内容保护检查失败：非排版对象或关系发生变化。");
            if (opt.GetValueOrDefault("numbering", "preserve") == "preserve" && !ContentEqual(new Package(input), produced)) throw new Exception("内容保护检查失败：字段、修订或公式文字发生变化。");
            if (errors.Count > 0) warnings.Add($"输入已有 {errors.Count} 项兼容性/结构问题，未新增；详见 sourceSchemaErrors。不能称为完全通过结构验证。");
            var report = new
            {
                version = 1, profile = cfg.Name, inputSha256 = HashFile(input), outputSha256 = HashFile(temp),
                contentPreserved = opt.GetValueOrDefault("numbering", "preserve") == "preserve",
                numbering = opt.GetValueOrDefault("numbering", "preserve"), objectPreservation = true,
                fonts = cfg.ResolvedFonts, paragraphs = items.Select(x => new { x.Id, x.Role, x.Reason, x.Spans, x.LineBreaks }),
                layout = new { cfg.LineMode, cfg.FixedLinePt, cfg.AlignToGrid, cfg.Rows, cfg.Columns, cfg.BodyPt, cfg.TitlePt },
                warnings = warnings.Distinct().ToArray(), sourceSchemaErrors = beforeErrors, outputSchemaErrors = errors,
                newSchemaErrors = added, formatErrors = violations,
                visualReview = "pending — 必须渲染并逐页检查后才可声明版面验证通过"
            };
            // Create-new semantics also protect against an output appearing during conversion.
            using (var sourceStream = File.OpenRead(temp))
            using (var destination = new FileStream(output, FileMode.CreateNew, FileAccess.Write)) sourceStream.CopyTo(destination);
            try { using var writer = new StreamWriter(new FileStream(reportPath, FileMode.CreateNew, FileAccess.Write), new UTF8Encoding(false)); writer.Write(JsonSerializer.Serialize(report, Json)); }
            catch { File.Delete(output); throw; }
            return new { output, report = reportPath, warnings = warnings.Count, schema = errors.Count == 0 ? "passed" : "source-issues-preserved", visualReview = "pending" };
        }
        finally { if (File.Exists(temp)) File.Delete(temp); }
    }

    static void Renumber(XElement p, ParagraphInfo item, int[] counters)
    {
        if (p.Elements().Any(x => x.Name != W + "pPr" && x.Name != W + "r") || OwnRuns(p).Any(r => r.Elements().Any(x => x.Name != W + "rPr" && x.Name != W + "t"))) throw new Exception($"{item.Id} 标题有复杂对象，不能安全重编号；请保留原编号。");
        int level = int.Parse(item.Role[^1..]) - 1;
        counters[level]++; for (int i = level + 1; i < 4; i++) counters[i] = 0;
        string chinese = Chinese(counters[level]);
        string prefix = level switch { 0 => chinese + "、", 1 => "（" + chinese + "）", 2 => counters[level] + ".", _ => "（" + counters[level] + "）" };
        string text = Regex.Replace(Text(p), @"^\s*(?:[一二三四五六七八九十百零〇]+、|[（(][一二三四五六七八九十百零〇\d]+[）)]|\d+[.．、])\s*", "");
        var runs = OwnRuns(p).ToList();
        var rp = runs.FirstOrDefault()?.Element(W + "rPr");
        foreach (var run in runs) run.Remove();
        p.Add(new XElement(W + "r", rp == null ? null : new XElement(rp), new XElement(W + "t", new XAttribute(XNamespace.Xml + "space", "preserve"), prefix + text)));
        p.Element(W + "pPr")?.Elements(W + "numPr").Remove();
    }
    static void ValidateProfile(Profile cfg)
    {
        var keys = new HashSet<string>(new[] { "title", "body", "heading1", "heading2", "addressee", "page", "latin" });
        if (cfg.ResolvedFonts == null || !keys.SetEquals(cfg.ResolvedFonts.Keys) || cfg.ResolvedFonts.Values.Any(string.IsNullOrWhiteSpace)) throw new Exception("配置必须解析 title/body/heading1/heading2/addressee/page/latin 七种字体。");
        if (cfg.LineMode is not ("fixed" or "grid") || cfg.SignatureMode is not ("unsealed" or "sample")) throw new Exception("lineMode 只接受 fixed/grid；signatureMode 只接受 unsealed/sample。");
        var positive = new[] { cfg.PageWidthMm, cfg.PageHeightMm, cfg.FixedLinePt, (double)cfg.BodyPt, cfg.TitlePt, cfg.PageNumberPt, cfg.Rows, cfg.Columns };
        var nonnegative = new[] { cfg.MarginTopMm, cfg.MarginBottomMm, cfg.MarginLeftMm, cfg.MarginRightMm, cfg.HeaderMm, cfg.FooterMm };
        if (positive.Any(x => !double.IsFinite(x) || x <= 0) || nonnegative.Any(x => !double.IsFinite(x) || x < 0) || cfg.PageWidthMm <= cfg.MarginLeftMm + cfg.MarginRightMm || cfg.PageHeightMm <= cfg.MarginTopMm + cfg.MarginBottomMm || cfg.ContentWidth <= 0 || cfg.ContentHeight <= 0) throw new Exception("配置尺寸、字号、行列数及行距必须有效，页边距不能占满页面。");
    }
    static XElement? EffectiveNumbering(Package doc, XElement p)
    {
        var pp = p.Element(W + "pPr");
        var styles = doc.GetXml("word/styles.xml")?.Root?.Elements(W + "style").ToDictionary(s => Val(s, "styleId"), s => s) ?? new();
        string id = Val(pp?.Element(W + "pStyle"));
        if (id == "") id = styles.Values.FirstOrDefault(s => Val(s, "type") == "paragraph" && Val(s, "default") is "1" or "true") is XElement def ? Val(def, "styleId") : "";
        var merged = E("numPr");
        void Merge(XElement? num) { if (num != null) foreach (var child in num.Elements()) if (merged.Element(child.Name) == null) merged.Add(new XElement(child)); }
        Merge(pp?.Element(W + "numPr"));
        var seen = new HashSet<string>();
        while (id != "" && seen.Add(id) && styles.TryGetValue(id, out var s))
        {
            Merge(s.Element(W + "pPr")?.Element(W + "numPr"));
            id = Val(s.Element(W + "basedOn"));
        }
        if (!merged.HasElements) return null;
        var children = merged.Elements().OrderBy(x => x.Name.LocalName switch { "ilvl" => 0, "numId" => 1, "numberingChange" => 2, _ => 3 }).ToArray();
        merged.ReplaceNodes(children); return merged;
    }
    static void PreserveInheritedNumbering(Package doc, XElement p)
    {
        if (EffectiveNumbering(doc, p) is XElement num) { var pp = Ensure(p, "pPr"); Set(pp, num); Order(pp); }
    }
    static void ProtectPreservedStyles(Package doc, IEnumerable<XElement> paragraphs)
    {
        var xml = doc.GetXml("word/styles.xml"); if (xml == null) return;
        var styles = xml.Root!.Elements(W + "style").ToDictionary(s => Val(s, "styleId"), s => s);
        var clones = new Dictionary<string, string>();
        bool NeedsCopy(string id, HashSet<string> seen)
        {
            if (!seen.Add(id) || !styles.TryGetValue(id, out var s)) return false;
            return id.StartsWith("GW_") || s.Elements().Where(x => x.Name == W + "basedOn" || x.Name == W + "link").Any(x => NeedsCopy(Val(x), seen));
        }
        string Copy(string id)
        {
            if (clones.TryGetValue(id, out var existing)) return existing;
            if (!styles.TryGetValue(id, out var s)) return id;
            string hash = Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(s.ToString(SaveOptions.DisableFormatting))))[..16];
            string target = "GWP_" + hash; clones[id] = target;
            if (styles.ContainsKey(target)) return target;
            var clone = new XElement(s); clone.SetAttributeValue(W + "styleId", target); clone.Attribute(W + "default")?.Remove();
            foreach (var dependency in clone.Elements().Where(x => x.Name == W + "basedOn" || x.Name == W + "link" || x.Name == W + "next")) dependency.SetAttributeValue(W + "val", Copy(Val(dependency)));
            xml.Root.Add(clone); return target;
        }
        foreach (var p in paragraphs)
            foreach (var reference in p.Descendants().Where(x => x.Name == W + "pStyle" || x.Name == W + "rStyle").ToArray())
                if (Val(reference) != "GW_LayoutBreak" && NeedsCopy(Val(reference), new())) reference.SetAttributeValue(W + "val", Copy(Val(reference)));
        doc.PutXml("word/styles.xml", xml);
    }
    static string Chinese(int n) { if (n > 99) throw new Exception("自动中文标题编号首版支持1—99。"); string d = "零一二三四五六七八九"; return n < 10 ? d[n].ToString() : (n < 20 ? "" : d[n / 10].ToString()) + "十" + (n % 10 == 0 ? "" : d[n % 10].ToString()); }

    static void ApplyParagraph(XElement p, ParagraphInfo item, Profile cfg, List<ParagraphInfo> all, int index)
    {
        if (item.Role is "preserve" or "blank") return;
        if (item.Role == "table")
        {
            var tp = Ensure(p, "pPr"); Set(tp, E("snapToGrid", 0)); Order(tp); return;
        }
        var pp = Ensure(p, "pPr");
        foreach (string name in new[] { "pStyle", "spacing", "ind", "jc", "outlineLvl", "snapToGrid", "contextualSpacing", "pBdr", "shd", "textAlignment", "adjustRightInd", "autoSpaceDE", "autoSpaceDN", "keepNext", "keepLines", "widowControl" }) pp.Elements(W + name).Remove();
        Set(pp, E("pStyle", "GW_" + item.Role));
        foreach (var node in ParagraphProps(item.Role, cfg).Elements()) Set(pp, new XElement(node));
        // Existing intentional blank paragraphs already provide the requested separation.
        if (item.Role is "title" or "attachment-title" && index + 1 < all.Count && all[index + 1].Role == "blank") pp.Element(W + "spacing")!.SetAttributeValue(W + "after", 0);
        if (item.Role is "signature" or "attachment" && index > 0 && all[index - 1].Role == "blank") pp.Element(W + "spacing")!.SetAttributeValue(W + "before", 0);
        bool hasList = pp.Element(W + "numPr") != null;
        if (hasList)
        {
            // List indentation is owned by numbering definitions; avoid doubled first-line offsets.
            pp.Elements(W + "ind").Remove();
        }
        if (item.Role is "signature" or "date" && cfg.SignatureMode == "unsealed")
        {
            int signature = index;
            if (item.Role == "date") { signature--; while (signature >= 0 && all[signature].Role == "blank") signature--; }
            string st = signature >= 0 && all[signature].Role == "signature" ? all[signature].ExpectedText.Trim() : "";
            double width = cfg.ContentWidth / 20.0 / cfg.BodyPt;
            double left = Math.Max(0, width - 2 - TextWidth(st.Length > 0 ? st : item.ExpectedText.Trim()));
            if (item.Role == "date")
            {
                double dateWidth = TextWidth(item.ExpectedText.Trim());
                left = st.Length > 0 && TextWidth(st) >= dateWidth ? left + 2 : Math.Max(0, width - 2 - dateWidth);
            }
            Set(pp, new XElement(W + "ind", new XAttribute(W + "left", (int)Math.Round(left * cfg.BodyPt * 20)), new XAttribute(W + "firstLine", 0)));
            Set(pp, E("jc", "left"));
            if (item.Role == "signature")
            {
                Set(pp, new XElement(W + "ind", new XAttribute(W + "right", cfg.BodyPt * 40), new XAttribute(W + "firstLine", 0)));
                Set(pp, E("jc", "right"));
            }
        }
        if (item.Role == "attachment")
        {
            var m = Regex.Match(item.ExpectedText, @"^附件\s*[：:]\s*(?:\d+[.．、]\s*)?");
            if (m.Success)
            {
                int first = 2 * cfg.BodyPt * 20, label = (int)Math.Round(TextWidth(m.Value) * cfg.BodyPt * 20);
                Set(pp, new XElement(W + "ind", new XAttribute(W + "left", first + label), new XAttribute(W + "hanging", label)));
            }
        }
        if (item.Role == "attachment-label") Set(pp, E("pageBreakBefore", 1));
        // A paragraph mark can otherwise retain the sample's old font/highlight.
        var markRevision = pp.Element(W + "rPr")?.Element(W + "rPrChange");
        var mark = FontProps(item.Role, cfg);
        if (markRevision != null) mark.Add(new XElement(markRevision));
        Set(pp, mark);
        Order(pp);
    }
    static double TextWidth(string t) => t.EnumerateRunes().Sum(r => r.Value is >= 65 and <= 90 ? 0.72 : r.Value < 128 ? 0.5 : 1.0);

    internal static XElement ParagraphProps(string role, Profile cfg)
    {
        bool title = role is "title" or "attachment-title";
        bool heading = role.StartsWith("heading");
        int pitch = cfg.LineMode == "grid" ? (int)Math.Round(cfg.ContentHeight / (double)cfg.Rows) : (int)Math.Round(cfg.FixedLinePt * 20);
        var pp = E("pPr");
        pp.Add(E("keepNext", title || heading || role is "signature" or "attachment-label" ? 1 : 0), E("keepLines", title || heading || role is "signature" or "date" ? 1 : 0), E("widowControl", 0));
        bool grid = cfg.LineMode == "grid" && !title && role is not ("signature" or "date");
        pp.Add(E("autoSpaceDE", 0), E("autoSpaceDN", 0), E("adjustRightInd", 0), E("snapToGrid", (grid || cfg.AlignToGrid) && !title && role is not ("signature" or "date") ? 1 : 0));
        pp.Add(new XElement(W + "spacing", new XAttribute(W + "before", role is "attachment" or "signature" ? pitch : 0), new XAttribute(W + "after", title ? pitch : role == "attachment-label" ? pitch : 0), new XAttribute(W + "line", title ? (int)(Math.Max(cfg.TitlePt + 6, cfg.FixedLinePt) * 20) : grid ? 240 : pitch), new XAttribute(W + "lineRule", grid ? "auto" : "exact")));
        bool indent = role is "body" or "note" || heading;
        pp.Add(new XElement(W + "ind", new XAttribute(W + "left", 0), new XAttribute(W + "right", 0), new XAttribute(W + "firstLine", indent ? cfg.BodyPt * 40 : 0), new XAttribute(W + "firstLineChars", indent ? 200 : 0)));
        string align = title ? "center" : role is "addressee" or "attachment-label" ? "left" : "both";
        if (role is "signature" or "date")
        {
            align = cfg.SignatureMode == "sample" ? "center" : "right";
            if (cfg.SignatureMode == "sample") Set(pp, new XElement(W + "ind", new XAttribute(W + "firstLine", cfg.BodyPt * 160), new XAttribute(W + "firstLineChars", 800)));
        }
        pp.Add(E("jc", align), E("outlineLvl", heading ? int.Parse(role[^1..]) - 1 : 9));
        Order(pp); return pp;
    }

    internal static XElement FontProps(string role, Profile cfg)
    {
        string fontRole = role switch { "title" or "attachment-title" => "title", "heading1" or "attachment-label" => "heading1", "heading2" => "heading2", "addressee" => "addressee", "page" => "page", _ => "body" };
        int pt = role is "title" or "attachment-title" ? cfg.TitlePt : role == "page" ? cfg.PageNumberPt : cfg.BodyPt;
        var rp = E("rPr");
        rp.Add(new XElement(W + "rFonts", new XAttribute(W + "ascii", cfg.ResolvedFonts["latin"]), new XAttribute(W + "hAnsi", cfg.ResolvedFonts["latin"]), new XAttribute(W + "cs", cfg.ResolvedFonts["latin"]), new XAttribute(W + "eastAsia", cfg.ResolvedFonts[fontRole]), new XAttribute(W + "hint", "eastAsia")));
        if (role == "page") foreach (var a in rp.Element(W + "rFonts")!.Attributes().Where(x => new[] { "ascii", "hAnsi", "cs" }.Contains(x.Name.LocalName))) a.Value = cfg.ResolvedFonts["page"];
        rp.Add(E("color", "000000"), E("sz", pt * 2), E("szCs", pt * 2));
        if ((cfg.LineMode == "grid" || cfg.Name == "gbt9704") && pt == cfg.BodyPt && role is not ("signature" or "date"))
        {
            // 28 full-width 16pt glyphs need 448pt, slightly wider than the
            // 156mm text area. Word does not contract glyphs via a negative
            // docGrid charSpace alone. A <=0.25pt tracking adjustment supplies
            // the required 28-cell advance without changing the font size.
            int tracking = (int)Math.Floor((cfg.ContentWidth / 20.0 / cfg.Columns - cfg.BodyPt) * 20);
            rp.Add(E("spacing", tracking));
        }
        Order(rp); return rp;
    }
    static void ApplyRun(XElement run, string role, Profile cfg)
    {
        // Drawing-only runs and field instructions have no normal text to restyle.
        if (!run.Elements().Any(x => x.Name == W + "t" || x.Name == W + "delText")) return;
        var rp = Ensure(run, "rPr");
        foreach (string name in new[] { "rFonts", "color", "sz", "szCs", "spacing", "w", "kern", "highlight", "shd", "bdr" }) rp.Elements(W + name).Remove();
        foreach (var node in FontProps(role, cfg).Elements()) Set(rp, new XElement(node));
        if (role != "body" && role != "note" && role != "attachment")
        {
            Set(rp, E("b", 0)); Set(rp, E("bCs", 0));
        }
        // Keep emphasis, underline, superscript, links, fields and revision containers.
        Set(rp, E("snapToGrid", (cfg.LineMode == "grid" || cfg.AlignToGrid) && role is not ("title" or "attachment-title" or "signature" or "date") ? 1 : 0));
        Order(rp);
    }

    static void ApplySpans(XElement p, List<SpanInfo> spans)
    {
        if (spans.Count == 0) return;
        SplitRunChildren(p);
        int position = 0, fieldDepth = 0;
        foreach (var run in OwnRuns(p).ToList())
        {
            fieldDepth += run.Elements(W + "fldChar").Count(x => Val(x, "fldCharType") == "begin");
            string rt = string.Concat(run.Elements(W + "t").Select(x => x.Value));
            int length = Len(rt), extra = run.Elements().Where(x => x.Name != W + "t").Sum(x => Len(ElementText(x)));
            bool control = run.Elements().Any(x => x.Name != W + "rPr" && x.Name != W + "t" && x.Name != W + "lastRenderedPageBreak");
            if (control && spans.Any(s => extra > 0 ? s.Start < position + extra && s.Start + s.Length > position : s.Start < position && s.Start + s.Length > position)) throw new Exception("段内题头范围不能跨越制表符、换行、字段或复杂对象。");
            var hits = spans.Where(s => s.Start < position + length && s.Start + s.Length > position).ToList();
            if (hits.Count > 0)
            {
                if (control || fieldDepth > 0 || run.Ancestors(W + "fldSimple").Any()) throw new Exception("段内题头范围触及字段或复杂对象，请缩小范围。");
                var cuts = new SortedSet<int> { 0, length };
                foreach (var h in hits) { cuts.Add(Math.Max(0, h.Start - position)); cuts.Add(Math.Min(length, h.Start + h.Length - position)); }
                var points = cuts.ToArray();
                for (int i = 0; i < points.Length - 1; i++)
                {
                    var clone = new XElement(run); clone.Elements(W + "t").Remove();
                    clone.Add(new XElement(W + "t", new XAttribute(XNamespace.Xml + "space", "preserve"), Slice(rt, points[i], points[i + 1] - points[i])));
                    var span = hits.FirstOrDefault(h => h.Start <= position + points[i] && h.Start + h.Length >= position + points[i + 1]);
                    if (span != null) clone.SetAttributeValue("data-gw-role", span.Role);
                    run.AddBeforeSelf(clone);
                }
                run.Remove();
            }
            position += length + extra;
            fieldDepth = Math.Max(0, fieldDepth - run.Elements(W + "fldChar").Count(x => Val(x, "fldCharType") == "end"));
        }
    }

    static void InsertTitleBreaks(XElement p, List<int> breaks)
    {
        SplitRunChildren(p);
        var remaining = breaks.ToHashSet();
        int position = 0, fieldDepth = 0;
        foreach (var run in OwnRuns(p).ToList())
        {
            fieldDepth += run.Elements(W + "fldChar").Count(x => Val(x, "fldCharType") == "begin");
            string rt = string.Concat(run.Elements(W + "t").Select(x => x.Value));
            int length = Len(rt);
            var hits = remaining.Where(x => length > 0 && x >= position && x <= position + length).Order().ToArray();
            if (hits.Length > 0)
            {
                if (run.Elements().Any(x => x.Name != W + "rPr" && x.Name != W + "t") || fieldDepth > 0 || run.Ancestors(W + "fldSimple").Any()) throw new Exception("标题断行触及复杂对象，请另选断行位置。");
                var cuts = new SortedSet<int>(hits.Select(x => x - position)) { 0, length };
                var points = cuts.ToArray();
                for (int i = 0; i < points.Length; i++)
                {
                    if (remaining.Remove(position + points[i]))
                    {
                        var br = new XElement(W + "r", new XElement(W + "rPr", E("rStyle", "GW_LayoutBreak")), new XElement(W + "br"));
                        run.AddBeforeSelf(br);
                    }
                    if (i + 1 < points.Length && points[i + 1] > points[i])
                    {
                        var clone = new XElement(run); clone.Elements(W + "t").Remove();
                        clone.Add(new XElement(W + "t", new XAttribute(XNamespace.Xml + "space", "preserve"), Slice(rt, points[i], points[i + 1] - points[i]))); run.AddBeforeSelf(clone);
                    }
                }
                run.Remove();
            }
            position += length + run.Elements().Where(x => x.Name != W + "t").Sum(x => Len(ElementText(x)));
            fieldDepth = Math.Max(0, fieldDepth - run.Elements(W + "fldChar").Count(x => Val(x, "fldCharType") == "end"));
        }
        if (remaining.Count != 0) throw new Exception("标题断行未能定位到普通文字边界。");
    }

    static void SplitRunChildren(XElement p)
    {
        foreach (var run in OwnRuns(p).ToList())
        {
            var children = run.Elements().Where(x => x.Name != W + "rPr").ToArray();
            if (children.Length < 2) continue;
            foreach (var child in children)
            {
                var clone = new XElement(run); clone.Elements().Where(x => x.Name != W + "rPr").Remove();
                clone.Add(new XElement(child)); run.AddBeforeSelf(clone);
            }
            run.Remove();
        }
    }

    static void AddStyles(XDocument styles, Profile cfg)
    {
        if (!styles.Root!.Elements(W + "style").Any(s => Val(s, "styleId") == "GW_LayoutBreak")) styles.Root.Add(new XElement(W + "style", new XAttribute(W + "type", "character"), new XAttribute(W + "customStyle", 1), new XAttribute(W + "styleId", "GW_LayoutBreak"), E("name", "公文排版标题换行")));
        foreach (var role in Roles.Where(x => x is not ("preserve" or "table" or "blank")))
        {
            string id = "GW_" + role;
            styles.Root!.Elements(W + "style").Where(s => Val(s, "styleId") == id).Remove();
            var s = new XElement(W + "style", new XAttribute(W + "type", "paragraph"), new XAttribute(W + "customStyle", 1), new XAttribute(W + "styleId", id), E("name", "公文-" + role), E("next", "GW_body"), E("qFormat"), ParagraphProps(role, cfg), FontProps(role, cfg));
            styles.Root.Add(s);
        }
        foreach (var role in new[] { "body", "heading1", "heading2", "heading3", "heading4" })
        {
            string id = "GW_char_" + role;
            styles.Root!.Elements(W + "style").Where(s => Val(s, "styleId") == id).Remove();
            styles.Root.Add(new XElement(W + "style", new XAttribute(W + "type", "character"), new XAttribute(W + "customStyle", 1), new XAttribute(W + "styleId", id), E("name", "公文题头-" + role), FontProps(role, cfg)));
        }
        styles.Root!.Elements(W + "style").Where(s => Val(s, "styleId") == "GW_PageNumber").Remove();
        styles.Root.Add(new XElement(W + "style", new XAttribute(W + "type", "paragraph"), new XAttribute(W + "customStyle", 1), new XAttribute(W + "styleId", "GW_PageNumber"), E("name", "公文页码"), FontProps("page", cfg)));
        // Do not change Normal or docDefaults, since tables/notes may inherit them.
    }

    static List<SchemaIssue> Validate(string file)
    {
        using var doc = WordprocessingDocument.Open(file, false);
        return new OpenXmlValidator(FileFormatVersions.Microsoft365).Validate(doc).Select(x => new SchemaIssue(x.Id, x.Part?.Uri.ToString() ?? "", x.Path?.XPath ?? "", x.Description)).ToList();
    }
    internal static List<string> FormatCheck(Package doc, Profile? profile = null)
    {
        var issues = new List<string>();
        var styleNodes = doc.GetXml("word/styles.xml")?.Root?.Elements(W + "style").ToDictionary(s => Val(s, "styleId"), s => s) ?? new();
        foreach (var p in Paras(doc))
        {
            string id = Val(p.Element(W + "pPr")?.Element(W + "pStyle"));
            if (id.StartsWith("GW_") && !styleNodes.ContainsKey(id)) issues.Add("缺少公文段落样式 " + id);
            if (id.StartsWith("GW_heading") && !int.TryParse(Val(p.Element(W + "pPr")?.Element(W + "outlineLvl")), out _)) issues.Add("标题缺少大纲层级");
            if (id.StartsWith("GW_") && styleNodes.TryGetValue(id, out var ps))
            {
                foreach (var run in OwnRuns(p).Where(r => r.Elements(W + "t").Any() || r.Elements(W + "delText").Any()))
                {
                    var rp = run.Element(W + "rPr");
                    string character = Val(rp?.Element(W + "rStyle"));
                    var expected = character.StartsWith("GW_char_") && styleNodes.TryGetValue(character, out var cs) ? cs.Element(W + "rPr") : ps.Element(W + "rPr");
                    foreach (string property in new[] { "sz", "szCs", "color" }) if (Val(rp?.Element(W + property)) != Val(expected?.Element(W + property))) issues.Add(id + " 文字属性不一致：" + property);
                    foreach (string attr in new[] { "eastAsia", "ascii", "hAnsi" }) if (Val(rp?.Element(W + "rFonts"), attr) != Val(expected?.Element(W + "rFonts"), attr)) issues.Add(id + " 字体不一致：" + attr);
                }
                foreach (string property in new[] { "line", "lineRule" })
                    if (Val(p.Element(W + "pPr")?.Element(W + "spacing"), property) != Val(ps.Element(W + "pPr")?.Element(W + "spacing"), property)) issues.Add(id + " 行距与配置不一致：" + property);
                if (Val(p.Element(W + "pPr")?.Element(W + "snapToGrid")) != Val(ps.Element(W + "pPr")?.Element(W + "snapToGrid"))) issues.Add(id + " 网格对齐与配置不一致");
            }
        }
        foreach (var sec in doc.Document.Descendants(W + "sectPr"))
        {
            var refs = sec.Elements(W + "footerReference").ToArray();
            if (refs.GroupBy(x => Val(x, "type")).Any(g => g.Count() != 1)) issues.Add("重复页脚引用");
            if (profile != null)
            {
                foreach (var pair in new Dictionary<string, double> { ["top"] = profile.MarginTopMm, ["bottom"] = profile.MarginBottomMm, ["left"] = profile.MarginLeftMm, ["right"] = profile.MarginRightMm })
                    if (Val(sec.Element(W + "pgMar"), pair.Key) != Mm(pair.Value).ToString()) issues.Add("页边距与配置不一致：" + pair.Key);
                int pitch = profile.LineMode == "grid" ? (int)Math.Round(profile.ContentHeight / (double)profile.Rows) : (int)Math.Round(profile.FixedLinePt * 20);
                if (Val(sec.Element(W + "docGrid"), "linePitch") != pitch.ToString()) issues.Add("行网格与配置不一致");
            }
        }
        return issues.Distinct().ToList();
    }
    static object Inventory(Package doc) => new
    {
        paragraphs = Paras(doc).Count(), tables = doc.Document.Descendants(W + "tbl").Count(), drawings = doc.Document.Descendants(W + "drawing").Count(),
        equations = doc.Document.Descendants().Count(x => x.Name.LocalName == "oMath"), hyperlinks = doc.Document.Descendants(W + "hyperlink").Count(),
        insertions = doc.Document.Descendants(W + "ins").Count(), deletions = doc.Document.Descendants(W + "del").Count(),
        comments = doc.GetXml("word/comments.xml")?.Root?.Elements(W + "comment").Count() ?? 0,
        footnotes = doc.Document.Descendants(W + "footnoteReference").Count(), endnotes = doc.Document.Descendants(W + "endnoteReference").Count(),
        media = doc.Files.Where(x => x.Key.StartsWith("word/media/")).ToDictionary(x => x.Key, x => Convert.ToHexString(SHA256.HashData(x.Value)))
    };
    // Compare semantic structure, allowing only the transformations this engine owns.
    // Ordinary run splitting/font changes and drawing dimensions are formatting;
    // control characters, fields, references, revisions, math and textbox text are content.
    static string SemanticContent(Package doc)
    {
        IEnumerable<XNode> Visit(XElement node)
        {
            if (node.Name == W + "r")
            {
                bool layoutBreak = Val(node.Element(W + "rPr")?.Element(W + "rStyle")) == "GW_LayoutBreak"
                    && node.Elements().Count(x => x.Name != W + "rPr") == 1 && node.Element(W + "br") is XElement br && !br.HasAttributes;
                if (layoutBreak) yield break;
                foreach (var child in node.Elements()) foreach (var result in Visit(child)) yield return result;
                yield break;
            }
            if (node.Name == W + "t") { yield return new XText(node.Value); yield break; }
            if (node.Name == W + "lastRenderedPageBreak") yield break;
            if (node.Name == W + "pPr" || node.Name == W + "rPr" || node.Name == W + "sectPr")
            {
                foreach (var change in node.Elements().Where(x => x.Name.Namespace == W && x.Name.LocalName.EndsWith("PrChange"))) yield return CanonicalElement(change);
                yield break;
            }
            var output = new XElement(node.Name);
            foreach (var attr in node.Attributes().Where(a => !a.IsNamespaceDeclaration).OrderBy(a => a.Name.ToString(), StringComparer.Ordinal))
            {
                if (attr.Name.LocalName.StartsWith("data-gw-")) continue;
                bool drawingSize = node.Name.NamespaceName is "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" or "http://schemas.openxmlformats.org/drawingml/2006/main" && node.Name.LocalName is "extent" or "ext" && attr.Name.LocalName is "cx" or "cy";
                bool tableSize = node.Name.Namespace == W && (node.Name.LocalName is "tblW" or "tcW" or "gridCol" && attr.Name == W + "w" || node.Name.LocalName == "trHeight" && attr.Name == W + "hRule");
                if (!drawingSize && !tableSize) output.Add(new XAttribute(attr));
            }
            if (node.Name == W + "p")
            {
                var num = EffectiveNumbering(doc, node);
                if (num != null && Val(num.Element(W + "numId")) is not ("" or "0"))
                {
                    if (num.Element(W + "ilvl") == null) num.AddFirst(E("ilvl", 0));
                    output.Add(CanonicalElement(num));
                }
            }
            foreach (var child in node.Nodes())
            {
                if (child is XElement e) foreach (var result in Visit(e)) output.Add(result);
                else if (child is XText text && !node.HasElements) output.Add(new XText(text.Value));
            }
            yield return output;
        }
        return new XElement("semantic", Visit(doc.Document.Root!.Element(W + "body")!)).ToString(SaveOptions.DisableFormatting);
    }
    static XElement CanonicalElement(XElement node) => new(node.Name,
        node.Attributes().Where(a => !a.IsNamespaceDeclaration).OrderBy(a => a.Name.ToString(), StringComparer.Ordinal).Select(a => new XAttribute(a)),
        node.Nodes().Where(n => n is XElement || n is XText && !node.HasElements).Select(n => n is XElement e ? (XNode)CanonicalElement(e) : new XText(((XText)n).Value)));
    static bool ContentEqual(Package before, Package after) => SemanticContent(before) == SemanticContent(after) && ObjectsEqual(before, after);
    static bool ObjectsEqual(Package before, Package after)
    {
        if (JsonSerializer.Serialize(Inventory(before), Json) != JsonSerializer.Serialize(Inventory(after), Json)) return false;
        // Untouched parts (comments, revisions in notes, embedded files, images, etc.) must remain byte-identical.
        foreach (var pair in before.Files)
        {
            if (pair.Key is "word/document.xml" or "word/styles.xml" or "word/settings.xml" or "word/_rels/document.xml.rels" or "[Content_Types].xml" || pair.Key.StartsWith("word/gongwen-footer-")) continue;
            if (!after.Files.TryGetValue(pair.Key, out var value) || !pair.Value.SequenceEqual(value)) return false;
        }
        var relBefore = (before.GetXml("word/_rels/document.xml.rels")?.Root?.Elements() ?? Enumerable.Empty<XElement>()).Where(x => !((string?)x.Attribute("Id") ?? "").StartsWith("rIdGongwen")).Select(x => x.ToString(SaveOptions.DisableFormatting)).ToArray();
        var relAfter = (after.GetXml("word/_rels/document.xml.rels")?.Root?.Elements() ?? Enumerable.Empty<XElement>()).Where(x => !((string?)x.Attribute("Id") ?? "").StartsWith("rIdGongwen")).Select(x => x.ToString(SaveOptions.DisableFormatting)).ToHashSet();
        return relBefore.All(relAfter.Contains) && LayoutEngine.HeaderFooterContentEqual(before, after);
    }

    internal static void Order(XElement p)
    {
        string order = p.Name.LocalName switch
        {
            "pPr" => "pStyle keepNext keepLines pageBreakBefore framePr widowControl numPr suppressLineNumbers pBdr shd tabs suppressAutoHyphens kinsoku wordWrap overflowPunct topLinePunct autoSpaceDE autoSpaceDN bidi adjustRightInd snapToGrid spacing ind contextualSpacing mirrorIndents suppressOverlap jc textDirection textAlignment textboxTightWrap outlineLvl divId cnfStyle rPr sectPr pPrChange",
            "rPr" => "rStyle rFonts b bCs i iCs caps smallCaps strike dstrike outline shadow emboss imprint noProof snapToGrid vanish webHidden color spacing w kern position sz szCs highlight u effect bdr shd fitText vertAlign rtl cs em lang eastAsianLayout specVanish oMath rPrChange",
            "sectPr" => "headerReference footerReference footnotePr endnotePr type pgSz pgMar paperSrc pgBorders lnNumType pgNumType cols formProt vAlign noEndnote titlePg textDirection bidi rtlGutter docGrid printerSettings sectPrChange",
            "settings" => "writeProtection view zoom removePersonalInformation removeDateAndTime doNotDisplayPageBoundaries displayBackgroundShape printPostScriptOverText printFractionalCharacterWidth printFormsData embedTrueTypeFonts embedSystemFonts saveSubsetFonts saveFormsData mirrorMargins alignBordersAndEdges bordersDoNotSurroundHeader bordersDoNotSurroundFooter gutterAtTop hideSpellingErrors hideGrammaticalErrors activeWritingStyle proofState formsDesign attachedTemplate linkStyles stylePaneFormatFilter stylePaneSortMethod documentType mailMerge revisionView trackRevisions doNotTrackMoves doNotTrackFormatting documentProtection autoFormatOverride styleLockTheme styleLockQFSet defaultTabStop autoHyphenation consecutiveHyphenLimit hyphenationZone doNotHyphenateCaps showEnvelope summaryLength clickAndTypeStyle defaultTableStyle evenAndOddHeaders bookFoldRevPrinting bookFoldPrinting bookFoldPrintingSheets drawingGridHorizontalSpacing drawingGridVerticalSpacing displayHorizontalDrawingGridEvery displayVerticalDrawingGridEvery doNotUseMarginsForDrawingGrid drawingGridHorizontalOrigin drawingGridVerticalOrigin doNotShadeFormData noPunctuationKerning characterSpacingControl printTwoOnOne strictFirstAndLastChars noLineBreaksAfter noLineBreaksBefore savePreviewPicture doNotValidateAgainstSchema saveInvalidXml doNotIncludeSubdocsInStats ignoreMixedContent alwaysShowPlaceholderText doNotDemarcateInvalidXml saveXmlDataOnly useXSLTWhenSaving saveThroughXslt showXMLTags alwaysMergeEmptyNamespace updateFields hdrShapeDefaults footnotePr endnotePr compat docVars rsids mathPr uiCompat97To2003 attachedSchema themeFontLang clrSchemeMapping doNotIncludeSubdocsInStats doNotAutoCompressPictures forceUpgrade captions readModeInkLockDown smartTagType schemaLibrary shapeDefaults doNotEmbedSmartTags decimalSymbol listSeparator",
            _ => ""
        };
        if (order == "") return;
        var keys = order.Split(' ');
        var nodes = p.Elements().Select((e, i) => (e, i, rank: Array.IndexOf(keys, e.Name.LocalName))).OrderBy(x => x.rank < 0 ? 1000 : x.rank).ThenBy(x => x.i).Select(x => x.e).ToArray();
        p.RemoveNodes(); p.Add(nodes);
    }

    static void Layout(Package doc, Profile cfg, List<string> warnings) => LayoutEngine.Apply(doc, cfg, warnings);
}

internal sealed class ParagraphInfo
{
    public string Id { get; set; } = "";
    public string ExpectedText { get; set; } = "";
    public string Role { get; set; } = "body";
    public string Reason { get; set; } = "";
    public string Style { get; set; } = "";
    public bool InTable { get; set; }
    public List<SpanInfo> Spans { get; set; } = new();
    public List<int> LineBreaks { get; set; } = new();
    public List<object> Runs { get; set; } = new();
}
internal sealed class SpanInfo
{
    public int Start { get; set; }
    public int Length { get; set; }
    public string Role { get; set; } = "heading2";
    public string? ExpectedText { get; set; }
}
internal sealed class StructureMap
{
    public int Version { get; set; }
    public string SourceSha256 { get; set; } = "";
    public List<ParagraphInfo> Paragraphs { get; set; } = new();
}
internal sealed class Profile
{
    public string Name { get; set; } = "";
    public double PageWidthMm { get; set; } = 210;
    public double PageHeightMm { get; set; } = 297;
    public double MarginTopMm { get; set; } = 37;
    public double MarginBottomMm { get; set; } = 35;
    public double MarginLeftMm { get; set; } = 28;
    public double MarginRightMm { get; set; } = 26;
    public double HeaderMm { get; set; } = 15;
    public double FooterMm { get; set; } = 26;
    public int Rows { get; set; } = 22;
    public int Columns { get; set; } = 28;
    public string LineMode { get; set; } = "fixed";
    public double FixedLinePt { get; set; } = 28.9;
    public bool AlignToGrid { get; set; } = true;
    public int BodyPt { get; set; } = 16;
    public int TitlePt { get; set; } = 22;
    public int PageNumberPt { get; set; } = 14;
    public string SignatureMode { get; set; } = "unsealed";
    public Dictionary<string, string> ResolvedFonts { get; set; } = new();
    public int ContentWidth => Program.Mm(PageWidthMm - MarginLeftMm - MarginRightMm);
    public int ContentHeight => Program.Mm(PageHeightMm - MarginTopMm - MarginBottomMm);
}
internal sealed record SchemaIssue(string Id, string Part, string Path, string Description)
{
    public string Key => Id + "|" + Part + "|" + Path + "|" + Description;
}
internal sealed class Package
{
    public Dictionary<string, byte[]> Files { get; } = new();
    public XDocument Document { get; }
    public Package(string path)
    {
        using var zip = ZipFile.OpenRead(path);
        foreach (var entry in zip.Entries)
        {
            if (entry.FullName.EndsWith('/')) continue;
            using var s = entry.Open(); using var ms = new MemoryStream(); s.CopyTo(ms); Files.Add(entry.FullName.Replace('\\', '/'), ms.ToArray());
        }
        Document = GetXml("word/document.xml") ?? throw new Exception("文件不是可读取的 DOCX。");
    }
    public XDocument? GetXml(string name) => Files.TryGetValue(name, out var bytes) ? XDocument.Load(new MemoryStream(bytes), LoadOptions.PreserveWhitespace) : null;
    public void PutXml(string name, XDocument xml)
    {
        using var ms = new MemoryStream(); using (var writer = System.Xml.XmlWriter.Create(ms, new() { Encoding = new UTF8Encoding(false), Indent = false })) xml.Save(writer); Files[name] = ms.ToArray();
    }
    public void EnsureContentType(string part, string type)
    {
        var ct = GetXml("[Content_Types].xml")!;
        var old = ct.Root!.Elements(Program.CT + "Override").FirstOrDefault(e => (string?)e.Attribute("PartName") == part);
        if (old == null) ct.Root.Add(new XElement(Program.CT + "Override", new XAttribute("PartName", part), new XAttribute("ContentType", type)));
        else old.SetAttributeValue("ContentType", type);
        PutXml("[Content_Types].xml", ct);
    }
    public string EnsureMainRelationship(string kind, string target, string? requestedId = null)
    {
        var rels = GetXml("word/_rels/document.xml.rels") ?? new XDocument(new XElement(Program.Pkg + "Relationships"));
        string type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/" + kind;
        var existing = requestedId == null ? rels.Root!.Elements().FirstOrDefault(e => (string?)e.Attribute("Type") == type && (string?)e.Attribute("Target") == target) : rels.Root!.Elements().FirstOrDefault(e => (string?)e.Attribute("Id") == requestedId);
        if (existing != null)
        {
            if ((string?)existing.Attribute("Type") != type) throw new Exception("关系 ID 冲突。");
            existing.SetAttributeValue("Target", target); PutXml("word/_rels/document.xml.rels", rels); return (string)existing.Attribute("Id")!;
        }
        string id = requestedId ?? "rIdGongwen" + kind;
        if (rels.Root!.Elements().Any(e => (string?)e.Attribute("Id") == id)) throw new Exception("关系 ID 冲突。");
        rels.Root.Add(new XElement(Program.Pkg + "Relationship", new XAttribute("Id", id), new XAttribute("Type", type), new XAttribute("Target", target)));
        PutXml("word/_rels/document.xml.rels", rels); return id;
    }
    public string? RelationshipTarget(string id)
    {
        string? target = (string?)GetXml("word/_rels/document.xml.rels")?.Root?.Elements().FirstOrDefault(x => (string?)x.Attribute("Id") == id)?.Attribute("Target");
        if (target == null) return null;
        return new Uri(new Uri("http://package/word/document.xml"), target).AbsolutePath.TrimStart('/');
    }
    public void Save(string path)
    {
        using var zip = ZipFile.Open(path, ZipArchiveMode.Create);
        foreach (var (name, bytes) in Files) { using var stream = zip.CreateEntry(name, CompressionLevel.Optimal).Open(); stream.Write(bytes); }
    }
}
