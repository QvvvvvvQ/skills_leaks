// Blueprint.cs - Technical Product Manual Generation Template
// Demonstrates OpenXML SDK for hardware/software documentation
// Structure: Specs → Installation → Troubleshooting → Appendix

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace DocumentFoundry;

/// <summary>
/// Product manual generator demonstrating technical documentation patterns.
/// Output: A structured manual with specifications, procedures, and reference tables.
/// </summary>
public static class Blueprint
{
    // ═══════════════════════════════════════════════════════════════════════
    // Theme Configuration - swap entire palette by changing this binding
    // ═══════════════════════════════════════════════════════════════════════
    private static class Theme
    {
        public const string HeadingColor = ColorSchemes.Verdant.Text.Primary;
        public const string SubheadColor = ColorSchemes.Verdant.Text.Secondary;
        public const string BodyColor = ColorSchemes.Verdant.Text.Body;
        public const string CaptionColor = ColorSchemes.Verdant.Text.Muted;
        public const string AccentColor = ColorSchemes.Verdant.Decorative.Muted.Primary;
        public const string HighlightColor = ColorSchemes.Verdant.Decorative.Muted.Accent;
        public const string GridColor = ColorSchemes.Verdant.UI.Border;
        public const string HeaderBg = ColorSchemes.Verdant.UI.TableHeader;
    }

    // ═══════════════════════════════════════════════════════════════════════
    // Document Dimensions (A4 metric)
    // ═══════════════════════════════════════════════════════════════════════
    private const int DocWidthTwips = 11906;   // 210mm in twips
    private const int DocHeightTwips = 16838;  // 297mm in twips
    private const long DocWidthEmu = 7560000;  // 210mm in EMU
    private const long DocHeightEmu = 10692000; // 297mm in EMU

    /// <summary>
    /// Main entry point - generates complete product manual.
    /// </summary>
    public static void Generate(string outputPath, string assetDir)
    {
        using var document = WordprocessingDocument.Create(outputPath, WordprocessingDocumentType.Document);
        var mainPart = document.AddMainDocumentPart();
        mainPart.Document = new Document(new Body());
        var body = mainPart.Document.Body!;

        // Initialize document infrastructure
        RegisterStyles(mainPart);
        RegisterNumbering(mainPart);

        // Load backdrop images
        var titleBgId = LoadImagePart(mainPart, Path.Combine(assetDir, "title_backdrop.png"));
        var contentBgId = LoadImagePart(mainPart, Path.Combine(assetDir, "content_watermark.png"));
        var endBgId = LoadImagePart(mainPart, Path.Combine(assetDir, "closing_flourish.png"));

        // ═══════ Title Page ═══════
        RenderTitlePage(body, titleBgId);

        // ═══════ Contents ═══════
        RenderContentsPage(body, contentBgId, mainPart);

        // ═══════ Main Sections ═══════
        RenderMainContent(document, body, contentBgId, mainPart, assetDir);

        // ═══════ Back Page ═══════
        RenderBackPage(body, endBgId);

        // Enable field auto-update
        EnableFieldRefresh(mainPart);

        document.Save();
        Console.WriteLine($"Manual generated: {outputPath}");
    }

    // ═══════════════════════════════════════════════════════════════════════
    // Style Registration
    // ═══════════════════════════════════════════════════════════════════════
    private static void RegisterStyles(MainDocumentPart mainPart)
    {
        var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
        stylesPart.Styles = new Styles();

        // Base paragraph style
        stylesPart.Styles.Append(new Style(
            new StyleName { Val = "Normal" },
            new StyleParagraphProperties(
                new SpacingBetweenLines { After = "200", Line = "312", LineRule = LineSpacingRuleValues.Auto }
            ),
            new StyleRunProperties(
                new RunFonts { Ascii = "Segoe UI", HighAnsi = "Segoe UI", EastAsia = "Source Han Sans SC" },
                new FontSize { Val = "22" },
                new Color { Val = Theme.BodyColor }
            )
        )
        { Type = StyleValues.Paragraph, StyleId = "Normal", Default = true });

        // Chapter heading
        stylesPart.Styles.Append(new Style(
            new StyleName { Val = "heading 1" },
            new BasedOn { Val = "Normal" },
            new StyleParagraphProperties(
                new KeepNext(),
                new KeepLines(),
                new PageBreakBefore(),
                new SpacingBetweenLines { Before = "480", After = "240", Line = "240", LineRule = LineSpacingRuleValues.Auto },
                new OutlineLevel { Val = 0 }
            ),
            new StyleRunProperties(
                new Bold(),
                new Color { Val = Theme.HeadingColor },
                new FontSize { Val = "36" }
            )
        )
        { Type = StyleValues.Paragraph, StyleId = "Heading1" });

        // Section heading
        stylesPart.Styles.Append(new Style(
            new StyleName { Val = "heading 2" },
            new BasedOn { Val = "Normal" },
            new StyleParagraphProperties(
                new KeepNext(),
                new KeepLines(),
                new SpacingBetweenLines { Before = "320", After = "160" },
                new OutlineLevel { Val = 1 }
            ),
            new StyleRunProperties(
                new Bold(),
                new Color { Val = Theme.SubheadColor },
                new FontSize { Val = "28" }
            )
        )
        { Type = StyleValues.Paragraph, StyleId = "Heading2" });

        // Subsection heading
        stylesPart.Styles.Append(new Style(
            new StyleName { Val = "heading 3" },
            new BasedOn { Val = "Normal" },
            new StyleParagraphProperties(
                new KeepNext(),
                new KeepLines(),
                new SpacingBetweenLines { Before = "200", After = "120" },
                new OutlineLevel { Val = 2 }
            ),
            new StyleRunProperties(
                new Bold(),
                new Color { Val = Theme.CaptionColor },
                new FontSize { Val = "24" }
            )
        )
        { Type = StyleValues.Paragraph, StyleId = "Heading3" });

        // Figure/table caption
        stylesPart.Styles.Append(new Style(
            new StyleName { Val = "Caption" },
            new BasedOn { Val = "Normal" },
            new StyleParagraphProperties(
                new KeepLines(),
                new Justification { Val = JustificationValues.Center },
                new SpacingBetweenLines { Before = "60", After = "320" }
            ),
            new StyleRunProperties(
                new Italic(),
                new Color { Val = Theme.CaptionColor },
                new FontSize { Val = "20" }
            )
        )
        { Type = StyleValues.Paragraph, StyleId = "Caption" });

        // TOC level 1
        stylesPart.Styles.Append(new Style(
            new StyleName { Val = "toc 1" },
            new BasedOn { Val = "Normal" },
            new StyleParagraphProperties(
                new Tabs(new TabStop { Val = TabStopValues.Right, Leader = TabStopLeaderCharValues.Dot, Position = 9072 }),
                new SpacingBetweenLines { Before = "160", After = "60" }
            ),
            new StyleRunProperties(
                new Bold(),
                new Color { Val = Theme.HeadingColor }
            )
        )
        { Type = StyleValues.Paragraph, StyleId = "TOC1" });

        // TOC level 2
        stylesPart.Styles.Append(new Style(
            new StyleName { Val = "toc 2" },
            new BasedOn { Val = "Normal" },
            new StyleParagraphProperties(
                new Tabs(new TabStop { Val = TabStopValues.Right, Leader = TabStopLeaderCharValues.Dot, Position = 9072 }),
                new SpacingBetweenLines { Before = "40", After = "40" },
                new Indentation { Left = "360" }
            ),
            new StyleRunProperties(
                new Color { Val = Theme.CaptionColor }
            )
        )
        { Type = StyleValues.Paragraph, StyleId = "TOC2" });
    }

    // ═══════════════════════════════════════════════════════════════════════
    // Numbering Registration
    // ═══════════════════════════════════════════════════════════════════════
    private static void RegisterNumbering(MainDocumentPart mainPart)
    {
        var numberingPart = mainPart.AddNewPart<NumberingDefinitionsPart>();
        numberingPart.Numbering = BuildNumberingDefinition();
    }

    // ═══════════════════════════════════════════════════════════════════════
    // Image Loading
    // ═══════════════════════════════════════════════════════════════════════
    private static string LoadImagePart(MainDocumentPart mainPart, string imagePath)
    {
        var imagePart = mainPart.AddImagePart(ImagePartType.Png);
        using var fs = new FileStream(imagePath, FileMode.Open, FileAccess.Read);
        imagePart.FeedData(fs);
        return mainPart.GetIdOfPart(imagePart);
    }

    // ═══════════════════════════════════════════════════════════════════════
    // Full-Page Background Anchor
    // ═══════════════════════════════════════════════════════════════════════
    private static Drawing CreatePageBackdrop(string relId, uint drawingId, string name)
    {
        return new Drawing(
            new DW.Anchor(
                new DW.SimplePosition { X = 0, Y = 0 },
                new DW.HorizontalPosition(new DW.PositionOffset("0"))
                { RelativeFrom = DW.HorizontalRelativePositionValues.Page },
                new DW.VerticalPosition(new DW.PositionOffset("0"))
                { RelativeFrom = DW.VerticalRelativePositionValues.Page },
                new DW.Extent { Cx = DocWidthEmu, Cy = DocHeightEmu },
                new DW.EffectExtent { LeftEdge = 0, TopEdge = 0, RightEdge = 0, BottomEdge = 0 },
                new DW.WrapNone(),
                new DW.DocProperties { Id = drawingId, Name = name },
                new DW.NonVisualGraphicFrameDrawingProperties(
                    new A.GraphicFrameLocks { NoChangeAspect = true }
                ),
                new A.Graphic(
                    new A.GraphicData(
                        new PIC.Picture(
                            new PIC.NonVisualPictureProperties(
                                new PIC.NonVisualDrawingProperties { Id = 0, Name = $"{name}.png" },
                                new PIC.NonVisualPictureDrawingProperties()
                            ),
                            new PIC.BlipFill(
                                new A.Blip { Embed = relId },
                                new A.Stretch(new A.FillRectangle())
                            ),
                            new PIC.ShapeProperties(
                                new A.Transform2D(
                                    new A.Offset { X = 0, Y = 0 },
                                    new A.Extents { Cx = DocWidthEmu, Cy = DocHeightEmu }
                                ),
                                new A.PresetGeometry { Preset = A.ShapeTypeValues.Rectangle }
                            )
                        )
                    )
                    { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" }
                )
            )
            {
                DistanceFromTop = 0,
                DistanceFromBottom = 0,
                DistanceFromLeft = 0,
                DistanceFromRight = 0,
                SimplePos = false,
                RelativeHeight = 251658240,
                BehindDoc = true,
                Locked = false,
                LayoutInCell = true,
                AllowOverlap = true
            }
        );
    }

    // ═══════════════════════════════════════════════════════════════════════
    // Title Page - Product Manual Cover
    // ═══════════════════════════════════════════════════════════════════════
    private static void RenderTitlePage(Body body, string bgRelId)
    {
        body.Append(new Paragraph(new Run(CreatePageBackdrop(bgRelId, 1, "TitleBg"))));

        // Vertical spacing to position content
        body.Append(new Paragraph(
            new ParagraphProperties(new SpacingBetweenLines { Before = "4800" }),
            new Run()
        ));

        // Product name
        body.Append(new Paragraph(
            new ParagraphProperties(
                new Justification { Val = JustificationValues.Left },
                new Indentation { Left = "1200", Right = "1200" },
                new SpacingBetweenLines { After = "240" }
            ),
            new Run(
                new RunProperties(
                    new FontSize { Val = "72" },
                    new Bold(),
                    new Color { Val = Theme.HeadingColor }
                ),
                new Text("NexGen IoT Gateway")
            )
        ));

        // Product subtitle
        body.Append(new Paragraph(
            new ParagraphProperties(
                new Justification { Val = JustificationValues.Left },
                new Indentation { Left = "1200", Right = "1200" },
                new SpacingBetweenLines { After = "480" }
            ),
            new Run(
                new RunProperties(
                    new FontSize { Val = "28" },
                    new Color { Val = Theme.AccentColor }
                ),
                new Text("Model NG-5000 Series")
            )
        ));

        // Document type
        body.Append(new Paragraph(
            new ParagraphProperties(
                new Justification { Val = JustificationValues.Left },
                new Indentation { Left = "1200" },
                new SpacingBetweenLines { After = "2400" }
            ),
            new Run(
                new RunProperties(
                    new FontSize { Val = "24" },
                    new Spacing { Val = 40 },
                    new Color { Val = Theme.CaptionColor }
                ),
                new Text("INSTALLATION & CONFIGURATION MANUAL")
            )
        ));

        // Version info
        body.Append(new Paragraph(
            new ParagraphProperties(
                new Justification { Val = JustificationValues.Left },
                new Indentation { Left = "1200" },
                new SpacingBetweenLines { After = "80" }
            ),
            new Run(
                new RunProperties(
                    new FontSize { Val = "20" },
                    new Color { Val = Theme.CaptionColor }
                ),
                new Text("Document Version 3.2.1")
            )
        ));

        body.Append(new Paragraph(
            new ParagraphProperties(
                new Justification { Val = JustificationValues.Left },
                new Indentation { Left = "1200" }
            ),
            new Run(
                new RunProperties(
                    new FontSize { Val = "20" },
                    new Color { Val = Theme.CaptionColor }
                ),
                new Text("Last Updated: January 2025")
            )
        ));

        // Section break
        body.Append(new Paragraph(
            new ParagraphProperties(
                new SectionProperties(
                    new SectionType { Val = SectionMarkValues.NextPage },
                    new PageSize { Width = (UInt32Value)(uint)DocWidthTwips, Height = (UInt32Value)(uint)DocHeightTwips },
                    new PageMargin { Top = 0, Right = 0, Bottom = 0, Left = 0, Header = 0, Footer = 0 }
                )
            )
        ));
    }

    // ═══════════════════════════════════════════════════════════════════════
    // Contents Page
    // ═══════════════════════════════════════════════════════════════════════
    private static void RenderContentsPage(Body body, string bgRelId, MainDocumentPart mainPart)
    {
        body.Append(new Paragraph(
            new ParagraphProperties(new ParagraphStyleId { Val = "Heading1" }),
            new BookmarkStart { Id = "0", Name = "_TocRoot" },
            new Run(new Text("Contents")),
            new BookmarkEnd { Id = "0" }
        ));

        body.Append(new Paragraph(
            new ParagraphProperties(new SpacingBetweenLines { After = "320" }),
            new Run(
                new RunProperties(
                    new Color { Val = Theme.CaptionColor },
                    new FontSize { Val = "18" }
                ),
                new Text("Select any entry to navigate. Update page numbers via right-click \u2192 \u201cUpdate Field\u201d.")
            )
        ));

        // TOC field
        body.Append(new Paragraph(
            new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }),
            new Run(new FieldCode(" TOC \\o \"1-3\" \\h \\z \\u ") { Space = SpaceProcessingModeValues.Preserve }),
            new Run(new FieldChar { FieldCharType = FieldCharValues.Separate })
        ));

        // Placeholder entries
        string[,] entries = {
            { "Safety Information", "1", "3" },
            { "Package Contents", "1", "4" },
            { "Hardware Specifications", "1", "5" },
            { "Electrical Requirements", "2", "5" },
            { "Environmental Limits", "2", "6" },
            { "Installation Procedure", "1", "7" },
            { "Mounting Options", "2", "7" },
            { "Network Configuration", "1", "9" },
            { "Troubleshooting Guide", "1", "11" },
            { "Warranty Terms", "1", "13" }
        };

        foreach (var i in Enumerable.Range(0, entries.GetLength(0)))
        {
            var level = entries[i, 1];
            var styleId = level == "1" ? "TOC1" : "TOC2";

            body.Append(new Paragraph(
                new ParagraphProperties(new ParagraphStyleId { Val = styleId }),
                new Run(new Text(entries[i, 0])),
                new Run(new TabChar()),
                new Run(new Text(entries[i, 2]))
            ));
        }

        body.Append(new Paragraph(
            new Run(new FieldChar { FieldCharType = FieldCharValues.End })
        ));

        // Section break
        body.Append(new Paragraph(
            new ParagraphProperties(
                new SectionProperties(
                    new SectionType { Val = SectionMarkValues.NextPage },
                    new PageSize { Width = (UInt32Value)(uint)DocWidthTwips, Height = (UInt32Value)(uint)DocHeightTwips },
                    new PageMargin { Top = 1800, Right = 1440, Bottom = 1440, Left = 1440, Header = 720, Footer = 720 }
                )
            )
        ));
    }

    // ═══════════════════════════════════════════════════════════════════════
    // Main Content Sections
    // ═══════════════════════════════════════════════════════════════════════
    private static void RenderMainContent(WordprocessingDocument doc, Body body, string bgRelId, MainDocumentPart mainPart, string assetDir)
    {
        // Setup header with watermark
        var headerPart = mainPart.AddNewPart<HeaderPart>();
        var headerId = mainPart.GetIdOfPart(headerPart);

        var headerImgPart = headerPart.AddImagePart(ImagePartType.Png);
        using (var fs = new FileStream(Path.Combine(assetDir, "content_watermark.png"), FileMode.Open))
            headerImgPart.FeedData(fs);
        var headerImgRelId = headerPart.GetIdOfPart(headerImgPart);

        headerPart.Header = new Header(
            new Paragraph(new Run(CreatePageBackdrop(headerImgRelId, 2, "ContentWatermark"))),
            new Paragraph(
                new ParagraphProperties(
                    new Justification { Val = JustificationValues.Right },
                    new SpacingBetweenLines { Before = "0", After = "0" }
                ),
                new Run(
                    new RunProperties(
                        new FontSize { Val = "18" },
                        new Color { Val = Theme.CaptionColor }
                    ),
                    new Text("NexGen NG-5000 Installation Manual")
                )
            )
        );

        // Setup footer with pagination
        var footerPart = mainPart.AddNewPart<FooterPart>();
        var footerId = mainPart.GetIdOfPart(footerPart);
        var footerPara = new Paragraph(
            new ParagraphProperties(new Justification { Val = JustificationValues.Center })
        );
        footerPara.Append(CreatePageNumberRun());
        footerPara.Append(new Run(new Text(" of ") { Space = SpaceProcessingModeValues.Preserve }));
        footerPara.Append(CreateTotalPagesRun());
        footerPart.Footer = new Footer(footerPara);

        // ══════════════════════════════════════════════════════════════════
        // Chapter 1: Safety Information
        // ══════════════════════════════════════════════════════════════════
        body.Append(CreateH1("Safety Information", "_Sec001"));

        body.Append(new Paragraph(
            new Run(new Text("Before installing or operating the NG-5000 gateway, read all safety instructions carefully. Failure to follow these guidelines may result in equipment damage, personal injury, or voiding of warranty coverage."))
        ));

        body.Append(CreateH2("Electrical Safety"));

        body.Append(CreateBulletItem(1, "Voltage Compatibility", "Verify input voltage matches your regional power supply (100-240V AC, 50/60Hz auto-sensing)"));
        body.Append(CreateBulletItem(1, "Grounding Requirement", "Always connect the protective earth terminal before applying power"));
        body.Append(CreateBulletItem(1, "Surge Protection", "Install an appropriate surge protector rated for at least 2000V transient suppression"));

        body.Append(CreateH2("Environmental Precautions"));

        body.Append(new Paragraph(
            new Run(new Text("The NG-5000 is rated for indoor installation only. Do not expose the unit to direct sunlight, moisture, or corrosive atmospheres. Maintain minimum clearances of 50mm on all sides for adequate ventilation."))
        ));

        body.Append(new Paragraph(new Run(new Break { Type = BreakValues.Page })));

        // ══════════════════════════════════════════════════════════════════
        // Chapter 2: Hardware Specifications
        // ══════════════════════════════════════════════════════════════════
        body.Append(CreateH1("Hardware Specifications", "_Sec002"));

        body.Append(CreateH2("Electrical Requirements"));

        body.Append(new Paragraph(
            new ParagraphProperties(new KeepNext()),
            new Run(new Text("The following table summarizes power and electrical characteristics:"))
        ));

        body.Append(CreateSpecTable());

        body.Append(new Paragraph(
            new ParagraphProperties(new ParagraphStyleId { Val = "Caption" }),
            new Run(new Text("Table 1: Electrical Specifications"))
        ));

        body.Append(CreateH2("Environmental Limits"));

        body.Append(CreateEnvTable());

        body.Append(new Paragraph(
            new ParagraphProperties(new ParagraphStyleId { Val = "Caption" }),
            new Run(new Text("Table 2: Operating Environment"))
        ));

        // Cross-reference to performance chart
        var chartRef = new Paragraph(
            new ParagraphProperties(new KeepNext(), new SpacingBetweenLines { Before = "200" }),
            new Run(new Text("Throughput performance under varying temperature conditions is illustrated in "))
        );
        foreach (var r in CreateHyperlinkField("Fig1", "Figure 1"))
            chartRef.Append(r);
        chartRef.Append(new Run(new Text(".")));
        body.Append(chartRef);

        // Performance bar chart
        CreateBarChart(body, mainPart, "Fig1");

        body.Append(new Paragraph(
            new ParagraphProperties(new ParagraphStyleId { Val = "Caption" }),
            new BookmarkStart { Id = "50", Name = "Fig1" },
            new Run(new Text("Figure 1: Throughput vs. Ambient Temperature")),
            new BookmarkEnd { Id = "50" }
        ));

        body.Append(new Paragraph(new Run(new Break { Type = BreakValues.Page })));

        // ══════════════════════════════════════════════════════════════════
        // Chapter 3: Installation Procedure
        // ══════════════════════════════════════════════════════════════════
        body.Append(CreateH1("Installation Procedure", "_Sec003"));

        body.Append(CreateH2("Mounting Options"));

        var mountPara = new Paragraph(
            new Run(new Text("The NG-5000 supports three mounting configurations. Select the appropriate method based on your deployment environment."))
        );
        body.Append(mountPara);
        AttachFootnoteToRun(doc, mountPara, "DIN rail adapter (part #NG-DIN-01) sold separately.");

        body.Append(CreateBulletItem(1, "DIN Rail Mount", "Snap onto standard 35mm DIN rail using integrated clips"));
        body.Append(CreateBulletItem(1, "Wall Mount", "Secure with four M4 screws through keyhole slots (hardware included)"));
        body.Append(CreateBulletItem(1, "Desktop Placement", "Attach rubber feet to base and place on flat surface"));

        body.Append(CreateH2("Connection Sequence"));

        body.Append(new Paragraph(
            new Run(new Text("Follow the steps below in order to ensure proper initialization:"))
        ));

        body.Append(CreateNumberedStep(1, "Attach all antenna cables before applying power"));
        body.Append(CreateNumberedStep(2, "Connect Ethernet cable to the WAN port (RJ-45, leftmost)"));
        body.Append(CreateNumberedStep(3, "Insert SIM card (nano-SIM) with contacts facing down"));
        body.Append(CreateNumberedStep(4, "Apply power via the DC barrel jack or terminal block"));
        body.Append(CreateNumberedStep(5, "Wait for STATUS LED to show solid green (approx. 90 seconds)"));

        body.Append(new Paragraph(new Run(new Break { Type = BreakValues.Page })));

        // ══════════════════════════════════════════════════════════════════
        // Chapter 4: Network Configuration
        // ══════════════════════════════════════════════════════════════════
        body.Append(CreateH1("Network Configuration", "_Sec004"));

        body.Append(new Paragraph(
            new Run(new Text("Access the web-based configuration interface by navigating to https://192.168.8.1 from a connected device. Default credentials are admin / admin (change immediately after first login)."))
        ));

        body.Append(CreateH2("Protocol Support Matrix"));

        body.Append(CreateProtocolTable());

        body.Append(new Paragraph(
            new ParagraphProperties(new ParagraphStyleId { Val = "Caption" }),
            new Run(new Text("Table 3: Supported Industrial Protocols"))
        ));

        // Protocol distribution chart
        var protoRef = new Paragraph(
            new ParagraphProperties(new KeepNext(), new SpacingBetweenLines { Before = "200" }),
            new Run(new Text("Typical protocol utilization across deployments is shown in "))
        );
        foreach (var r in CreateHyperlinkField("Fig2", "Figure 2"))
            protoRef.Append(r);
        protoRef.Append(new Run(new Text(".")));
        body.Append(protoRef);

        CreatePieChart(body, mainPart, "Fig2");

        body.Append(new Paragraph(
            new ParagraphProperties(new ParagraphStyleId { Val = "Caption" }),
            new BookmarkStart { Id = "51", Name = "Fig2" },
            new Run(new Text("Figure 2: Protocol Distribution in Field Deployments")),
            new BookmarkEnd { Id = "51" }
        ));

        body.Append(new Paragraph(new Run(new Break { Type = BreakValues.Page })));

        // ══════════════════════════════════════════════════════════════════
        // Chapter 5: Troubleshooting
        // ══════════════════════════════════════════════════════════════════
        body.Append(CreateH1("Troubleshooting Guide", "_Sec005"));

        body.Append(new Paragraph(
            new Run(new Text("Consult the table below for common issues and recommended resolutions. If the problem persists after following these steps, contact technical support with your device serial number and firmware version."))
        ));

        body.Append(CreateTroubleshootingTable());

        body.Append(new Paragraph(
            new ParagraphProperties(new ParagraphStyleId { Val = "Caption" }),
            new Run(new Text("Table 4: Common Issues and Resolutions"))
        ));

        // Section properties with header/footer
        body.Append(new Paragraph(
            new ParagraphProperties(
                new SectionProperties(
                    new SectionType { Val = SectionMarkValues.NextPage },
                    new HeaderReference { Type = HeaderFooterValues.Default, Id = headerId },
                    new FooterReference { Type = HeaderFooterValues.Default, Id = footerId },
                    new PageSize { Width = (UInt32Value)(uint)DocWidthTwips, Height = (UInt32Value)(uint)DocHeightTwips },
                    new PageMargin { Top = 1800, Right = 1440, Bottom = 1440, Left = 1440, Header = 720, Footer = 720 }
                )
            )
        ));
    }

    // ═══════════════════════════════════════════════════════════════════════
    // Back Page
    // ═══════════════════════════════════════════════════════════════════════
    private static void RenderBackPage(Body body, string bgRelId)
    {
        body.Append(new Paragraph(new Run(CreatePageBackdrop(bgRelId, 3, "ClosingBg"))));

        body.Append(new Paragraph(
            new ParagraphProperties(new SpacingBetweenLines { Before = "8000" }),
            new Run()
        ));

        // Company name
        body.Append(new Paragraph(
            new ParagraphProperties(
                new Justification { Val = JustificationValues.Center },
                new SpacingBetweenLines { After = "320" }
            ),
            new Run(
                new RunProperties(
                    new FontSize { Val = "36" },
                    new Bold(),
                    new Color { Val = Theme.HeadingColor }
                ),
                new Text("Nexigen Industrial Technologies")
            )
        ));

        // Support contact
        body.Append(new Paragraph(
            new ParagraphProperties(
                new Justification { Val = JustificationValues.Center },
                new SpacingBetweenLines { After = "160" }
            ),
            new Run(
                new RunProperties(
                    new FontSize { Val = "20" },
                    new Color { Val = Theme.CaptionColor }
                ),
                new Text("support@nexigen-tech.io \u00b7 +1 (408) 555-2900")
            )
        ));

        // Legal notice
        body.Append(new Paragraph(
            new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
            new Run(
                new RunProperties(
                    new FontSize { Val = "16" },
                    new Color { Val = Theme.CaptionColor }
                ),
                new Text("\u00a9 2025 Nexigen Industrial Technologies Ltd. All rights reserved.")
            )
        ));

        // Final section properties
        body.Append(new SectionProperties(
            new PageSize { Width = (UInt32Value)(uint)DocWidthTwips, Height = (UInt32Value)(uint)DocHeightTwips },
            new PageMargin { Top = 0, Right = 0, Bottom = 0, Left = 0, Header = 0, Footer = 0 }
        ));
    }

    // ═══════════════════════════════════════════════════════════════════════
    // Heading Helpers
    // ═══════════════════════════════════════════════════════════════════════
    private static Paragraph CreateH1(string text, string bookmarkName)
    {
        var bmId = bookmarkName.Replace("_Sec", "");
        return new Paragraph(
            new ParagraphProperties(new ParagraphStyleId { Val = "Heading1" }),
            new BookmarkStart { Id = bmId, Name = bookmarkName },
            new Run(new Text(text)),
            new BookmarkEnd { Id = bmId }
        );
    }

    private static Paragraph CreateH2(string text)
    {
        return new Paragraph(
            new ParagraphProperties(new ParagraphStyleId { Val = "Heading2" }),
            new Run(new Text(text))
        );
    }

    private static Paragraph CreateBulletItem(int listId, string term, string definition)
    {
        return new Paragraph(
            new ParagraphProperties(
                new NumberingProperties(
                    new NumberingLevelReference { Val = 0 },
                    new NumberingId { Val = listId }
                )
            ),
            new Run(new RunProperties(new Bold()), new Text(term)),
            new Run(new Text(" \u2014 " + definition))
        );
    }

    private static Paragraph CreateNumberedStep(int stepNum, string instruction)
    {
        return new Paragraph(
            new ParagraphProperties(
                new Indentation { Left = "360" },
                new SpacingBetweenLines { Before = "80", After = "80" }
            ),
            new Run(
                new RunProperties(new Bold(), new Color { Val = Theme.AccentColor }),
                new Text($"Step {stepNum}: ")
            ),
            new Run(new Text(instruction))
        );
    }

    // ═══════════════════════════════════════════════════════════════════════
    // Tables
    // ═══════════════════════════════════════════════════════════════════════
    private static Table CreateSpecTable()
    {
        var table = new Table();

        table.Append(new TableProperties(
            new TableWidth { Width = "5000", Type = TableWidthUnitValues.Pct },
            new TableBorders(
                new TopBorder { Val = BorderValues.Single, Size = 8, Color = Theme.GridColor },
                new BottomBorder { Val = BorderValues.Single, Size = 8, Color = Theme.GridColor },
                new LeftBorder { Val = BorderValues.Nil },
                new RightBorder { Val = BorderValues.Nil },
                new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4, Color = Theme.GridColor },
                new InsideVerticalBorder { Val = BorderValues.Nil }
            ),
            new TableCellMarginDefault(
                new TopMargin { Width = "100", Type = TableWidthUnitValues.Dxa },
                new TableCellLeftMargin { Width = 160, Type = TableWidthValues.Dxa },
                new BottomMargin { Width = "100", Type = TableWidthUnitValues.Dxa },
                new TableCellRightMargin { Width = 160, Type = TableWidthValues.Dxa }
            )
        ));

        table.Append(new TableGrid(
            new GridColumn { Width = "3600" },
            new GridColumn { Width = "6400" }
        ));

        table.Append(CreateTableHeader(new[] { "Parameter", "Specification" }));
        table.Append(CreateTableDataRow(new[] { "Input Voltage", "100-240V AC, 50/60Hz (auto-sensing)" }));
        table.Append(CreateTableDataRow(new[] { "Power Consumption", "12W typical, 18W maximum" }));
        table.Append(CreateTableDataRow(new[] { "DC Input Option", "12-48V DC via terminal block" }));
        table.Append(CreateTableDataRow(new[] { "Backup Battery", "Optional 3.7V 2600mAh Li-ion (4hr standby)" }));

        return table;
    }

    private static Table CreateEnvTable()
    {
        var table = new Table();

        table.Append(new TableProperties(
            new TableWidth { Width = "5000", Type = TableWidthUnitValues.Pct },
            new TableBorders(
                new TopBorder { Val = BorderValues.Single, Size = 8, Color = Theme.GridColor },
                new BottomBorder { Val = BorderValues.Single, Size = 8, Color = Theme.GridColor },
                new LeftBorder { Val = BorderValues.Nil },
                new RightBorder { Val = BorderValues.Nil },
                new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4, Color = Theme.GridColor },
                new InsideVerticalBorder { Val = BorderValues.Nil }
            ),
            new TableCellMarginDefault(
                new TopMargin { Width = "100", Type = TableWidthUnitValues.Dxa },
                new TableCellLeftMargin { Width = 160, Type = TableWidthValues.Dxa },
                new BottomMargin { Width = "100", Type = TableWidthUnitValues.Dxa },
                new TableCellRightMargin { Width = 160, Type = TableWidthValues.Dxa }
            )
        ));

        table.Append(new TableGrid(
            new GridColumn { Width = "3600" },
            new GridColumn { Width = "6400" }
        ));

        table.Append(CreateTableHeader(new[] { "Condition", "Range" }));
        table.Append(CreateTableDataRow(new[] { "Operating Temperature", "-20\u00b0C to +70\u00b0C" }));
        table.Append(CreateTableDataRow(new[] { "Storage Temperature", "-40\u00b0C to +85\u00b0C" }));
        table.Append(CreateTableDataRow(new[] { "Relative Humidity", "5% to 95% (non-condensing)" }));
        table.Append(CreateTableDataRow(new[] { "Ingress Protection", "IP40 (optional IP65 enclosure)" }));

        return table;
    }

    private static Table CreateProtocolTable()
    {
        var table = new Table();

        table.Append(new TableProperties(
            new TableWidth { Width = "5000", Type = TableWidthUnitValues.Pct },
            new TableBorders(
                new TopBorder { Val = BorderValues.Single, Size = 8, Color = Theme.GridColor },
                new BottomBorder { Val = BorderValues.Single, Size = 8, Color = Theme.GridColor },
                new LeftBorder { Val = BorderValues.Nil },
                new RightBorder { Val = BorderValues.Nil },
                new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4, Color = Theme.GridColor },
                new InsideVerticalBorder { Val = BorderValues.Nil }
            ),
            new TableCellMarginDefault(
                new TopMargin { Width = "100", Type = TableWidthUnitValues.Dxa },
                new TableCellLeftMargin { Width = 160, Type = TableWidthValues.Dxa },
                new BottomMargin { Width = "100", Type = TableWidthUnitValues.Dxa },
                new TableCellRightMargin { Width = 160, Type = TableWidthValues.Dxa }
            )
        ));

        table.Append(new TableGrid(
            new GridColumn { Width = "2800" },
            new GridColumn { Width = "3600" },
            new GridColumn { Width = "3600" }
        ));

        table.Append(CreateTableHeader(new[] { "Protocol", "Mode", "Max Devices" }));
        table.Append(CreateTableDataRow(new[] { "Modbus RTU/TCP", "Master/Slave", "247" }));
        table.Append(CreateTableDataRow(new[] { "OPC-UA", "Client/Server", "Unlimited" }));
        table.Append(CreateTableDataRow(new[] { "MQTT", "Publish/Subscribe", "N/A" }));
        table.Append(CreateTableDataRow(new[] { "BACnet IP", "Client", "128" }));

        return table;
    }

    private static Table CreateTroubleshootingTable()
    {
        var table = new Table();

        table.Append(new TableProperties(
            new TableWidth { Width = "5000", Type = TableWidthUnitValues.Pct },
            new TableBorders(
                new TopBorder { Val = BorderValues.Single, Size = 8, Color = Theme.GridColor },
                new BottomBorder { Val = BorderValues.Single, Size = 8, Color = Theme.GridColor },
                new LeftBorder { Val = BorderValues.Nil },
                new RightBorder { Val = BorderValues.Nil },
                new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4, Color = Theme.GridColor },
                new InsideVerticalBorder { Val = BorderValues.Nil }
            ),
            new TableCellMarginDefault(
                new TopMargin { Width = "100", Type = TableWidthUnitValues.Dxa },
                new TableCellLeftMargin { Width = 160, Type = TableWidthValues.Dxa },
                new BottomMargin { Width = "100", Type = TableWidthUnitValues.Dxa },
                new TableCellRightMargin { Width = 160, Type = TableWidthValues.Dxa }
            )
        ));

        table.Append(new TableGrid(
            new GridColumn { Width = "3200" },
            new GridColumn { Width = "3200" },
            new GridColumn { Width = "3600" }
        ));

        table.Append(CreateTableHeader(new[] { "Symptom", "Likely Cause", "Resolution" }));
        table.Append(CreateTableDataRow(new[] { "No power LED", "DC supply issue", "Check polarity and voltage" }));
        table.Append(CreateTableDataRow(new[] { "WAN LED off", "Cable/port fault", "Try alternate cable or port" }));
        table.Append(CreateTableDataRow(new[] { "Slow throughput", "Thermal throttling", "Improve ventilation" }));
        table.Append(CreateTableDataRow(new[] { "Config reset", "Firmware corruption", "Reflash via recovery mode" }));

        return table;
    }

    private static TableRow CreateTableHeader(string[] cells)
    {
        var row = new TableRow();
        row.Append(new TableRowProperties(
            new TableHeader(),
            new TableRowHeight { Val = 380, HeightType = HeightRuleValues.AtLeast }
        ));

        foreach (var text in cells)
        {
            row.Append(new TableCell(
                new TableCellProperties(
                    new TableCellWidth { Width = "0", Type = TableWidthUnitValues.Auto },
                    new Shading { Val = ShadingPatternValues.Clear, Fill = Theme.HeaderBg },
                    new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Center }
                ),
                new Paragraph(
                    new ParagraphProperties(
                        new Justification { Val = JustificationValues.Center },
                        new SpacingBetweenLines { Before = "0", After = "0" }
                    ),
                    new Run(
                        new RunProperties(new Bold(), new Color { Val = Theme.HeadingColor }, new FontSize { Val = "21" }),
                        new Text(text)
                    )
                )
            ));
        }

        return row;
    }

    private static TableRow CreateTableDataRow(string[] cells)
    {
        var row = new TableRow();
        row.Append(new TableRowProperties(
            new TableRowHeight { Val = 340, HeightType = HeightRuleValues.AtLeast }
        ));

        for (int i = 0; i < cells.Length; i++)
        {
            row.Append(new TableCell(
                new TableCellProperties(
                    new TableCellWidth { Width = "0", Type = TableWidthUnitValues.Auto },
                    new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Center }
                ),
                new Paragraph(
                    new ParagraphProperties(
                        new Justification { Val = i == 0 ? JustificationValues.Left : JustificationValues.Center },
                        new SpacingBetweenLines { Before = "0", After = "0" }
                    ),
                    new Run(
                        new RunProperties(new Color { Val = Theme.BodyColor }, new FontSize { Val = "21" }),
                        new Text(cells[i])
                    )
                )
            ));
        }

        return row;
    }

    // ═══════════════════════════════════════════════════════════════════════
    // Charts
    // ═══════════════════════════════════════════════════════════════════════
    private static void CreateBarChart(Body body, MainDocumentPart mainPart, string bookmarkName)
    {
        var chartPart = mainPart.AddNewPart<ChartPart>();
        string chartId = mainPart.GetIdOfPart(chartPart);

        chartPart.ChartSpace = BuildBarChartSpace();

        long chartW = 5040000;  // ~14 cm
        long chartH = 3024000;  // ~8.4 cm

        var drawing = new Drawing(
            new DW.Inline(
                new DW.Extent { Cx = chartW, Cy = chartH },
                new DW.EffectExtent { LeftEdge = 0, TopEdge = 0, RightEdge = 0, BottomEdge = 0 },
                new DW.DocProperties { Id = 20, Name = "ThroughputChart" },
                new DW.NonVisualGraphicFrameDrawingProperties(
                    new A.GraphicFrameLocks { NoChangeAspect = true }
                ),
                new A.Graphic(
                    new A.GraphicData(
                        new C.ChartReference { Id = chartId }
                    )
                    { Uri = "http://schemas.openxmlformats.org/drawingml/2006/chart" }
                )
            )
            { DistanceFromTop = 0, DistanceFromBottom = 0, DistanceFromLeft = 0, DistanceFromRight = 0 }
        );

        body.Append(new Paragraph(
            new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
            new Run(drawing)
        ));
    }

    private static C.ChartSpace BuildBarChartSpace()
    {
        var chartSpace = new C.ChartSpace();
        chartSpace.AddNamespaceDeclaration("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");
        chartSpace.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

        var chart = new C.Chart();
        var plotArea = new C.PlotArea();

        var barChart = new C.BarChart(
            new C.BarDirection { Val = C.BarDirectionValues.Column },
            new C.BarGrouping { Val = C.BarGroupingValues.Clustered },
            new C.VaryColors { Val = false }
        );

        // Throughput data series
        var series = new C.BarChartSeries();
        series.Append(new C.Index { Val = 0 });
        series.Append(new C.Order { Val = 0 });
        series.Append(new C.SeriesText(new C.NumericValue("Throughput (Mbps)")));

        series.Append(new C.ChartShapeProperties(
            new A.SolidFill(new A.RgbColorModelHex { Val = Theme.AccentColor })
        ));

        string[] temps = { "-20\u00b0C", "0\u00b0C", "25\u00b0C", "50\u00b0C", "70\u00b0C" };
        double[] throughput = { 92, 98, 100, 95, 82 };

        var catData = new C.CategoryAxisData();
        var strRef = new C.StringReference();
        var strCache = new C.StringCache();
        strCache.Append(new C.PointCount { Val = (uint)temps.Length });
        for (int i = 0; i < temps.Length; i++)
            strCache.Append(new C.StringPoint(new C.NumericValue(temps[i])) { Index = (uint)i });
        strRef.Append(strCache);
        catData.Append(strRef);
        series.Append(catData);

        var valData = new C.Values();
        var numRef = new C.NumberReference();
        var numCache = new C.NumberingCache();
        numCache.Append(new C.FormatCode("0"));
        numCache.Append(new C.PointCount { Val = (uint)throughput.Length });
        for (int i = 0; i < throughput.Length; i++)
            numCache.Append(new C.NumericPoint(new C.NumericValue(throughput[i].ToString())) { Index = (uint)i });
        numRef.Append(numCache);
        valData.Append(numRef);
        series.Append(valData);

        barChart.Append(series);
        barChart.Append(new C.AxisId { Val = 100 });
        barChart.Append(new C.AxisId { Val = 101 });

        plotArea.Append(barChart);

        plotArea.Append(new C.CategoryAxis(
            new C.AxisId { Val = 100 },
            new C.Scaling(new C.Orientation { Val = C.OrientationValues.MinMax }),
            new C.Delete { Val = false },
            new C.AxisPosition { Val = C.AxisPositionValues.Bottom },
            new C.TickLabelPosition { Val = C.TickLabelPositionValues.NextTo },
            new C.CrossingAxis { Val = 101 },
            new C.Crosses { Val = C.CrossesValues.AutoZero },
            new C.AutoLabeled { Val = true }
        ));

        plotArea.Append(new C.ValueAxis(
            new C.AxisId { Val = 101 },
            new C.Scaling(new C.Orientation { Val = C.OrientationValues.MinMax }),
            new C.Delete { Val = false },
            new C.AxisPosition { Val = C.AxisPositionValues.Left },
            new C.MajorGridlines(),
            new C.NumberingFormat { FormatCode = "0", SourceLinked = false },
            new C.TickLabelPosition { Val = C.TickLabelPositionValues.NextTo },
            new C.CrossingAxis { Val = 100 },
            new C.Crosses { Val = C.CrossesValues.AutoZero }
        ));

        chart.Append(plotArea);
        chart.Append(new C.Legend(
            new C.LegendPosition { Val = C.LegendPositionValues.Bottom },
            new C.Overlay { Val = false }
        ));
        chart.Append(new C.PlotVisibleOnly { Val = true });
        chartSpace.Append(chart);

        return chartSpace;
    }

    private static void CreatePieChart(Body body, MainDocumentPart mainPart, string bookmarkName)
    {
        var chartPart = mainPart.AddNewPart<ChartPart>();
        string chartId = mainPart.GetIdOfPart(chartPart);

        chartPart.ChartSpace = BuildPieChartSpace();

        long chartW = 4320000;  // ~12 cm
        long chartH = 3240000;  // ~9 cm

        var drawing = new Drawing(
            new DW.Inline(
                new DW.Extent { Cx = chartW, Cy = chartH },
                new DW.EffectExtent { LeftEdge = 0, TopEdge = 0, RightEdge = 0, BottomEdge = 0 },
                new DW.DocProperties { Id = 21, Name = "ProtocolPie" },
                new DW.NonVisualGraphicFrameDrawingProperties(
                    new A.GraphicFrameLocks { NoChangeAspect = true }
                ),
                new A.Graphic(
                    new A.GraphicData(
                        new C.ChartReference { Id = chartId }
                    )
                    { Uri = "http://schemas.openxmlformats.org/drawingml/2006/chart" }
                )
            )
            { DistanceFromTop = 0, DistanceFromBottom = 0, DistanceFromLeft = 0, DistanceFromRight = 0 }
        );

        body.Append(new Paragraph(
            new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
            new Run(drawing)
        ));
    }

    private static C.ChartSpace BuildPieChartSpace()
    {
        var chartSpace = new C.ChartSpace();
        chartSpace.AddNamespaceDeclaration("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");
        chartSpace.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

        var chart = new C.Chart();
        var plotArea = new C.PlotArea();

        var pieChart = new C.PieChart(new C.VaryColors { Val = true });

        var series = new C.PieChartSeries();
        series.Append(new C.Index { Val = 0 });
        series.Append(new C.Order { Val = 0 });
        series.Append(new C.SeriesText(new C.NumericValue("Protocol Usage")));

        string[] colors = {
            ColorSchemes.Verdant.Decorative.Muted.Primary,
            ColorSchemes.Verdant.Decorative.Muted.Secondary,
            ColorSchemes.Verdant.Decorative.Muted.Accent,
            ColorSchemes.Verdant.Text.Body
        };
        string[] protocols = { "Modbus", "OPC-UA", "MQTT", "BACnet" };
        double[] usage = { 42, 28, 19, 11 };

        for (uint i = 0; i < colors.Length; i++)
        {
            series.Append(new C.DataPoint(
                new C.Index { Val = i },
                new C.Bubble3D { Val = false },
                new C.ChartShapeProperties(
                    new A.SolidFill(new A.RgbColorModelHex { Val = colors[i] })
                )
            ));
        }

        var catData = new C.CategoryAxisData();
        var strRef = new C.StringReference();
        var strCache = new C.StringCache();
        strCache.Append(new C.PointCount { Val = (uint)protocols.Length });
        for (int i = 0; i < protocols.Length; i++)
            strCache.Append(new C.StringPoint(new C.NumericValue(protocols[i])) { Index = (uint)i });
        strRef.Append(strCache);
        catData.Append(strRef);
        series.Append(catData);

        var valData = new C.Values();
        var numRef = new C.NumberReference();
        var numCache = new C.NumberingCache();
        numCache.Append(new C.FormatCode("0"));
        numCache.Append(new C.PointCount { Val = (uint)usage.Length });
        for (int i = 0; i < usage.Length; i++)
            numCache.Append(new C.NumericPoint(new C.NumericValue(usage[i].ToString())) { Index = (uint)i });
        numRef.Append(numCache);
        valData.Append(numRef);
        series.Append(valData);

        pieChart.Append(series);
        plotArea.Append(pieChart);
        chart.Append(plotArea);

        chart.Append(new C.Legend(
            new C.LegendPosition { Val = C.LegendPositionValues.Right },
            new C.Overlay { Val = false }
        ));
        chart.Append(new C.PlotVisibleOnly { Val = true });
        chartSpace.Append(chart);

        return chartSpace;
    }

    // ═══════════════════════════════════════════════════════════════════════
    // Document Infrastructure
    // ═══════════════════════════════════════════════════════════════════════
    private static void EnableFieldRefresh(MainDocumentPart mainPart)
    {
        var settingsPart = mainPart.AddNewPart<DocumentSettingsPart>();
        settingsPart.Settings = new Settings(
            new UpdateFieldsOnOpen { Val = true },
            new DisplayBackgroundShape()
        );
    }

    private static Numbering BuildNumberingDefinition()
    {
        return new Numbering(
            new AbstractNum(
                new Level(
                    new NumberingFormat { Val = NumberFormatValues.Bullet },
                    new LevelText { Val = "\u2022" },
                    new LevelJustification { Val = LevelJustificationValues.Left },
                    new ParagraphProperties(new Indentation { Left = "720", Hanging = "360" })
                ) { LevelIndex = 0 }
            ) { AbstractNumberId = 1 },
            new NumberingInstance(new AbstractNumId { Val = 1 }) { NumberID = 1 }
        );
    }

    private static Run CreatePageNumberRun()
    {
        return new Run(
            new FieldChar { FieldCharType = FieldCharValues.Begin },
            new FieldCode(" PAGE ") { Space = SpaceProcessingModeValues.Preserve },
            new FieldChar { FieldCharType = FieldCharValues.Separate },
            new Text("1"),
            new FieldChar { FieldCharType = FieldCharValues.End }
        );
    }

    private static Run CreateTotalPagesRun()
    {
        return new Run(
            new FieldChar { FieldCharType = FieldCharValues.Begin },
            new FieldCode(" NUMPAGES ") { Space = SpaceProcessingModeValues.Preserve },
            new FieldChar { FieldCharType = FieldCharValues.Separate },
            new Text("1"),
            new FieldChar { FieldCharType = FieldCharValues.End }
        );
    }

    private static void AttachFootnoteToRun(WordprocessingDocument doc, Paragraph para, string noteText)
    {
        var mainPart = doc.MainDocumentPart!;

        if (mainPart.FootnotesPart == null)
        {
            var fnPart = mainPart.AddNewPart<FootnotesPart>();
            fnPart.Footnotes = new Footnotes(
                new Footnote(
                    new Paragraph(new Run(new SeparatorMark()))
                ) { Type = FootnoteEndnoteValues.Separator, Id = -1 },
                new Footnote(
                    new Paragraph(new Run(new ContinuationSeparatorMark()))
                ) { Type = FootnoteEndnoteValues.ContinuationSeparator, Id = 0 }
            );
        }

        var footnotes = mainPart.FootnotesPart!.Footnotes!;
        int nextId = (int)(footnotes.Elements<Footnote>().Max(fn => fn.Id?.Value ?? 0) + 1);

        footnotes.Append(new Footnote(
            new Paragraph(
                new Run(
                    new RunProperties(new VerticalTextAlignment { Val = VerticalPositionValues.Superscript }),
                    new FootnoteReferenceMark()
                ),
                new Run(new Text(" " + noteText) { Space = SpaceProcessingModeValues.Preserve })
            )
        ) { Id = nextId });

        para.Append(new Run(
            new RunProperties(new VerticalTextAlignment { Val = VerticalPositionValues.Superscript }),
            new FootnoteReference { Id = nextId }
        ));
    }

    private static IEnumerable<Run> CreateHyperlinkField(string bookmarkName, string displayText)
    {
        yield return new Run(new FieldChar { FieldCharType = FieldCharValues.Begin });
        yield return new Run(new FieldCode($" REF {bookmarkName} \\h ") { Space = SpaceProcessingModeValues.Preserve });
        yield return new Run(new FieldChar { FieldCharType = FieldCharValues.Separate });
        yield return new Run(
            new RunProperties(new Color { Val = Theme.AccentColor }),
            new Text(displayText)
        );
        yield return new Run(new FieldChar { FieldCharType = FieldCharValues.End });
    }
}
