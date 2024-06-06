using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.IO;
using Exercise.WordBuilder.Models;
using DocumentFormat.OpenXml.Vml.Spreadsheet;

namespace Exercise.WordBuilder
{
    internal static class WordGenerator
    {

        public static async Task<byte[]> GenerateDocument(DocumentData document)
        {

            if (File.Exists("Document.docx"))
            {
                File.Delete("Document.docx");
            }

            using (WordprocessingDocument wordDocument = WordprocessingDocument.Create("Document.docx", WordprocessingDocumentType.Document))
            //MemoryStream _Ms = new MemoryStream();
            //using (WordprocessingDocument wordDocument = WordprocessingDocument.Create(_Ms, WordprocessingDocumentType.Document))
            {
                MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();
                mainPart.Document = new Document();
                Body body = new Body();
                mainPart.Document.Append(body);

                GenerateStyleDefinition_Heading1(wordDocument, mainPart);

                Paragraph title = new Paragraph();
                Run titleRun = new Run();

                RunProperties runTitleProperties = new RunProperties();
                runTitleProperties.Append(new Bold());
                runTitleProperties.Append(new Italic());
                runTitleProperties.Append(new FontSize() { Val = "48" });
                //runTitleProperties.Append(new HorizontalTextAlignment("center"));     // This is not working
                runTitleProperties.Append(new TextAlignment() { Val = VerticalTextAlignmentValues.Center });
                titleRun.Append(runTitleProperties);

                titleRun.Append(new Text(document.Title));
                title.Append(titleRun);
                body.Append(title);

                if (document.IncludeTableOfContents)
                {
                    Paragraph para = new Paragraph();
                    Run run = new Run();
                    RunProperties runProperties = new RunProperties();
                    Bold bold = new Bold();
                    runProperties.Append(bold);
                    run.Append(runProperties);
                    run.Append(new Text("Inhaltsverzeichnis"));
                    para.Append(run);
                    body.Append(para);

                    foreach (var part in document.TableOfContents)
                    {
                        Paragraph para2 = new Paragraph();
                        Run run2 = new Run();
                        run2.Append(new Text(part));
                        para2.Append(run2);
                        body.Append(para2);
                    }
                }

                if (!string.IsNullOrEmpty(document.Preface))
                {
                    Paragraph prefaceTitle = new Paragraph();
                    Run prefacetitleRun = new Run();
                    prefacetitleRun.Append(new Break() { Type = BreakValues.Page });

                    RunProperties runProperties = new RunProperties();
                    runProperties.Append(new Bold());
                    prefacetitleRun.Append(runProperties);

                    prefacetitleRun.Append(new Text("Vorwort"));
                    prefaceTitle.Append(prefacetitleRun);
                    body.Append(prefaceTitle);

                    Paragraph preface = new Paragraph();
                    Run prefaceRun = new Run();

                    runProperties = new RunProperties();
                    runProperties.Append(new Italic());
                    prefaceRun.Append(runProperties);

                    prefaceRun.Append(new Text(document.Preface));
                    preface.Append(prefaceRun);
                    body.Append(preface);
                }

                foreach (var part in document.Chapters)
                {
                    Paragraph chapter = new Paragraph();
                    if (document.ChapterNumbering)
                    {
                        ParagraphProperties paragraphProperties = new ParagraphProperties();
                        NumberingProperties numberingProperties = new NumberingProperties();
                        NumberingLevelReference numberingLevelReference = new NumberingLevelReference() { Val = 0 };
                        NumberingId numberingId = new NumberingId() { Val = 1 };
                        numberingProperties.Append(numberingLevelReference);
                        numberingProperties.Append(numberingId);
                        paragraphProperties.Append(numberingProperties);

                        string styleId = GetStyleIdFromStyleName(wordDocument, "Heading 1");
                        paragraphProperties.ParagraphStyleId = new ParagraphStyleId() { Val = styleId };

                        chapter.Append(paragraphProperties);

                    }
                    Run chapterRun = new Run();
                    chapterRun.Append(new Break() { Type = BreakValues.Page });

                    RunProperties runProperties = new RunProperties();
                    runProperties.Append(new Bold());
                    chapterRun.Append(runProperties);

                    chapterRun.Append(new Text(part.Header));
                    chapter.Append(chapterRun);
                    body.Append(chapter);

                    if (!string.IsNullOrEmpty(part.Teaser))
                    {
                        Paragraph teaser = new Paragraph();
                        Run teaserRun = new Run();
                        teaserRun.Append(new Text(part.Teaser));
                        teaser.Append(teaserRun);
                        body.Append(teaser);
                    }

                    foreach (var paragraph in part.Paragraphs)
                    {
                        Paragraph text = new Paragraph();
                        Run textRun = new Run();
                        textRun.Append(new Text(paragraph));
                        text.Append(textRun);
                        body.Append(text);
                    }
                }

                mainPart.Document.Save();
                //wordDocument.Package.Flush();
            }

            byte[] content = await System.IO.File.ReadAllBytesAsync("Document.docx");
            return content;
        }


        static void GenerateStyleDefinition_Heading1(WordprocessingDocument doc, MainDocumentPart mainPart)
        {
            Styles? styles = mainPart.StyleDefinitionsPart?.Styles ?? AddStylesPartToPackage(doc).Styles;

            if (styles != null && !IsStyleIdInDocument(doc, "Heading1"))
            {
                Style style = new Style()
                {
                    Type = StyleValues.Paragraph,
                    StyleId = "Heading1",
                    CustomStyle = false, 
                };
                style.Append(new StyleName() { Val = "Heading 1" });
                style.Append(new BasedOn() { Val = "Absatz-Standardschriftart" });      // muss hierauf basieren (für Inhaltsverzeichnis)
                style.Append(new NextParagraphStyle() { Val = "Normal" });

                StyleRunProperties styleRunProperties1 = new StyleRunProperties();
                styleRunProperties1.Append(new Bold());
                //styleRunProperties1.Append(new Italic());
                //styleRunProperties1.Append(new RunFonts() { Ascii = "Lucida Console" });
                styleRunProperties1.Append(new FontSize() { Val = "30" });  // Sizes are in half-points. Oy!
                styleRunProperties1.Append(new Color() { Val = "055D83" });

                style.Append(styleRunProperties1);
                styles.Append(style);
            }
        }

        static StyleDefinitionsPart AddStylesPartToPackage(WordprocessingDocument? doc)
        {
            StyleDefinitionsPart part;

            if (doc?.MainDocumentPart is null)
            {
                throw new ArgumentNullException("MainDocumentPart is null.");
            }

            part = doc.MainDocumentPart.AddNewPart<StyleDefinitionsPart>();
            Styles root = new Styles();
            root.Save(part);
            return part;
        }

        static bool IsStyleIdInDocument(WordprocessingDocument doc, string styleid)
        {
            // Get access to the Styles element for this document.
            Styles? s = doc.MainDocumentPart?.StyleDefinitionsPart?.Styles;

            if (s is null)
            {
                return false;
            }

            // Check that there are styles and how many.
            int n = s.Elements<Style>().Count();

            if (n == 0)
            {
                return false;
            }

            // Look for a match on styleid.
            Style? style = s.Elements<Style>()
                .Where(st => (st.StyleId is not null && st.StyleId == styleid) && (st.Type is not null && st.Type == StyleValues.Paragraph))
                .FirstOrDefault();
            if (style is null)
            {
                return false;
            }

            return true;
        }

        static string? GetStyleIdFromStyleName(WordprocessingDocument doc, string styleName)
        {
            StyleDefinitionsPart? stylePart = doc.MainDocumentPart?.StyleDefinitionsPart;
            string? styleId = stylePart?.Styles?.Descendants<StyleName>()
                .Where(s =>
                {
                    OpenXmlElement? p = s.Parent;
                    EnumValue<StyleValues>? styleValue = p is null ? null : ((Style)p).Type;

                    return s.Val is not null && s.Val.Value is not null && s.Val.Value.Equals(styleName) &&
                    (styleValue is not null && styleValue == StyleValues.Paragraph);
                })
                .Select(n =>
                {

                    OpenXmlElement? p = n.Parent;
                    return p is null ? null : ((Style)p).StyleId;
                }).FirstOrDefault();

            return styleId;
        }

    }
}
