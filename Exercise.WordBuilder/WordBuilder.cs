using Exercise.WordBuilder.Models;

namespace Exercise.WordBuilder
{
    public class WordBuilder
    {
        private DocumentData document = new();

        public WordBuilder()
        {
            this.Reset();
        }

        public void Reset()
        {
            this.document = new();
        }

        public void AddTitle(string title)
        {
            document.Title = title;
        }

        public void AddPreface(string preface)
        {
            document.Preface = preface;
        }

        public Guid AddChapter(string header, string text, string? teaser = null)
        {
            var chapter = new Chapter
            {
                Header = header,
                Teaser = teaser,
            };
            chapter.Paragraphs.Add(text);

            document.Chapters.Add(chapter);

            return chapter.Id;
        }

        public void AddParagraphToChapter(Guid chapterId, string paragraph)
        {
            document.Chapters.FirstOrDefault(c=> c.Id == chapterId)?.Paragraphs.Add(paragraph);
        }

        public void IncludeTableOfContents(bool include)
        {
            document.IncludeTableOfContents = include;
        }
        public void UseChapterNumbering(bool use)
        {
            document.ChapterNumbering = use;
        }

        public async Task<byte[]> Build()
        {
            GenerateTableOfContents();
            return await WordGenerator.GenerateDocument(document);
        }

        /// <summary>
        /// Inhaltsverzeichnis erstellen
        /// </summary>
        private void GenerateTableOfContents()
        {
            int currentPage = 2;
            document.TableOfContents.Clear();

            foreach (var part in document.Chapters)
            {
                document.TableOfContents.Add(part.Header.PadRight(57, '.') + currentPage.ToString().PadLeft(3, '.'));
                currentPage += part.Pages;
            }
        }

    }
}
