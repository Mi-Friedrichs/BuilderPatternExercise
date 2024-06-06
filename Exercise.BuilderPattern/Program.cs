namespace Exercise.BuilderPattern
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Hello, World! Generate your own word document");

            Exercise.WordBuilder.WordBuilder wordBuilder = new();

            wordBuilder.AddTitle("My Exercise Document");
            wordBuilder.AddPreface("Und hier ein kleines Vorwort");
            
            Guid chapterID = wordBuilder.AddChapter("Chapter 1", "This is the first chapter of my document");
            wordBuilder.AddParagraphToChapter(chapterID, "This is the second paragraph of the first chapter");
            wordBuilder.AddParagraphToChapter(chapterID, "Lorem ipsum dolor sit amet, consetetur sadipscing elitr, sed diam nonumy eirmod tempor invidunt ut labore et dolore magna aliquyam erat, sed diam voluptua. At vero eos et accusam et justo duo dolores et ea rebum. Stet clita kasd gubergren, no sea takimata sanctus est Lorem ipsum dolor sit amet. Lorem ipsum dolor sit amet, consetetur sadipscing elitr, sed diam nonumy eirmod tempor invidunt ut labore et dolore magna aliquyam erat, sed diam voluptua. At vero eos et accusam et justo duo dolores et ea rebum. Stet clita kasd gubergren, no sea takimata sanctus est Lorem ipsum dolor sit amet.");
            wordBuilder.AddParagraphToChapter(chapterID, "This is the fourth paragraph of the first chapter");

            chapterID = wordBuilder.AddChapter("Chapter 2", "This is the second chapter of my document");
            wordBuilder.AddParagraphToChapter(chapterID, "Lorem ipsum dolor sit amet, consetetur sadipscing elitr, sed diam nonumy eirmod tempor invidunt ut labore et dolore magna aliquyam erat, sed diam voluptua. At vero eos et accusam et justo duo dolores et ea rebum. Stet clita kasd gubergren, no sea takimata sanctus est Lorem ipsum dolor sit amet. Lorem ipsum dolor sit amet, consetetur sadipscing elitr, sed diam nonumy eirmod tempor invidunt ut labore et dolore magna aliquyam erat, sed diam voluptua. At vero eos et accusam et justo duo dolores et ea rebum. Stet clita kasd gubergren, no sea takimata sanctus est Lorem ipsum dolor sit amet.");
            wordBuilder.AddParagraphToChapter(chapterID, "At vero eos et accusam et justo duo dolores et ea rebum. Stet clita kasd gubergren, no sea takimata sanctus est Lorem ipsum dolor sit amet. Lorem ipsum dolor sit amet, consetetur sadipscing elitr, sed diam nonumy eirmod tempor invidunt ut labore et dolore magna aliquyam erat, sed diam voluptua. At vero eos et accusam et justo duo dolores et ea rebum. Stet clita kasd gubergren, no sea takimata sanctus est Lorem ipsum dolor sit amet.");

            wordBuilder.IncludeTableOfContents(true);
            wordBuilder.UseChapterNumbering(true);

            Console.Write("Building document...");

            byte[] content = wordBuilder.Build().Result;

            File.WriteAllBytes(@"c:\Temp\ExerciseBuilderPattern.docx", content);

            Console.WriteLine("\r\nDone.");
        }
    }
}
