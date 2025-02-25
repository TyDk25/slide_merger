using ShapeCrawler;

class Program
{
    static void Main()
    {
        // Call the method to process PowerPoint files
        GetFiles();
    }

    static void GetFiles()
    {
        // Define the directory containing PowerPoint files
        var files = Directory.EnumerateFiles(@"/path/to/presentation/folder", "*.pptx");

        // Create a new presentation
        var pres = new Presentation();

        // Remove the first slide from the new presentation
        pres.Slides.Remove(pres.Slides[0]);

        foreach (var file in files)
        {
            // Open each PowerPoint file
            var ppts = new Presentation(file);

            // Get the first slide from the current PowerPoint file
            var slides = ppts.Slides[0];

            // Add the slide to the new presentation
            pres.Slides.Add(slides);

            // Save the updated presentation
            pres.SaveAs(@"/path/to/output/folder/merged_presentation.pptx");
        }
    }
}
