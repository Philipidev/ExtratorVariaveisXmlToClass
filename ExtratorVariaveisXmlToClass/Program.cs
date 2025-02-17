using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Text;

namespace ExtratorVariaveisXmlToClass
{
    internal class Program
    {
        static void Main(string[] args)
        {
            // Path to your Word document template
            string caminhoWord = @"C:\Users\phili\Downloads\TEMPLATE_APENAS_REGISTREO_COPEL (1).docx";

            // Extract all tag values from the Word file.
            List<string> tagValues = ExtractTagValues(caminhoWord);

            // Get the file name (without extension) to use as the class name.
            string className = Path.GetFileNameWithoutExtension(caminhoWord);

            // Generate the C# class code.
            string classCode = GenerateClassCode(className, tagValues);

            // Output the generated class code.
            Console.WriteLine(classCode);
            Console.ReadKey();
        }

        /// <summary>
        /// Opens the Word document and extracts the values from each w:tag element.
        /// </summary>
        /// <param name="wordFilePath">The full path to the Word document.</param>
        /// <returns>A list of tag values found in the document.</returns>
        static List<string> ExtractTagValues(string wordFilePath)
        {
            List<string> tagValues = new List<string>();

            // Open the Word document for read-only access.
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(wordFilePath, false))
            {
                // Get the main document part.
                MainDocumentPart mainPart = wordDoc.MainDocumentPart;

                // Find all the Tag elements in the document.
                var tags = mainPart.Document.Descendants<Tag>();

                foreach (var tag in tags)
                {
                    // The tag's value is stored in the Val attribute.
                    if (tag.Val != null)
                    {
                        string tagVal = tag.Val.Value;
                        if (!string.IsNullOrEmpty(tagVal) && !tagValues.Contains(tagVal))
                        {
                            tagValues.Add(tagVal);
                        }
                    }
                }
            }

            return tagValues;
        }

        /// <summary>
        /// Generates a C# class definition as a string using the provided class name and properties.
        /// </summary>
        /// <param name="className">The name of the class to generate.</param>
        /// <param name="properties">A list of property names.</param>
        /// <returns>A string containing the C# class code.</returns>
        static string GenerateClassCode(string className, List<string> properties)
        {
            StringBuilder sb = new StringBuilder();
            sb.AppendLine($"public class {className}");
            sb.AppendLine("{");

            foreach (string prop in properties)
            {
                sb.AppendLine($"    public string {prop} {{ get; set; }}");
            }

            sb.AppendLine("}");
            return sb.ToString();
        }
    }
}
