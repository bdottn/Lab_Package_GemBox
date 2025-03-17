using GemBox.Document;

namespace Lab.Package.GemBox
{
    public class Document
    {
        private const string LicenseKey = "FREE-LIMITED-KEY";

        public Document()
        {
            // If using the Professional version, put your serial key below.
            ComponentInfo.SetLicense(LicenseKey);
        }

        public void HelloWorld()
        {
            // 1. 建立 DocumentModel
            // 2. 建立 Section，DocumentModel 加入此 Section
            // 3. 建立 Paragraph，Section 加入此 Paragraph
            // 4. 建立 Run，Paragraph 加入此 Run
            // 5. 儲存 DocumentModel，寫入檔案路徑

            DocumentModel document = new DocumentModel();

            Section section = new Section(document);
            document.Sections.Add(section);

            Paragraph paragraph = new Paragraph(document);
            section.Blocks.Add(paragraph);

            Run run = new Run(document, "Hello World!");
            paragraph.Inlines.Add(run);

            document.Save("HelloWorld.docx");
        }

        public void Reading()
        {
            // 1. 建立 DocumentModel，讀取檔案路徑
            // 2. 讀取 DocumentModel 內 Paragraph 元素
            // 3. 讀取 Paragraph 元素內 Run 元素
            // 4. 讀取 Run 元素內文字與格式
            // 5. 判斷 Run 元素是否為粗體，是則將文字轉換為 'Mathematical Bold Italic' Unicode 字元
            // 6. 寫入文字至檔案

            var document = DocumentModel.Load("Input/Reading.docx");

            using (var writer = File.CreateText("Output.txt"))
            {
                // Iterate through all Paragraph elements in the Word document.
                foreach (Paragraph paragraph in document.GetChildElements(true, ElementType.Paragraph))
                {
                    // Iterate through all Run elements in the Paragraph element.
                    foreach (Run run in paragraph.GetChildElements(true, ElementType.Run))
                    {
                        string text = run.Text;
                        CharacterFormat format = run.CharacterFormat;

                        // Replace text with bold formatting to 'Mathematical Bold Italic' Unicode characters.
                        // For instance, "ABC" to "𝑨𝑩𝑪".
                        if (format.Bold)
                        {
                            text = string.Concat(text.Select(
                                c => c >= 'A' && c <= 'Z' ? char.ConvertFromUtf32(119847 + c) :
                                     c >= 'a' && c <= 'z' ? char.ConvertFromUtf32(119841 + c) :
                                     c.ToString()));
                        }

                        writer.Write(text);
                    }

                    writer.WriteLine();
                }
            }
        }
    }
}