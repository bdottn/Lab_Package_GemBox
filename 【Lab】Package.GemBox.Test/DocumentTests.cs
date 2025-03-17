namespace Lab.Package.GemBox.Test
{
    public class DocumentTests
    {
        [Fact]
        public void HelloWorld_Test()
        {
            var document = new Document();

            document.HelloWorld();
        }

        [Fact]
        public void Reading_Test()
        {
            var document = new Document();

            document.Reading();
        }

        [Fact]
        public void CreateChart_Test()
        {
            var document = new Document();

            document.CreateChart();
        }
    }
}