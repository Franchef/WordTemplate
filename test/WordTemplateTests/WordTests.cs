using WordTemplateDomain;

namespace WordTemplateTests
{
    public class WordTests
    {
        [Theory]
        [InlineData(@"C:\Azure\Francesco\WordTemplate\TestTemplate.docx")]
        public void OpenAndGetFields(string filePath)
        {
            var sut = new WordTemplateBinder(filePath);

            var fields = sut.GetFields();

            Assert.NotEmpty(fields);
        }

        [Theory]
        [InlineData(@"C:\Azure\Francesco\WordTemplate\TestTemplate.docx")]
        public void ReplaceAndSave(string filePath)
        {
            var sut = new WordTemplateBinder(filePath);

            sut.ReplaceFields(new Dictionary<string, string> {
                { "{cliente.nome}", "Francesco" },
                { "{cliente.cognome}", "Carbone" },
            });
        }
    }
}