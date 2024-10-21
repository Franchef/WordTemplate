using System.Reflection;
using WordTemplateDomain;

namespace WordTemplateTests
{
    public class WordTests
    {
        private string GetFullPath(string fileName)
        {
            return Path.Combine(
                Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location),
                fileName
            );
        }

        [Theory]
        [InlineData(@"TestTemplate.docx")]
        public void OpenAndGetFields(string filePath)
        {
            using var sut = new WordTemplateBinder(GetFullPath(filePath));

            var fields = sut.GetFields();

            Assert.NotEmpty(fields);
        }

        [Theory]
        [InlineData(@"TestTemplate.docx")]
        public void ReplaceAndSave(string filePath)
        {
            using var sut = new WordTemplateBinder(GetFullPath(filePath));
            sut.ReplaceFields(new Dictionary<string, string> {
                { "{cliente.nome}", "Francesco" },
                { "{cliente.cognome}", "Carbone" },
            });
        }
    }
}