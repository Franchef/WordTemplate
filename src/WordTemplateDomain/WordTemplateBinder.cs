using System.Text.RegularExpressions;

using Microsoft.Office.Interop.Word;

namespace WordTemplateDomain
{

    public class WordTemplateBinder : IDisposable
    {
        private readonly string _file;
        private readonly Application _app;
        private readonly Document _document;
        private readonly Regex _regex;
        public WordTemplateBinder(string filePath)
        {
            _file = filePath;
            _app = new Application();
            _document = _app.Documents.Open(filePath, ReadOnly: true);
            _regex = new Regex(@"\{[^}]*\}");
        }

        public void Dispose()
        {
            _app.Quit();
        }

        public IEnumerable<string> GetFields()
        {
            return _regex.Matches(_document.Content.Text)
                .Cast<Match>()
                .Select(m => m.Value)
                .Distinct();
        }

        private string GetNewFileName(string newFileName = null!)
        {
            if (string.IsNullOrWhiteSpace(newFileName))
            {
                return Path.Combine(
                    Path.GetDirectoryName(_file),
                    $"{Path.GetFileNameWithoutExtension(_file)}_compilato.{Path.GetExtension(_file)}"
                );
            }
            else
            {
                return Path.Combine(
                   Path.GetDirectoryName(_file),
                   $"{newFileName}.{Path.GetExtension(_file)}"
               );
            }
        }

        public void ReplaceFields(Dictionary<string, string> fieldReplacements, string newFileName = null!)
        {
            foreach (var fieldReplacement in fieldReplacements)
                FindAndReplace(fieldReplacement.Key, fieldReplacement.Value);

            var text = _document.Content.Text;
            //_document.SaveAs2(_document.Name);
            _document.SaveAs2(GetNewFileName(newFileName));
            _document.Close();
        }

        private void FindAndReplace(object findText, object replaceWithText)
        {
            //options
            object matchCase = false;
            object matchWholeWord = true;
            object matchWildCards = false;
            object matchSoundsLike = false;
            object matchAllWordForms = false;
            object forward = true;
            object format = false;
            object matchKashida = false;
            object matchDiacritics = false;
            object matchAlefHamza = false;
            object matchControl = false;
            object read_only = false;
            object visible = true;
            object replace = 2;
            object wrap = 1;
            //execute find and replace
            _document.Content.Find.Execute(ref findText, ref matchCase, ref matchWholeWord,
                ref matchWildCards, ref matchSoundsLike, ref matchAllWordForms, ref forward, ref wrap, ref format, ref replaceWithText, ref replace,
                ref matchKashida, ref matchDiacritics, ref matchAlefHamza, ref matchControl);
        }
    }
}
