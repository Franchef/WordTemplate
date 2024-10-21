using System.Text.RegularExpressions;

using Microsoft.Office.Interop.Word;

namespace WordTemplateDomain
{

    public class WordTemplateBinder
    {
        private readonly Application _app;
        private readonly Document _document;
        private readonly Regex _regex;
        private readonly string _newFileName;
        public WordTemplateBinder(string filePath)
        {
            _app = new Application();
            _newFileName = Path.Combine(
                Path.GetDirectoryName(filePath),
                $"{Path.GetFileNameWithoutExtension(filePath)}_compilato.{Path.GetExtension(filePath)}"
            );
            _document = _app.Documents.Open(filePath, ReadOnly: true);
            _regex = new Regex(@"\{[^}]*\}");
        }

        public IEnumerable<string> GetFields()
        {
            foreach (Match match in _regex.Matches(_document.Content.Text))
            {
                yield return match.Value;
            }
        }

        public void ReplaceFields(Dictionary<string, string> fieldReplacements)
        {
            foreach (var fieldReplacement in fieldReplacements)
                FindAndReplace(fieldReplacement.Key, fieldReplacement.Value);

            var text = _document.Content.Text;
            //_document.SaveAs2(_document.Name);
            _document.SaveAs2(_newFileName);
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
