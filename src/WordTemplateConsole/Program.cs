using Spectre.Console;
using WordTemplateDomain;

internal class Program
{
    private static void Main(string[] args)
    {
        args = new string[] { @"D:\Git\Franchef\WordTemplate\test\WordTemplateTests\TestTemplate.docx" };
        AnsiConsole.MarkupLine("[dodgerblue2]Word template utility[/]");
        if (args.Length > 0 && CanOpenFile(args[0]))
        {
            var wtb = new WordTemplateBinder(args[0]);
            var fields = wtb.GetFields();
            AnsiConsole.MarkupLine($"[chartreuse3]Numero di campi trovati: {fields.Count()}[/]");
            foreach (var field in fields)
            {
                AnsiConsole.MarkupLine($"[chartreuse3]{field}[/]");
            }
        }
        else
            AnsiConsole.MarkupLine("[orange3]Nessun file valido passato[/]");
    }

    private static bool CanOpenFile(string filePath)
    {
        if (!File.Exists(filePath))
        {
            AnsiConsole.MarkupLine($"[orange3]Il file {filePath} non esiste[/]");
            return false;
        }
        if (!Path.GetExtension(filePath).Equals(".doc", StringComparison.InvariantCultureIgnoreCase)  && !Path.GetExtension(filePath).Equals(".docx", StringComparison.InvariantCultureIgnoreCase))
        {
            AnsiConsole.MarkupLine($"[orange3]Estensione {Path.GetExtension(filePath)} non valida[/]");
            return false;
        }
        return true;
    }
}