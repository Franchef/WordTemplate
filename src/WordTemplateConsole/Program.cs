using Spectre.Console;
using WordTemplateDomain;

internal class Program
{
    private static async Task Main(string[] args)
    {
        AnsiConsole.MarkupLine("[dodgerblue2]Word template utility[/]");
        if (args.Length > 0 && CanOpenFile(args[0]))
        {
            var wtb = new WordTemplateBinder(args[0]);
            var fields = wtb.GetFields();
            var replacements = new Dictionary<string, string>();
            AnsiConsole.MarkupLine($"[chartreuse3]Numero di campi trovati: {fields.Count()}[/]");

            var fieldGroups = new FieldGroup();

            foreach (var field in fields)
                fieldGroups.AddField(field);

            foreach (var kv in GetReplacementsRecursive(fieldGroups))
                replacements.Add(kv.Key, kv.Value);


            foreach (var replacement in replacements)
            {
                AnsiConsole.MarkupLine($"[chartreuse3]{replacement.Key}[/] [green]{replacement.Value}[/]");
            }

            bool rivedi = true;
            while (rivedi)
            {
                rivedi = AnsiConsole.Prompt(
                    new TextPrompt<bool>("Modificare qualcosa?")
                        .AddChoice(true)
                        .AddChoice(false)
                        .DefaultValue(true)
                        .WithConverter(choice => choice ? "y" : "n")
                );

                if (rivedi)
                {
                    var field = AnsiConsole.Prompt(
                           new SelectionPrompt<string>()
                            .Title("Seleziona sostituzione da [green]modificare[/]")
                            .PageSize(10)
                            .MoreChoicesText("[grey](Freccia su o freccia giù per scorrere([/]")
                            .AddChoices(fields)
                        );
                    var value = AnsiConsole.Prompt(new TextPrompt<string>($"[chartreuse3]{field}[/] "));
                    replacements[field] = value;
                }
            }
            AnsiConsole.MarkupLine($"[dodgerblue2]Sto creando [/]{wtb.GetNewFileName()}");
            wtb.ReplaceFields(replacements);
            AnsiConsole.MarkupLine($"[dodgerblue2]Fatto[/]");
        }
        else
            AnsiConsole.MarkupLine("[orange3]Nessun file valido passato[/]");
    }

    private static IEnumerable<KeyValuePair<string, string>> GetReplacementsRecursive(FieldGroup fieldGroup)
    {
        AnsiConsole.MarkupLine($"[dodgerblue2]{fieldGroup.Name}[/]");
        foreach (var field in fieldGroup.Fields)
        {
            var value = AnsiConsole.Prompt(new TextPrompt<string>($"[chartreuse3]{field}[/] "));
            yield return new KeyValuePair<string, string>(field, value);
        }
        foreach (var subGroup in fieldGroup.SubGroups)
            foreach (var kev in GetReplacementsRecursive(subGroup))
                yield return kev;
    }

    public record FieldGroup
    {
        public FieldGroup()
        {

        }
        public string Name { get; init; }

        public List<string> Fields { get; init; } = new List<string>();

        public void AddField(string field)
        {
            if (field.Contains("."))
            {
                var groups = field.TrimStart('{').TrimEnd('}').Split('.');
                if (!SubGroups.Any(sg => sg.Name == groups[0]))
                {
                    SubGroups.Add(new FieldGroup { Name = groups[0] });
                }
                var subgroup = SubGroups.Single(sg => sg.Name == groups[0]);
                subgroup.AddSubField(field, groups.Skip(1).ToArray());
            }
            else
            {
                Fields.Add(field);
            }
        }

        protected void AddSubField(string field, params string[] groups)
        {
            if (groups.Length == 1)
            {
                Fields.Add(field);
            }
            else
            {
                if (!SubGroups.Any(sg => sg.Name == groups[0]))
                {
                    SubGroups.Add(new FieldGroup { Name = groups[0] });
                }
                var subgroup = SubGroups.Single(sg => sg.Name == groups[0]);
                subgroup.AddSubField(field, groups.Skip(1).ToArray());
            }
        }

        public List<FieldGroup> SubGroups { get; init; } = new List<FieldGroup>();
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