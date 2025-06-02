// See https://aka.ms/new-console-template for more information
using ClosedXML.Excel;

Console.WriteLine("Veuillez entrer le chemin absolu du fichier Excel (.xlsx) :");
var excelPath = Console.ReadLine();

if (string.IsNullOrWhiteSpace(excelPath) || !File.Exists(excelPath))
{
    Console.WriteLine("Chemin invalide ou fichier introuvable.");
    return;
}

using var workbook = new XLWorkbook(excelPath);
foreach (var worksheet in workbook.Worksheets)
{
    var csvPath = Path.Combine(Path.GetDirectoryName(excelPath)!,
        $"{Path.GetFileNameWithoutExtension(excelPath)}_{worksheet.Name}.csv");
    using var writer = new StreamWriter(csvPath, false, System.Text.Encoding.UTF8);
    foreach (var row in worksheet.RowsUsed())
    {
        var values = row.CellsUsed().Select(cell =>
        {
            var val = cell.GetValue<string>() ?? string.Empty;
            if (val.Contains('"')) val = val.Replace("\"", "\"\"");
            if (val.Contains(';') || val.Contains('"') || val.Contains('\n'))
                val = $"\"{val}\"";
            return val;
        });
        writer.WriteLine(string.Join(";", values));
    }
    Console.WriteLine($"Feuille '{worksheet.Name}' exportée vers : {csvPath}");
}

