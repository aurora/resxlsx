#!/usr/bin/env dotnet-script
#r "nuget: ClosedXML, 0.104.2"

using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using ClosedXML.Excel;

string[] args = Environment.GetCommandLineArgs().Skip(2).ToArray();

string command = args.Length > 0 ? args[0] : "";
string directory = args.Length > 1 ? args[1] : "";
string excelFile = args.Length > 2 ? args[2] : "";

if (args.Length < 3) {
    Console.WriteLine("Usage: resxlsx [export|import] <resource-path> <excel-file>");
} else if (command == "export")
    ExportToExcel(directory, excelFile);
else if (command == "import")
    ImportFromExcel(excelFile, directory);
else
    Console.WriteLine($"Unknown command '{command}'. Use 'export' or 'import'.");

static void ExportToExcel(string resxDirectory, string excelFile)
{
    var files = Directory.GetFiles(resxDirectory, "*.resx");
    var translations = new Dictionary<string, Dictionary<string, string>>();
    var languages = new HashSet<string> { "Neutral" };
    
    foreach (var file in files)
    {
        string lang = GetLanguageFromFilename(file);
        languages.Add(lang);        
        var doc = XDocument.Load(file);
        
        foreach (var data in doc.Descendants("data"))
        {
            string key = data.Attribute("name")?.Value;
            string value = data.Element("value")?.Value;
            if (key == null) continue;
            
            if (!translations.ContainsKey(key))
                translations[key] = new Dictionary<string, string>();
            
            translations[key][lang] = value;
        }
    }
    
    var orderedLanguages = languages.OrderBy(x => x == "Neutral" ? 0 : 1).ThenBy(x => x);

    var workbook = new XLWorkbook();
    var sheet = workbook.Worksheets.Add("Translations");
    sheet.Cell(1, 1).Value = "Key";
    int col = 2;
    foreach (var lang in orderedLanguages)
        sheet.Cell(1, col++).Value = lang;
    
    int row = 2;
    foreach (var key in translations.Keys.OrderBy(x => x))
    {
        sheet.Cell(row, 1).Value = key;
        col = 2;
        foreach (var lang in orderedLanguages)
        {
            translations[key].TryGetValue(lang, out string value);
            sheet.Cell(row, col++).Value = value ?? "";
        }
        row++;
    }
    
    workbook.SaveAs(excelFile);
    Console.WriteLine("Export completed.");
}

static void ImportFromExcel(string excelFile, string resxDirectory)
{
    var workbook = new XLWorkbook(excelFile);
    var sheet = workbook.Worksheet("Translations");
    var headers = sheet.Row(1).CellsUsed().Select(c => c.Value.ToString()).ToList();
    var translations = new Dictionary<string, Dictionary<string, string>>();
    
    foreach (var row in sheet.RowsUsed().Skip(1))
    {
        string key = row.Cell(1).GetString();
        if (!translations.ContainsKey(key))
            translations[key] = new Dictionary<string, string>();
        
        for (int i = 1; i < headers.Count; i++)
        {
            string lang = headers[i];
            string value = row.Cell(i + 1).GetString();
            translations[key][lang] = value;
        }
    }
    
    foreach (var lang in headers.Skip(1))
    {
        var fileName = Path.Combine(resxDirectory, lang == "Neutral" ? "Messages.resx" : $"Messages.{lang}.resx");
        var doc = new XDocument(new XElement("root"));
        
        foreach (var entry in translations)
        {
            if (!entry.Value.TryGetValue(lang, out string value)) continue;
            doc.Root.Add(new XElement("data", new XAttribute("name", entry.Key), new XElement("value", value)));
        }
        
        doc.Save(fileName);
    }
    
    Console.WriteLine("Import completed.");
}

static string GetLanguageFromFilename(string filename)
{
    var name = Path.GetFileNameWithoutExtension(filename);
    var parts = name.Split('.');
    return parts.Length > 1 ? parts.Last() : "Neutral";
}
