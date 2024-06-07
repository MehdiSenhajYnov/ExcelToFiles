// See https://aka.ms/new-console-template for more information

using System.Security;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.SharePoint.Client;
using File = System.IO.File;

namespace Lang
{
    static class Program
    {

        // User Configuration
        static string excelFilePath = @"C:\Users\mehdi\Desktop\test.xlsx";
        static string _outputFolder = @"C:\Users\mehdi\Desktop\Travaill\Travaill\SepapayAdmin\SepapayAdminFront\src\Lang";
        static string _wantedExcelPageName = "Code";
        
        
        
        enum columnType
        {
            LanguageHelper,
            French,
            FrenchBelgium,
            English,
            Dutch,
            German,
            Italian,
            Spanish,
            Romanian,
            Polish,
            Russian,
            Swedish,
            Greek,
            Littuanian,
            Serbian

        }

        // Program Configuration*
        static Dictionary<columnType, string> typeContent = new Dictionary<columnType, string>();
        static List<string?> _allfilesnames = Directory.GetFiles(_outputFolder, "*.ts").Select(Path.GetFileName).ToList();

        static void Main(string[] args)
        {
            LoadTypeContent();
            
            PrepareFiles();
            ProcessTypeContent();
            EndFiles();
        }
        
        private const string abstractPrepare = "export abstract class LanguageHelper {\n"; 
        private const string overridePrepare = "import { LanguageHelper } from './constantsText.LanguageHelper';\n\nexport class $1 extends LanguageHelper {\n";
        private static void PrepareFiles()
        {
            foreach (string? filename in _allfilesnames)
            {
                if (filename == null) continue;
                if (!filename.StartsWith("constantsText.")) continue;
                string lang = filename.Replace("constantsText.", "").Replace(".ts", "");
                if (lang == "LanguageHelper")
                {
                    File.WriteAllText(Path.Combine(_outputFolder, filename), abstractPrepare);
                }
                else
                {
                    string langPrepare = overridePrepare.Replace("$1", "Language"+lang);
                    File.WriteAllText(Path.Combine(_outputFolder, filename), langPrepare);
                }
            }            
        }

        private static void ProcessTypeContent()
        {
            foreach (var (keyType, valueContent) in typeContent)
            {
                    
                foreach (string? filename in _allfilesnames)
                {
                    if (filename == null) continue;
                    if (!filename.StartsWith("constantsText.")) continue;
                    string lang = filename.Replace("constantsText.", "").Replace(".ts", "");
                    if (keyType.ToString() == lang)
                    {
                        File.AppendAllText(Path.Combine(_outputFolder, filename),valueContent);
                    }
                }
                        
            }
        }
        
        private static void EndFiles()
        {
            foreach (string? filename in _allfilesnames)
            {
                if (filename == null) continue;
                if (!filename.StartsWith("constantsText.")) continue;
                File.AppendAllText(Path.Combine(_outputFolder, filename), "}");
            }   
        }


        private static void LoadTypeContent()
        {
            
            int typeCounter = 0;
            if (!File.Exists(excelFilePath))
            {
                Console.WriteLine("File not found");
                return;
            }

            using (SpreadsheetDocument doc = SpreadsheetDocument.Open(excelFilePath, false))
            {
                WorkbookPart workbookPart = doc.WorkbookPart;
                if (!workbookPart.Workbook.Descendants<Sheet>().Any())
                {
                    Console.WriteLine("No sheet found in the workbook");
                    return;
                }

                foreach (Sheet sheet in workbookPart.Workbook.Descendants<Sheet>())
                {
                    if (sheet.Name != _wantedExcelPageName)
                    {
                        continue;
                    }
                    
                    WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);
                    if (!worksheetPart.Worksheet.Elements<SheetData>().Any())
                    {
                        Console.WriteLine("No data found in the sheet");
                        return;
                    }

                    foreach (SheetData sheetData in worksheetPart.Worksheet.Elements<SheetData>())
                    {
                        
                        if (!sheetData.Elements<Row>().Any())
                        {
                            Console.WriteLine("No data found in the sheet");
                            return;
                        }
                        foreach (Row row in sheetData.Elements<Row>())
                        {
                            foreach (Cell cell in row.Elements<Cell>())
                            {
                                string cellValue = cell.CellValue.Text;

                                if (cellValue.StartsWith("abstract"))
                                {
                                    typeCounter = 0;
                                } else if (cellValue.StartsWith("override"))
                                {
                                    typeCounter++;
                                }
                                else
                                {
                                    continue;
                                }
                                columnType type = (columnType) typeCounter;
                                if (!typeContent.ContainsKey(type))
                                {
                                    typeContent.Add(type, "");
                                }
                                typeContent[type] += "    " + cellValue + "\n";
                                
                            }
                        }
                    
                    }
                    
                }
                
            }
        }

        
    }
}
