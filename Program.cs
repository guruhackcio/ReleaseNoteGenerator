using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;
using System;
using System.IO;

class Program
{
    static void Main()
    {
        string mainPath = "E:\\Visual Studio Proj\\ReleaseNoteGenerator\\";
        string excelPath = $"{mainPath}ReleaseNote Generator.xlsx"; // Replace with the actual path to your .xlsx file
        string outputPath = $"{mainPath}\\_software\\";
        string sheet = "Sheet1";

        string codeLangsColumn = "A";
        string tlContentsColumn = "B";
        
        List<string> codeLangs = new List<string>();
        List<string> tlContents = new List<string>();
        List<string> fileNames = new List<string>();

        try
        {
            string resourceType = ReadCellValue(excelPath, sheet, "C2");
            string swVersion = ReadCellValue(excelPath, sheet, "D2");
            string hwVersion = ReadCellValue(excelPath, sheet, "E2");
            if (string.IsNullOrWhiteSpace(resourceType))
            {
                Console.WriteLine("Resource Type is empty, aborting operation...");
                return;
            }

            if (string.IsNullOrWhiteSpace(swVersion))
            {
                Console.WriteLine("Software Version is empty, aborting operation...");
                return;
            }

            if (string.IsNullOrWhiteSpace(hwVersion))
            {
                Console.WriteLine("Hardware Version is empty, aborting operation...");
                return;
            }

            string subFolder = swVersion.Substring(0, 3);

            codeLangs = ReadIterateCellValue(excelPath, sheet, codeLangsColumn, 1);
            tlContents = ReadIterateCellValue(excelPath, sheet, tlContentsColumn, 1);

            for (int i=0; i < codeLangs.Count; i++)
            {
                if (string.Equals(resourceType, "Snape", StringComparison.OrdinalIgnoreCase) || string.Equals(resourceType, "Mesh", StringComparison.OrdinalIgnoreCase))
                {
                    fileNames.Add($"ReleaseNotes-v{swVersion}-H{hwVersion}.{codeLangs[i]}.txt");
                }
                else if (string.Equals(resourceType, "Beam", StringComparison.OrdinalIgnoreCase))
                {
                    fileNames.Add($"ReleaseNotes-v{swVersion}.{codeLangs[i]}.txt");
                }
                else
                {
                    Console.WriteLine($"Invalid Resource Type name : {resourceType}");
                    return;
                }
            }

            for (int i = 0; i< fileNames.Count; i++)
            {
                GenerateTxtFile(fileNames[i], tlContents[i], outputPath, resourceType, subFolder);
            }

            Console.WriteLine("Text file generation complete.");
            return;
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex.ToString());
        }
    }

    static List<string> ReadIterateCellValue(string filePath, string sheetName, string column, int startingRow)
    {
        List<string> cells = new List<string>();
        using (var package = new OfficeOpenXml.ExcelPackage(new FileInfo(filePath)))
        {
            var worksheet = package.Workbook.Worksheets[sheetName];
            int index = 0;
            while (true)
            {
                string cellCode = $"{column}{startingRow}";
                string cellValue = worksheet.Cells[cellCode].Value?.ToString();
                if (string.IsNullOrEmpty(cellValue))
                    break;

                cells.Add(cellValue);
                index++;
                startingRow++;
            }

            return cells;
        }
    }

    static string ReadCellValue(string filePath, string sheetName, string cellAddress)
    {
        // Using EPPlus ver 4 with LGPL license to read excel values.

        using (var package = new OfficeOpenXml.ExcelPackage(new FileInfo(filePath)))
        {
            var worksheet = package.Workbook.Worksheets[sheetName];
            var cell = worksheet.Cells[cellAddress];
            return cell.Value?.ToString();
        }
    }

    static void GenerateTxtFile(string fileName, string content, string outputPath, string resourceType, string subFolder)
    {
        string folderPath = Path.Combine(outputPath, resourceType.ToLower(), subFolder);
        Directory.CreateDirectory(folderPath);

        string filePath = Path.Combine(folderPath, fileName);
        using (StreamWriter writer = File.CreateText(filePath))
        {
            writer.WriteLine(content);
        }
    }

    public void UploadToCloud(string fileName, string localPath)
    {
        string connectionString = "DefaultEndpointsProtocol=https;AccountName=nodesoftwarenotification;AccountKey=/WX4wcR9FcN6T3mKQAE5s8ft+Nx2ShJVLjPYc9iudi2hBsV8RG8to9syA2eAa37uDd0QbbOk2R7w+AStUvcZiA==;EndpointSuffix=core.windows.net";
        string containerName = "container1";

    }
}