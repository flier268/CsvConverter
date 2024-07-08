using System.Text;
using MiniExcelLibs;
using MiniExcelLibs.Csv;

namespace XlsxToCsvBase;

public static class Base
{
    public static async Task Run(string[] args, Encoding encoding)
    {
        if (args.Length == 0 || args.Contains("-help", StringComparer.OrdinalIgnoreCase))
        {
            Console.WriteLine("請在命令列輸入一個或多個 Excel 檔案的路徑，或直接將檔案拖曳至執行檔上。");
            Console.WriteLine("Press any key to exit...");
            Console.ReadLine();
            Environment.Exit(0);
        }

        var anyExecute = false;
        string[] supportFormat = [".xlsx", ".xls"];
        foreach (string arg in args)
        {
            if (supportFormat.Contains(Path.GetExtension(arg), StringComparer.OrdinalIgnoreCase))
            {
                Console.WriteLine($"讀取 \"{arg}\"");

                await using var stream = File.Open(arg, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);

                // 讀取Excel並轉換為DataTable
                var dt = MiniExcel.QueryAsDataTable(stream);

                string csvPath = Path.ChangeExtension(arg, "csv");

                // 使用UTF8 BOM寫入csv
                await using var sw = new StreamWriter(csvPath, false, encoding);
                // 遍寫入CSV

                var config = new CsvConfiguration
                {
                    StreamWriterFunc = (x) => new(x, encoding)
                };
                
                await MiniExcel.SaveAsAsync(sw.BaseStream, dt, configuration: config, excelType: ExcelType.CSV);
                
                Console.WriteLine($"已輸出到 {csvPath}, 編碼：{encoding.WebName}");
                anyExecute = true;
            }
        }

        if (!anyExecute)
        {
            Console.WriteLine("請在命令列輸入一個或多個 Excel 檔案的路徑，或直接將檔案拖曳至執行檔上。");
            Console.WriteLine("Press any key to exit...");
            Console.ReadLine();
            Environment.Exit(0);
        }


        if (!args.Contains("-s", StringComparer.OrdinalIgnoreCase))
        {
            Console.WriteLine("Press any key to exit...");
            Console.ReadLine();
        }
    }
}