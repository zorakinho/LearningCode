using System;
using System.IO;
using OfficeOpenXml;

class Program
{
    static void Main(string[] args)
    {
        string folderPath = @"C:\Users\robert.alves\source\aulas\AulasEuCodo\ReadFiles\imobfile";

        using (var package = new ExcelPackage(new FileInfo(@"C:\Users\robert.alves\source\aulas\AulasEuCodo\ReadFiles\dadostoemail.xlsx")))
        {
            ExcelWorksheet worksheet = package.Workbook.Worksheets[0]; // Assume que a planilha está na primeira guia

            int rowCount = worksheet.Dimension.Rows;

            foreach (var filePath in Directory.GetFiles(folderPath))
            {
                string fileName = Path.GetFileNameWithoutExtension(filePath);
                bool foundMatch = false;

                Console.Write($"Unidade: {fileName}, ");

                for (int row = 2; row <= rowCount; row++)
                {
                    string unidade = worksheet.Cells[row, 2].Value?.ToString();

                    if (fileName == unidade)
                    {
                        string bloco = worksheet.Cells[row, 1].Value?.ToString();
                        string corretor = worksheet.Cells[row, 3].Value?.ToString();
                        string cliente = worksheet.Cells[row, 4].Value?.ToString();
                        string gerente = worksheet.Cells[row, 5].Value?.ToString();

                        Console.WriteLine($"Bloco: {bloco}, Unidade: {unidade}, Corretor: {corretor}, Cliente: {cliente}, Gerente: {gerente}");

                        foundMatch = true;
                        break; // Não é necessário continuar a busca se encontramos um correspondente
                    }
                }

                if (!foundMatch)
                {
                    Console.WriteLine("Arquivo não vinculado");
                }
            }
        }
    }
}
