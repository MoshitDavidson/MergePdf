using Microsoft.Extensions.Configuration;
using System;
using System.IO;

namespace MergePdf
{
    class Program
    {
        static IConfigurationRoot config;
        static void Main(string[] args)
        {
            var builder = new ConfigurationBuilder()
                .SetBasePath(Directory.GetCurrentDirectory())
                .AddJsonFile("appsettings.json", optional: false);

            config = builder.Build();
            Console.WriteLine("Hi Liran!");

            Console.WriteLine("Which fund do you need?");
            int fundNum = int.Parse(Console.ReadLine());

            Console.WriteLine("which quarter of the year? (1,2,3 or 4), and press enter");
            int quarter = int.Parse(Console.ReadLine());

            Console.WriteLine("Please insert the number of the rows in excel file, and press enter");
            int investorsCount = int.Parse(Console.ReadLine());

            string excelFolder = GetValueFromConfiguration("ExcelFolder");
            ParseExcelHandler parseExcelHandler = new(excelFolder);
            Console.WriteLine("Please insert the excel file name and press enter");
            string excelFileName = Console.ReadLine();
            var dic = parseExcelHandler.ParseExcelFile(excelFileName, investorsCount);
            
            string outputFolderPdf  = GetValueFromConfiguration("OutputFolderPdf");
            string inputFolderPdf = GetValueFromConfiguration("InputFolderPdf");

            MergePdfHandler mergePdfHandler = new(outputFolderPdf, inputFolderPdf);
            
            string pdf0 = GetValueFromConfiguration("Pdf_0");
            foreach (var item in mergePdfHandler.CombineTwoPdfs(pdf0, dic, fundNum, quarter))
            {
                Console.WriteLine($"File {item} created and saved!");
            }
            Console.WriteLine($"Finish all pdf!! press any key to close");
            Console.ReadLine();
        }

        static string GetValueFromConfiguration(string key)
        {
            return config.GetSection(key).Value;
        }
    }
}


