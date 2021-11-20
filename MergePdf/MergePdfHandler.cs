using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using System;
using System.Collections.Generic;

namespace MergePdf
{
    public class MergePdfHandler
    {
        private string _outputFolder;
        private string _inputFolder;
        public MergePdfHandler(string outputFolder, string inputFolder)
        {
            _outputFolder = outputFolder;
            _inputFolder = inputFolder;
        }
        public IEnumerable<string> CombineTwoPdfs(string firstPdf, Dictionary<int, List<string>> investorsNames, int fundNum, int quarter)
        {
            foreach (var item in investorsNames)
            {
                foreach (var name in item.Value)
                {
                    PdfDocument outputPDFDocument = new();
                    CombinePdfFileToOutputPdfFile(outputPDFDocument, firstPdf);
                    CombinePdfFileToOutputPdfFile(outputPDFDocument, $"{_inputFolder}\\{item.Key}.pdf");
                    string fullNameToSave = GenerateFullCombinedFileName(quarter, name, fundNum);
                    outputPDFDocument.Save($"{_outputFolder}\\{fullNameToSave}.pdf");
                    yield return fullNameToSave;

                }
            }
        }

        private void CombinePdfFileToOutputPdfFile(PdfDocument outputPDFDocument, string pdfFile)
        {
            PdfDocument inputPDFDocument = PdfReader.Open(pdfFile, PdfDocumentOpenMode.Import);
            outputPDFDocument.Version = inputPDFDocument.Version;
            foreach (PdfPage page in inputPDFDocument.Pages)
                outputPDFDocument.AddPage(page);
        }

        private string GenerateFullCombinedFileName(int quarter, string investorName, int fundNum)
        {
            DateTime dateTime = GetDateTimeByQuarter(quarter);
            string fundNumStr = string.Empty;
            string text = "Financial Statements as of";
            switch (fundNum)
            {
                case 1:
                    fundNumStr = "I";
                    text = "combined FS";
                    //fileName = $"Fund I-{investorName}-combined FS {dateTime:ddMMyy}";
                    break;
                case 2:
                    fundNumStr = "II";
                    //fileName = $"IIF II-{investorName}-Financial Statements as of {dateTime:ddMMyy}";
                    break;
                case 3:
                    fundNumStr = "III";
                    //fileName = $"IIF III-{investorName}-Financial Statements as of {dateTime:ddMMyy}";
                    break;
                case 4:
                    fundNumStr = "IV";
                    //fileName = $"IIF IV-{investorName}-Financial Statements as of {dateTime:ddMMyy}";
                    break;
            }
            string fileName = $"IIF {fundNumStr}-{investorName}-{text} {dateTime:ddMMyy}";
            return fileName;
        }

        private DateTime GetDateTimeByQuarter(int quarter) =>
        quarter switch
        {
            1 => new DateTime(DateTime.Now.Year, 3, DateTime.DaysInMonth(DateTime.Now.Year, 3)),
            2 => new DateTime(DateTime.Now.Year, 6, DateTime.DaysInMonth(DateTime.Now.Year, 6)),
            3 => new DateTime(DateTime.Now.Year, 9, DateTime.DaysInMonth(DateTime.Now.Year, 9)),
            4 => new DateTime(DateTime.Now.Year, 12, DateTime.DaysInMonth(DateTime.Now.Year, 12)),
        };
    }
}
