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
            string prefixFullName = string.Empty;
            string text = "Financial Statements as of";
            switch (fundNum)
            {
                case 1:
                    prefixFullName = "Fund I";
                    text = "combined FS";
                    break;
                case 2:
                    prefixFullName = "IIF II";
                    break;
                case 3:
                    prefixFullName = "IIF III";
                    break;
                case 4:
                    prefixFullName = "IIF IV";
                    break;
            }
            string fileName = $"{prefixFullName}-{investorName}-{text} {dateTime:ddMMyy}";
            return fileName;
        }

        private DateTime GetDateTimeByQuarter(int quarter)
        {
            int currentYear = DateTime.Now.Year;
            int lastYear = DateTime.Now.AddYears(-1).Year;
            DateTime dateTimeResult = quarter switch
            {
                1 => new DateTime(currentYear, 3, DateTime.DaysInMonth(currentYear, 3)),
                2 => new DateTime(currentYear, 6, DateTime.DaysInMonth(currentYear, 6)),
                3 => new DateTime(currentYear, 9, DateTime.DaysInMonth(currentYear, 9)),
                4 => new DateTime(lastYear, 12, DateTime.DaysInMonth(lastYear, 12)),
                _ => DateTime.Now,
            };
            return dateTimeResult;
        }
    }
}
