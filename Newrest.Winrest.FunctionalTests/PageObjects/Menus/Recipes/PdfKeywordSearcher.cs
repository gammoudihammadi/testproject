using System;
using System.IO;
using System.Collections.Generic;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas.Parser;

public class PdfKeywordSearcher
{
    public bool FindAllPdfsWithKeyword(string directoryPath, string guestType)
    {
        List<string> pdfFilesWithKeyword = new List<string>();
        string keyword = $"Type of guest {guestType}";
        foreach (string file in Directory.GetFiles(directoryPath))
        {
            if (file.EndsWith(".pdf"))
            {
                bool found = FindKeywordInPdf(file, keyword);
                if (found)
                {
                    pdfFilesWithKeyword.Add(file);
                }
            }
        }
        if (pdfFilesWithKeyword.Count == 0)
        {
            throw new FileNotFoundException($"No PDF files containing '{keyword}' were found in directory: {directoryPath}");
        }
        return Directory.GetFiles(directoryPath).Length == pdfFilesWithKeyword.Count;
    }
    public bool FindAllPdfsWithKeywordMeal(string directoryPath, string guestType)
    {
        List<string> pdfFilesWithKeyword = new List<string>();
        string keyword = guestType;
        foreach (string file in Directory.GetFiles(directoryPath))
        {
            if (file.EndsWith(".pdf"))
            {
                bool found = FindKeywordInPdf(file, keyword);
                if (found)
                {
                    pdfFilesWithKeyword.Add(file);
                }
            }
        }
        if (pdfFilesWithKeyword.Count == 0)
        {
            throw new FileNotFoundException($"No PDF files containing '{keyword}' were found in directory: {directoryPath}");
        }
        return Directory.GetFiles(directoryPath).Length == pdfFilesWithKeyword.Count;
    }
    public bool FindKeywordInPdf(string pdfFilePath, string keyword)
    {
        using (PdfReader reader = new PdfReader(pdfFilePath))
        using (PdfDocument pdfDoc = new PdfDocument(reader))
        {
            for (int page = 1; page <= pdfDoc.GetNumberOfPages(); page++)
            {
                string text = PdfTextExtractor.GetTextFromPage(pdfDoc.GetPage(page));
                if (text.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    return true;
                }
            }
        }
        return false;
    }
}