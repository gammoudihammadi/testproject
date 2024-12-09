using iText.Kernel.Pdf.Canvas.Parser;
using iText.Kernel.Pdf;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Content;
using iText.Kernel.Pdf.Canvas.Parser.Data;
using iText.Kernel.Pdf.Canvas.Parser.Listener;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Shared
{
    public class PdfTableExtractor
    {
        public List<Dictionary<string, string>> ExtractTableData(string pdfFilePath, List<string> columns)
        {
            List<Dictionary<string, string>> tableData = new List<Dictionary<string, string>>();

            using (var document = UglyToad.PdfPig.PdfDocument.Open(pdfFilePath))
            {
                foreach (var page in document.GetPages())
                {
                    var words = page.GetWords().ToList();
                    List<string> columnHeaders = new List<string>();
                    bool isTableStarted = false;

                    foreach (var word in words)
                    {
                        if (isTableStarted)
                        {
                            if (columnHeaders.Contains(word.Text))
                            {
                                var row = ExtractRow(words, word, columnHeaders);
                                if (row.Count > 0)
                                {
                                    tableData.Add(row);
                                }
                            }
                        }
                        else if (columns.Contains(word.Text))
                        {
                            columnHeaders.Add(word.Text);
                            if (columnHeaders.Count == columns.Count)
                            {
                                isTableStarted = true;
                            }
                        }
                    }
                }
            }

            return tableData;
        }

        private Dictionary<string, string> ExtractRow(List<Word> words, Word startWord, List<string> columnHeaders)
        {
            Dictionary<string, string> rowData = new Dictionary<string, string>();
            int startIndex = words.IndexOf(startWord);
            int columnCount = columnHeaders.Count;

            for (int i = 0; i < columnCount; i++)
            {
                if (startIndex + i < words.Count)
                {
                    string columnHeader = columnHeaders[i];
                    string cellValue = words[startIndex + i].Text;
                    rowData[columnHeader] = cellValue;
                }
            }

            return rowData;
        }
        public void ExtractDataFromPdf(string pdfFilePath)
        {
            using (var document = UglyToad.PdfPig.PdfDocument.Open(pdfFilePath))
            {
                foreach (var page in document.GetPages())
                {
                    var words = page.GetWords().ToList();

                    // Example: Extract specific data from the words list
                    var flightDate = GetWordAfter(words, "Flight date");
                    var flightNumber = GetWordAfter(words, "Flight number");
                    var etd = GetWordAfter(words, "ETD");
                    var agreementNumber = GetWordAfter(words, "Agreement number");

                    // Add logic to handle other fields...

                    Console.WriteLine($"Flight Date: {flightDate}");
                    Console.WriteLine($"Flight Number: {flightNumber}");
                    Console.WriteLine($"ETD: {etd}");
                    Console.WriteLine($"Agreement Number: {agreementNumber}");
                }
            }
        }

        private string GetWordAfter(List<Word> words, string key)
        {
            for (int i = 0; i < words.Count; i++)
            {
                if (words[i].Text.Equals(key, StringComparison.OrdinalIgnoreCase))
                {
                    return (i + 1 < words.Count) ? words[i + 1].Text : string.Empty;
                }
            }
            return string.Empty;
        }
        public List<(string Text, float X, float Y)> ExtractTextWithPositions(string pdfPath)
        {
            var result = new List<(string Text, float X, float Y)>();

            using (PdfReader pdfReader = new PdfReader(pdfPath))
            using (iText.Kernel.Pdf.PdfDocument pdfDocument = new iText.Kernel.Pdf.PdfDocument(pdfReader))
            {
                for (int i = 1; i <= pdfDocument.GetNumberOfPages(); i++)
                {
                    PdfPage page = pdfDocument.GetPage(i);
                    var strategy = new SimpleTextExtractionStrategyWithPositions();
                    string textOnPage = PdfTextExtractor.GetTextFromPage(page, strategy);
                    List<List<string>> table1 = ConvertTextToTable(textOnPage);
                    List<List<TableCell>> table = ConvertTextToTableWithPositions(textOnPage);

                    // Afficher le tableau résultant
                    foreach (var row in table)
                    {
                        var rowNew = string.Join(" ", row);
                    }

                    foreach (var textChunk in strategy.TextChunks)
                    {
                        result.Add((textChunk.Text, textChunk.Position.Get(0), textChunk.Position.Get(1)));
                    }
                }
            }

            return result;
        }
        public class SimpleTextExtractionStrategyWithPositions : LocationTextExtractionStrategy
        {
            public List<(string Text, iText.Kernel.Geom.Vector Position)> TextChunks { get; } = new List<(string Text, iText.Kernel.Geom.Vector Position)>();

            public override void EventOccurred(IEventData data, EventType type)
            {
                base.EventOccurred(data, type);

                if (type == EventType.RENDER_TEXT)
                {
                    TextRenderInfo renderInfo = (TextRenderInfo)data;
                    TextChunks.Add((renderInfo.GetText(), renderInfo.GetDescentLine().GetStartPoint()));
                }
            }
        }
        public List<List<string>> ConvertTextToTable(string text)
        {
            List<List<string>> table = new List<List<string>>();

            // Séparer le texte en lignes
            string[] lines = text.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.RemoveEmptyEntries);

            foreach (string line in lines)
            {
                // Séparer chaque ligne en colonnes (en supposant que les colonnes sont séparées par des espaces)
                string[] columns = line.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);

                // Ajouter les colonnes à la table
                table.Add(new List<string>(columns));
            }

            return table;
        }
        public class TableCell
        {
            public string Text { get; set; }
            public float X { get; set; }
            public float Y { get; set; }

            public TableCell(string text, float x, float y)
            {
                Text = text;
                X = x;
                Y = y;
            }
        }
        public List<List<TableCell>> ConvertTextToTableWithPositions(string text)
        {
            List<List<TableCell>> table = new List<List<TableCell>>();
            float y = 0;
            float x = 0;
            // Séparer le texte en lignes
            string[] lines = text.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.RemoveEmptyEntries);

            foreach (string line in lines)
            {
                x = 0;
                // Séparer chaque ligne en colonnes (en supposant que les colonnes sont séparées par des espaces)
                string[] columns = line.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);

                List<TableCell> rowCells = new List<TableCell>();
                foreach (string column in columns)
                {
                    // Ici vous devez déterminer comment obtenir les positions X et Y pour chaque cellule
                    //x = 0; // Remplacez par la valeur correcte
                    //y = 0; // Remplacez par la valeur correcte

                    TableCell cell = new TableCell(column, x, y);
                    rowCells.Add(cell);
                    x = x+1;
                }


                // Ajouter les cellules de la ligne à la table
                table.Add(rowCells);
                y = y+1;
            }

            return table;
        }
    }
}
