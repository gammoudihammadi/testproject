using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Shared
{
    public class OpenXmlExcel
    {
        //private static object worksheet;
        //public interface Alignment;
        //public static object SpreadsheetVerticalAlignment { get; private set; }

        public static bool ReadAllDataInColumn(string columnName, string sheetName, string fileName, string value, bool premiereLigne=false)
        {
            bool valueBool = true;
            using (FileStream fs = new FileStream(fileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                // Ouverture du document à exploiter
                using (SpreadsheetDocument document = SpreadsheetDocument.Open(fs, false))
                {

                    WorkbookPart workbookPart = document.WorkbookPart;
                    IEnumerable<Sheet> sheets = workbookPart.Workbook.Descendants<Sheet>().Where(s => s.Name == sheetName);
                    if (sheets.Count() == 0)
                    {
                        // La sheet n'existe pas
                        return false;
                    }

                    WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheets.First().Id);
                    SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();

                    int rowNumber = sheetData.Elements<Row>().Count();
                    if (premiereLigne)
                    {
                        // seulement sur la première ligne
                        rowNumber = 2;
                    }
                    int columnNumber = 0;

                    // Si le nnom du fichier contient "CyclesCalendar", on ne prend pas la première ligne
                    Row row = fileName.Contains("CyclesCalendar") ? sheetData.Elements<Row>().Reverse().Skip(1).First() : sheetData.Elements<Row>().First();

                    foreach (Cell c in row)
                    {
                        if (c.DataType != null && c.DataType == CellValues.SharedString)
                        {
                            var stringId = Convert.ToInt32(c.InnerText);
                            if (workbookPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(stringId).InnerText.Equals(columnName))
                            {
                                // Récupération du numéro de la colonne contenant l'information recherchée
                                columnNumber = stringId + 1;
                                break;
                            }
                        }
                    }

                    if (columnNumber != 0)
                    {
                        int startIndex = GetStartIndexForReadData(fileName);

                        // On récupère la lettre correspondant au numéro de la colonne de la donnée récherchée
                        string columnLetter = GetColumnName(columnNumber);

                        // Lecture des cellules de chaque ligne à partir de la quatrième ligne (pour n'avoir que les données)
                        for (int i = startIndex; i <= rowNumber; i++)
                        {
                            // Construction du nom de la cellule Excel et récupoération de la ligne en cours
                            string cellName = columnLetter + i;
                            IEnumerable<Row> rows = sheetData.Elements<Row>().Where(r => r.RowIndex == i);

                            // Récupération de la cellule à lire
                            Cell cell = rows.First().Elements<Cell>().Where(c => string.Compare(c.CellReference.Value, cellName, true) == 0).FirstOrDefault();
                            if (cell != null && cell.DataType != null && cell.DataType == CellValues.SharedString)
                            {
                                var stringId = Convert.ToInt32(cell.InnerText);
                                var contenu = workbookPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(stringId).InnerText;
                                if (!contenu.Contains(value))
                                {
                                    valueBool = false;
                                    break;
                                }
                            }
                        }
                    }
                    else
                    {
                        valueBool = false;
                    }

                    // On ferme la ressource
                    document.Close();
                }
                fs.Close();
            }
            return valueBool;
        }

        public static int GetExportResultNumber(string sheetName, string fileName)
        {
            int resultNumber = 0;

            using (FileStream fs = new FileStream(fileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                // Ouverture du document à exploiter
                using (SpreadsheetDocument document = SpreadsheetDocument.Open(fs, false))
                {

                    WorkbookPart workbookPart = document.WorkbookPart;
                    IEnumerable<Sheet> sheets = workbookPart.Workbook.Descendants<Sheet>().Where(s => s.Name == sheetName);
                    if (sheets.Count() == 0)
                    {
                        // La sheet n'existe pas
                        return 0;
                    }

                    WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheets.First().Id);
                    SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();

                    // On récupère le nombre de lignes actives de la sheet et la première ligne
                    resultNumber = sheetData.Elements<Row>().Count() - 1;

                    // On ferme la ressource
                    document.Close();
                }
                fs.Close();
            }

            return resultNumber;
        }

        public static List<string> GetValuesInList(string columnName, string sheetName, string fileName, string decimalSeparator = ",")
        {
            
            List<string> resultList = new List<string>();
            using (FileStream fs = new FileStream(fileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                // Ouverture du document à exploiter
                using (SpreadsheetDocument document = SpreadsheetDocument.Open(fs, false))
                {

                    WorkbookPart workbookPart = document.WorkbookPart;
                    IEnumerable<Sheet> sheets = workbookPart.Workbook.Descendants<Sheet>().Where(s => s.Name == sheetName);
                    if (sheets.Count() == 0)
                    {
                        // La sheet n'existe pas
                        return resultList;
                    }

                    WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheets.First().Id);
                    SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();

                    // On récupère le nombre de lignes actives de la sheet et la première ligne
                    int rowNumber = sheetData.Elements<Row>().Count();
                    Row row = sheetData.Elements<Row>().First();
                    int columnNumber = 0;

                    // Récupération de la colonne dont le contenu correspond à celui en pramètre
                    int offset = 0;
                    foreach (Cell c in row)
                    {
                        offset++;
                        if (c.DataType != null && c.DataType == CellValues.SharedString)
                        {
                            var stringId = Convert.ToInt32(c.InnerText);
                            if (workbookPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(stringId).InnerText.Equals(columnName))
                            {
                                // Récupération du numéro de la colonne contenant l'information recherchée
                                columnNumber = offset;// stringId + 1;
                                break;
                            }
                        }
                    }

                    if (columnNumber != 0)
                    {
                        // On récupère la lettre correspondant au numéro de la colonne de la donnée récherchée
                        string columnLetter = GetColumnName(columnNumber);

                        // Lecture des cellules de chaque ligne à partir de la seconde ligne (pour n'avoir que les données)
                        for (int i = 2; i <= rowNumber; i++)
                        {
                            // Construction du nom de la cellule Excel et récupoération de la ligne en cours
                            string cellName = columnLetter + i;
                            IEnumerable<Row> rows = sheetData.Elements<Row>().Where(r => r.RowIndex == i);

                            // Récupération de la cellule à lire
                            Cell cell = rows.First().Elements<Cell>().Where(c => string.Compare(c.CellReference.Value, cellName, true) == 0).FirstOrDefault();

                            if (cell == null)
                            {
                                resultList.Add("");
                            }
                            else if (cell.DataType != null && cell.DataType == CellValues.SharedString)
                            {
                                var stringId = Convert.ToInt32(cell.InnerText);
                                var contenu = workbookPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(stringId).InnerText;
                                if (contenu == null)
                                {
                                    resultList.Add(null);
                                }
                                else
                                {
                                    resultList.Add(contenu);
                                }
                            }
                            else if (cell.DataType != null && cell.DataType == CellValues.Boolean)
                            {
                                if (cell.CellValue?.Text == "0")
                                {
                                    resultList.Add("FALSE");
                                }
                                else
                                {
                                    resultList.Add("TRUE");
                                }
                            }
                            else if (cell.DataType == null)
                            {
                                var contenu = cell.CellValue?.Text;
                                if (contenu == null)
                                {
                                    resultList.Add(null);
                                }
                                else
                                {
                                    Console.WriteLine("#### decimalSeparator : " + decimalSeparator);
                                    
                                    CultureInfo ci = decimalSeparator.Equals(",") ? new CultureInfo("fr-FR") : new CultureInfo("en-US");
                                    Console.WriteLine("### contenu " + contenu);
                                    // Excel est anglais ?
                                    double simpleContenu = double.Parse(contenu, new CultureInfo("en-US"));
                                    Console.WriteLine("### simpleContenu " + simpleContenu);
                                    double roundContenu = Math.Round(simpleContenu, 4);
                                    Console.WriteLine("### roundContenu " + roundContenu);
                                    contenu = roundContenu.ToString(ci);
                                    Console.WriteLine("### contenu final " + contenu);
                                    resultList.Add(contenu);
                                }
                            }
                        }
                    }
                    // On ferme la ressource
                    document.Close();
                }
                fs.Close();
            }

            return resultList;
        }

        public static List<string> GetValuesOnListDeliveryRound(string columnName, string sheetName, string fileName, string value)
        {
            List<string> resultList = new List<string>();
            using (FileStream fs = new FileStream(fileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                // Ouverture du document à exploiter
                using (SpreadsheetDocument document = SpreadsheetDocument.Open(fs, false))
                {

                    WorkbookPart workbookPart = document.WorkbookPart;
                    IEnumerable<Sheet> sheets = workbookPart.Workbook.Descendants<Sheet>().Where(s => s.Name == sheetName);
                    if (sheets.Count() == 0)
                    {
                        // La sheet n'existe pas
                        return resultList;
                    }

                    WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheets.First().Id);
                    SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();

                    int rowNumber = sheetData.Elements<Row>().Count();
                    int columnNumber = 0;

                    Row row = sheetData.Elements<Row>().First();
                    foreach (Cell c in row)
                    {
                        if (c.DataType != null && c.DataType == CellValues.SharedString)
                        {
                            var stringId = Convert.ToInt32(c.InnerText);
                            if (workbookPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(stringId).InnerText.Equals(columnName))
                            {
                                // Récupération du numéro de la colonne contenant l'information recherchée
                                columnNumber = stringId + 1;
                                break;
                            }
                        }
                    }

                    if (columnNumber != 0)
                    {
                        // On récupère la lettre correspondant au numéro de la colonne de la donnée récherchée
                        string columnLetter = GetColumnName(columnNumber);

                        // Lecture des cellules de chaque ligne à partir de la quatrième ligne (pour n'avoir que les données)
                        for (int i = 4; i <= rowNumber; i++)
                        {
                            // Construction du nom de la cellule Excel et récupoération de la ligne en cours
                            string cellName = columnLetter + i;
                            IEnumerable<Row> rows = sheetData.Elements<Row>().Where(r => r.RowIndex == i);

                            // Récupération de la cellule à lire
                            Cell cell = rows.First().Elements<Cell>().Where(c => string.Compare(c.CellReference.Value, cellName, true) == 0).FirstOrDefault();
                            if (cell != null && cell.DataType != null && cell.DataType == CellValues.SharedString)
                            {
                                var stringId = Convert.ToInt32(cell.InnerText);
                                var contenu = workbookPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(stringId).InnerText;
                                if (contenu.Equals(value))
                                {
                                    resultList.Add(contenu);
                                }
                            }
                        }
                    }
                    else
                    {
                        return resultList;
                    }

                    // On ferme la ressource
                    document.Close();
                }
                fs.Close();
            }
            return resultList;
        }

        public static void WriteDataInColumn(string columnName, string sheetName, string fileName, string value, CellValues cellValue, bool premiereLigne = false)
        {
            using (FileStream fs = new FileStream(fileName, FileMode.Open, FileAccess.ReadWrite, FileShare.ReadWrite))
            {
                // Ouverture du document à exploiter
                using (SpreadsheetDocument document = SpreadsheetDocument.Open(fs, true))
                {
                    WorkbookPart workbookPart = document.WorkbookPart;
                    IEnumerable<Sheet> sheets = workbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>().Where(s => s.Name == sheetName);
                    if (sheets.Count() == 0)
                    {
                        // La sheet n'existe pas
                        return;
                    }

                    SharedStringTablePart shareStringPart;
                    if (document.WorkbookPart.GetPartsOfType<SharedStringTablePart>().Count() > 0)
                    {
                        shareStringPart = document.WorkbookPart.GetPartsOfType<SharedStringTablePart>().First();
                    }
                    else
                    {
                        shareStringPart = document.WorkbookPart.AddNewPart<SharedStringTablePart>();
                    }

                    // Insert the text into the SharedStringTablePart.
                    int index = 0;
                    if (cellValue == CellValues.SharedString)
                    {
                        index = InsertSharedStringItem(value, shareStringPart);
                    }
                    else if (cellValue == CellValues.Boolean)
                    {
                        value = ModifyValueBoolean(value);
                    }

                    WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheets.First().Id.Value);
                    SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();

                    // On récupère le nombre de lignes actives de la sheet et la première ligne
                    int rowNumber = sheetData.Elements<Row>().Count();
                    if (premiereLigne)
                    {
                        // seulement sur la première ligne
                        rowNumber = 2;
                    }
                    Row row = sheetData.Elements<Row>().First();
                    int columnNumber = 0;

                    // Récupération de la colonne dont le contenu correspond à celui en paramètre
                    foreach (Cell c in row)
                    {
                        if (c.DataType != null && c.DataType == CellValues.SharedString)
                        {
                            var stringId = Convert.ToInt32(c.InnerText);
                            if (workbookPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(stringId).InnerText.Equals(columnName))
                            {
                                // Récupération du numéro de la colonne contenant l'information recherchée
                                columnNumber = stringId + 1;
                                break;
                            }
                        }
                    }

                    if (columnNumber != 0)
                    {
                        // On récupère la lettre correspondant au numéro de la colonne de la donnée récherchée
                        string columnLetter = GetColumnName(columnNumber);
                        
                        // Lecture des cellules de chaque ligne à partir de la seconde ligne (pour n'avoir que les données)
                        for (int i = 2; i <= rowNumber; i++)
                        {
                            // Construction du nom de la cellule Excel et récupoération de la ligne en cours
                            string cellName = columnLetter + i;
                            IEnumerable<Row> rows = sheetData.Elements<Row>().Where(r => r.RowIndex == i);

                            // Récupération de la cellule à lire
                            Cell cell = rows.First().Elements<Cell>().Where(c => string.Compare(c.CellReference.Value, cellName, true) == 0).FirstOrDefault();

                            if (cell == null)
                            {
                                cell = new Cell() { CellReference = cellName, DataType = cellValue, InnerXml = "<x:v xmlns:x=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">42</x:v>" };

                                foreach (var cellOrder in rows.First().Elements<Cell>())
                                {
                                    Regex r = new Regex("([A-Z]+)[0-9]+");
                                    Match m = r.Match(cellOrder.CellReference.Value);
                                    string colonneOrder = m.Groups[1].Value;

                                    if (colonneOrder.Length < columnLetter.Length)
                                    {
                                        // nothing
                                    } else if (string.Compare(colonneOrder, columnLetter) == 1)
                                    {
                                        cellOrder.InsertBeforeSelf<Cell>(cell);
                                        break;
                                    }
                                }
                                
                            }
                            else
                            {
                                if (cellValue== CellValues.Date)
                                {
                                    //cell.DataType = new EnumValue<CellValues>(CellValues.Date);
                                }
                                else
                                {
                                    cell.DataType = cellValue;
                                }
                            }

                            if (cellValue == CellValues.SharedString)
                            {

                                cell.CellValue = new CellValue(index.ToString());
                            }
                            else if (cellValue == CellValues.Date) {
                                cell.CellValue = new CellValue(value);
                            }
                            else
                            {
                                cell.InnerXml = cell.InnerXml.Replace(cell.InnerText, value);
                                
                                cell.CellValue = new CellValue(cell.InnerText);

                            }

                            worksheetPart.Worksheet.Save();
                        }
                    }

                    // On ferme la ressource
                    document.Close();
                }
                fs.Close();
            }
        }

        public static void WriteDataInCell(string columnName, string idColumnName, string id, string sheetName, string fileName, string value, CellValues cellValue)
        {
            using (FileStream fs = new FileStream(fileName, FileMode.Open, FileAccess.ReadWrite, FileShare.ReadWrite))
            {
                // Ouverture du document à exploiter
                using (SpreadsheetDocument document = SpreadsheetDocument.Open(fs, true))
                {
                    WorkbookPart workbookPart = document.WorkbookPart;
                    IEnumerable<Sheet> sheets = workbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>().Where(s => s.Name == sheetName);
                    if (sheets.Count() == 0)
                    {
                        // La sheet n'existe pas
                        return;
                    }

                    SharedStringTablePart shareStringPart;
                    if (document.WorkbookPart.GetPartsOfType<SharedStringTablePart>().Count() > 0)
                    {
                        shareStringPart = document.WorkbookPart.GetPartsOfType<SharedStringTablePart>().First();
                    }
                    else
                    {
                        shareStringPart = document.WorkbookPart.AddNewPart<SharedStringTablePart>();
                    }

                    // Insert the text into the SharedStringTablePart.
                    int index = 0;
                    if (cellValue == CellValues.SharedString)
                    {
                        index = InsertSharedStringItem(value, shareStringPart);
                    }
                    else if (cellValue == CellValues.Boolean)
                    {
                        value = ModifyValueBoolean(value);
                    }

                    WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheets.First().Id.Value);
                    SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();

                    // On récupère le nombre de lignes et de colonnes actives de la sheet et la première ligne
                    Row row = sheetData.Elements<Row>().First();
                    int columnNumber = 0;
                    int idColumnNumber = 0;

                    int rowNumber = sheetData.Elements<Row>().Count();

                    // Récupération des colonnes dont le contenu correspondent à celui de la colonne à modifier et à celui de l'id
                    foreach (Cell c in row)
                    {
                        if (c.DataType != null && c.DataType == CellValues.SharedString)
                        {
                            var stringId = Convert.ToInt32(c.InnerText);
                            if (workbookPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(stringId).InnerText.Equals(columnName))
                            {
                                // Récupération du numéro de la colonne contenant l'information recherchée
                                columnNumber = stringId + 1;
                            }
                        }

                        if (c.DataType != null && c.DataType == CellValues.SharedString)
                        {
                            var stringId = Convert.ToInt32(c.InnerText);
                            if (workbookPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(stringId).InnerText.Equals(idColumnName))
                            {
                                // Récupération du numéro de la colonne contenant l'information recherchée
                                idColumnNumber = stringId + 1;
                            }
                        }

                        if (columnNumber != 0 && idColumnNumber != 0)
                            break;
                    }

                    // Récupération de la cellule dont le contenu correspond à l'id récherché
                    string cellId = null;
                    if (idColumnNumber != 0)
                    {
                        string idColumnLetter = GetColumnName(idColumnNumber);

                        for (int i = 2; i <= rowNumber; i++)
                        {
                            // Construction du nom de la cellule Excel et récupoération de la ligne en cours
                            string cellName = idColumnLetter + i;

                            IEnumerable<Row> rows = sheetData.Elements<Row>().Where(r => r.RowIndex == i);

                            // Récupération de la cellule à lire
                            Cell cell = rows.First().Elements<Cell>().Where(c => string.Compare(c.CellReference.Value, cellName, true) == 0).FirstOrDefault();
                            
                            if (cell.DataType != null && cell.DataType == CellValues.SharedString)
                            {
                                var stringId = Convert.ToInt32(cell.InnerText);
                                var contenu = workbookPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(stringId).InnerText;
                                if (contenu.Equals(id))
                                {
                                    cellId = i.ToString();
                                    break;
                                }
                            }
                            else
                            {
                                if (cell.InnerText.Equals(id))
                                {
                                    cellId = i.ToString();
                                    break;
                                }
                            }
                        }

                        if (columnNumber != 0)
                        {
                            // On récupère la lettre correspondant au numéro de la colonne de la donnée récherchée
                            string columnLetter = GetColumnName(columnNumber);

                            string cellName = columnLetter + cellId;

                            // Récupération de la cellule à modifier                       
                            Row theRow = sheetData.Elements<Row>().Where(r => r.RowIndex == cellId).FirstOrDefault();
                            // Récupération de la cellule à lire
                            var comparer = new ExcelColumnComparer();
                            var indexInTheRow = theRow.Elements<Cell>().Where(c => comparer.Compare(c.CellReference.Value, cellName) <= 0).Count();
                            Cell cell = theRow.Elements<Cell>().Where(c => string.Compare(c.CellReference.Value, cellName, true) == 0).FirstOrDefault();
                            // détruire
                            //if (cell != null)
                            //{
                            //    theRow.RemoveChild(cell);
                            //    cell = null;
                            //}

                            if (cell == null)
                            {
                                if (cellValue == CellValues.SharedString)
                                {
                                    cell = new Cell()
                                    {
                                        CellReference = cellName,
                                        CellValue = new CellValue(index.ToString()),
                                        DataType = CellValues.SharedString
                                    };
                                } else
                                {
                                    cell = new Cell()
                                    {
                                        CellReference = cellName,
                                        CellValue = new CellValue(value),
                                    };
                                }
                                theRow.InsertAt(cell, indexInTheRow);
                                

                            }

                            if (cellValue == CellValues.SharedString)
                            {
                                cell.CellValue = new CellValue(index.ToString());
                            }
                            else
                            {
                                if (!string.IsNullOrEmpty(cell.InnerText) && !string.IsNullOrEmpty(cell.InnerXml)) {
                                    cell.InnerXml = cell.InnerXml.Replace(cell.InnerText, value);
                                    cell.CellValue = new CellValue(cell.InnerText);
                                }
                            }


                            worksheetPart.Worksheet.Save();
                            Thread.Sleep(1000);
                        }
                    }
                    // On ferme la ressource
                    document.Close();
                }
                fs.Close();
            }
        }

        public static List<string> ReadFirstRow(string sheetName, string fileName)
        {
            List<string> resultList = new List<string>();

            using (FileStream fs = new FileStream(fileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                using (SpreadsheetDocument document = SpreadsheetDocument.Open(fs, false))
                {
                    WorkbookPart workbookPart = document.WorkbookPart;
                    IEnumerable<Sheet> sheets = workbookPart.Workbook.Descendants<Sheet>().Where(s => s.Name == sheetName);
                    if (!sheets.Any())
                    {
                        // La feuille n'existe pas
                        return resultList;
                    }

                    WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheets.First().Id);
                    SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().FirstOrDefault();

                    if (sheetData == null)
                    {
                        // Les données de la feuille n'existent pas
                        return resultList;
                    }

                    // On récupère la première ligne
                    Row row = sheetData.Elements<Row>().FirstOrDefault();

                    if (row != null)
                    {
                        foreach (Cell cell in row.Elements<Cell>())
                        {
                            string cellValue = GetCellValue(cell, workbookPart);
                            resultList.Add(cellValue);
                        }
                    }
                }
            }

            return resultList;
        }

        public static string GetCellValue(Cell cell, WorkbookPart workbookPart)
        {
            if (cell == null || cell.CellValue == null)
                return string.Empty;

            string value = cell.CellValue.InnerText;

            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
            {
                return workbookPart.SharedStringTablePart.SharedStringTable
                    .Elements<SharedStringItem>().ElementAt(int.Parse(value)).InnerText;
            }

            return value;
        }

        //_______________________________________________________Utilitaire____________________________________________________________

        private static int GetStartIndexForReadData(string fileName)
        {
            // Détermination de l'index de début de lecture des données
            if (fileName.Contains("export-rounddelivery"))
                return 4;
            else if (fileName.Contains("CyclesCalendar"))
                return 7;
            else
                return 2;
        }

        private static string GetColumnName(int columnId)
        {
            if (columnId < 1)
            {
                throw new Exception("The column # can't be less then 1.");
            }
            columnId--;
            if (columnId >= 0 && columnId < 26)
                return ((char)('A' + columnId)).ToString();
            else if (columnId > 25)
                return GetColumnName(columnId / 26) + GetColumnName(columnId % 26 + 1);
            else
                throw new Exception("Invalid Column #" + (columnId + 1).ToString());
        }

        private static int InsertSharedStringItem(string text, SharedStringTablePart shareStringPart)
        {
            // If the part does not contain a SharedStringTable, create one.
            if (shareStringPart.SharedStringTable == null)
            {
                shareStringPart.SharedStringTable = new SharedStringTable();
            }

            int i = 0;

            // Iterate through all the items in the SharedStringTable. If the text already exists, return its index.
            foreach (SharedStringItem item in shareStringPart.SharedStringTable.Elements<SharedStringItem>())
            {
                if (item.InnerText == text)
                {
                    return i;
                }

                i++;
            }

            // The text does not exist in the part. Create the SharedStringItem and return its index.
            shareStringPart.SharedStringTable.AppendChild(new SharedStringItem(new DocumentFormat.OpenXml.Spreadsheet.Text(text)));
            shareStringPart.SharedStringTable.Save();

            return i;
        }

        private static string ModifyValueBoolean(string value)
        {
            string newValue;

            switch (value)
            {
                case "FAUX":
                    newValue = "0";
                    break;
                case "VRAI":
                    newValue = "1";
                    break;
                default:
                    newValue = "1";
                    break;
            }

            return newValue;
        }

        public static void DuplicateFirstLine(string sheetName, string fileName)
        {
            // insérer une première ligne vide
            using (FileStream fs = new FileStream(fileName, FileMode.Open, FileAccess.ReadWrite, FileShare.ReadWrite))
            {
                // Ouverture du document à exploiter
                using (SpreadsheetDocument document = SpreadsheetDocument.Open(fs, true))
                {
                    WorkbookPart workbookPart = document.WorkbookPart;
                    IEnumerable<Sheet> sheets = workbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>().Where(s => s.Name == sheetName);
                    if (sheets.Count() == 0)
                    {
                        // La sheet n'existe pas
                        return;
                    }

                    WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheets.First().Id);
                    // Get the sheetData cell table.  
                    SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();

                    // Add a row to the cell table.
                    var allRows = sheetData.Elements<Row>().ToList<Row>();
                    var newRow = (Row)allRows[1].CloneNode(true);
                    if (allRows.Count>1)
                    {
                        // on décale tout en bas sauf l'entete
                        for (int i = allRows.Count-1; i>=1; i--)
                        {
                            uint newRowIndex = allRows[i].RowIndex + 1;
                            allRows[i].RowIndex = newRowIndex;
                            foreach (var cell in allRows[i].Elements<Cell>())
                            {
                                Regex r = new Regex("([A-Z]+)[0-9]+");
                                Match m = r.Match(cell.CellReference.ToString());
                                string ColumnName = m.Groups[1].Value;
                                cell.CellReference = ColumnName + newRowIndex.ToString();
                            }
                        }
                    }
                    sheetData.InsertAt(newRow, 1);

                    worksheetPart.Worksheet.Save();
                    // On ferme la ressource
                    document.Close();
                }
                fs.Close();
            }
        }

        public static void DuplicateLastLine(string sheetName, string fileName)
        {
            // insérer une première ligne vide
            using (FileStream fs = new FileStream(fileName, FileMode.Open, FileAccess.ReadWrite, FileShare.ReadWrite))
            {
                // Ouverture du document à exploiter
                using (SpreadsheetDocument document = SpreadsheetDocument.Open(fs, true))
                {
                    WorkbookPart workbookPart = document.WorkbookPart;
                    IEnumerable<Sheet> sheets = workbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>().Where(s => s.Name == sheetName);
                    if (sheets.Count() == 0)
                    {
                        // La sheet n'existe pas
                        return;
                    }

                    WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheets.First().Id);
                    // Get the sheetData cell table.  
                    SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();

                    // Add a row to the cell table.
                    var allRows = sheetData.Elements<Row>().ToList<Row>();
                    var newRow = (Row)allRows[allRows.Count-1].CloneNode(true);
                    if (allRows.Count > 1)
                    {
                        // on décale tout en bas sauf l'entete
                        for (int i = allRows.Count - 1; i >= 1; i--)
                        {
                            uint newRowIndex = allRows[i].RowIndex + 1;
                            allRows[i].RowIndex = newRowIndex;
                            foreach (var cell in allRows[i].Elements<Cell>())
                            {
                                Regex r = new Regex("([A-Z]+)[0-9]+");
                                Match m = r.Match(cell.CellReference.ToString());
                                string ColumnName = m.Groups[1].Value;
                                cell.CellReference = ColumnName + newRowIndex.ToString();
                            }
                        }
                    }
                    newRow.RowIndex = 2;
                    foreach (var cell in newRow.Elements<Cell>())
                    {
                        Regex r = new Regex("([A-Z]+)[0-9]+");
                        Match m = r.Match(cell.CellReference.ToString());
                        string ColumnName = m.Groups[1].Value;
                        cell.CellReference = ColumnName + "2";
                    }

                    sheetData.InsertAt(newRow, 1);

                    worksheetPart.Worksheet.Save();
                    // On ferme la ressource
                    document.Close();
                }
                fs.Close();
            }
        }

        public static void AddRow(string sheetName, string fileName, object[] rowData)
        {
            using (FileStream fs = new FileStream(fileName, FileMode.Open, FileAccess.ReadWrite, FileShare.ReadWrite))
            {
                // Open the document to manipulate
                using (SpreadsheetDocument document = SpreadsheetDocument.Open(fs, true))
                {
                    WorkbookPart workbookPart = document.WorkbookPart;
                    IEnumerable<Sheet> sheets = workbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>().Where(s => s.Name == sheetName);
                    if (sheets.Count() == 0)
                    {
                        // The sheet doesn't exist
                        return;
                    }



                    WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheets.First().Id.Value);
                    SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();



                    // Create a new row
                    Row newRow = new Row();



                    // Populate the cells with data
                    foreach (var cellValue in rowData)
                    {
                        Cell newCell = new Cell();



                        if (cellValue is bool boolValue)
                        {
                            // Handle bool value
                            newCell.DataType = CellValues.Boolean;
                            newCell.CellValue = new CellValue(boolValue ? "1" : "0");
                        }
                        else if (cellValue is string stringValue)
                        {
                            // Handle string value
                            newCell.DataType = CellValues.String;
                            newCell.CellValue = new CellValue(stringValue);
                        }



                        newRow.AppendChild(newCell);
                    }



                    // Append the new row to the sheet
                    sheetData.AppendChild(newRow);



                    // Save the changes to the worksheet
                    worksheetPart.Worksheet.Save();
                    Thread.Sleep(1000);
                }
            }
        }
        public static SheetData GetSheetDataByName(WorkbookPart workbookPart, string sheetName)
        {
            var sheet = workbookPart.Workbook.Descendants<Sheet>().FirstOrDefault(s => s.Name == sheetName);
            if (sheet == null)
            {
                throw new ArgumentException($"Sheet with name '{sheetName}' not found.");
            }

            WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);
            return worksheetPart.Worksheet.GetFirstChild<SheetData>();
        }
        public static bool VerifyExportedMenus(List<string> expectedMenus, List<Dictionary<string, string>> exportedMenus)
        {
            List<string> exportedMenuNames = exportedMenus
                .Where(menu => menu.ContainsKey("Name"))
                .Select(menu => menu["Name"].Trim().ToLower())
                .ToList();

            var sortedExpectedMenus = expectedMenus.Select(m => m.Trim().ToLower()).OrderBy(m => m).ToList();
            var sortedExportedMenus = exportedMenuNames.OrderBy(m => m).ToList();

            return sortedExpectedMenus.SequenceEqual(sortedExportedMenus);
        }



        public static string GetCellValue(WorkbookPart workbookPart, Cell cell)
        {
            if (cell == null || cell.CellValue == null)
            {
                return string.Empty;
            }

            string cellValue = cell.CellValue.Text;
            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
            {
                return workbookPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(int.Parse(cellValue)).InnerText;
            }

            return cellValue;
        }
        public static List<Dictionary<string, string>> ReadExportedMenus(string filePath)
        {
            var menus = new List<Dictionary<string, string>>();

            using (SpreadsheetDocument document = SpreadsheetDocument.Open(filePath, false))
            {
                WorkbookPart workbookPart = document.WorkbookPart;
                SheetData sheetData = workbookPart.WorksheetParts.First().Worksheet.Elements<SheetData>().First();

                // Assumer que la première ligne contient les en-têtes de colonne
                Row headerRow = sheetData.Elements<Row>().First();
                var headerCells = headerRow.Elements<Cell>().ToArray();
                var columnNames = headerCells.Select(c => GetCellValue(workbookPart, c)).ToArray();

                // Parcourir les autres lignes pour récupérer les données
                foreach (Row row in sheetData.Elements<Row>().Skip(1)) // Skip the header row
                {
                    var cells = row.Elements<Cell>().ToArray();
                    var menu = new Dictionary<string, string>();

                    for (int i = 0; i < cells.Length; i++)
                    {
                        string columnName = columnNames[i];  // Nom de la colonne (ex: "Id", "Name", etc.)
                        string cellValue = GetCellValue(workbookPart, cells[i]);  // Valeur de la cellule

                        menu[columnName] = cellValue;  // Ajouter au dictionnaire
                    }

                    menus.Add(menu);
                }
            }

            return menus;
        }


    }
}

