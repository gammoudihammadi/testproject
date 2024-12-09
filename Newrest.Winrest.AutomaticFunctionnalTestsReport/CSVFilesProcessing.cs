using Microsoft.VisualBasic.FileIO;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;

namespace Newrest.Winrest.AutomaticFunctionnalTestsReport
{
    public class CSVFilesProcessing
    {
        public List<FileInfo> SearchForCSVFiles(string CSVDirectory, int nbCSV)
        {
            List<FileInfo> correctDownloadFiles = new List<FileInfo>();

            DirectoryInfo taskDirectory = new DirectoryInfo(CSVDirectory);            

            foreach (var file in taskDirectory.GetFiles())
            {
                //  Test REGEX
                if (IsCSVFileNameCorrect(file.Name))
                {
                    correctDownloadFiles.Add(file);
                }
            }

            if (correctDownloadFiles.Count != nbCSV)
            {
                throw new Exception("Tous les fichiers CSV n'ont pas été générés.");
            }

            return correctDownloadFiles;
        }

        public bool IsCSVFileNameCorrect(string filePath)
        {
            // "Accounting_TestResults.csv"

            Regex r = new Regex("^(.*)_TestResults.csv$", RegexOptions.IgnoreCase | RegexOptions.Singleline);
            Match m = r.Match(filePath);

            return m.Success;
        }


        // ____________________________________________________________ exploitation fichiers CSV __________________________________________________________
        public void ExtractInfosFromCSVFiles(string fileName, Dictionary<string, TestAuto> testsList)
        {

            using (TextFieldParser csvReader = new TextFieldParser(fileName))
            {
                // Déclaration des infos pour les CSV
                csvReader.CommentTokens = new string[] { "#" };
                csvReader.SetDelimiters(new string[] { "," });
                csvReader.HasFieldsEnclosedInQuotes = true;

                // On lit la première ligne pour extraire les index des colonnes à conserver
                string[] header = csvReader.ReadFields();

                int indexTestName = Array.IndexOf(header, "testCaseName");
                int indexEnv = Array.IndexOf(header, "Environment");
                int indexStatus = Array.IndexOf(header, "outcome");
                int indexErrorMessage = Array.IndexOf(header, "errorMessage");

                if (indexTestName != -1 || indexEnv != -1 || indexStatus != -1 || indexErrorMessage != -1)
                {
                    throw new Exception("Le format du fichier " + fileName + " est incorrect, les éléments nécessaires au remplissage du rapport ne sont pas disponibles.");
                }

                // On parcourt le reste du document
                while (!csvReader.EndOfData)
                {
                    // Read current line fields, pointer moves to the next line.
                    string[] fields = csvReader.ReadFields();
                    TestAuto test = new TestAuto
                    {
                        NomTest = fields[indexTestName],
                        Environnement = fields[indexEnv],
                        Status = fields[indexStatus],
                        ErrorMessage = fields[indexErrorMessage]
                    };

                    testsList.Add(test.NomTest, test);
                }

            }

        }

        public void TestImportCSVAcEPPlus(string fileName)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            //set the formatting options
            ExcelTextFormat format = new ExcelTextFormat();
            format.Delimiter = ',';
            format.TextQualifier = '"';
            format.Culture = new CultureInfo(Thread.CurrentThread.CurrentCulture.ToString());
            format.Culture.DateTimeFormat.ShortDatePattern = "dd-mm-yyyy";
            format.Encoding = new UTF8Encoding();

            //read the CSV file from disk
            FileInfo file = new FileInfo(fileName);

            //create a new Excel package
            using (ExcelPackage excelPackage = new ExcelPackage())
            {
                //create a WorkSheet
                ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Sheet 1");

                //load the CSV data into cell A1
                var lignes = worksheet.Cells["A1"].LoadFromText(file, format);
            }
        }
    }
}
