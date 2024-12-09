using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Warehouse.OutputForm;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;


namespace Newrest.Winrest.FunctionalTests.PageObjects.Shared
{
    public class GenericExport : PageBase
    {
        public string downloadsPath = "C:\\ChromeDriverDownloads";
        public GenericExport(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        public void ClearDownload()
        {
            string[] filenames = Directory.GetFiles(downloadsPath, "*", SearchOption.TopDirectoryOnly);
            foreach (string fName in filenames)
            {
                try
                {
                    File.Delete(fName);
                }
                catch (IOException ioe)
                {
                    Console.WriteLine("cannot delete " + ioe.ToString());
                }
            }
        }

        public virtual void ShowExtendedMenu(string extendedButtonXPath)
        {
            var extendedButton = WaitForElementExists(By.XPath(extendedButtonXPath));
            var actions = new Actions(_webDriver);
            actions.MoveToElement(extendedButton).Perform();
            WaitForLoad();
        }

        public FileInfo GetExportExcelFile(FileInfo[] taskFiles, FileNamePattern fileNamePattern)
        {
            List<FileInfo> correctDownloadFiles = new List<FileInfo>();

            foreach (var file in taskFiles)
            {
                // Test REGEX
                if (IsExcelFileCorrect(file.Name, fileNamePattern))
                {
                    correctDownloadFiles.Add(file);
                }
            }

            if (correctDownloadFiles.Count == 0)
            {
                return null;
            }

            var time = correctDownloadFiles[0].LastWriteTimeUtc;
            var correctFile = correctDownloadFiles[0];

            correctDownloadFiles.ForEach(file =>
            {
                if (time < file.LastWriteTimeUtc)
                {
                    time = file.LastWriteTimeUtc;
                    correctFile = file;
                }
            });

            return correctFile;
        }

        public bool IsExcelFileCorrect(string filePath, FileNamePattern fileNamePattern)
        {
            WaitPageLoading();
            string mois = "(?:0[1-9]|1[0-2])";         // mois
            string space = "(\\s)";                    // Espace
            string annee = "\\d{4}";                   // annee YYYY
            string jour = "[0-3]\\d";                  // jour
            string heure = "(?:0[0-9]|1[0-9]|2[0-3])"; // heure
            string minutes = "[0-5]\\d";               // minutes
            string secondes = "[0-5]\\d";              // secondes

            string pattern;
            if (fileNamePattern.HasDate)
            {
                pattern = "^" + fileNamePattern.Pattern + space + annee + "-" + mois + "-" + jour + space + heure + "-" + minutes + "-" + secondes + ".xlsx$";
            }
            else
            {
                pattern = "^" + fileNamePattern.Pattern + ".xlsx$";
            }

            Regex regex = new Regex(pattern, RegexOptions.IgnoreCase | RegexOptions.Singleline);
            return regex.IsMatch(filePath);
        }

        public FileInfo IsGenerated(FileNamePattern fileNamePattern)
        {
            int maxRetries = 5;

            WaitPageLoading();
            WaitForLoad();

            for (int i = 0; i < maxRetries; i++)
            {
                DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
                FileInfo[] taskFiles = taskDirectory.GetFiles();

                var matchingFiles = taskFiles.Where(file =>
                    file.Name.Contains(fileNamePattern.Pattern));

                if (matchingFiles.Any())
                {
                    return matchingFiles.First();
                }

                WaitPageLoading();
                WaitForLoad();
            }

            return null;
        }

        public int resultNumber(string fileName, string sheetName)
        {
            WaitPageLoading();
            var filePath = Path.Combine(downloadsPath, fileName);

            // Exploitation du fichier Excel
            return OpenXmlExcel.GetExportResultNumber(sheetName, filePath);
        }

        public class FileNamePattern
        {
            public string Pattern { get; set; }
            public bool HasDate { get; set; }
        }
        public bool IsEqualValue(string columnname, string sheetName, string fileName, string value)
        {
            WaitPageLoading();
            return OpenXmlExcel.ReadAllDataInColumn(columnname, sheetName, fileName, value);
        }
    }
}
