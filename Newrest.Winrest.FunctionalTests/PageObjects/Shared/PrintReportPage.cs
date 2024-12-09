using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;
using System.Collections.Generic;
using System;
using System.IO;
using System.IO.Compression;
using System.Threading;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Content;
using UglyToad.PdfPig.DocumentLayoutAnalysis;
using OpenQA.Selenium.Support.UI;
using System.Linq;
using DocumentFormat.OpenXml.Bibliography;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Shared
{
    public class PrintReportPage : PageBase
    {
        public PrintReportPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        // ___________________________________________ Constantes ____________________________________________________
        private const string REPORT = "embed";
        private const string PRINT_BUTTON = "//*[@id=\"header-print-button\"]";
        private const string PRINT_POPUP = "//h3[text() = 'Print list']";
        private const string PRINT_ALL = "//*[contains(text(),'Print ALL')]/parent::a";
        private const string CLEAR = "//*[contains(text(),'Clear')]/parent::a";
        private const string REFRESHBUTTON = "/html/body/div[15]/div[2]/div/a[3]";
        private const string STATUSXPATH = "/html/body/div[15]/div[2]/table/tbody/tr[2]/td[3]";

        // ___________________________________________ Variables _____________________________________________________
        [FindsBy(How = How.XPath, Using = PRINT_BUTTON)]
        private IWebElement _printButton;

        [FindsBy(How = How.XPath, Using = PRINT_ALL)]
        private IWebElement _printAllButton;
        [FindsBy(How = How.XPath, Using = CLEAR)]
        private IWebElement _clearButton;
        [FindsBy(How = How.XPath, Using = REFRESHBUTTON)]
        private IWebElement _refreshButton;
        [FindsBy(How = How.XPath, Using = STATUSXPATH)]
        private IWebElement _statusXpath;

        // ___________________________________________ Méthodes ______________________________________________________

        public bool IsReportGenerated()
        {
            try
            {
                WaitForElementIsVisible(By.TagName(REPORT));
                return true;
            }
            catch
            {
                return false;
            }
        }


        public string PrintAllZip(string downloadsPath, string DocFileNamePdfBegin, string DocFileNameZipBegin, bool withRefresh = true)
        {
            if (withRefresh)
            {
                _webDriver.Navigate().Refresh();
            }

            // on clique sur le bouton impression
            //if (!isElementVisible(By.XPath(PRINT_POPUP)))
            //{
            //   _printButton = WaitForElementIsVisible(By.XPath(PRINT_BUTTON));
            //   _printButton.Click();
            //    WaitForLoad();
            //}
            ClickPrintButton();

            if (IsFileLoadedVisible(By.CssSelector("[target='_blank'][class='far fa-file-pdf']")) ||
                IsFileLoadedVisible(By.CssSelector("[target='_blank'][class='far fa-file-excel']")))
            {

                _printAllButton = WaitForElementIsVisible(By.XPath(PRINT_ALL));
                _printAllButton.Click();
                WaitLoading();
            }


            WaitForDownload();
            //Close();
            ClosePrintButton();
            // parcourir le dossier Download
            string trouve = "";
            foreach (string file in Directory.GetFiles(downloadsPath))
            {
                if (file.Contains(DocFileNameZipBegin) && file.EndsWith(".zip"))
                {
                    trouve = file;
                }
            }
            Assert.IsTrue(trouve.Length > 0, "Dir " + downloadsPath + " fichier zip " + DocFileNameZipBegin + " non trouvé");

            ZipFile.ExtractToDirectory(trouve, downloadsPath);
            trouve = "";
            int counter = 0;

            while (trouve == "" && counter < 10)
            {
                foreach (string file in Directory.GetFiles(downloadsPath))
                {
                    if (file.Contains(DocFileNamePdfBegin) && file.EndsWith(".pdf"))
                    {
                        trouve = file;
                    }
                }
                counter++;
                if (trouve == "" && counter < 10)
                {
                    WaitForLoad();
                }
            }
            Assert.IsTrue(trouve.Length > 0, "Dir " + downloadsPath + " (fichier zip " + DocFileNameZipBegin + ") fichier pdf " + DocFileNamePdfBegin + " non trouvé");
            return trouve;
        }
        // Purge pour printAllZip (service d'impression blankdoc)
        public void Purge(string downloadsPath, string docFileNamePdfBegin, string docFileNameZipBegin)
        {
            string[] files = Directory.GetFiles(downloadsPath);
            foreach (string f in files)
            {
                FileInfo fi = new FileInfo(f);
                if (fi.Name.StartsWith(docFileNamePdfBegin) || fi.Name.StartsWith(docFileNameZipBegin))
                {
                    fi.Delete();
                }
            }
        }
        public bool VerifyNewPositions(FileInfo fi, string pandp, string positionDetail)
        {
            var pandpExist = false;
            var positionDetailExist = false;
            PdfDocument document = PdfDocument.Open(fi.FullName);
            var pages = document.GetPages();
            foreach (var page in pages)
            {
                var words = page.GetWords();
                foreach (var word in words)
                {

                    if (word.Text == pandp)
                    {
                        pandpExist = true;
                    }
                    if (word.Text == positionDetail)
                    {
                        positionDetailExist = true;
                    }
                }
                if (pandpExist && positionDetailExist)
                {
                    return true;
                }
            }
            return false;

        }
        public int FindWordIndex(IEnumerable<string> words, string targetWord)
        {
            int index = -1;
            int currentIndex = 0;
            foreach (var word in words)
            {
                if (word.Equals(targetWord, StringComparison.OrdinalIgnoreCase))
                {
                    index = currentIndex;
                    break;
                }
                currentIndex++;
            }
            return index;
        }
        public int FindWordIndex(IEnumerable<string> words, string targetWord, int startIndex = 0)
        {
            WaitLoading();
            int index = -1;
            int currentIndex = 0;

            foreach (var word in words)
            {
                if (currentIndex >= startIndex && word.Equals(targetWord, StringComparison.OrdinalIgnoreCase))
                {
                    index = currentIndex;
                    break;
                }
                currentIndex++;
            }

            return index;
        }
        public static int FindWordIndex(IEnumerable<IReadOnlyList<TextBlock>> textBlocks, string targetWord)
        {
            int index = -1;
            int currentIndex = 0;
            foreach (var block in textBlocks)
            {
                foreach (var textBlock in block)
                {
                    if (textBlock.Text.Equals(targetWord))
                    {
                        index = currentIndex;
                        return index; // Return the index of the first occurrence
                    }
                    currentIndex++;
                }
            }
            return index;
        }
        public bool VerifyIfNewWindowIsOpened()
        {
        
            var wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(60));
            if (_webDriver.WindowHandles.Count == 1)
            {
                return false;
            }
            return true;
        }
        public void ExtractZipOverwrite(string zipFilePath, string destinationDirectory)
        {
            using (ZipArchive archive = ZipFile.OpenRead(zipFilePath))
            {
                foreach (ZipArchiveEntry entry in archive.Entries)
                {
                    string destinationPath = Path.Combine(destinationDirectory, entry.FullName);

                    // Check if the file exists and delete it to avoid the exception
                    if (File.Exists(destinationPath))
                    {
                        File.Delete(destinationPath);
                    }

                    // Create the directory if needed
                    if (entry.FullName.EndsWith("/"))
                    {
                        Directory.CreateDirectory(destinationPath);
                    }
                    else
                    {
                        entry.ExtractToFile(destinationPath);
                    }
                }
            }
        }

        public string PrintAllZipAllPdf(string downloadsPath, string DocFileNameZipBegin, bool withRefresh = true)
        {

            if (withRefresh)
            {
                _webDriver.Navigate().Refresh();
            }

            // on clique sur le bouton impression
            if (!isElementVisible(By.XPath(PRINT_POPUP)))
            {
                _printButton = WaitForElementIsVisible(By.XPath(PRINT_BUTTON));
                _printButton.Click();
                WaitForLoad();
            }


            if (IsFileLoadedVisible(By.CssSelector("[target='_blank'][class='far fa-file-pdf']")) ||
                IsFileLoadedVisible(By.CssSelector("[target='_blank'][class='far fa-file-excel']")))
            {

                _printAllButton = WaitForElementIsVisible(By.XPath(PRINT_ALL));
                _printAllButton.Click();
                WaitLoading();
            }

            WaitForDownload();
            Close();
            ClosePrintButton();

            // Parcourir le dossier Download pour les fichiers .zip
            string trouve = "";
            foreach (string file in Directory.GetFiles(downloadsPath))
            {
                if (file.Contains(DocFileNameZipBegin) && file.EndsWith(".zip"))
                {
                    trouve = file;
                }
            }

            Assert.IsTrue(trouve.Length > 0, "Dir " + downloadsPath + " fichier zip " + DocFileNameZipBegin + " non trouvé");
            // Extract the first zip file
            ExtractZipOverwrite(trouve, downloadsPath);
            // Recheck if there are additional zip files and extract all of them
            bool moreZipsFound;
            do
            {
                moreZipsFound = false;
                foreach (string file in Directory.GetFiles(downloadsPath))
                {
                    if (file.EndsWith(".zip"))
                    {
                        ExtractZipOverwrite(file, downloadsPath);
                        File.Delete(file);  // Optionally delete the zip file after extraction
                        moreZipsFound = true;  // Found another zip, so repeat the loop
                    }
                }
            }
            while (moreZipsFound);
            // Find all PDF files after extracting all the zip files
            trouve = "";
            int counter = 0;

            while (trouve == "" && counter < 10)
            {
                foreach (string file in Directory.GetFiles(downloadsPath))
                {
                    if (file.EndsWith(".pdf"))
                    {
                        trouve = file;  // Find any PDF file
                    }
                }
                counter++;
                if (trouve == "" && counter < 10)
                {
                    WaitForLoad();
                }
            }

            Assert.IsTrue(trouve.Length > 0, "Dir " + downloadsPath + " fichier pdf non trouvé");
            return trouve;
        }
        public void ClickRefreshUntilFinished()
        {
            while (true)
            {
                _refreshButton = WaitForElementIsVisible(By.XPath(REFRESHBUTTON));
                _refreshButton.Click();
                Thread.Sleep(2000);
                WaitForLoad();
                string _statusXpath = WaitForElementIsVisible(By.XPath(STATUSXPATH)).Text;
                if (_statusXpath.Equals("Finished", StringComparison.OrdinalIgnoreCase))
                {
                    break;
                }
            }
        }
    }
}


   