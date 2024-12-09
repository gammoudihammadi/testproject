using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Text.RegularExpressions;
using System.Threading;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Accounting.Reporting
{
    public class ReportingPage : PageBase
    {
        public ReportingPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        // ___________________________________________ Constantes ____________________________________________________
        private const string REPORTING_SELECT = "drop-down-report-type";
        private const string SITE_SELECT = "drop-down-site";
        private const string DATE_PRODUCTION = "date-picker";
        private const string FILTER_DATE_FROM_PRODUCTION = "date-from-picker";
        private const string FILTER_DATE_TO_PRODUCTION = "date-to-picker";
        private const string SEARCH_BUTTON = "btn-get-haccp-records";
        private const string REPORTS_LINES = "//*[@id=\"table-records\"]/tbody/tr[*]";
        //private const string REPORT_DOWNLOAD_ICON = "//*[@id=\"table-records\"]/tbody/tr[{0}]//span[contains(@class, 'glyphicon glyphicon-open-file')]";
        private const string REPORT_DOWNLOAD_ICON = "//*[@id='table-records']/tbody/tr[{0}]/td[7]/button | //*[@id='open-file-icon']";
        private const string REPORT_DOWNLOAD_ICON_2 = "//*[@id=\"table-records\"]/tbody/tr[{0}]/td/button/span[@class='fas fa-file-upload']";

        // ___________________________________________ Variables _____________________________________________________
        [FindsBy(How = How.Id, Using = REPORTING_SELECT)]
        private IWebElement _reportingSelect;

        [FindsBy(How = How.Id, Using = SITE_SELECT)]
        private IWebElement _siteSelect;

        [FindsBy(How = How.Id, Using = DATE_PRODUCTION)]
        private IWebElement _dateProduction;

        [FindsBy(How = How.Id, Using = FILTER_DATE_FROM_PRODUCTION)]
        private IWebElement _dateFromProduction;

        [FindsBy(How = How.Id, Using = FILTER_DATE_TO_PRODUCTION)]
        private IWebElement _dateToProduction;

        [FindsBy(How = How.Id, Using = SEARCH_BUTTON)]
        private IWebElement _searchButton;

        [FindsBy(How = How.XPath, Using = REPORT_DOWNLOAD_ICON)]
        private IWebElement _reportDownload;

        // ___________________________________________ Méthodes ______________________________________________________

        public void FillReportingPage(string reportType, string site)
        {
            _reportingSelect = WaitForElementIsVisible(By.Id(REPORTING_SELECT), nameof(REPORTING_SELECT));
            SelectElement reporting = new SelectElement(_reportingSelect);
            reporting.SelectByText(reportType);

            _siteSelect = WaitForElementIsVisible(By.Id(SITE_SELECT), nameof(SITE_SELECT));
            SelectElement site2 = new SelectElement(_siteSelect);
            site2.SelectByText(site);

            try
            {
                _dateFromProduction = WaitForElementIsVisible(By.Id(FILTER_DATE_FROM_PRODUCTION), nameof(FILTER_DATE_FROM_PRODUCTION));
                _dateFromProduction.Clear();
                _dateFromProduction.SetValue(ControlType.DateTime, DateUtils.Now);
                _dateFromProduction.SendKeys(Keys.Tab);
                WaitForLoad();

                _dateToProduction = WaitForElementIsVisible(By.Id(FILTER_DATE_TO_PRODUCTION), nameof(FILTER_DATE_TO_PRODUCTION));
                _dateToProduction.Clear();
                _dateToProduction.SetValue(ControlType.DateTime, DateUtils.Now);
                _dateToProduction.SendKeys(Keys.Tab);
                WaitForLoad();
                Thread.Sleep(1000);
            }
            catch
            {
                //PATCH 2022.0302.1-P18
                _dateProduction = WaitForElementIsVisible(By.Id(DATE_PRODUCTION), nameof(DATE_PRODUCTION));
                _dateProduction.Clear();
                _dateProduction.SendKeys(DateUtils.Now.Date.ToString("dd/MM/yyyy"));
                WaitForLoad();
            }

            var searchButton = WaitForElementIsVisible(By.Id(SEARCH_BUTTON), nameof(SEARCH_BUTTON));
            searchButton.Click();
            WaitForLoad();
        }

        public void PrintDownload(int offset)
        {
            IWebElement _reportDownload;
            //_reportDownload = WaitForElementIsVisible(By.XPath(string.Format("//*[@id=\"table-records\"]/tbody/tr/td[7]/button/span[contains(@class, 'glyphicon glyphicon-open-file')]"
            //    , offset)), nameof(REPORT_DOWNLOAD_ICON));
            //if(isElementVisible(By.XPath(string.Format(REPORT_DOWNLOAD_ICON, offset))))
            //{
            //    _reportDownload = WaitForElementIsVisible(By.XPath(string.Format(REPORT_DOWNLOAD_ICON, offset)));
            //}
            //else
            //{
                _reportDownload = WaitForElementIsVisible(By.XPath(string.Format(REPORT_DOWNLOAD_ICON_2, offset)));
            //}
            Actions action = new Actions(_webDriver);
            action.MoveToElement(_reportDownload).Perform();
            _reportDownload.Click();
            WaitForLoad();
            WaitForDownload();
        }

        public int TableHasDocument(string DocTitle)
        {
            var table = _webDriver.FindElements(By.XPath(REPORTS_LINES));
            var x = _webDriver.FindElements(By.XPath("//*[@id=\"table-records\"]/tbody/tr[*]"));
            //"WRFR user01 HACCP3 Sanitization 2022/04/01 à 05h40"
            List<string> reportsList = new List<string>();

            foreach (var line in table)
            {
                //  Test REGEX
                if (line.Text.Contains(DocTitle))
                {
                    reportsList.Add(line.Text);
                }
            }

            if (reportsList.Count == 0)
            {
                return -1;
            }

            //Get last element of list and its index
            string last = reportsList[reportsList.Count - 1];
            for (int i = table.Count; i>0; i--)
            {
                
                if (table[i-1].Text.Contains(last))
                {
                    return i;
                }
            }
            return -1;
        }
        public bool TableHasCommentary(string DocCommentary)
        {
            var tableau = _webDriver.FindElements(By.XPath(REPORTS_LINES));
            bool trouve = false;
            int offset = 0;
            foreach (var ligne in tableau)
            {
                offset++;
                if (ligne.Text.Contains(DocCommentary))
                {
                    trouve = true;
                    break;
                }
            }
            return trouve;
        }
        public string DevPathDocTypeOfRecord()
        {
            // en Dev : ##### HACCP #####, en Path : HACCP
            string xPathHACCP = "//*[@id=\"drop-down-report-type\"]/option[contains(text(),'HACCP')]";
            return WaitForElementIsVisible(By.XPath(xPathHACCP)).Text;
        }

        public string PrintAllZip(string downloadsPath, string DocFileNamePdfBegin, string DocFileNameZipBegin)
        {
            string trouve = "";

            try
            {
                //Ouvre
                ClickPrintButton();

                // Y a t'il des document à télécharger ? - sinon print All fait une redirection vers une page blanche
                try
                {
                    string tableau = "//*/div[@class='popover-content']/table/tbody/tr";
                    var tableauAffiche = _webDriver.FindElements(By.XPath(tableau));
                    if (tableauAffiche.Count <= 1)
                    {
                        return null;
                    }
                }
                catch
                {
                    // tableau non affiché (saturation print service ?)
                    return null;
                }

                // télécharge un ZIP contenant un PDF
                string xPath = "//*[contains(text(),'Print ALL')]/parent::a";
                var PrintAll = WaitForElementIsVisible(By.XPath(xPath));
                PrintAll.Click();
                //Ferme
                ClickPrintButton();
                Thread.Sleep(2000);
                // parcourir le dossier Download

                foreach (string file in Directory.GetFiles(downloadsPath))
                {
                    // All_files_20220225_093850.zip
                    if (file.Contains(DocFileNameZipBegin) && file.EndsWith(".zip"))
                    {
                        trouve = file;
                    }
                }
                Assert.IsTrue(trouve.Length > 0, "Dir " + downloadsPath + " fichier zip " + DocFileNameZipBegin + " non trouvé");
                ZipFile.ExtractToDirectory(trouve, downloadsPath);
                trouve = "";
                foreach (string file in Directory.GetFiles(downloadsPath))
                {
                    //HACCP3 Test_25022022_-_438654_ - _20220225100147.pdf
                    if (file.Contains(DocFileNamePdfBegin) && file.EndsWith(".pdf"))
                    {
                        trouve = file;
                    }
                }
                Assert.IsTrue(trouve.Length > 0, "Dir " + downloadsPath + " (fichier zip " + DocFileNameZipBegin + ") fichier pdf " + DocFileNamePdfBegin + " non trouvé");
                Thread.Sleep(5000);
            }
            catch
            {
                trouve = null;
            }
            return trouve;
        }

        public FileInfo FindProdInFlightPDF(string downloadsPath, string docFileNamePdfBegin)
        {
            string[] files = Directory.GetFiles(downloadsPath);
            foreach (string f in files)
            {
                FileInfo fi = new FileInfo(f);
                if (fi.Name.StartsWith(docFileNamePdfBegin))
                {
                    return fi;
                }
            }
            return null;

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

        public void FixPDFMagic(string fullName)
        {
            int lengthToRemove = "data:application/pdf;base64,".Length;
            byte[] x = File.ReadAllBytes(fullName); //opens the file
            byte[] temp = new byte[x.Length - lengthToRemove]; //creates new byte array with the new size (minus 55 bytes)
            long tempx = 0; // couter for the new array position
            for (long i = lengthToRemove; i < x.LongLength; i++) // for loop which starts after 55 byte. 
            {
                temp[tempx] = x[i]; // rewritting the new file with byte from 56 and on..
                tempx++; // the counter for the new byte array
            }
            File.WriteAllBytes(fullName, temp); //after the loop, the new file is written
        }

        public string GetHACCPReportPdf(FileInfo[] taskFiles, string fileNameStart)
        {
            List<FileInfo> correctDownloadFiles = new List<FileInfo>();

            foreach (var file in taskFiles)
            {
                //  Test REGEX
                if (IsHACCPPDFFileCorrect(file.Name, fileNameStart))
                {
                    correctDownloadFiles.Add(file);
                }
            }
            if (correctDownloadFiles.Count <= 0)
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

            return correctFile.FullName;
        }

        public bool IsHACCPPDFFileCorrect(string filePath, string fileNameStart)
        {
            // HACCP3 Sanitization6_28032022_073730

            string mois = "(?:0[1-9]|1[0-2])";         // mois
            string annee = "\\d{4}";                   // annee YYYY
            string jour = "[0-3]\\d";                  // jour
            string heure = "(?:0[0-9]|1[0-9]|2[0-3])"; // heure
            string minutes = "[0-5]\\d";               // minutes
            string secondes = "[0-5]\\d";              // secondes

            Regex r = new Regex(fileNameStart + "6_" + jour + mois + annee + "_" + heure + minutes + secondes, RegexOptions.IgnoreCase | RegexOptions.Singleline);
            Match m = r.Match(filePath);

            return m.Success;
        }
    }
}
