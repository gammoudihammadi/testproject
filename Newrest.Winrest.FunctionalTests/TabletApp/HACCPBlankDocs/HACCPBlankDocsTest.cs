using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects;
using Newrest.Winrest.FunctionalTests.PageObjects.Accounting.Reporting;
using Newrest.Winrest.FunctionalTests.PageObjects.TabletApp;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Content;

namespace Newrest.Winrest.FunctionalTests.TabletApp
{
    [TestClass]
    public class HACCPBlankDocsTest : TestBase
    {
        private const int Timeout = 600000;
        /// <summary>
        /// Mise en place du paramétrage pour HACCP
        /// </summary>
        [TestMethod]
        [Priority(0)]
        [Timeout(Timeout)]
        public void TB_HACCP_SetConfigWinrest()
        {
            //Arrange
            var homePage = LogInAsAdmin();

            //Check all documents in application settings
            var applicationSettings = homePage.GoToApplicationSettings();

            //Add documents & columns HACCP
            var HACCPDocumentSettings = applicationSettings.GoToHACCPDocumentSettings();
            //Attention ! Seul 4 documents sont paramétrés HACCP3 Sanitization, HACCP3 Thawing, HACCP3 Test, HACCP3 Modified texture
            HACCPDocumentSettings.AddAllDocumentsAndAllColumns();

            //Vérifier si documents bien cochés dans App settings au niveau de la ligne DocumentsDisplay
            var isDocumentsHACCP3Configured = HACCPDocumentSettings.IsHACCP_CONFIG();

            //Assert
            Assert.IsTrue(isDocumentsHACCP3Configured, "L'ajouter des documents à tester a échoué");
        }

        /**
         * Crée un blank doc
         */
        [TestMethod]
        [Timeout(Timeout)]
        public void TB_HACCP_BlankDocs_SelectDocumentHACCP()
        {
            // Prepare
            string DocTitle = "HACCP3 Sanitization";
            string DocSite = "GRO";

            // Arrange
            var homePage = LogInAsAdmin();

            TabletAppPage tabletAppPage = homePage.GotoTabletApp();

            //1.sélectionner un site sur la tablette
            tabletAppPage.Select("mat-select-0", DocSite);

            //2.aller sur ""blank docs""
            tabletAppPage.GotoTabletApp_BlankDoc();

            //3.sélectionner le document haccp
            tabletAppPage.Select("mat-select-2", DocTitle);
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void TB_HACCP_BlankDocs_SaveDocumentHACCP()
        {
            // Prepare
            string DocTitle = "HACCP3 Sanitization";
            string DocSite = "GRO";
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            //HACCP3 Sanitization_01032022_-_438857_-_20220301095317.pdf
            string DocFileNamePdfBegin = DocTitle + "_";
            //All_files_20220225_102148.zip
            string DocFileNameZipBegin = "All_files_";

            // Arrange
            var homePage = LogInAsAdmin();

            TabletAppPage tabletAppPage = homePage.GotoTabletApp();

            //1.sélectionner un site sur la tablette
            tabletAppPage.Select("mat-select-0", DocSite);

            //2.aller sur ""blank docs""
            tabletAppPage.GotoTabletApp_BlankDoc();

            //3.Sélectioner un doc haccp dans le menu déroulant
            tabletAppPage.Select("mat-select-2", DocTitle);

            //4.Cliquer sur next pour ouvrir le document
            tabletAppPage.ClickBaseButton("NEXT STEP");

            //5.choisir une date de production et cliquer sur print (date non sélectionnable)
            tabletAppPage.ClickButton("Print");

            // je suis dans un nouvel onglet !!!!
            tabletAppPage.GotoTabletApp_DocumentTab();

            //6.cliquer sur add line et remplir une ligne de données
            tabletAppPage.addLineSanitization("MyProduct", "1.3", "2.2", "2.3", "MyView", "MyCorrectiveAction", "MyPreparedBy");

            Thread.Sleep(2000);
            //7.cliquer sur ""save and validate"" puis sur ""create record""
            tabletAppPage.ClickButton("Save and validate");
            tabletAppPage.ClickButton("Validate");
            tabletAppPage.ClickButton("Ok");

            // verify Accouting>Reporting
            homePage.Navigate();
            homePage.ClearDownloads();

            ReportingPage reportingPage = homePage.GoToAccounting_Reporting();
            //HACCP3 Sanitization_01032022_-_438856_-_20220301094904.pdf
            reportingPage.Purge(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);

            string DocTypeOfReport = reportingPage.DevPathDocTypeOfRecord();
            reportingPage.FillReportingPage(DocTypeOfReport, DocSite);
            var offset = reportingPage.TableHasDocument(DocTitle);
            Assert.IsNotNull(offset, "Document non trouvé");

            // télécharger le document
            reportingPage.PrintDownload(offset);

            var correctDownloadedFile = reportingPage.PrintAllZip(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);

            if (correctDownloadedFile == null)
            {
                // On récupère les fichiers du répertoire de téléchargement
                DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
                FileInfo[] taskFiles = taskDirectory.GetFiles();

                // On recherche le fichier téléchargé
                correctDownloadedFile = reportingPage.GetHACCPReportPdf(taskFiles, DocTitle);
                Assert.IsNotNull(correctDownloadedFile, "Le report n'a pas été téléchargé.");
            }

            PdfDocument document = PdfDocument.Open(correctDownloadedFile);
            Page page1 = document.GetPage(1);
            IEnumerable<Word> words = page1.GetWords();
            Assert.AreEqual(1, words.Count(w => w.Text == "MyProduct"), "MyProduct non présent dans le Pdf");
            Assert.AreEqual(1, words.Count(w => w.Text == "1.3"), "1.3 non présent dans le Pdf");
            Assert.AreEqual(1, words.Count(w => w.Text == "2.2"), "2.2 non présent dans le Pdf");
            Assert.AreEqual(1, words.Count(w => w.Text == "2.3"), "2.3 non présent dans le Pdf");
            Assert.AreEqual(1, words.Count(w => w.Text == "MyView"), "Le commentaire n'est pas retrouvé dans le Pdf");
            Assert.AreEqual(1, words.Count(w => w.Text == "MyCorrectiveAction"), "MyCorrectiveAction non présent dans le Pdf");
            Assert.AreEqual(1, words.Count(w => w.Text == "MyPreparedBy"), "MyPreparedBy non présent dans le Pdf");
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void TB_HACCP_BlancDocs_CheckSaveComment()
        {
            // Prepare
            string DocTitle = "HACCP3 Sanitization";
            string DocCommentary = "Hello World BlancDocs !";
            string DocSite = "GRO";


            // Arrange
            var homePage = LogInAsAdmin();

            //Act
            TabletAppPage tabletAppPage = homePage.GotoTabletApp();

            //sélectionner un site sur la tablette
            tabletAppPage.Select("mat-select-0", DocSite);

            //aller sur ""blank docs""
            tabletAppPage.GotoTabletApp_BlankDoc();
            Thread.Sleep(2000);

            //2.Choisir le doc haccp, puis cliquer sur next step pour l'ouvrir
            //Sélectioner un doc haccp dans le menu déroulant
            // "Comment" s'efface si on change de select
            tabletAppPage.Select("mat-select-2", DocTitle);
            //1.saisir un commentaire dans la case ""comment""
            var option = tabletAppPage.WaitForElementIsVisible(By.Id("mat-input-0"));
            option.SendKeys(DocCommentary);
            Thread.Sleep(2000);

            //Cliquer sur next pour ouvrir le document
            tabletAppPage.ClickBaseButton("NEXT STEP");

            //3.cliquer sur le bouton print pour imprimer le rapport
            //choisir une date de production et cliquer sur print(date non sélectionnable)
            tabletAppPage.ClickButton("Print");

            // je suis dans un nouvel onglet !!!!
            tabletAppPage.GotoTabletApp_DocumentTab();
            Thread.Sleep(2000);

            tabletAppPage.ClickButton("Print");
            //cliquer sur ""Save and validate""
            tabletAppPage.ClickButton("Save and validate");
            tabletAppPage.ClickButton("Validate");
            tabletAppPage.WaitHACCPHorizontalProgressBar();
            tabletAppPage.ClickButton("Ok");


            //Vérifier que le commentaire est présent sur le print
            //1.Commentaire visible sur le print du document haccp
            //2.sur la partie winrest dans accounting => reporting que le commentaire est présent dans le champs ""commentary ""
            //(après avoir sélectionné le type de rapport ""HACCP"" et avoir choisi la même date de production et le même site)
            homePage.Navigate();
            ReportingPage reportingPage = homePage.GoToAccounting_Reporting();
            string DocTypeOfReport = reportingPage.DevPathDocTypeOfRecord();
            reportingPage.FillReportingPage(DocTypeOfReport, DocSite);
            bool hasCommentary = reportingPage.TableHasCommentary(DocCommentary);
            Assert.IsTrue(hasCommentary, "commentaire non présent dans le Accounting>Reporting");

        }

        [TestMethod]
        [Timeout(Timeout)]
        public void TB_HACCP_BlankDocs_SaveDocumentTemporary()
        {
            // Prepare
            string DocTitle = "HACCP3 Sanitization";
            string DocSite = "GRO";
            string DocFileName = "TempTest_wrfr.user01";
            string Product = "MyProduct";
            string Quantity = "1.3";
            string Disinfection1 = "2.2";
            string Disinfection2 = "2.3";
            string Comments = "MyView";
            string Corrective = "MyCorrectiveAction";
            string PreparedBy = "MyPreparedBy";

            // Arrange
            var homePage = LogInAsAdmin();

            TabletAppPage tabletAppPage = homePage.GotoTabletApp();

            //1. sélectionner un site sur la tablette
            tabletAppPage.Select("mat-select-0", DocSite);

            //2. aller sur ""blank docs""
            tabletAppPage.GotoTabletApp_BlankDoc();
            tabletAppPage.WaitPageLoading();

            //1. Sélectioner un doc haccp dans le menu déroulant
            tabletAppPage.SelectV3("mat-select-2", DocTitle);
            tabletAppPage.WaitPageLoading();

            tabletAppPage.Purge(6);
            tabletAppPage.Purge(4);

            homePage.Navigate();
            tabletAppPage = homePage.GotoTabletApp();
            tabletAppPage.Select("mat-select-0", DocSite);
            tabletAppPage.GotoTabletApp_BlankDoc();
            tabletAppPage.WaitPageLoading();

            //Sélectioner un doc haccp dans le menu déroulant
            tabletAppPage.SelectV3("mat-select-2", DocTitle);
            tabletAppPage.WaitPageLoading();

            //Cliquer sur next 
            tabletAppPage.ClickBaseButton("NEXT STEP");

            //choisir une date de production et cliquer sur print
            tabletAppPage.ClickButton("Print");

            tabletAppPage.GotoTabletApp_DocumentTab();
            tabletAppPage.WaitPageLoading();

            //remplir une ligne de données
            tabletAppPage.addLineSanitization(Product, Quantity, Disinfection1, Disinfection2, Comments, Corrective, PreparedBy);


            //Cliquer sur le bouton "save "
            tabletAppPage.ClickButton("Save");

            //Choisir un nom pour le doc et cliquer sur ""validate""
            tabletAppPage.FillFileName3(DocFileName);
            tabletAppPage.ClickButton("Validate");

            tabletAppPage.ClickButton("Yes");
            tabletAppPage.ClickButton("Ok");

            //Revenir vers la page d'acceuil de blank docs , sélectionner le doc haccp dans le menu déroulant
            tabletAppPage.Close();
            tabletAppPage.WaitPageLoading();
            homePage.Navigate();

            tabletAppPage = homePage.GotoTabletApp();
            //sélectionner un site sur la tablette
            tabletAppPage.Select("mat-select-0", DocSite);
            //aller sur ""blank docs""
            tabletAppPage.GotoTabletApp_BlankDoc();
            tabletAppPage.WaitPageLoading();

            //Dans la barre des ""temporary"" on devrait retrouver le doc temporary sauvegardé avec la bonne nomination"
            tabletAppPage.SelectV3("mat-select-2", DocTitle);

            tabletAppPage.SelectAction("mat-select-4", DocFileName, true);

            var docFile = tabletAppPage.GetFileTemporary();

            //Dans la barre des ""temporary"" on devrait retrouver le doc temporary sauvegardé avec la bonne nomination
            Assert.AreEqual(docFile, DocFileName, "Dans la barre des Temporary on n'a pas retrouvé le doc avec la bonne nomination");

            tabletAppPage.ClickBaseButton("NEXT STEP");
            tabletAppPage.GotoTabletApp_DocumentTab();

            // les données enregsitrés 
            Assert.AreEqual("MyProduct", tabletAppPage.CheckValue(0), "Docucment ne comprend pas : MyProduct");
            Assert.AreEqual("1.3", tabletAppPage.CheckValue(1), "Document ne comprend pas : MyProduct");
            Assert.AreEqual("2.2", tabletAppPage.CheckValue(2), "Document ne comprend pas : MyProduct");
            Assert.AreEqual("2.3", tabletAppPage.CheckValue(3), "Document ne comprend pas : MyProduct");

            Assert.AreEqual(true, tabletAppPage.CheckValueBox(1), "Document ne comprend pas : Rinsing");
            Assert.AreEqual(true, tabletAppPage.CheckValueBox(2), "Document ne comprend pas : Final");
         
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void TB_HACCP_BlankDocs_SaveDocumentTemporaryCheckDate()
        {
            // Prepare
            string DocTitle = "HACCP3 Thawing";
            string DocSite = "GRO";
            string DocFileName = "TempTest_Thawing.user01";
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            //HACCP3 Thawing6_01022023_081300.pdf
            string DocFileNamePdfBegin = DocTitle;
            //All_files_20220225_102148.zip
            string DocFileNameZipBegin = "All_files_";

            // Arrange
            var homePage = LogInAsAdmin();

            TabletAppPage tabletAppPage = homePage.GotoTabletApp();

            //1. sélectionner un site sur la tablette
            tabletAppPage.Select("mat-select-0", DocSite);

            //2. aller sur ""blank docs""
            tabletAppPage.GotoTabletApp_BlankDoc();
            Thread.Sleep(2000);

            //4. Sélectioner un doc haccp dans le menu déroulant
            tabletAppPage.SelectV3("mat-select-2", DocTitle);
            tabletAppPage.Purge(6);
            tabletAppPage.Purge(4);

            homePage.Navigate();
            tabletAppPage = homePage.GotoTabletApp();
            tabletAppPage.Select("mat-select-0", DocSite);
            tabletAppPage.GotoTabletApp_BlankDoc();
            Thread.Sleep(2000);

            //Sélectioner un doc haccp dans le menu déroulant
            tabletAppPage.SelectV3("mat-select-2", DocTitle);
            Thread.Sleep(2000);

            //5. Cliquer sur next pour ouvrir le document
            tabletAppPage.ClickBaseButton("NEXT STEP");

            //6. choisir une date de production et cliquer sur print
            tabletAppPage.ClickButton("Print");

            // je suis dans un nouvel onglet !!!!
            tabletAppPage.GotoTabletApp_DocumentTab();
            Thread.Sleep(2000);

            //7. renseigner une date dans la colonne où le format date est paramétré
            tabletAppPage.addLineThrowning("MyProduct", DateUtils.Now.Date.ToString("dd/MM/yyyy"), "7", "10", DateUtils.Now.Date.AddDays(-10).ToString("dd/MM/yyyy"), DateUtils.Now.Date.AddDays(10).ToString("dd/MM/yyyy"), "MyView", "MyCorrective", "MyPreparedBy");


            //8. cliquer sur save pour enregistrer un document temporary
            tabletAppPage.ClickButton("Save");

            //cliquer sur ""validate""
            tabletAppPage.FillFileName3(DocFileName);
            tabletAppPage.ClickButton("Validate");

            try
            {
                // erase old file (pourtant on a purge)
                tabletAppPage.ClickButton("Yes", true);
            }
            catch
            {
                // nouveau fichier
            }

            //une pop de validation apparait et puis cliquer sur OK
            tabletAppPage.ClickButton("Ok");

            //Revenir vers la page d'acceuil de blank docs , sélectionner le doc haccp dans le menu déroulant
            Thread.Sleep(2000);
            tabletAppPage.Close();
            homePage.Navigate();

            tabletAppPage = homePage.GotoTabletApp();
            //sélectionner un site sur la tablette
            tabletAppPage.Select("mat-select-0", DocSite);
            //aller sur ""blank docs""
            tabletAppPage.GotoTabletApp_BlankDoc();
            Thread.Sleep(2000);

            //9. récupérer ensuite le document temporary dans la page d'acceuil de blank docs dans la section "" temporary""
            tabletAppPage.SelectV3("mat-select-2", DocTitle);
            // envoi des event remplissant le second combobox
            //tabletAppPage.SelectAction("mat-select-2", DocTitle);
            tabletAppPage.SelectAction("mat-select-4", DocFileName, true);
            tabletAppPage.ClickBaseButton("NEXT STEP");
            // patch : tabletAppPage.ClickButton("Print");
            // je suis dans un nouvel onglet !!!!
            tabletAppPage.GotoTabletApp_DocumentTab();
            Thread.Sleep(2000);

            //10. vérifier que la date s'est bien enregistrée sur le document temporary
            IWebElement date;
            if (tabletAppPage.isElementVisible(By.Id("mat-input-1")))
            {
                date = tabletAppPage.WaitForElementIsVisible(By.Id("mat-input-1"));
                Assert.AreEqual(DateUtils.Now.Date.ToString("dd/MM/yyyy"), date.GetAttribute("value"));
            }
            else
            {
                date = tabletAppPage.WaitForElementIsVisible(By.XPath("/html/body/app-root/mat-sidenav-container/mat-sidenav-content/main/app-haccp-document/mat-toolbar/mat-toolbar-row/span[1]"));
                var dates = date.Text;
                int startIndex = dates.IndexOf(':') + 2;
                string dateSubstring = dates.Substring(startIndex);
                Assert.AreEqual(DateUtils.Now.Date.ToString("dd/MM/yyyy"), dateSubstring);
            }

            //11.puis cliquer sur print pour générer le rapport haccp en pdf
            tabletAppPage.ClickButton("Print");
            //12.vérifier que la date est correcte sur le rapport généré "
            //Vérifier que les dates renseignés sont correctes sur le rapport généré
            //et

            tabletAppPage.ClickButton("Save and validate");
            tabletAppPage.ClickButton("Validate");

            tabletAppPage.WaitHACCPHorizontalProgressBar();
            tabletAppPage.ClickButton("Ok");

            tabletAppPage.Close();
            //Dans winrest dans la partie accouting=> reporting:
            homePage.Navigate();
            homePage.ClearDownloads();
            ReportingPage reportingPage = homePage.GoToAccounting_Reporting();
            reportingPage.Purge(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);

            //1.Sélectionner type of report: HACCP
            string DocTypeOfReport = reportingPage.DevPathDocTypeOfRecord();
            //2.sélectionner le site et sélectionner la même production
            reportingPage.FillReportingPage(DocTypeOfReport, DocSite);

            //3.recupérer le rapport HACCP
            var offset = reportingPage.TableHasDocument(DocTitle);
            Assert.IsNotNull(offset, "Document non trouvé");
            //4.On retrouve bien les dates renseignées et les dates doivent être conformes dans le rapport généré dans le bon format également défini dans les global settings(settings => global settings => DateFormatDisplay)"
            // télécharger le document
            reportingPage.PrintDownload(offset);

            //var correctDownloadedFile = reportingPage.PrintAllZip(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);

            //if (correctDownloadedFile == null)
            //{
            // On récupère les fichiers du répertoire de téléchargement
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier téléchargé
            var correctDownloadedFile = reportingPage.GetHACCPReportPdf(taskFiles, DocTitle);
            Assert.IsNotNull(correctDownloadedFile, "Le report n'a pas été téléchargé.");
            //}

            PdfDocument document = PdfDocument.Open(correctDownloadedFile);
            Page page1 = document.GetPage(1);
            IEnumerable<Word> words = page1.GetWords();
            Assert.AreEqual(1, words.Count(w => w.Text == "MyProduct"), "MyProduct non présent dans le Pdf");
            Assert.AreEqual(2, words.Count(w => w.Text == DateUtils.Now.Date.ToString("dd/MM/yyyy")), DateUtils.Now.Date.ToString("dd/MM/yyyy") + " non présent dans le Pdf");
            Assert.AreEqual(1, words.Count(w => w.Text == "7"), "7 non présent dans le Pdf");
            Assert.AreEqual(1, words.Count(w => w.Text == "10"), "10 non présent dans le Pdf");
            Assert.AreEqual(1, words.Count(w => w.Text == DateUtils.Now.Date.AddDays(-10).ToString("dd/MM/yyyy")), DateUtils.Now.Date.AddDays(-10).ToString("dd/MM/yyyy") + " non présent dans le Pdf");
            Assert.AreEqual(1, words.Count(w => w.Text == DateUtils.Now.Date.AddDays(10).ToString("dd/MM/yyyy")), DateUtils.Now.Date.AddDays(10).ToString("dd/MM/yyyy") + " non présent dans le Pdf");
            Assert.AreEqual(1, words.Count(w => w.Text == "MyView"), "Le commentaire n'est pas retrouvé dans le Pdf");
            Assert.AreEqual(1, words.Count(w => w.Text == "MyCorrective"), "MyCorrective non présent dans le Pdf");
            Assert.AreEqual(1, words.Count(w => w.Text == "MyPreparedBy"), "MyPreparedBy non présent dans le Pdf");
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void TB_HACCP_BlankDocs_SaveDocumentCheckDate()
        {

            // Prepare
            string DocTitle = "HACCP3 Thawing";
            string DocSite = "GRO";
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            //HACCP3 Sanitization_01032022_-_438857_-_20220301095317.pdf
            string DocFileNamePdfBegin = DocTitle + "_";
            //All_files_20220225_102148.zip
            string DocFileNameZipBegin = "All_files_";

            // Arrange
            var homePage = LogInAsAdmin();

            TabletAppPage tabletAppPage = homePage.GotoTabletApp();

            //1. sélectionner un site sur la tablette
            tabletAppPage.Select("mat-select-0", DocSite);

            //2. aller sur ""blank doc""
            tabletAppPage.GotoTabletApp_BlankDoc();
            tabletAppPage.WaitPageLoading();

            //3. Sélectioner un doc haccp dans le menu déroulant
            tabletAppPage.Select("mat-select-2", DocTitle);
            // envoi des event remplissant le second combobox
            tabletAppPage.SelectAction("mat-select-2", DocTitle);

            //4. Cliquer sur next pour ouvrir le document
            tabletAppPage.ClickBaseButton("NEXT STEP");

            //5. choisir une date de production et cliquer sur print
            tabletAppPage.ClickButton("Print");

            tabletAppPage.GotoTabletApp_DocumentTab();
            tabletAppPage.WaitPageLoading();

            //6.renseigner une date dans la colonne où le format date est paramétré
            tabletAppPage.addLineThrowning("MyProduct", DateUtils.Now.Date.ToString("dd/MM/yyyy"), "7", "10", DateUtils.Now.Date.AddDays(-10).ToString("dd/MM/yyyy"), DateUtils.Now.Date.AddDays(10).ToString("dd/MM/yyyy"), "MyView", "MyCorrective", "MyPreparedBy");

            //7.puis cliquer sur ""save and validate"" puis cliquer sur ""validate""
            tabletAppPage.ClickButton("Save and validate");
            tabletAppPage.ClickButton("Validate");
            tabletAppPage.ClickButton("Ok");

            //Dans winrest dans la partie accouting=> reporting :
            homePage.Navigate();
            homePage.ClearDownloads();
            ReportingPage reportingPage = homePage.GoToAccounting_Reporting();
            reportingPage.Purge(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);

            //1. Sélectionner type of report : HACCP
            string DocTypeOfReport = reportingPage.DevPathDocTypeOfRecord();

            //2. sélectionner le site et sélectionner la même production date
            reportingPage.FillReportingPage(DocTypeOfReport, DocSite);
            var offset = reportingPage.TableHasDocument(DocTitle);
            //3. recupérer le rapport HACCP
            Assert.IsNotNull(offset, "Document non trouvé");

            // télécharger le document
            reportingPage.PrintDownload(offset);

            var correctDownloadedFile = reportingPage.PrintAllZip(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);
      
            // On récupère les fichiers du répertoire de téléchargement
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier téléchargé
            correctDownloadedFile = reportingPage.GetHACCPReportPdf(taskFiles, DocTitle);
            Assert.IsNotNull(correctDownloadedFile, "Le report n'a pas été téléchargé.");
           

            PdfDocument document = PdfDocument.Open(correctDownloadedFile);
            Page page1 = document.GetPage(1);
            IEnumerable<Word> words = page1.GetWords();
            Assert.AreEqual(1, words.Count(w => w.Text == "MyProduct"), "MyProduct non présent dans le Pdf");
            Assert.AreEqual(2, words.Count(w => w.Text == DateUtils.Now.Date.ToString("dd/MM/yyyy")), DateUtils.Now.Date.ToString("dd/MM/yyyy") + " non présent dans le Pdf");
            Assert.AreEqual(1, words.Count(w => w.Text == "7"), "7 non présent dans le Pdf");
            Assert.AreEqual(1, words.Count(w => w.Text == "10"), "10 non présent dans le Pdf");
            Assert.AreEqual(1, words.Count(w => w.Text == DateUtils.Now.Date.AddDays(-10).ToString("dd/MM/yyyy")), DateUtils.Now.Date.AddDays(-10).ToString("dd/MM/yyyy") + " non présent dans le Pdf");
            Assert.AreEqual(1, words.Count(w => w.Text == DateUtils.Now.Date.AddDays(10).ToString("dd/MM/yyyy")), DateUtils.Now.Date.AddDays(10).ToString("dd/MM/yyyy") + " non présent dans le Pdf");
            Assert.AreEqual(1, words.Count(w => w.Text == "MyView"), "MyView non présent dans le Pdf");
            Assert.AreEqual(1, words.Count(w => w.Text == "MyCorrective"), "MyCorrective non présent dans le Pdf");
            Assert.AreEqual(1, words.Count(w => w.Text == "MyPreparedBy"), "MyPreparedBy non présent dans le Pdf");
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void TB_HACCP_BlankDocs_SaveDocumentHACCPFavorite()
        {
            // Prepare
            string DocTitle = "HACCP3 Sanitization";
            string DocSite = "GRO";
            string DocFileName = "FavTest_wrfr.user01" + new Random().Next().ToString();
            string DocProduct = "FavProduct";
            string quantityProductEntered = "1.3";
            string disinfection1 = "2.2";
            string disinfection2 = "2.3";
            string commentProduct = "MyView";
            string correctionAction = "MyCorrectiveAction";
            string productPreparedBy = "MyPreparedBy";

            // Arrange
            var homePage = LogInAsAdmin();

            TabletAppPage tabletAppPage = homePage.GotoTabletApp();

            //1. sélectionner un site sur la tablette
            tabletAppPage.Select("mat-select-0", DocSite);

            //2. aller sur ""blank docs""
            tabletAppPage.GotoTabletApp_BlankDoc();
            Thread.Sleep(2000);

            //purge favorite
            tabletAppPage.Select("mat-select-2", DocTitle);
            tabletAppPage.SelectAction("mat-select-2", DocTitle);
            tabletAppPage.Purge(6);
            tabletAppPage.Purge(4);
            //2. aller sur ""blank docs""
            homePage.Navigate();
            tabletAppPage = homePage.GotoTabletApp();
            tabletAppPage.Select("mat-select-0", DocSite);
            tabletAppPage.GotoTabletApp_BlankDoc();
            Thread.Sleep(2000);

            //3. Sélectioner un doc haccp dans le menu déroulant
            tabletAppPage.Select("mat-select-2", DocTitle);

            //4. Cliquer sur next pour ouvrir le document
            tabletAppPage.ClickBaseButton("NEXT STEP");

            //5. choisir une date de production et cliquer sur print
            tabletAppPage.ClickButton("Print");

            // je suis dans un nouvel onglet !!!!
            tabletAppPage.GotoTabletApp_DocumentTab();
            Thread.Sleep(2000);

            //6. cliquer sur ""add line"" remplir une ligne de données
            tabletAppPage.addLineSanitization(DocProduct, quantityProductEntered, disinfection1, disinfection2, commentProduct, correctionAction, productPreparedBy);
            Thread.Sleep(2000);

            //7. Cliquer sur le bouton ""favorite "" , une pop up s'ouvre de confirmation , puis cliquer sur ""validate""
            tabletAppPage.ClickButton("Favorite");

            //6. cliquer sur ""validate""
            tabletAppPage.FillFileName3(DocFileName);
            tabletAppPage.ClickButton("Validate");

            Thread.Sleep(2000);
            // pas de message "voulez-vous écraser le fichier"

            //8.cliquer sur ""save and validate"" et puis cliquer sur ""save favorite"""
            tabletAppPage.ClickButton("Save and validate");
            tabletAppPage.ClickButton("Save favorite");
            tabletAppPage.ClickButton("Ok");


            //8.Revenir vers la page d'acceuil de blank docs , sélectionner le doc haccp dans le menu déroulant

            Thread.Sleep(2000);
            tabletAppPage.Close();
            homePage.Navigate();

            tabletAppPage = homePage.GotoTabletApp();
            //1. sélectionner un site sur la tablette
            tabletAppPage.Select("mat-select-0", DocSite);
            //2. aller sur ""blank docs""
            tabletAppPage.GotoTabletApp_BlankDoc();
            Thread.Sleep(2000);

            //9. Dans la barre des ""temporary"" on devrait retrouver le doc temporary sauvegardé avec la bonne nomination"
            tabletAppPage.Select("mat-select-2", DocTitle);
            // envoi des event remplissant le second combobox
            tabletAppPage.SelectAction("mat-select-2", DocTitle);
            tabletAppPage.SelectAction("mat-select-6", DocFileName, true);

            Thread.Sleep(2000);
            tabletAppPage.ClickBaseButton("NEXT STEP");
            tabletAppPage.ClickButton("Print");
            // je suis dans un nouvel onglet !!!!
            tabletAppPage.GotoTabletApp_DocumentTab();
            Thread.Sleep(2000);

            var columnProject = tabletAppPage.WaitForElementIsVisible(By.Id("mat-input-0"));
            var columnQuatity = tabletAppPage.WaitForElementIsVisible(By.Id("mat-input-1"));
            string firstColumn = columnProject.GetProperty("value");
            string quantitySaved = columnQuatity.GetProperty("value");
            Assert.AreEqual(DocProduct, firstColumn, "Document ne comprend pas " + DocProduct + " dans sa première colonne");
            Assert.AreEqual(quantityProductEntered, quantitySaved, "Document ne comprend pas " + quantityProductEntered + " dans sa première colonne");
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void TB_HACCP_BlankDocs_SaveDocumentFavoriteCheckDate()
        {
            // Prepare
            string DocTitle = "HACCP3 Thawing";
            string DocSite = "GRO";
            string DocFileName = "FavDateTest_Thawing.user01";

            // Arrange
            var homePage = LogInAsAdmin();

            TabletAppPage tabletAppPage = homePage.GotoTabletApp();

            //1. sélectionner un site sur la tablette
            tabletAppPage.Select("mat-select-0", DocSite);

            //2. aller sur ""blank docs""
            tabletAppPage.GotoTabletApp_BlankDoc();
            Thread.Sleep(2000);

            //4. Sélectioner un doc haccp dans le menu déroulant
            tabletAppPage.Select("mat-select-2", DocTitle);
            Thread.Sleep(2000);

            //5. Cliquer sur next pour ouvrir le document
            tabletAppPage.ClickBaseButton("NEXT STEP");

            //6. choisir une date de production et cliquer sur print
            tabletAppPage.ClickButton("Print");

            // je suis dans un nouvel onglet !!!!
            tabletAppPage.GotoTabletApp_DocumentTab();
            Thread.Sleep(2000);

            //7. renseigner une date dans la colonne où le format date est paramétré
            tabletAppPage.addLineThrowning("MyProduct", DateUtils.Now.Date.ToString("dd/MM/yyyy"), "7", "10", DateUtils.Now.Date.AddDays(-10).ToString("dd/MM/yyyy"), DateUtils.Now.Date.AddDays(10).ToString("dd/MM/yyyy"), "MyComment", "MyCorrective", "MyPreparedBy");


            //8. cliquer sur save pour enregistrer un document temporary
            tabletAppPage.ClickButton("Favorite");

            //8. puis cliquer sur ""favorite"" puis sur ""validate"" 
            tabletAppPage.FillFileName3(DocFileName);
            tabletAppPage.ClickButton("Validate");
            Thread.Sleep(2000);

            //puis cliquer sur ""save and validate"" et puis sur ""create record""
            tabletAppPage.ClickButton("Save and validate");
            tabletAppPage.ClickButton("Save favorite");
            tabletAppPage.WaitHACCPHorizontalProgressBar();
            tabletAppPage.ClickButton("Ok");


            //Revenir vers la page d'acceuil de blank docs , sélectionner le doc haccp dans le menu déroulant
            Thread.Sleep(2000);
            homePage.Navigate();

            tabletAppPage = homePage.GotoTabletApp();
            //sélectionner un site sur la tablette
            tabletAppPage.Select("mat-select-0", DocSite);
            //aller sur ""blank docs""
            tabletAppPage.GotoTabletApp_BlankDoc();
            Thread.Sleep(2000);

            //9. récupérer ensuite le document temporary dans la page d'acceuil de blank docs dans la section "" favorite ""
            tabletAppPage.Select("mat-select-2", DocTitle);
            // envoi des event remplissant le second combobox
            tabletAppPage.SelectAction("mat-select-2", DocTitle);
            tabletAppPage.SelectAction("mat-select-6", DocFileName, true);

            tabletAppPage.ClickBaseButton("NEXT STEP");
            // je suis dans un nouvel onglet !!!!
            tabletAppPage.GotoTabletApp_DocumentTab();
            Thread.Sleep(2000);

            //10. vérifier que la date s'est bien enregistrée sur le document temporary
            var date = tabletAppPage.WaitForElementIsVisible(By.Id("mat-input-1"));
            Assert.AreEqual(DateUtils.Now.Date.ToString("dd/MM/yyyy"), date.GetAttribute("value"), "la date n'est pas enregistrée sur le document favori");

            // retour à BlankDoc
            Thread.Sleep(2000);
            tabletAppPage.Close();
            homePage.Navigate();

            tabletAppPage = homePage.GotoTabletApp();
            //1. sélectionner un site sur la tablette
            tabletAppPage.Select("mat-select-0", DocSite);
            //2. aller sur ""blank docs""
            tabletAppPage.GotoTabletApp_BlankDoc();
            Thread.Sleep(2000);

            //purge favorite
            tabletAppPage.Select("mat-select-2", DocTitle);
            tabletAppPage.SelectAction("mat-select-2", DocTitle);
            tabletAppPage.Purge(6);
            tabletAppPage.Purge(4);
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void TB_HACCP_BlankDocs_CheckOpenDocumentHACCP()
        {
            // Prepare
            string DocTitle = "HACCP3 Thawing";
            string DocSite = "GRO";

            // Arrange
            var homePage = LogInAsAdmin();

            TabletAppPage tabletAppPage = homePage.GotoTabletApp();

            //1. sélectionner un site sur la tablette
            tabletAppPage.Select("mat-select-0", DocSite);

            //2. aller sur ""blank docs""
            tabletAppPage.GotoTabletApp_BlankDoc();
            Thread.Sleep(2000);

            //3.sélectionner un document haccp dans le menu déroulant
            tabletAppPage.Select("mat-select-2", DocTitle);
            Thread.Sleep(2000);

            //Cliquer sur NEXT STEP
            tabletAppPage.ClickBaseButton("NEXT STEP");

            //4.Sélectionner une date de production puis cliquer sur print 
            tabletAppPage.ClickButton("Print");

            // je suis dans un nouvel onglet !!!!
            tabletAppPage.GotoTabletApp_DocumentTab();

            // je suis dans un document (je peux faire Favorite puis Cancel)
            tabletAppPage.ClickButton("Favorite");
            tabletAppPage.ClickButton("Cancel");
            tabletAppPage.Close();
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void TB_HACCP_BlankDocs_SaveDocumentHACCPFromTemporary()
        {
            TB_HACCP_BlankDocs_SaveDocumentTemporary();
            // Prepare
            string DocTitle = "HACCP3 Sanitization";
            string DocSite = "GRO";
            string DocFileName = "TempTest_wrfr.user01";
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            //HACCP3 Sanitization_01032022_-_438857_-_20220301095317.pdf
            string DocFileNamePdfBegin = DocTitle + "_";
            //All_files_20220225_102148.zip
            string DocFileNameZipBegin = "All_files_";

            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            TabletAppPage tabletAppPage = homePage.GotoTabletApp();

            //1.sélectionner un site sur la tablette

            tabletAppPage.Select("mat-select-0", DocSite);

            //2.aller sur ""blank docs""
            tabletAppPage.GotoTabletApp_BlankDoc();
            Thread.Sleep(2000);

            //3.Sélectioner un doc haccp dans le menu déroulant
            tabletAppPage.SelectV3("mat-select-2", DocTitle);
            //tabletAppPage.SelectAction("mat-select-2", DocTitle);

            //4.sélectionner un doc temporary qui est enregistré à une date de production X
            tabletAppPage.SelectAction("mat-select-4", DocFileName, true);

            //et cliquer sur next step
            tabletAppPage.ClickBaseButton("NEXT STEP");

            // je suis dans un nouvel onglet !!!!
            tabletAppPage.GotoTabletApp_DocumentTab();

            // FIXME perte de précision 1.313, 2.222, 2.323

            //5. renseigner au moins une donnée de plus
            tabletAppPage.addLineSanitization("MyProduct2", "1.3", "2.2", "2.3", "MyView", "MyCorrectiveAction", "MyPreparedBy", 4);

            //6.cliquer sur ""save and validate"" puis sur ""create record""
            tabletAppPage.ClickButton("Save and validate");
            tabletAppPage.ClickButton("Validate");
            tabletAppPage.WaitPageLoading();
            tabletAppPage.ClickButton("Ok");

            //"Dans winrest dans la partie accouting=> reporting :


            // verify Accouting>Reporting
            homePage.Navigate();
            homePage.ClearDownloads();

            ReportingPage reportingPage = homePage.GoToAccounting_Reporting();
            reportingPage.Purge(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);

            //1.Sélectionner type of report: HACCP
            //2.sélectionner le site et sélectionner la production date X initiale sur laquelle était enregistré le doc temporary
            string DocTypeOfReport = reportingPage.DevPathDocTypeOfRecord();
            reportingPage.FillReportingPage(DocTypeOfReport, DocSite);

            //3.recupérer le rapport HACCP
            var offset = reportingPage.TableHasDocument(DocTitle);
            Assert.IsNotNull(offset, "Document non trouvé");

            // télécharger le document
            reportingPage.PrintDownload(offset);

            var correctDownloadedFile = reportingPage.PrintAllZip(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);

            if (correctDownloadedFile == null)
            {
                // On récupère les fichiers du répertoire de téléchargement
                DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
                FileInfo[] taskFiles = taskDirectory.GetFiles();

                // On recherche le fichier téléchargé
                correctDownloadedFile = reportingPage.GetHACCPReportPdf(taskFiles, DocTitle);
                Assert.IsNotNull(correctDownloadedFile, "Le report n'a pas été téléchargé.");
            }

            PdfDocument document = PdfDocument.Open(correctDownloadedFile);
            Page page1 = document.GetPage(1);
            IEnumerable<Word> words = page1.GetWords();
            Assert.AreEqual(1, words.Count(w => w.Text == "MyProduct"), "MyProduct non présent dans le Pdf");
            Assert.AreEqual(2, words.Count(w => w.Text == "1.3"), "1.3 non présent dans le Pdf");
            Assert.AreEqual(2, words.Count(w => w.Text == "2.2"), "2.2 non présent dans le Pdf");
            Assert.AreEqual(2, words.Count(w => w.Text == "2.3"), "2.3 non présent dans le Pdf");
            Assert.AreEqual(2, words.Count(w => w.Text == "MyView"), "Le commentaire n'est pas retrouvé dans le Pdf");
            Assert.AreEqual(2, words.Count(w => w.Text == "MyCorrectiveAction"), "MyCorrectiveAction non présent dans le Pdf");
            Assert.AreEqual(2, words.Count(w => w.Text == "MyPreparedBy"), "MyPreparedBy non présent dans le Pdf");
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void TB_HACCP_BlankDocs_SaveDocumentHACCPTemporaryFromFavorite()
        {
            // Prepare
            string DocTitle = "HACCP3 Sanitization";
            string DocSite = "MAD";
            string DocFileNameFAV = "TempTestFAV_wrfr.user01" + new Random().Next(1, 500).ToString();
            string DocFileName = "TempTest_wrfr.user01";

            // Arrange
            var homePage = LogInAsAdmin();

            TabletAppPage tabletAppPage = homePage.GotoTabletApp();

            //sélectionner un site sur la tablette
            tabletAppPage.Select("mat-select-0", DocSite);

            //aller sur ""blank docs""
            tabletAppPage.GotoTabletApp_BlankDoc();
            Thread.Sleep(2000);

            //Sélectioner un doc haccp dans le menu déroulant
            tabletAppPage.SelectV3("mat-select-2", DocTitle);
            Thread.Sleep(2000);
            // purge favorite
            tabletAppPage.Purge(6);
            tabletAppPage.Purge(4);

            homePage.Navigate();
            tabletAppPage = homePage.GotoTabletApp();
            tabletAppPage.Select("mat-select-0", DocSite);
            tabletAppPage.GotoTabletApp_BlankDoc();
            Thread.Sleep(2000);

            //Sélectioner un doc haccp dans le menu déroulant
            tabletAppPage.SelectV3("mat-select-2", DocTitle);
            Thread.Sleep(2000);

            //Cliquer sur next pour ouvrir le document
            tabletAppPage.ClickBaseButton("NEXT STEP");

            //choisir une date de production et cliquer sur print
            tabletAppPage.ClickButton("Print");

            // je suis dans un nouvel onglet !!!!
            tabletAppPage.GotoTabletApp_DocumentTab();
            Thread.Sleep(2000);

            //remplir une ligne de données
            tabletAppPage.addLineSanitization("MyProduct", "1.3", "2.2", "2.3", "MyView", "MyCorrectiveAction", "MyPreparedBy");


            //Cliquer sur le bouton ""Favorite"" => winrest propose un nom du favorite par défaut qui peut être modifié
            tabletAppPage.ClickButton("Favorite");

            //cliquer sur ""validate""
            tabletAppPage.FillFileName3(DocFileNameFAV);
            tabletAppPage.ClickButton("Validate");
            Thread.Sleep(2000);

            try
            {
                // erase old file (pourtant on a purgé)
                tabletAppPage.ClickButton("Yes", true);
            }
            catch
            {
                // nouveau fichier
            }

            //FIXME faut passer par "Save and validate" pour "sauvegarder" le favorite
            tabletAppPage.ClickButton("Save and validate");
            tabletAppPage.ClickButton("Save favorite");
            //une pop de validation apparait et puis cliquer sur OK
            tabletAppPage.ClickButton("Ok");

            //Revenir vers la page d'acceuil de blank docs , sélectionner le doc haccp dans le menu déroulant
            tabletAppPage.Close();
            Thread.Sleep(2000);
            homePage.Navigate();

            tabletAppPage = homePage.GotoTabletApp();
            //1. sélectionner un site sur la tablette
            tabletAppPage.Select("mat-select-0", DocSite);
            //2. aller sur ""blank docs""
            tabletAppPage.GotoTabletApp_BlankDoc();
            Thread.Sleep(2000);

            //3. Sélectioner un doc haccp dans le menu déroulant
            tabletAppPage.SelectV3("mat-select-2", DocTitle);
            // envoi des event remplissant le second combobox
            //tabletAppPage.SelectAction("mat-select-2", DocTitle);
            //4. sur la section favori, sélectionner un document favori...
            tabletAppPage.SelectAction("mat-select-6", DocFileNameFAV, true);

            //...puis l'ouvrir en cliquant sur next step
            tabletAppPage.ClickBaseButton("NEXT STEP");

            //3. choisir une date de production et cliquer sur print
            tabletAppPage.ClickButton("Print");

            // je suis dans un nouvel onglet !!!!
            tabletAppPage.GotoTabletApp_DocumentTab();
            tabletAppPage.WaitHACCPHorizontalProgressBar();

            // 5. renseigner au moins une donnée de plus
            tabletAppPage.addLineSanitization("MyProduct2", "1.3", "2.2", "2.3", "MyView", "MyCorrectiveAction", "MyPreparedBy", 4);
            //6. cliquer sur save, une pop up de confirmation s'ouvre et puis cliquer sur validate
            tabletAppPage.ClickButton("Save");

            //cliquer sur ""validate""
            tabletAppPage.FillFileName3(DocFileName);
            tabletAppPage.ClickButton("Validate");

            try
            {
                // erase old file
                tabletAppPage.ClickButton("Yes", true);
            }
            catch
            {
                // nouveau fichier
            }

            //une pop de validation apparait et puis cliquer sur OK
            tabletAppPage.ClickButton("Ok");
            tabletAppPage.Close();

            Thread.Sleep(2000);
            homePage.Navigate();

            tabletAppPage = homePage.GotoTabletApp();
            //1. sélectionner un site sur la tablette
            tabletAppPage.Select("mat-select-0", DocSite);
            //2. aller sur ""blank docs""
            tabletAppPage.GotoTabletApp_BlankDoc();
            Thread.Sleep(2000);

            tabletAppPage.SelectV3("mat-select-2", DocTitle);
            Thread.Sleep(2000);
            tabletAppPage.SelectAction("mat-select-4", DocFileName, true);

            tabletAppPage.ClickBaseButton("NEXT STEP");
            //choisir une date de production et cliquer sur print

            // je suis dans un nouvel onglet !!!!
            tabletAppPage.GotoTabletApp_DocumentTab();

            //...et avec les données enregsitrés y compris les nouvelles modifications
            Assert.AreEqual("MyProduct", tabletAppPage.CheckValue(0), "Docucment ne comprend pas : MyProduct");
            Assert.AreEqual("1.3", tabletAppPage.CheckValue(1), "Document ne comprend pas : MyProduct");
            Assert.AreEqual("2.2", tabletAppPage.CheckValue(2), "Document ne comprend pas : MyProduct");
            Assert.AreEqual("2.3", tabletAppPage.CheckValue(3), "Document ne comprend pas : MyProduct");

            Assert.AreEqual(true, tabletAppPage.CheckValueBox(1), "Document ne comprend pas : Rinsing");
            Assert.AreEqual(true, tabletAppPage.CheckValueBox(2), "Document ne comprend pas : Final");

            Assert.AreEqual("MyView", tabletAppPage.CheckNote(0, 8), "Note Comments");
            Assert.AreEqual("MyCorrectiveAction", tabletAppPage.CheckNote(1, 9), "Note Corrective");
            Assert.AreEqual("MyPreparedBy", tabletAppPage.CheckNote(2, 10), "Note PreparedBy");

            Assert.AreEqual("MyProduct2", tabletAppPage.CheckValue(4), "Document ne comprend pas : MyProduct");
            Assert.AreEqual("1.3", tabletAppPage.CheckValue(5), "Document ne comprend pas : MyProduct");
            Assert.AreEqual("2.2", tabletAppPage.CheckValue(6), "Document ne comprend pas : MyProduct");
            Assert.AreEqual("2.3", tabletAppPage.CheckValue(7), "Document ne comprend pas : MyProduct");

            Assert.AreEqual(false, tabletAppPage.CheckValueBox(3), "Document ne comprend pas : Rinsing");
            Assert.AreEqual(false, tabletAppPage.CheckValueBox(4), "Document ne comprend pas : Final");

            Assert.AreEqual("MyView", tabletAppPage.CheckNote(3, 11), "Note Comments");
            Assert.AreEqual("MyCorrectiveAction", tabletAppPage.CheckNote(4, 12), "Note Corrective");
            Assert.AreEqual("MyPreparedBy", tabletAppPage.CheckNote(5, 13), "Note PreparedBy");
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void TB_HACCP_BlankDocs_SaveDocumentTemporaryNameConflict()
        {
            // Prepare
            string DocTitle = "HACCP3 Sanitization";
            string DocSite = "GRO";
            string DocFileName = "TempTest_wrfr.user01";

            // Arrange
            var homePage = LogInAsAdmin();

            TabletAppPage tabletAppPage = homePage.GotoTabletApp();

            //1. sélectionner un site sur la tablette
            tabletAppPage.Select("mat-select-0", DocSite);

            //2. aller sur ""blank docs""
            tabletAppPage.GotoTabletApp_BlankDoc();
            Thread.Sleep(2000);

            //1. Sélectioner un doc haccp dans le menu déroulant
            tabletAppPage.SelectV3("mat-select-2", DocTitle);
            tabletAppPage.Purge(6);
            tabletAppPage.Purge(4);

            homePage.Navigate();
            tabletAppPage = homePage.GotoTabletApp();
            tabletAppPage.Select("mat-select-0", DocSite);
            tabletAppPage.GotoTabletApp_BlankDoc();
            Thread.Sleep(2000);

            //Sélectioner un doc haccp dans le menu déroulant
            tabletAppPage.SelectV3("mat-select-2", DocTitle);
            Thread.Sleep(2000);

            //2. Cliquer sur next pour ouvrir le document
            tabletAppPage.ClickBaseButton("NEXT STEP");

            //3. choisir une date de production et cliquer sur print
            tabletAppPage.ClickButton("Print");

            // je suis dans un nouvel onglet !!!!
            tabletAppPage.GotoTabletApp_DocumentTab();
            Thread.Sleep(2000);

            //4. remplir une ligne de données
            tabletAppPage.addLineSanitization("MyProduct", "1.3", "2.2", "2.3", "MyView", "MyCorrectiveAction", "MyPreparedBy");


            //5. Cliquer sur le bouton ""save "" => winrest propose un nom du temporary par défaut qui peut être modifié
            tabletAppPage.ClickButton("Save");

            //6. cliquer sur ""validate""
            tabletAppPage.FillFileName3(DocFileName);
            tabletAppPage.ClickButton("Validate");

            tabletAppPage.WaitHACCPHorizontalProgressBar();
            try
            {
                // erase old file (pourtant on a purge)
                tabletAppPage.ClickButton("Yes", true);
            }
            catch
            {
                // nouveau fichier
            }

            //7. une pop de validation apparait et puis cliquer sur OK
            tabletAppPage.ClickButton("Ok");

            //8.Revenir vers la page d'acceuil de blank docs , sélectionner le doc haccp dans le menu déroulant
            tabletAppPage.Close();
            Thread.Sleep(2000);
            homePage.Navigate();

            tabletAppPage = homePage.GotoTabletApp();
            //1. sélectionner un site sur la tablette
            tabletAppPage.Select("mat-select-0", DocSite);
            //2. aller sur ""blank docs""
            tabletAppPage.GotoTabletApp_BlankDoc();
            Thread.Sleep(2000);

            //9. Dans la barre des ""temporary"" on devrait retrouver le doc temporary sauvegardé avec la bonne nomination"
            tabletAppPage.SelectV3("mat-select-2", DocTitle);
            // envoi des event remplissant le second combobox
            //tabletAppPage.SelectAction("mat-select-2", DocTitle);

            //2. Cliquer sur next pour ouvrir le document
            tabletAppPage.ClickBaseButton("NEXT STEP");

            //3. choisir une date de production et cliquer sur print
            tabletAppPage.ClickButton("Print");

            // je suis dans un nouvel onglet !!!!
            tabletAppPage.GotoTabletApp_DocumentTab();
            Thread.Sleep(2000);

            //4. remplir une ligne de données - ecrasement
            tabletAppPage.addLineSanitization("MyProduct2NameConflict", "2.6", "2.9", "2.7", "MyViewNameConflict", "MyCorrectiveActionNameConflict", "MyPreparedByNameConflict");


            //5. Cliquer sur le bouton ""save "" => winrest propose un nom du temporary par défaut qui peut être modifié
            tabletAppPage.ClickButton("Save");

            //6. cliquer sur ""validate""
            tabletAppPage.FillFileName3(DocFileName);
            tabletAppPage.ClickButton("Validate");

            // erase old file
            tabletAppPage.ClickButton("Yes");

            //7. une pop de validation apparait et puis cliquer sur OK
            tabletAppPage.ClickButton("Ok");

            tabletAppPage.Close();
            Thread.Sleep(2000);
            homePage.Navigate();

            tabletAppPage = homePage.GotoTabletApp();
            //1. sélectionner un site sur la tablette
            tabletAppPage.Select("mat-select-0", DocSite);
            //2. aller sur ""blank docs""
            tabletAppPage.GotoTabletApp_BlankDoc();
            Thread.Sleep(2000);

            tabletAppPage.SelectV3("mat-select-2", DocTitle);
            // envoi des event remplissant le second combobox
            //tabletAppPage.SelectAction("mat-select-2", DocTitle);
            tabletAppPage.SelectAction("mat-select-4", DocFileName, true);

            //2. Cliquer sur next pour ouvrir le document
            tabletAppPage.ClickBaseButton("NEXT STEP");

            //3. choisir une date de production et cliquer sur print

            // en patch : tabletAppPage.ClickButton("Print");
            // je suis dans un nouvel onglet !!!!
            tabletAppPage.GotoTabletApp_DocumentTab();
            Thread.Sleep(2000);
            bool isProductOverwritten = tabletAppPage.CheckValue(0).Equals("MyProduct2NameConflict");
            bool isQuantityOverwritten = tabletAppPage.CheckValue(1).Equals("2.6");
            bool isDisinfection1Overwritten = tabletAppPage.CheckValue(2).Equals("2.9");
            bool isUpdatedData = isProductOverwritten && isQuantityOverwritten && isDisinfection1Overwritten;
            Assert.IsTrue(isUpdatedData, "L'ancien fichier n'est pas écrasé ni remplacé");
            //Assert.AreEqual("2.6", tabletAppPage.CheckValue(1), "L'ancien fichier n'est pas écrasé ni remplacé");
            //Assert.AreEqual("2.0", tabletAppPage.CheckValue(2), "L'ancien fichier n'est pas écrasé ni remplacé");

        }

        [TestMethod]
        [Timeout(Timeout)]
        public void TB_HACCP_BlankDocs_SaveDocumentCheckMultiDate()
        {
            string DocTitle = "HACCP3 Test";
            string DocSite = "GRO";
            //HACCP3 Sanitization_01032022_-_438857_-_20220301095317.pdf
            string DocFileNamePdfBegin = DocTitle + "_";
            //All_files_20220225_102148.zip
            string DocFileNameZipBegin = "All_files_";
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();


            var homePage = LogInAsAdmin();

            TabletAppPage tabletAppPage = homePage.GotoTabletApp();

            //1. sélectionner un site sur la tablette
            tabletAppPage.Select("mat-select-0", DocSite);

            //2. aller sur ""blank docs""
            tabletAppPage.GotoTabletApp_BlankDoc();
            Thread.Sleep(2000);

            //3. Sélectioner un doc haccp dans le menu déroulant
            tabletAppPage.Select("mat-select-2", DocTitle);
            tabletAppPage.SelectAction("mat-select-2", DocTitle);

            //4. Cliquer sur next pour ouvrir le document
            tabletAppPage.ClickBaseButton("NEXT STEP");
            //5. choisir une date de production et cliquer sur print
            tabletAppPage.ClickButton("Print");

            // je suis dans un nouvel onglet !!!!
            tabletAppPage.GotoTabletApp_DocumentTab();

            //6. renseigner des dates dans la colonne où le format date est paramétré
            tabletAppPage.addLineTest(false, true);

            // pour télécharger le pdf

            //7. puis cliquer sur ""save and validate"" puis cliquer sur ""validate""
            tabletAppPage.ClickButton("Save and validate");
            tabletAppPage.ClickButton("Validate");
            tabletAppPage.WaitHACCPHorizontalProgressBar();
            tabletAppPage.ClickButton("Ok");

            //"Dans winrest dans la partie accouting=> reporting :

            tabletAppPage.Close();
            // verify Accouting>Reporting
            homePage.Navigate();
            homePage.ClearDownloads();

            ReportingPage reportingPage = homePage.GoToAccounting_Reporting();
            reportingPage.Purge(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);

            //1.Sélectionner type of report: HACCP
            //2.sélectionner le site et sélectionner la production date X initiale sur laquelle était enregistré le doc temporary
            string DocTypeOfReport = reportingPage.DevPathDocTypeOfRecord();
            reportingPage.FillReportingPage(DocTypeOfReport, DocSite);

            //3.recupérer le rapport HACCP
            var offset = reportingPage.TableHasDocument(DocTitle);
            Assert.IsNotNull(offset, "Document non trouvé");

            // télécharger le document
            reportingPage.PrintDownload(offset);

            var correctDownloadedFile = reportingPage.PrintAllZip(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);

            if (correctDownloadedFile == null)
            {
                // On récupère les fichiers du répertoire de téléchargement
                DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
                FileInfo[] taskFiles = taskDirectory.GetFiles();

                // On recherche le fichier téléchargé
                correctDownloadedFile = reportingPage.GetHACCPReportPdf(taskFiles, DocTitle);
                Assert.IsNotNull(correctDownloadedFile, "Le report n'a pas été téléchargé.");
            }

            PdfDocument document = PdfDocument.Open(correctDownloadedFile);

            //4. On retrouve bien les dates renseignées et les dates doivent être conformes dans le rapport généré dans le bon format également défini dans les global settings (settings=> global settings=> DateFormatDisplay)"
            Page page1 = document.GetPage(1);
            IEnumerable<Word> words = page1.GetWords();
            // du 10 au 17 de ce mois
            //10/02/2022 , 11/02/2022 , 12/02/2022 , 13/02/2022 , 14/02/2022 , 15/02/2022 , 16/02/2022
            string yearAndMonth = DateUtils.Now.ToString("/MM/yyyy");
            // date de production tombant le 17
            if (DateUtils.Now.Day != 17)
            {
                Assert.IsFalse(words.Any(w => w.Text == (17 + yearAndMonth)), (17 + yearAndMonth) + " MultiDate trouvé dans " + correctDownloadedFile);
            }
            for (int i = 10; i < 17; i++)
            {
                Assert.IsTrue(words.Any(w => w.Text == (i + yearAndMonth)), (i + yearAndMonth) + " MultiDate non trouvé dans " + correctDownloadedFile);
            }
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void TB_HACCP_BlankDocs_SaveDocumentTemporaryCheckMultiDate()
        {
            string DocTitle = "HACCP3 Test";
            string DocSite = "GRO";
            string DocFileName = "TempTestMultiDate_wrfr.user01";
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            //HACCP3 Sanitization_01032022_-_438857_-_20220301095317.pdf
            string DocFileNamePdfBegin = DocTitle + "_";
            //All_files_20220225_102148.zip
            string DocFileNameZipBegin = "All_files_";

            var homePage = LogInAsAdmin();

            TabletAppPage tabletAppPage = homePage.GotoTabletApp();

            //1. sélectionner un site sur la tablette

            tabletAppPage.Select("mat-select-0", DocSite);

            //2. aller sur ""blank docs""
            tabletAppPage.GotoTabletApp_BlankDoc();
            Thread.Sleep(2000);

            //4. Sélectioner un doc haccp dans le menu déroulant
            tabletAppPage.SelectV3("mat-select-2", DocTitle);
            tabletAppPage.Purge(6);
            tabletAppPage.Purge(4);
            //tabletAppPage.SelectAction("mat-select-2", DocTitle);

            //5. Cliquer sur next pour ouvrir le document
            tabletAppPage.ClickBaseButton("NEXT STEP");
            //6. choisir une date de production et cliquer sur print
            tabletAppPage.ClickButton("Print");

            // je suis dans un nouvel onglet !!!!
            tabletAppPage.GotoTabletApp_DocumentTab();

            //7. renseigner plusieurs dates dans la colonne où le format date est paramétré
            tabletAppPage.addLineTest(false, true);

            //8. cliquer sur save pour enregistrer un document temporary
            tabletAppPage.ClickButton("Save");
            //cliquer sur ""validate""
            tabletAppPage.FillFileName0(DocFileName);
            tabletAppPage.ClickButton("Validate");

            //try
            //{
            // erase old file
            //tabletAppPage.WaitForLoad();
            //tabletAppPage.ClickButton("Yes");
            //}
            //catch
            //{
            // nouveau fichier
            //}

            //une pop de validation apparait et puis cliquer sur OK
            tabletAppPage.ClickButton("Yes");

            tabletAppPage.Close();
            Thread.Sleep(2000);
            homePage.Navigate();
            tabletAppPage = homePage.GotoTabletApp();
            tabletAppPage.Select("mat-select-0", DocSite);
            //9. récupérer ensuite le document temporary dans la page d'acceuil de blank docs dans la section "" temporary""
            tabletAppPage.GotoTabletApp_BlankDoc();
            Thread.Sleep(2000);
            tabletAppPage.SelectV3("mat-select-2", DocTitle);
            //tabletAppPage.SelectAction("mat-select-2", DocTitle);
            tabletAppPage.SelectAction("mat-select-4", DocFileName, true);

            tabletAppPage.ClickBaseButton("NEXT STEP");
            // patch : tabletAppPage.ClickButton("Print");
            // je suis dans un nouvel onglet !!!!
            tabletAppPage.GotoTabletApp_DocumentTab();
            Thread.Sleep(2000);
            //10. vérifier que les date sont bien enregistrées sur le document temporary
            Assert.IsTrue(tabletAppPage.CheckMultiDate(), "CheckMultiDate KO");

            tabletAppPage.Close();
            // verify Accouting>Reporting
            homePage.Navigate();
            homePage.ClearDownloads();

            ReportingPage reportingPage = homePage.GoToAccounting_Reporting();
            reportingPage.Purge(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);

            //1.Sélectionner type of report: HACCP
            //2.sélectionner le site et sélectionner la production date X initiale sur laquelle était enregistré le doc temporary
            string DocTypeOfReport = reportingPage.DevPathDocTypeOfRecord();
            reportingPage.FillReportingPage(DocTypeOfReport, DocSite);

            //3.recupérer le rapport HACCP
            var offset = reportingPage.TableHasDocument(DocTitle);
            Assert.IsNotNull(offset, "Document non trouvé");

            // télécharger le document
            reportingPage.PrintDownload(offset);

            var correctDownloadedFile = reportingPage.PrintAllZip(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);

            if (correctDownloadedFile == null)
            {
                // On récupère les fichiers du répertoire de téléchargement
                DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
                FileInfo[] taskFiles = taskDirectory.GetFiles();

                // On recherche le fichier téléchargé
                correctDownloadedFile = reportingPage.GetHACCPReportPdf(taskFiles, DocTitle);
                Assert.IsNotNull(correctDownloadedFile, "Le report n'a pas été téléchargé.");
            }

            PdfDocument document = PdfDocument.Open(correctDownloadedFile);

            //4. On retrouve bien les dates renseignées et les dates doivent être conformes dans le rapport généré dans le bon format également défini dans les global settings (settings=> global settings=> DateFormatDisplay)"

            Page page1 = document.GetPage(1);
            IEnumerable<Word> words = page1.GetWords();
            // du 10 au 17 de ce mois
            //10/02/2022 , 11/02/2022 , 12/02/2022 , 13/02/2022 , 14/02/2022 , 15/02/2022 , 16/02/2022
            string yearAndMonth = DateUtils.Now.ToString("/MM/yyyy");
            // date de production tombant le 17
            if (DateUtils.Now.Day != 17)
            {
                Assert.IsFalse(words.Any(w => w.Text == (17 + yearAndMonth)), (17 + yearAndMonth) + " MultiDate trouvé dans " + correctDownloadedFile);
            }
            for (int i = 10; i < 17; i++)
            {
                Assert.IsTrue(words.Any(w => w.Text == (i + yearAndMonth)), (i + yearAndMonth) + " MultiDate non trouvé dans " + correctDownloadedFile);
            }
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void TB_HACCP_BlankDocs_SaveDocumentFavoriteCheckMultiDate()
        {
            string DocTitle = "HACCP3 Test";
            string DocSite = "GRO";
            string DocFileName = "FavTestMultiDate_wrfr3.user01";
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            //HACCP3 Sanitization_01032022_-_438857_-_20220301095317.pdf
            string DocFileNamePdfBegin = DocTitle + "_";
            //All_files_20220225_102148.zip
            string DocFileNameZipBegin = "All_files_";

            var homePage = LogInAsAdmin();

            TabletAppPage tabletAppPage = homePage.GotoTabletApp();

            //1. sélectionner un site sur la tablette

            tabletAppPage.Select("mat-select-0", DocSite);

            //2. aller sur ""blank docs""
            tabletAppPage.GotoTabletApp_BlankDoc();

            //4. Sélectioner un doc haccp dans le menu déroulant
            tabletAppPage.Select("mat-select-2", DocTitle);
            tabletAppPage.SelectAction("mat-select-2", DocTitle);
            // purge favorite
            tabletAppPage.Purge(6);
            tabletAppPage.Purge(4);
            homePage.Navigate();
            tabletAppPage = homePage.GotoTabletApp();
            tabletAppPage.Select("mat-select-0", DocSite);
            tabletAppPage.GotoTabletApp_BlankDoc();
            tabletAppPage.Select("mat-select-2", DocTitle);
            tabletAppPage.SelectAction("mat-select-2", DocTitle);

            //5. Cliquer sur next pour ouvrir le document
            tabletAppPage.ClickBaseButton("NEXT STEP");
            //6. choisir une date de production et cliquer sur print
            tabletAppPage.ClickButton("Print");

            // je suis dans un nouvel onglet !!!!
            tabletAppPage.GotoTabletApp_DocumentTab();
            Thread.Sleep(2000);

            //7. renseigner plusieurs dates dans la colonne où le format date est paramétré
            tabletAppPage.addLineTest(false, true);
            Thread.Sleep(2000);

            //8. puis cliquer sur ""favorite"" puis sur ""validate"" puis cliquer sur ""save and validate"" et puis sur ""create record""

            tabletAppPage.SavefavoriteDocument(DocFileName);

            tabletAppPage.Close();
            //9. récupérer ensuite le document temporary dans la page d'acceuil de blank docs dans la section "" favorite ""
            tabletAppPage.Select("mat-select-2", DocTitle);
            tabletAppPage.SelectAction("mat-select-6", DocFileName, true);

            tabletAppPage.ClickBaseButton("NEXT STEP");
            tabletAppPage.ClickButton("Print");
            Thread.Sleep(2000);
            // je suis dans un nouvel onglet !!!!
            tabletAppPage.GotoTabletApp_DocumentTab();
            Thread.Sleep(2000);

            //10. vérifier que les dates sont bien enregistrées sur le document favori
            Assert.IsTrue(tabletAppPage.CheckMultiDate(), "CheckMultiDate KO");

            tabletAppPage.Close();
            // verify Accouting>Reporting
            homePage.Navigate();
            homePage.ClearDownloads();

            ReportingPage reportingPage = homePage.GoToAccounting_Reporting();
            reportingPage.Purge(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);

            //1.Sélectionner type of report: HACCP
            //2.sélectionner le site et sélectionner la production date X initiale sur laquelle était enregistré le doc temporary
            string DocTypeOfReport = reportingPage.DevPathDocTypeOfRecord();
            reportingPage.FillReportingPage(DocTypeOfReport, DocSite);

            //3.recupérer le rapport HACCP
            var offset = reportingPage.TableHasDocument(DocTitle);
            Assert.IsNotNull(offset, "Document non trouvé");

            // télécharger le document
            reportingPage.PrintDownload(offset);

            var correctDownloadedFile = reportingPage.PrintAllZip(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);

            if (correctDownloadedFile == null)
            {
                // On récupère les fichiers du répertoire de téléchargement
                DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
                FileInfo[] taskFiles = taskDirectory.GetFiles();

                // On recherche le fichier téléchargé
                correctDownloadedFile = reportingPage.GetHACCPReportPdf(taskFiles, DocTitle);
                Assert.IsNotNull(correctDownloadedFile, "Le report n'a pas été téléchargé.");
            }

            PdfDocument document = PdfDocument.Open(correctDownloadedFile);

            //4. On retrouve bien les dates renseignées et les dates doivent être conformes dans le rapport généré dans le bon format également défini dans les global settings (settings=> global settings=> DateFormatDisplay)"

            Page page1 = document.GetPage(1);
            IEnumerable<Word> words = page1.GetWords();
            // du 10 au 17 de ce mois
            //10/02/2022 , 11/02/2022 , 12/02/2022 , 13/02/2022 , 14/02/2022 , 15/02/2022 , 16/02/2022
            string yearAndMonth = DateUtils.Now.ToString("/MM/yyyy");
            // date de production tombant le 17
            if (DateUtils.Now.Day != 17)
            {
                Assert.IsFalse(words.Any(w => w.Text == (17 + yearAndMonth)), (17 + yearAndMonth) + " MultiDate trouvé dans " + correctDownloadedFile);
            }
            for (int i = 10; i < 17; i++)
            {
                Assert.IsTrue(words.Any(w => w.Text == (i + yearAndMonth)), (i + yearAndMonth) + " MultiDate non trouvé dans " + correctDownloadedFile);
            }
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void TB_HACCP_BlankDocs_PrintFinalDocument()
        {
            // Prepare
            string DocTitle = "HACCP3 Sanitization";
            string DocSite = "GRO";
            string DocFileName = "HACCP3 Sanitization_" + DocSite + "_" + DateUtils.Now.Date.ToString("ddMMyyyy") + "_" + DateUtils.Now.Date.ToString("ddMMyyyy") + ".pdf";
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();

            // Arrange
            var homePage = LogInAsAdmin();

            string directory = TestContext.Properties["DownloadsPath"].ToString();
            foreach (string f in Directory.GetFiles(directory))
            {
                //purge
                if (f.Contains("HACCP3 Sanitization") && f.EndsWith(".pdf"))
                {
                    File.Delete(f);
                }
            }

            TabletAppPage tabletAppPage = homePage.GotoTabletApp();

            tabletAppPage.Purge(downloadsPath, "HACCP3 Sanitization_" + DocSite + "_");

            //1. sélectionner un site sur la tablette
            tabletAppPage.Select("mat-select-0", DocSite);

            //2. aller sur ""blank docs""
            tabletAppPage.GotoTabletApp_BlankDoc();
            tabletAppPage.WaitPageLoading();

            //3. Sélectioner un doc haccp dans le menu déroulant
            tabletAppPage.Select("mat-select-2", DocTitle);
            tabletAppPage.WaitPageLoading();

            //4. Cliquer sur next pour ouvrir le document
            tabletAppPage.ClickBaseButton("NEXT STEP");

            //5. choisir une date de production et cliquer sur print
            tabletAppPage.ClickButton("Print");

            // je suis dans un nouvel onglet !!!!
            tabletAppPage.GotoTabletApp_DocumentTab();
            tabletAppPage.WaitPageLoading();

            //6. cliquer sur ""add line"" remplir une ligne de données
            tabletAppPage.addLineSanitization("MyProduct", "1.3", "2.2", "2.3", "MyView", "MyCorrectiveAction", "MyPreparedBy");
            tabletAppPage.WaitPageLoading();
            //7. Cliquer sur le bouton ""print ""
            tabletAppPage.ClickButton("Print");
            FileInfo fi = new FileInfo(Path.Combine(downloadsPath, DocFileName));
            int counter = 0;
            while (!fi.Exists && counter < 20)
            {
                tabletAppPage.WaitPageLoading();
                fi.Refresh();
                counter++;
            }
            Assert.IsTrue(fi.Exists, DocFileName + " non généré");
            //1. le rapport haccp est généré en PDF avec toutes les données pré remplies
            //2. On retrouve dans le rapport haccp pdf les même données remplies dans le document haccp
            PdfDocument document = PdfDocument.Open(fi.FullName);
            Page page1 = document.GetPage(1);
            IEnumerable<Word> words = page1.GetWords();
            List<Word> cherchePdf = words.Where(word => word.Text == "MyProduct").ToList<Word>();
            Assert.IsTrue(cherchePdf.Count == 1, "MyProduct non trouvé dans " + fi.FullName);
            cherchePdf = words.Where(word => word.Text == "2.2").ToList<Word>();
            Assert.IsTrue(cherchePdf.Count == 1, "2.2 non trouvé dans " + fi.FullName);
            cherchePdf = words.Where(word => word.Text == "MyView").ToList<Word>();
            Assert.IsTrue(cherchePdf.Count == 1, "Le commentaire n'est pas retrouvé dans " + fi.FullName);
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void TB_HACCP_BlankDocs_PrintTemporaryDocument()
        {
            // Prepare
            string DocTitle = "HACCP3 Sanitization";
            string DocSite = "GRO";
            string DocFileName = "HACCP3 Sanitization_" + DocSite + "_" + DateUtils.Now.Date.ToString("ddMMyyyy") + "_" + DateUtils.Now.Date.ToString("ddMMyyyy") + ".pdf";
            if (DateUtils.IsBeforeMidnight())
            {
                DocFileName = "HACCP3 Sanitization_" + DocSite + "_" + DateTime.Now.Date.ToString("ddMMyyyy") + "_" + DateTime.Now.Date.ToString("ddMMyyyy") + ".pdf";
            }
            string DocTempName = "Temp.user";
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();

            // Arrange
            var homePage = LogInAsAdmin();

            TabletAppPage tabletAppPage = homePage.GotoTabletApp();

            tabletAppPage.Purge(downloadsPath, "HACCP3 Sanitization_" + DocSite + "_");

            //sélectionner un site sur la tablette
            tabletAppPage.Select("mat-select-0", DocSite);

            //aller sur ""blank docs""
            tabletAppPage.GotoTabletApp_BlankDoc();
            Thread.Sleep(2000);

            //Sélectioner un doc haccp dans le menu déroulant
            tabletAppPage.Select("mat-select-2", DocTitle);
            //purge favorite
            tabletAppPage.Purge(6);
            tabletAppPage.Purge(4);

            homePage.Navigate();
            tabletAppPage = homePage.GotoTabletApp();
            tabletAppPage.Select("mat-select-0", DocSite);
            tabletAppPage.GotoTabletApp_BlankDoc();
            Thread.Sleep(2000);

            //Sélectioner un doc haccp dans le menu déroulant
            tabletAppPage.SelectV3("mat-select-2", DocTitle);
            Thread.Sleep(2000);

            //Cliquer sur next pour ouvrir le document
            tabletAppPage.ClickBaseButton("NEXT STEP");

            //choisir une date de production et cliquer sur print
            tabletAppPage.ClickButton("Print");

            // je suis dans un nouvel onglet !!!!
            tabletAppPage.GotoTabletApp_DocumentTab();
            Thread.Sleep(2000);

            //cliquer sur ""add line"" remplir une ligne de données
            tabletAppPage.addLineSanitization("MyProduct", "1.3", "2.2", "2.3", "MyView", "MyCorrectiveAction", "MyPreparedBy");
            Thread.Sleep(2000);
            //Cliquer sur le bouton ""Save ""
            tabletAppPage.ClickButton("Save");
            tabletAppPage.FillFileName3(DocTempName);

            tabletAppPage.ClickButton("Validate");

            try
            {
                // erase old file (pourtant on a purge)
                tabletAppPage.ClickButton("Yes", true);
            }
            catch
            {
                // nouveau fichier
            }

            //une pop de validation apparait et puis cliquer sur OK
            tabletAppPage.ClickButton("Ok");

            //Revenir vers la page d'acceuil de blank docs , sélectionner le doc haccp dans le menu déroulant
            tabletAppPage.Close();
            Thread.Sleep(2000);
            homePage.Navigate();

            tabletAppPage = homePage.GotoTabletApp();
            //1. sélectionner un site sur la tablette
            tabletAppPage.Select("mat-select-0", DocSite);
            //2. aller sur ""blank doc""
            tabletAppPage.GotoTabletApp_BlankDoc();
            Thread.Sleep(2000);

            //3. Sélectioner un doc haccp dans le menu déroulant
            tabletAppPage.SelectV3("mat-select-2", DocTitle);
            // envoi des event remplissant le second combobox
            //tabletAppPage.SelectAction("mat-select-2", DocTitle);
            //4. sélectionner un doc temporary
            tabletAppPage.SelectAction("mat-select-4", DocTempName, true);

            tabletAppPage.ClickBaseButton("NEXT STEP");
            Thread.Sleep(2000);
            // je suis dans un nouvel onglet !!!!
            tabletAppPage.GotoTabletApp_DocumentTab();
            Thread.Sleep(2000);

            //5.Cliquer sur le bouton ""print """

            tabletAppPage.ClickButton("Print");
            FileInfo fi = new FileInfo(Path.Combine(downloadsPath, DocFileName));
            int counter = 0;
            while (!fi.Exists && counter < 10)
            {
                Thread.Sleep(1000);
                fi.Refresh();
                counter++;
            }
            Assert.IsTrue(fi.Exists, DocFileName + " non généré");
            //1. le rapport haccp est généré en PDF avec toutes les données pré remplies
            //2. On retrouve dans le rapport haccp pdf les même données remplies dans le document haccp
            PdfDocument document = PdfDocument.Open(fi.FullName);
            Page page1 = document.GetPage(1);
            IEnumerable<Word> words = page1.GetWords();
            List<Word> cherchePdf = words.Where(word => word.Text == "MyProduct").ToList<Word>();
            Assert.IsTrue(cherchePdf.Count == 1, "MyProduct non trouvé dans " + fi.FullName);
            cherchePdf = words.Where(word => word.Text == "2.2").ToList<Word>();
            Assert.IsTrue(cherchePdf.Count == 1, "2.2 non trouvé dans " + fi.FullName);
            cherchePdf = words.Where(word => word.Text == "MyView").ToList<Word>();
            Assert.IsTrue(cherchePdf.Count == 1, "Le commentaire n'est pas retrouvé dans " + fi.FullName);
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void TB_HACCP_BlankDocs_CheckBoxCNA()
        {
            // Prepare
            string docTitle = "HACCP3 Test";
            string docSite = "GRO";
            string docFileName = $"{docTitle}_{docSite}_{DateUtils.Now.Date:ddMMyyyy}_{DateUtils.Now.Date:ddMMyyyy}.pdf";
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();

            // Arrange
            var homePage = LogInAsAdmin();
            var tabletAppPage = homePage.GotoTabletApp();

            // Purger les fichiers existants
            tabletAppPage.Purge(downloadsPath, $"{docTitle}_{docSite}_");

            // Sélectionner le site et accéder à la page des documents vierges
            tabletAppPage.Select("mat-select-0", docSite);
            tabletAppPage.GotoTabletApp_BlankDoc();
            tabletAppPage.WaitPageLoading();

            // Sélectionner un document HACCP dans le menu déroulant
            tabletAppPage.Select("mat-select-2", docTitle);
            tabletAppPage.WaitPageLoading();

            // Cliquer sur "Next Step" pour ouvrir le document
            tabletAppPage.ClickBaseButton("NEXT STEP");

            // Choisir une date de production et imprimer
            tabletAppPage.ClickButton("Print");
            tabletAppPage.GotoTabletApp_DocumentTab();
            tabletAppPage.WaitPageLoading();

            // Ajouter des lignes de données
            tabletAppPage.addLineTest(true, true);
            tabletAppPage.addLineTest(false, true, 3);

            // Vérifier les valeurs des cases à cocher CNA
            Assert.IsTrue(tabletAppPage.CheckValueBox(1), "La case CNA de la ligne 1 n'est pas correcte");
            Assert.IsFalse(tabletAppPage.CheckValueBox(3), "La case CNA de la ligne 3 est incorrectement cochée");

            // Imprimer et vérifier que le fichier PDF est généré
            tabletAppPage.ClickButton("Print");
            FileInfo pdfFile = new FileInfo(Path.Combine(downloadsPath, docFileName));
            int retries = 0;

            while (!pdfFile.Exists && retries < 20)
            {
                tabletAppPage.WaitPageLoading();
                pdfFile.Refresh();
                retries++;
            }

            Assert.IsTrue(pdfFile.Exists, $"{docFileName} non généré");

            // Vérifier le contenu du fichier PDF
            PdfDocument document = PdfDocument.Open(pdfFile.FullName);
            Page page = document.GetPage(1);
            IEnumerable<Word> words = page.GetWords();

            // Vérifier les occurrences de "C"
            int countC = words.Count(word => word.Text == "C");
            Assert.IsTrue(countC == 3, $"C non trouvé ou incorrect ({countC}/3) dans {pdfFile.FullName}");

            // Vérifier les occurrences de "NA"
            int countNA = words.Count(word => word.Text == "NA");
            Assert.IsTrue(countNA == 1, $"NA non trouvé ou incorrect ({countNA}/1) dans {pdfFile.FullName}");
        }


        [TestMethod]
        [Timeout(Timeout)]
        public void TB_HACCP_BlankDocs_CheckBoxCNATemporary()
        {
            // Préparation des données
            string docTitle = "HACCP3 Test";
            string docSite = "GRO";
            string docTempName = "TempCNAv2.user";
            string docFileName = $"{docTitle}_{docSite}_{DateUtils.Now.Date:ddMMyyyy}_{DateUtils.Now.Date:ddMMyyyy}.pdf";
            string downloadsPath = TestContext.Properties["DownloadsPath"].ToString();

            // Connexion et navigation initiale
            var homePage = LogInAsAdmin();
            var tabletAppPage = homePage.GotoTabletApp();
            tabletAppPage.Purge(downloadsPath, $"{docTitle}_{docSite}_");

            // Sélection du site et navigation vers "blank docs"
            tabletAppPage.Select("mat-select-0", docSite);
            tabletAppPage.GotoTabletApp_BlankDoc();
            tabletAppPage.WaitPageLoading();

            // Purge des documents temporaires existants
            tabletAppPage.SelectV3("mat-select-2", docTitle);
            tabletAppPage.Purge(6); // Favoris
            tabletAppPage.Purge(4); // Documents temporaires

            // Rafraîchissement de l'application et sélection du document
            homePage.Navigate();
            tabletAppPage = homePage.GotoTabletApp();
            tabletAppPage.Select("mat-select-0", docSite);
            tabletAppPage.GotoTabletApp_BlankDoc();
            tabletAppPage.WaitPageLoading();
            tabletAppPage.SelectV3("mat-select-2", docTitle);
            tabletAppPage.WaitPageLoading();

            // Ouverture et impression du document
            tabletAppPage.ClickBaseButton("NEXT STEP");
            tabletAppPage.ClickButton("Print");
            tabletAppPage.GotoTabletApp_DocumentTab();
            tabletAppPage.WaitHACCPHorizontalProgressBar();

            // Ajout d'une ligne et sauvegarde du document comme temporaire
            tabletAppPage.addLineTest(true, true);
            tabletAppPage.ClickButton("Save");
            tabletAppPage.FillFileName0(docTempName);
            tabletAppPage.ClickButton("Validate");

            // Gestion de la confirmation de remplacement
            try
            {
                tabletAppPage.ClickButton("Yes", true);
            }
            catch
            {
                // Pas de fichier existant, aucune action requise
            }

            // Confirmation de l'enregistrement
            tabletAppPage.ClickButton("Ok");

            // Navigation vers les documents temporaires
            tabletAppPage.Close();
            tabletAppPage.WaitPageLoading();
            homePage.Navigate();
            tabletAppPage = homePage.GotoTabletApp();
            tabletAppPage.Select("mat-select-0", docSite);
            tabletAppPage.GotoTabletApp_BlankDoc();
            tabletAppPage.SelectV3("mat-select-2", docTitle);
            tabletAppPage.SelectAction("mat-select-4", docTempName, true);

            // Réouverture du document temporaire
            tabletAppPage.ClickBaseButton("NEXT STEP");
            tabletAppPage.GotoTabletApp_DocumentTab();
            tabletAppPage.addLineTest(false, true, 1, 2);

            // Vérification des cases à cocher
            Assert.IsTrue(tabletAppPage.CheckValueBox(1), "Checkbox pour la ligne 1 devrait être cochée.");
            Assert.IsFalse(tabletAppPage.CheckValueBox(3), "Checkbox pour la ligne 3 ne devrait pas être cochée.");

            // Génération du PDF
            tabletAppPage.ClickButton("Print");
            tabletAppPage.WaitForDownloadHACCP();

            // Validation de la génération du fichier
            FileInfo pdfFile = new FileInfo(Path.Combine(downloadsPath, docFileName));
            int retryCount = 0;

            while (!pdfFile.Exists && retryCount < 10)
            {
                tabletAppPage.WaitPageLoading();   
                pdfFile.Refresh();
                retryCount++;
            }

            Assert.IsTrue(pdfFile.Exists, $"{docFileName} n'a pas été généré.");

            // Validation du contenu du PDF
            PdfDocument document = PdfDocument.Open(pdfFile.FullName);
            Page page1 = document.GetPage(1);
            IEnumerable<Word> words = page1.GetWords();

            // Vérification des occurrences "C"
            int countC = words.Count(word => word.Text == "C");
            Assert.IsTrue(countC == 3, $"Attendu 3 occurrences de 'C', trouvé {countC} dans {docFileName}.");

            // Vérification des occurrences "NA"
            int countNA = words.Count(word => word.Text == "NA");
            Assert.IsTrue(countNA == 1, $"Attendu 1 occurrence de 'NA', trouvé {countNA} dans {docFileName}.");
        }


        [TestMethod]
        [Timeout(Timeout)]
        public void TB_HACCP_BlankDocs_CheckBoxCNAFavorite()
        {
            // Prepare
            string DocTitle = "HACCP3 Test";
            string DocSite = "GRO";
            string today = DateUtils.Now.Date.ToString("ddMMyyyy");
            string DocFileName = $"HACCP3 Test_{DocSite}_{today}_{today}.pdf";
            string DocFavName = "FavCNA.user";
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();

            // Arrange
            var homePage = LogInAsAdmin();

            TabletAppPage tabletAppPage = homePage.GotoTabletApp();

            tabletAppPage.Purge(downloadsPath, "HACCP3 Test_" + DocSite + "_");

            //1. sélectionner un site sur la tablette
            tabletAppPage.Select("mat-select-0", DocSite);

            //2. aller sur ""blank docs""
            tabletAppPage.GotoTabletApp_BlankDoc();
            tabletAppPage.WaitPageLoading();

            //purge favorite
            tabletAppPage.SelectV3("mat-select-2", DocTitle);
            //tabletAppPage.SelectAction("mat-select-2", DocTitle);
            tabletAppPage.Purge(6);
            tabletAppPage.Purge(4);
            //2. aller sur ""blank docs""
            homePage.Navigate();
            tabletAppPage = homePage.GotoTabletApp();
            tabletAppPage.Select("mat-select-0", DocSite);
            tabletAppPage.GotoTabletApp_BlankDoc();
            tabletAppPage.WaitPageLoading();

            //4. Sélectioner un doc haccp dans le menu déroulant
            tabletAppPage.SelectV3("mat-select-2", DocTitle);
            tabletAppPage.WaitPageLoading();

            //5. Cliquer sur next pour ouvrir le document
            tabletAppPage.ClickBaseButton("NEXT STEP");

            //6. choisir une date de production et cliquer sur print
            tabletAppPage.ClickButton("Print");

            // je suis dans un nouvel onglet !!!!
            tabletAppPage.GotoTabletApp_DocumentTab();
            tabletAppPage.WaitPageLoading();

            //7. Dans la colonne où la checkbox est paramétrée, cocher une case
            tabletAppPage.addLineTest(true, true);

            tabletAppPage.SavefavoriteDocument(DocFavName);

            //Revenir vers la page d'acceuil de blank docs , sélectionner le doc haccp dans le menu déroulant
            tabletAppPage.Close();
            tabletAppPage.WaitPageLoading();
            homePage.Navigate();

            tabletAppPage = homePage.GotoTabletApp();
            //sélectionner un site sur la tablette
            tabletAppPage.Select("mat-select-0", DocSite);
            //aller sur ""blank docs""
            tabletAppPage.GotoTabletApp_BlankDoc();
            tabletAppPage.WaitPageLoading();

            //9. récupérer ensuite le document favori dans la page d'acceuil de blank docs dans la section "" favorite ""...
            tabletAppPage.SelectV3("mat-select-2", DocTitle);
            // envoi des event remplissant le second combobox
            //tabletAppPage.SelectAction("mat-select-2", DocTitle);
            tabletAppPage.SelectAction("mat-select-6", DocFavName, true);

            // ...et ouvrir le document en cliquant sur next step
            tabletAppPage.ClickBaseButton("NEXT STEP");

            //choisir une date de production et cliquer sur print
            tabletAppPage.ClickButton("Print");

            // je suis dans un nouvel onglet !!!!
            tabletAppPage.GotoTabletApp_DocumentTab();
            tabletAppPage.WaitPageLoading();

            tabletAppPage.addLineTest(false, true, 3);

            //10. vérifier que la case cochée reste bien cochée sur le document favori
            // check CNA
            Assert.IsTrue(tabletAppPage.CheckValueBox(1));
            Assert.IsFalse(tabletAppPage.CheckValueBox(3));

            tabletAppPage.WaitPageLoading();
            //11. puis cliquer sur print pour générer le rapport haccp en pdf
            tabletAppPage.ClickButton("Print");
            FileInfo fi = new FileInfo(Path.Combine(downloadsPath, DocFileName));
            int counter = 0;
            while (!fi.Exists && counter < 30)
            {
                tabletAppPage.WaitPageLoading();
                fi.Refresh();
                counter++;
            }
            Assert.IsTrue(fi.Exists, DocFileName + " non généré");
            //12. vérifier que la case cochée renvoie bien l'information ""C"" et que les autres cases non cochés renvoient l'information ""NA""
            //1. la case cochée dans l'enregistrement reste bien cochée sur le document temporary
            //2.Dans le print, la case cochée doit renvoyer l'information ""C"" et les autres cases non cochées renvoient l'information ""NA""
            PdfDocument document = PdfDocument.Open(fi.FullName);
            Page page1 = document.GetPage(1);
            IEnumerable<Word> words = page1.GetWords();
            List<Word> cherchePdf = words.Where(word => word.Text == "C").ToList<Word>();
            Assert.IsTrue(cherchePdf.Count == 3, "C non trouvé dans " + fi.FullName);
            cherchePdf = words.Where(word => word.Text == "NA").ToList<Word>();
            Assert.IsTrue(cherchePdf.Count == 1, "NA non trouvé dans " + fi.FullName);
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void TB_HACCP_BlankDocs_CheckBoxCNC()
        {
            // Prepare
            string DocTitle = "HACCP3 Test";
            string DocSite = "GRO";
            string DocFileName = "HACCP3 Test_" + DocSite + "_" + DateUtils.Now.ToString("ddMMyyyy") + "_" + DateUtils.Now.ToString("ddMMyyyy") + ".pdf";
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();

            // Arrange
            var homePage = LogInAsAdmin();

            TabletAppPage tabletAppPage = homePage.GotoTabletApp();

            tabletAppPage.Purge(downloadsPath, "HACCP3 Test_" + DocSite + "_");

            //1. sélectionner un site sur la tablette
            tabletAppPage.Select("mat-select-0", DocSite);

            //2. aller sur ""blank docs""
            tabletAppPage.GotoTabletApp_BlankDoc();
            tabletAppPage.WaitPageLoading();

            //3. Sélectioner un doc haccp dans le menu déroulant
            tabletAppPage.Select("mat-select-2", DocTitle);
            tabletAppPage.WaitPageLoading();

            //4. Cliquer sur next pour ouvrir le document
            tabletAppPage.ClickBaseButton("NEXT STEP");

            //5. choisir une date de production et cliquer sur print
            tabletAppPage.ClickButton("Print");

            // je suis dans un nouvel onglet !!!!
            tabletAppPage.GotoTabletApp_DocumentTab();
            tabletAppPage.WaitPageLoading();

            //6. cliquer sur ""add line"" remplir une ligne de données
            tabletAppPage.addLineTest(true, false);
            tabletAppPage.addLineTest(true, true, 3);

            // check CNC
            Assert.IsFalse(tabletAppPage.CheckValueBox(2));
            Assert.IsTrue(tabletAppPage.CheckValueBox(4));

            Thread.Sleep(2000);
            //7. Cliquer sur le bouton ""print ""

            tabletAppPage.ClickButton("Print");
            FileInfo fi = new FileInfo(Path.Combine(downloadsPath, DocFileName));
            int counter = 0;
            while (!fi.Exists && counter < 20)
            {
                tabletAppPage.WaitPageLoading();
                fi.Refresh();
                counter++;
            }
            Assert.IsTrue(fi.Exists, DocFileName + " non généré");

            //1. le rapport haccp est généré en PDF avec toutes les données pré remplies
            //2. On retrouve dans le rapport haccp pdf les même données remplies dans le document haccp
            PdfDocument document = PdfDocument.Open(fi.FullName);
            Page page1 = document.GetPage(1);
            IEnumerable<Word> words = page1.GetWords();
            List<Word> cherchePdf = words.Where(word => word.Text == "C").ToList<Word>();
            Assert.IsTrue(cherchePdf.Count == 3, "C non trouvé dans " + fi.FullName);
            cherchePdf = words.Where(word => word.Text == "NC").ToList<Word>();
            Assert.IsTrue(cherchePdf.Count == 1, "NC non trouvé dans " + fi.FullName);
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void TB_HACCP_BlankDocs_CheckBoxCNCTempory()
        {
            // Prepare
            string DocTitle = "HACCP3 Test";
            string DocSite = "GRO";
            string DocFileName = "HACCP3 Test_" + DocSite + "_" + DateUtils.Now.Date.ToString("ddMMyyyy") + "_" + DateUtils.Now.Date.ToString("ddMMyyyy") + ".pdf";
            string DocTempName = "TempCNC.user";
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();

            // Arrange
            var homePage = LogInAsAdmin();

            TabletAppPage tabletAppPage = homePage.GotoTabletApp();

            tabletAppPage.Purge(downloadsPath, "HACCP3 Test_" + DocSite + "_");

            //1. sélectionner un site sur la tablette
            tabletAppPage.Select("mat-select-0", DocSite);

            //2. aller sur ""blank docs""
            tabletAppPage.GotoTabletApp_BlankDoc();
            tabletAppPage.WaitPageLoading();

            //4. Sélectioner un doc haccp dans le menu déroulant
            tabletAppPage.SelectV3("mat-select-2", DocTitle);
            tabletAppPage.Purge(6);
            tabletAppPage.Purge(4);

            homePage.Navigate();
            tabletAppPage = homePage.GotoTabletApp();
            tabletAppPage.Select("mat-select-0", DocSite);
            tabletAppPage.GotoTabletApp_BlankDoc();
            tabletAppPage.WaitPageLoading();

            //Sélectioner un doc haccp dans le menu déroulant
            tabletAppPage.SelectV3("mat-select-2", DocTitle);
            tabletAppPage.WaitPageLoading();

            //5. Cliquer sur next pour ouvrir le document
            tabletAppPage.ClickBaseButton("NEXT STEP");

            //6. choisir une date de production et cliquer sur print
            tabletAppPage.ClickButton("Print");

            // je suis dans un nouvel onglet !!!!
            tabletAppPage.GotoTabletApp_DocumentTab();
            tabletAppPage.WaitPageLoading();

            //7. Dans la colonne où la checkbox est paramétrée, cocher une case
            tabletAppPage.addLineTest(true, false);

            //8. cliquer sur save pour enregistrer un document temporary
            tabletAppPage.ClickButton("Save");

            //cliquer sur ""validate""
            tabletAppPage.FillFileName0(DocTempName);
            tabletAppPage.ClickButton("Validate");

            try
            {
                // erase old file
                tabletAppPage.ClickButton("Yes", true);
            }
            catch
            {
                // nouveau fichier
            }

            //une pop de validation apparait et puis cliquer sur OK
            tabletAppPage.ClickButton("Ok");

            //Revenir vers la page d'acceuil de blank docs , sélectionner le doc haccp dans le menu déroulant
            tabletAppPage.Close();
            Thread.Sleep(2000);
            homePage.Navigate();

            tabletAppPage = homePage.GotoTabletApp();
            //sélectionner un site sur la tablette
            tabletAppPage.Select("mat-select-0", DocSite);
            //aller sur ""blank docs""
            tabletAppPage.GotoTabletApp_BlankDoc();
            Thread.Sleep(2000);

            //9. récupérer ensuite le document temporary dans la page d'acceuil de blank docs dans la section "" temporary""
            tabletAppPage.SelectV3("mat-select-2", DocTitle);
            // envoi des event remplissant le second combobox
            //tabletAppPage.SelectAction("mat-select-2", DocTitle);
            tabletAppPage.SelectAction("mat-select-4", DocTempName, true);

            //Cliquer sur next pour ouvrir le document
            tabletAppPage.ClickBaseButton("NEXT STEP");

            //choisir une date de production et cliquer sur print

            // je suis dans un nouvel onglet !!!!
            tabletAppPage.GotoTabletApp_DocumentTab();
            tabletAppPage.WaitPageLoading();

            tabletAppPage.addLineTest(true, true, 3);

            //10. vérifier que la case cochée reste bien cochée sur le document temporary
            // check CNC
            Assert.IsFalse(tabletAppPage.CheckValueBox(2));
            Assert.IsTrue(tabletAppPage.CheckValueBox(4));

            tabletAppPage.WaitPageLoading();
            //11. puis cliquer sur print pour générer le rapport haccp en pdf
            tabletAppPage.ClickButton("Print");
            FileInfo fi = new FileInfo(Path.Combine(downloadsPath, DocFileName));
            int counter = 0;
            while (!fi.Exists && counter < 50)
            {
                Thread.Sleep(1000);
                fi.Refresh();
                counter++;
            }
            Assert.IsTrue(fi.Exists, DocFileName + " non généré");
            //12. vérifier que la case cochée renvoie bien l'information ""C"" et que les autres cases non cochés renvoient l'information ""NC"""
            //1. la case cochée dans l'enregistrement reste bien cochée sur le document temporary
            //2.Dans le print, la case cochée doit renvoyer l'information ""C"" et les autres cases non cochées renvoient l'information ""NC""
            PdfDocument document = PdfDocument.Open(fi.FullName);
            Page page1 = document.GetPage(1);
            IEnumerable<Word> words = page1.GetWords();
            Assert.AreEqual(1, words.Count(w => w.Text == "NC"), "NC non présent dans le Pdf");
            Assert.AreEqual(3, words.Count(w => w.Text == "C"), "C non présent dans le Pdf");
            Assert.AreEqual(0, words.Count(w => w.Text == "NA"), "NA présent dans le Pdf");
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void TB_HACCP_BlankDocs_CheckBoxCNCFavorite()
        {
            // Prepare
            string DocTitle = "HACCP3 Test";
            string DocSite = "GRO";
            string DocFileName = "HACCP3 Test_" + DocSite + "_" + DateUtils.Now.Date.ToString("ddMMyyyy") + "_" + DateUtils.Now.Date.ToString("ddMMyyyy") + ".pdf";
            string DocFavName = "FavCNC.user";
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();

            // Arrange
            var homePage = LogInAsAdmin();

            TabletAppPage tabletAppPage = homePage.GotoTabletApp();

            tabletAppPage.Purge(downloadsPath, "HACCP3 Test_" + DocSite + "_");

            //1. sélectionner un site sur la tablette
            tabletAppPage.Select("mat-select-0", DocSite);

            //2. aller sur ""blank docs""
            tabletAppPage.GotoTabletApp_BlankDoc();
            tabletAppPage.WaitPageLoading();

            //purge favorite
            tabletAppPage.Select("mat-select-2", DocTitle);
            tabletAppPage.SelectAction("mat-select-2", DocTitle);
            tabletAppPage.Purge(6);
            tabletAppPage.Purge(4);
            //2. aller sur ""blank docs""
            homePage.Navigate();
            tabletAppPage = homePage.GotoTabletApp();
            tabletAppPage.Select("mat-select-0", DocSite);
            tabletAppPage.GotoTabletApp_BlankDoc();
            tabletAppPage.WaitPageLoading();

            //4. Sélectioner un doc haccp dans le menu déroulant
            tabletAppPage.Select("mat-select-2", DocTitle);
            tabletAppPage.WaitPageLoading();

            //5. Cliquer sur next pour ouvrir le document
            tabletAppPage.ClickBaseButton("NEXT STEP");

            //6. choisir une date de production et cliquer sur print
            tabletAppPage.ClickButton("Print");

            // je suis dans un nouvel onglet !!!!
            tabletAppPage.GotoTabletApp_DocumentTab();
            tabletAppPage.WaitHACCPHorizontalProgressBar();

            //7. Dans la colonne où la checkbox est paramétrée, cocher une case
            tabletAppPage.addLineTest(true, false);

            tabletAppPage.SavefavoriteDocument(DocFavName);

            //Revenir vers la page d'acceuil de blank docs , sélectionner le doc haccp dans le menu déroulant
            tabletAppPage.Close();
            tabletAppPage.WaitPageLoading();
            homePage.Navigate();

            tabletAppPage = homePage.GotoTabletApp();
            //sélectionner un site sur la tablette
            tabletAppPage.Select("mat-select-0", DocSite);
            //aller sur ""blank docs""
            tabletAppPage.GotoTabletApp_BlankDoc();
            tabletAppPage.WaitPageLoading();

            //9. récupérer ensuite le document favori dans la page d'acceuil de blank docs dans la section "" favorite ""...
            tabletAppPage.Select("mat-select-2", DocTitle);
            // envoi des event remplissant le second combobox
            tabletAppPage.SelectAction("mat-select-2", DocTitle);
            tabletAppPage.SelectAction("mat-select-6", DocFavName, true);

            // ...et ouvrir le document en cliquant sur next step
            tabletAppPage.ClickBaseButton("NEXT STEP");

            //choisir une date de production et cliquer sur print
            tabletAppPage.ClickButton("Print");

            // je suis dans un nouvel onglet !!!!
            tabletAppPage.GotoTabletApp_DocumentTab();
            tabletAppPage.WaitHACCPHorizontalProgressBar();

            tabletAppPage.addLineTest(true, true, 3);

            //10. vérifier que la case cochée reste bien cochée sur le document favori
            // check CNC
            Assert.IsFalse(tabletAppPage.CheckValueBox(2));
            Assert.IsTrue(tabletAppPage.CheckValueBox(4));

            tabletAppPage.WaitPageLoading();
            //11. puis cliquer sur print pour générer le rapport haccp en pdf
            tabletAppPage.ClickButton("Print");
            tabletAppPage.WaitPageLoading();
            FileInfo fi = new FileInfo(downloadsPath + "\\" + DocFileName);
            int counter = 0;
            while (counter < 10 && !fi.Exists)
            {
                fi.Refresh();
                if (fi.Exists)
                {
                    break;
                }
                counter++;
                tabletAppPage.WaitPageLoading();
            }
            fi.Refresh();
            Assert.IsTrue(fi.Exists, DocFileName + " non généré");
            //12. vérifier que la case cochée renvoie bien l'information ""C"" et que les autres cases non cochés renvoient l'information ""NC""
            //1. la case cochée dans l'enregistrement reste bien cochée sur le document favori
            //2.Dans le print, la case cochée doit renvoyer l'information ""C"" et les autres cases non cochées renvoient l'information ""NC""
            PdfDocument document = PdfDocument.Open(fi.FullName);
            Page page1 = document.GetPage(1);
            IEnumerable<Word> words = page1.GetWords();
            Assert.AreEqual(1, words.Count(w => w.Text == "NC"), "NC non présent dans le Pdf");
            Assert.AreEqual(3, words.Count(w => w.Text == "C"), "C non présent dans le Pdf");
            Assert.AreEqual(0, words.Count(w => w.Text == "NA"), "NA présent dans le Pdf");
        }
        [TestMethod]
        [Timeout(Timeout)]
        public void TB_HACCP_BlankDocs_PrintFavoriteDocument()
        {
            // Prepare
            string DocTitle = "HACCP3 Sanitization";
            string DocSite = "GRO";
            string DocFileNameStart = DocTitle + "_" + DocSite;
            string DocFavName = "FavPrint.user";
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();

            // Arrange
            var homePage = LogInAsAdmin();

            TabletAppPage tabletAppPage = homePage.GotoTabletApp();

            tabletAppPage.Purge(downloadsPath, "HACCP3 Sanitization_" + DocSite + "_");

            //sélectionner un site sur la tablette
            tabletAppPage.Select("mat-select-0", DocSite);

            //aller sur ""blank docs""
            tabletAppPage.GotoTabletApp_BlankDoc();
            tabletAppPage.WaitPageLoading();

            //Sélectioner un doc haccp dans le menu déroulant
            tabletAppPage.Select("mat-select-2", DocTitle);
            tabletAppPage.WaitPageLoading();

            //Cliquer sur next pour ouvrir le document
            tabletAppPage.ClickBaseButton("NEXT STEP");

            //choisir une date de production et cliquer sur print
            tabletAppPage.ClickButton("Print");

            // je suis dans un nouvel onglet !!!!
            tabletAppPage.GotoTabletApp_DocumentTab();
            tabletAppPage.WaitPageLoading();

            //cliquer sur ""add line"" remplir une ligne de données
            tabletAppPage.addLineSanitization("MyProduct", "1.3", "2.2", "2.3", "MyView", "MyCorrectiveAction", "MyPreparedBy");
            tabletAppPage.WaitPageLoading();
            //Cliquer sur le bouton ""Save ""
            tabletAppPage.ClickButton("Favorite");
            //cliquer sur ""validate""
            tabletAppPage.FillFileName3(DocFavName);
            tabletAppPage.ClickButton("Validate");

            tabletAppPage.WaitPageLoading();
            // pas de message "voulez-vous écraser le fichier"

            //cliquer sur ""save and validate"" et puis cliquer sur ""save favorite"""
            tabletAppPage.ClickButton("Save and validate");
            tabletAppPage.ClickButton("Save favorite");
            tabletAppPage.WaitHACCPHorizontalProgressBar();
            tabletAppPage.ClickButton("Ok");

            //Revenir vers la page d'acceuil de blank docs , sélectionner le doc haccp dans le menu déroulant
            tabletAppPage.Close();
            tabletAppPage.WaitPageLoading();
            homePage.Navigate();

            tabletAppPage = homePage.GotoTabletApp();
            //1. sélectionner un site sur la tablette
            tabletAppPage.Select("mat-select-0", DocSite);
            //2. aller sur ""blank doc""
            tabletAppPage.GotoTabletApp_BlankDoc();
            tabletAppPage.WaitPageLoading();

            //3. Sélectioner un doc haccp dans le menu déroulant
            tabletAppPage.Select("mat-select-2", DocTitle);
            // envoi des event remplissant le second combobox
            tabletAppPage.SelectAction("mat-select-2", DocTitle);
            //4. sélectionner un doc favori
            tabletAppPage.SelectAction("mat-select-6", DocFavName, true);

            tabletAppPage.ClickBaseButton("NEXT STEP");
            tabletAppPage.ClickButton("Print");
            tabletAppPage.WaitPageLoading();
            // je suis dans un nouvel onglet !!!!
            tabletAppPage.GotoTabletApp_DocumentTab();
            tabletAppPage.WaitHACCPHorizontalProgressBar();

            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();
            foreach (var f in taskFiles)
            {
                f.Delete();
            }
            //5. Cliquer sur le bouton ""print """
            tabletAppPage.ClickButton("Print");

            // On récupère les fichiers du répertoire de téléchargement

            taskDirectory = new DirectoryInfo(downloadsPath);
            taskFiles = null;
            int counter = 10;
            while (taskDirectory.GetFiles().Length == 0 && counter > 0)
            {
                taskDirectory.Refresh();
                tabletAppPage.WaitPageLoading();
                counter--;
            }

            FileInfo correctDownloadedFile = taskDirectory.GetFiles()[0];
            // On recherche le fichier téléchargé
            //var correctDownloadedFile = tabletAppPage.GetReportPdf(taskFiles, DocFileNameStart);
            Assert.IsNotNull(correctDownloadedFile, "Le report n'a pas été téléchargé.");

            //1. le rapport haccp est généré en PDF avec toutes les données pré remplies
            //2. On retrouve dans le rapport haccp pdf les même données remplies dans le document haccp
            PdfDocument document = PdfDocument.Open(correctDownloadedFile.FullName);
            Page page1 = document.GetPage(1);
            IEnumerable<Word> words = page1.GetWords();
            List<Word> cherchePdf = words.Where(word => word.Text == "MyProduct").ToList<Word>();
            Assert.IsTrue(cherchePdf.Count == 1, "MyProduct non trouvé dans " + correctDownloadedFile.FullName);
            cherchePdf = words.Where(word => word.Text == "2.2").ToList<Word>();
            Assert.IsTrue(cherchePdf.Count == 1, "2.2 non trouvé dans " + correctDownloadedFile.FullName);
            cherchePdf = words.Where(word => word.Text == "MyView").ToList<Word>();
            Assert.IsTrue(cherchePdf.Count == 1, "Le commentaire n'est pas retrouvé dans " + correctDownloadedFile.FullName);
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void TB_HACCP_BlankDocs_SaveAndValidateDocumentTemporaryFromFavorite()
        {
            // Prepare
            string DocTitle = "HACCP3 Sanitization";
            string DocSite = "GRO";
            string DocTempName = "TempReady.user";
            string DocFavName = "FavReady.user";
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();

            //avoir un document favori enregistré pour le document haccp en question

            // Arrange
            var homePage = LogInAsAdmin();

            TabletAppPage tabletAppPage = homePage.GotoTabletApp();

            tabletAppPage.Purge(downloadsPath, "HACCP3 Sanitization_" + DocSite + "_");

            //sélectionner un site sur la tablette
            tabletAppPage.Select("mat-select-0", DocSite);

            //aller sur ""blank docs""
            tabletAppPage.GotoTabletApp_BlankDoc();

            //Sélectioner un doc haccp dans le menu déroulant
            tabletAppPage.SelectV3("mat-select-2", DocTitle);
            //tabletAppPage.SelectAction("mat-select-2", DocTitle);
            tabletAppPage.Purge(6);
            tabletAppPage.Purge(4);

            homePage.Navigate();
            tabletAppPage = homePage.GotoTabletApp();
            tabletAppPage.Select("mat-select-0", DocSite);
            tabletAppPage.GotoTabletApp_BlankDoc();
            Thread.Sleep(2000);

            //Sélectioner un doc haccp dans le menu déroulant
            tabletAppPage.SelectV3("mat-select-2", DocTitle);
            Thread.Sleep(2000);

            //Cliquer sur next pour ouvrir le document
            tabletAppPage.ClickBaseButton("NEXT STEP");

            //choisir une date de production et cliquer sur print
            tabletAppPage.ClickButton("Print");

            // je suis dans un nouvel onglet !!!!
            tabletAppPage.GotoTabletApp_DocumentTab();
            tabletAppPage.WaitHACCPHorizontalProgressBar();

            //cliquer sur ""add line"" remplir une ligne de données
            tabletAppPage.addLineSanitization("MyProduct", "1.3", "2.2", "2.3", "MyView", "MyCorrectiveAction", "MyPreparedBy");
            //Cliquer sur le bouton ""Save ""
            tabletAppPage.ClickButton("Favorite");
            //cliquer sur ""validate""
            tabletAppPage.FillFileName3(DocFavName);
            tabletAppPage.ClickButton("Validate");
            tabletAppPage.WaitForElementIsVisible(By.XPath("//*/span[contains(text(),'stars')]"));
            tabletAppPage.WaitForLoad();
            tabletAppPage.ClickButton("Save and validate");
            tabletAppPage.ClickButton("Cancel");
            tabletAppPage.ClickButton("Save and validate");
            //tabletAppPage.ClickButton("Validate");
            tabletAppPage.ClickButton("Save favorite");
            // pas de message "voulez-vous écraser le fichier"
            tabletAppPage.WaitHACCPHorizontalProgressBar();
            tabletAppPage.ClickButton("Ok");

            tabletAppPage.Close();

            //1. sélectionner un site sur la tablette
            homePage.Navigate();
            tabletAppPage = homePage.GotoTabletApp();
            tabletAppPage.Select("mat-select-0", DocSite);

            //2.aller sur ""blank docs""
            tabletAppPage.GotoTabletApp_BlankDoc();

            //3.Sélectioner un doc haccp dans le menu déroulant
            tabletAppPage.SelectV3("mat-select-2", DocTitle);
            //tabletAppPage.SelectAction("mat-select-2", DocTitle);

            //4.sur la section favori, sélectionner un document favori puis l'ouvrir en cliquant sur next step
            tabletAppPage.SelectAction("mat-select-6", DocFavName, true);
            tabletAppPage.ClickBaseButton("NEXT STEP");
            //choisir une date de production et cliquer sur print (date non sélectionnable)
            tabletAppPage.ClickButton("Print");
            // je suis dans un nouvel onglet !!!!
            tabletAppPage.GotoTabletApp_DocumentTab();
            tabletAppPage.WaitHACCPHorizontalProgressBar();

            //5.renseigner au moins une donnée de plus
            tabletAppPage.addLineSanitization("MyProduct2", "1.3", "2.2", "2.3", "MyView", "MyCorrectiveAction", "MyPreparedBy", 4);

            //6.cliquer sur save and validate
            // save temporary ? + validate
            tabletAppPage.ClickButton("Save");
            tabletAppPage.FillFileName3(DocTempName);
            tabletAppPage.ClickButton("Validate");

            //enregistrement des blank docs
            // site tablette
            homePage.Navigate();
            tabletAppPage = homePage.GotoTabletApp();
            tabletAppPage.Select("mat-select-0", DocSite);
            // aller sur BlankDoc
            tabletAppPage.GotoTabletApp_BlankDoc();
            // select doc
            tabletAppPage.SelectV3("mat-select-2", DocTitle);
            //tabletAppPage.SelectAction("mat-select-2", DocTitle);
            tabletAppPage.WaitForLoad();
            tabletAppPage.SelectAction("mat-select-4", DocTempName, true);
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void TB_HACCP_BlankDocs_OpenOldTemporyDocument()
        {
       
            //Prepare
            string DocTitle = "HACCP3 Modified texture";
            string DocSite = "GRO";
            string DocFileName = "TmpModifiedTexture-" + DateUtils.Now.ToString("yyyy-MM-dd");
            string DocFileNameHier = "TmpModifiedTexture-" + DateUtils.Now.AddDays(-1).ToString("yyyy-MM-dd");
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string DocFileNamePdfBegin = DocTitle + "_";
            string DocFileNameZipBegin = "All_files_";

            // Arrange
            var homePage = LogInAsAdmin();

            TabletAppPage tabletAppPage = homePage.GotoTabletApp();
            try
            {
                //1. sélectionner le site sur la tablette
                tabletAppPage.Select("mat-select-0", DocSite);
                //2.aller sur "blank docs"
                tabletAppPage.GotoTabletApp_BlankDoc();
                //3.Sélectionner un doc haccp dans le menu déroulant(HACCP3 Modified texture)
                tabletAppPage.SelectV3("mat-select-2", DocTitle);
                //tabletAppPage.SelectAction("mat-select-2", DocTitle);
                //4.Sélectionner le doc temporary de la veille 
                tabletAppPage.WaitForLoad();
                tabletAppPage.SelectAction("mat-select-4", DocFileNameHier, true);

                //et cliquer sur next step
                tabletAppPage.ClickBaseButton("NEXT STEP");
                //tabletAppPage.ClickButton("Print");
                //5.Cliquer sur le bouton "save and validate" puis "validate"
                // je suis dans un nouvel onglet !!!!
                tabletAppPage.GotoTabletApp_DocumentTab();
                tabletAppPage.WaitHACCPHorizontalProgressBar();
                tabletAppPage.ClickButton("Save and validate");
                tabletAppPage.ClickButton("Validate");
                tabletAppPage.WaitHACCPHorizontalProgressBar();
                tabletAppPage.ClickButton("Ok");
                tabletAppPage.Close();
            }
            catch
            {
                // batch non lancé hier.
                Console.WriteLine("batch non lancé hier");
            }

            homePage.Navigate();
            tabletAppPage = homePage.GotoTabletApp();
            //1.sélectionner un site sur la tablette
            tabletAppPage.Select("mat-select-0", DocSite);
            //2.aller sur ""blank docs""
            tabletAppPage.GotoTabletApp_BlankDoc();

            //3. Sélectionner un doc haccp dans le menu déroulant
            //(j'ai crée HACCP3 Modified texture pour ne pas se marcher dessus avec les autres documents)
            tabletAppPage.SelectV3("mat-select-2", DocTitle);

            tabletAppPage.Purge(6);
            tabletAppPage.Purge(4);

            //tabletAppPage.SelectAction("mat-select-2", DocTitle);
            //3. choisir la date d'aujourd'hui, remplir une ligne
            tabletAppPage.ClickBaseButton("NEXT STEP");
            tabletAppPage.ClickButton("Print");
            // je suis dans un nouvel onglet !!!!
            tabletAppPage.GotoTabletApp_DocumentTab();
            tabletAppPage.WaitHACCPHorizontalProgressBar();
            //4. remplir une ligne de données
            //tabletAppPage.addLineSanitization("MyProduct2", "1.3", "2.2", "2.3", "MyView", "MyCorrectiveAction", "MyPreparedBy", 4);
            //tabletAppPage.addLineThrowning("MyProduct", DateUtils.Now.Date.ToString("dd/MM/yyyy"), "7", "10", DateUtils.Now.Date.AddDays(-10).ToString("dd/MM/yyyy"), DateUtils.Now.Date.AddDays(10).ToString("dd/MM/yyyy"), "MyView", "MyCorrective", "MyPreparedBy");
            tabletAppPage.addLineModifiedTexture("MyProduct3", "0205AM", "1006PM", "2.0", "3.3", "MyComment", "MyCorrective", "MyPreparedBy", 0);
            //5. Cliquer sur le bouton "save"
            tabletAppPage.ClickButton("Save");
            //puis "validate"
            tabletAppPage.FillFileName3(DocFileName);
            tabletAppPage.ClickButton("Validate");
            tabletAppPage.ClickButton("Yes");
            tabletAppPage.ClickButton("Ok");

            // verify Accouting>Reporting
            homePage.Navigate();
            homePage.ClearDownloads();

            ReportingPage reportingPage = homePage.GoToAccounting_Reporting();
            reportingPage.Purge(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);
            string DocTypeOfReport = reportingPage.DevPathDocTypeOfRecord();
            reportingPage.FillReportingPage(DocTypeOfReport, DocSite);
            var offset = reportingPage.TableHasDocument(DocTitle);

            Assert.IsNotNull(offset, "Document non trouvé");
            
        }
    }
}
