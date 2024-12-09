using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Flights.Crew;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using System;
using System.IO;
using System.Web.UI.WebControls;

namespace Newrest.Winrest.FunctionalTests.Flights
{
    [TestClass]
    public class CrewTest : TestBase
    {
        private const int _timeout = 60 * 10 * 1000;
        [TestMethod]
        [Timeout(_timeout)]
        public void FL_CREW_Create()
        {
            Random rnd = new Random();

            var homePage = LogInAsAdmin();

            CrewPage crewPage = homePage.GoToFlights_CrewPage();
            var empNumber = rnd.Next().ToString();
            var numberCrewsInHeader = crewPage.GetNumberCrewsInHeader();
            var userId = crewPage.GenerateString();

            crewPage.CreateNewCrewMember("VLM AIRLINES", "Test", "Test", empNumber, "TestLog", userId, "test**test");
            var numberAfterCreateCrews = crewPage.GetNumberCrewsInHeader();

            var AfterCreateCrewsNumber = int.Parse(numberAfterCreateCrews);
            var CrewsInHeaderNumber = (int.Parse(numberCrewsInHeader) + 1);

            Assert.AreEqual(AfterCreateCrewsNumber, CrewsInHeaderNumber, "La creation de crew ne fonctionne pas.");
            crewPage.Filter(CrewPage.FilterType.Search, empNumber);
            crewPage.DeleteCrew();
        }


        [TestMethod]
        [Timeout(_timeout)]
        public void FL_CREW_Delete_Crew()
        {
            Random rnd = new Random();
            string empNumber = rnd.Next(100, 1000).ToString();
            string airlineEdition = TestContext.Properties["customerName1Trolley"].ToString();
            string displayName = "Test";
            string firstName = "Test";
            string logonName = "TestLog";
            string password = "test**test";

            var homePage = LogInAsAdmin();
            CrewPage crewPage = homePage.GoToFlights_CrewPage();
            var userId = crewPage.GenerateString();
            crewPage.CreateNewCrewMember(airlineEdition, displayName, firstName, empNumber, logonName, userId, password);
            crewPage.ResetFilters();
            crewPage.Filter(CrewPage.FilterType.Search, empNumber);
            crewPage.DeleteCrew();
            crewPage.ResetFilters();
            crewPage.Filter(CrewPage.FilterType.Search, empNumber);
            var numberAfterDeleteCrews = crewPage.CheckTotalNumber();
            var checkDelete = numberAfterDeleteCrews.Equals(0);

            Assert.IsTrue(checkDelete, "La suppresion de crew ne fonctionne pas.");
        }


        [TestMethod]
        [Timeout(_timeout)]
        public void FL_CREW_Filter_Search()
        {
            var homePage = LogInAsAdmin();

            CrewPage crewPage = homePage.GoToFlights_CrewPage();
            crewPage.Filter(CrewPage.FilterType.ShowAll, true);
            var numberCrewsInHeader = crewPage.GetNumberCrewsInHeader();
            crewPage.Filter(CrewPage.FilterType.Search, "23");
            var numberAfterFilterCrews = crewPage.GetNumberCrewsInHeader();
            Assert.AreNotEqual(numberAfterFilterCrews, numberCrewsInHeader, "search filter ne fonctionne pas.");
        }


        [TestMethod]
        [Timeout(_timeout)]
        public void FL_CREW_Filter_ShowAll()
        {
            var homePage = LogInAsAdmin();
            var filterValue = true;

            CrewPage crewPage = homePage.GoToFlights_CrewPage();
            crewPage.ResetFilters();
            crewPage.Filter(CrewPage.FilterType.ShowOnlyActive, filterValue);
            var numberCrewsActive = crewPage.CheckTotalNumber();
            crewPage.Filter(CrewPage.FilterType.ShowOnlyInactive, filterValue);
            var numberCrewsInactive = crewPage.CheckTotalNumber();
            crewPage.Filter(CrewPage.FilterType.ShowAll, filterValue);
            var allCrews = crewPage.CheckTotalNumber();
            var activeAndInactiveCrews = numberCrewsActive + numberCrewsInactive;
            Assert.AreEqual(allCrews, activeAndInactiveCrews, String.Format(MessageErreur.FILTRE_ERRONE, "Show All"));
            crewPage.ResetFilters();
        }


        [TestMethod]
        [Timeout(_timeout)]
        public void FL_CREW_Filter_ShowOnlyActive()
        {
            var homePage = LogInAsAdmin();
            var value = true;

            CrewPage crewPage = homePage.GoToFlights_CrewPage();
            crewPage.Filter(CrewPage.FilterType.ShowOnlyActive, value);
            var verifyFilter = crewPage.VerifyShowOnlyActiveFilter();
            Assert.IsTrue(verifyFilter, String.Format(MessageErreur.FILTRE_ERRONE, "Show Only Active"));
            crewPage.ResetFilters();
        }


        [TestMethod]
        [Timeout(_timeout)]
        public void FL_CREW_Filter_ShowOnlyInactive()
        {
            var homePage = LogInAsAdmin();
            var value = true;

            CrewPage crewPage = homePage.GoToFlights_CrewPage();
            crewPage.Filter(CrewPage.FilterType.ShowOnlyInactive, value);
            var verifyFilter = crewPage.VerifyShowOnlyInActiveFilter();
            Assert.IsTrue(verifyFilter, String.Format(MessageErreur.FILTRE_ERRONE, "Show Only Inactive"));
            crewPage.ResetFilters();
        }


        [TestMethod]
        [Timeout(_timeout)]
        public void FL_CREW_ResetFilter()
        {
            object value = null;
            var trueValue = true;
            var falseValue = false;
            string randomNumber = new Random().Next().ToString();

            var homePage = LogInAsAdmin();

            CrewPage crewPage = homePage.GoToFlights_CrewPage();

            crewPage.Filter(CrewPage.FilterType.Search, randomNumber);
            crewPage.Filter(CrewPage.FilterType.ShowOnlyInactive, trueValue);

            crewPage.ResetFilters();

            value = crewPage.GetFilterValue(CrewPage.FilterType.Search);
            var verifySearch = value.Equals("");
            Assert.IsTrue(verifySearch, "La fonctionnalité ResetFilter ne fonctionne pas au niveau du filtrage Search");

            value = crewPage.GetFilterValue(CrewPage.FilterType.ShowAll);
            Assert.AreEqual(falseValue, value, "La fonctionnalité ResetFilter ne fonctionne pas au niveau du filtrage Show All");

            value = crewPage.GetFilterValue(CrewPage.FilterType.ShowOnlyActive);
            Assert.AreEqual(trueValue, value, "La fonctionnalité ResetFilter ne fonctionne pas au niveau du filtrage Show Only Active");

            value = crewPage.GetFilterValue(CrewPage.FilterType.ShowOnlyInactive);
            Assert.AreEqual(falseValue, value, "La fonctionnalité ResetFilter ne fonctionne pas au niveau du filtrage Show Only Inactive");
        }


        [TestMethod]
        [Timeout(_timeout)]
        public void FL_CREW_ModifyDetail()
        {
            Random rnd = new Random();
            var empNumber = rnd.Next(100, 1000).ToString();
            string airlineInitiale = TestContext.Properties["CustomerLP"].ToString();
            string airlineEdit = TestContext.Properties["customerName1Trolley"].ToString();
            string displayName = "displayNameTest";
            string firstName = "firstNameTest";
            string logName = "logNameTest";
            string password = "passwordTest";
            bool value = true;
            string logNameEdit = "logNameEdit";

            var homePage = LogInAsAdmin();
            CrewPage crewPage = homePage.GoToFlights_CrewPage();
            var userId = crewPage.GenerateString();
            crewPage.CreateNewCrewMember(airlineInitiale, displayName, firstName, empNumber, logName, userId, password);
            try
            {
                crewPage.Filter(CrewPage.FilterType.ShowAll, value);
                crewPage.Filter(CrewPage.FilterType.Search, empNumber);
                crewPage.SelectFirstCrew();
                crewPage.EditDetailCrew(airlineEdit, logNameEdit);
                crewPage.BackToList();
                crewPage.Filter(CrewPage.FilterType.Search, empNumber);
                crewPage.SelectFirstCrew();
                var verifyEdit = crewPage.VerifyModifyDetail(airlineEdit, logNameEdit);
                Assert.IsTrue(verifyEdit, "Les données ne sont pas modifiés.");
            }
            finally
            {
                crewPage.BackToList();
                crewPage.Filter(CrewPage.FilterType.Search, empNumber);
                crewPage.DeleteCrew();
            }
        }


        [TestMethod]
        [Timeout(_timeout)]
        public void FL_CREW_Export()
        {
            string downloadsPath = TestContext.Properties["DownloadsPath"].ToString();

            var homePage = LogInAsAdmin();

            CrewPage crewPage = homePage.GoToFlights_CrewPage();
            homePage.ClearDownloads();

            crewPage.ResetFilters();
            var numbersCrewGridList = crewPage.GetListNumbersCrew();
            foreach (var file in new DirectoryInfo(downloadsPath).GetFiles())
            {
                file.Delete();
            }
            crewPage.Export();
            // On récupère les fichiers du répertoire de téléchargement
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            var correctDownloadedFile = crewPage.GetExportExcelFile(taskFiles);
            Assert.IsNotNull(correctDownloadedFile);

            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadsPath, fileName);
            var numbersExcel = OpenXmlExcel.GetValuesInList("Employee number", "Crews", filePath);
            //verifier les données 
            var verify = crewPage.VerifyExcel(numbersCrewGridList, numbersExcel);
            Assert.IsTrue(verify, "les donées du grid et du fichier Excel ne sont pas identiques");
        }


        [TestMethod]
        [Timeout(_timeout)]
        public void FL_CREW_Import()
        {
            // Répertoire du fichier à importer
            var path = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase).Substring(6);
            var excelPath = path + "\\PageObjects\\Flights\\Crew\\ImportCrew_test.xlsx"; //1 NewTest VLM
            Assert.IsTrue(new FileInfo(excelPath).Exists, "Fichier " + excelPath + "non trouvé");

            var homePage = LogInAsAdmin();

            CrewPage crewPage = homePage.GoToFlights_CrewPage();
            //homePage.ClearDownloads();

            //si crew crée au précédent test non supprimé
            crewPage.Filter(CrewPage.FilterType.Search, "NewTest");
            if (crewPage.GetNumberCrewsInHeader() != "0")
            {
                crewPage.DeleteCrew();
            };

            crewPage.ResetFilters();
            crewPage.Filter(CrewPage.FilterType.ShowAll, true);
            var numberCrewsInHeader = crewPage.GetNumberCrewsInHeader();

            crewPage.Import(excelPath);
            crewPage.Filter(CrewPage.FilterType.ShowAll, true);
            bool isAddedFileVerif = crewPage.IsAddedFileVerif(int.Parse(numberCrewsInHeader));
            Assert.IsTrue(isAddedFileVerif, "ce fichier n'est pas importé.");

            //delete crew created during test
            crewPage.Filter(CrewPage.FilterType.Search, "NewTest");
            crewPage.DeleteCrew();
        }
    }
}