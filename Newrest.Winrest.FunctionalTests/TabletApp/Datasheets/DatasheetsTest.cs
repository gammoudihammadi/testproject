using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.TabletApp;
using Newrest.Winrest.FunctionalTests.PageObjects;


namespace Newrest.Winrest.FunctionalTests.TabletApp.Datasheets
{
    [TestClass]
    public class DatasheetsTest : TestBase
    {
        private const int Timeout = 600000;

        [TestMethod]
        [Timeout(Timeout)]
        public void TB_DATA_Access()
        {
            var homePage = LogInAsAdmin();
            TabletAppPage tabletAppPage = homePage.GotoTabletApp();
            var isDisplayed = tabletAppPage.isTabletAppPageDisplayed();
            Assert.IsTrue(isDisplayed," la page TabletApp n'est pas affichée");
           
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void TB_DATA_SearchGuestType()
        {
            string guestType = TestContext.Properties["Prodman_Needs_GuestType1"].ToString();

            var homePage = LogInAsAdmin();
            TabletAppPage tabletAppPage = homePage.GotoTabletApp();
            DatasheetTabletAppPage DatasheetPage = tabletAppPage.GotoTabletApp_Datasheet();
            var isGuestTypeExist = DatasheetPage.CheckGuestTypeExist(guestType);
            Assert.IsTrue(isGuestTypeExist, "Le guestType n'existe pas ou il ne correspond pas au filtre");
        }
    }
}
