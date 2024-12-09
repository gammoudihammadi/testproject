using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects;
using Newrest.Winrest.FunctionalTests.PageObjects.Parameters.Sites;


namespace Newrest.Winrest.FunctionalTests.Parameters
{
    [TestClass]
    public class FiscalEntitiesTest : TestBase
    {
        private const int Timeout = 600000;

        [TestMethod]
        [Timeout(Timeout)]
        public void PA_SI_FE_SiteFiscalEntitySelectionTest()
        {
            HomePage homePage = LogInAsAdmin();
            ParametersSites sitesPage = homePage.GoToParameters_Sites();
            sitesPage.ClickOnFiscalEntitiesTab();
            int count_before = sitesPage.GetLinesFiscalEntitiesCount();
            ParametersSitesModalPage modalCreate = sitesPage.ClickOnNewFiscalEntityModal();
            modalCreate.CreateNewFiscalEntity("A name", "A code" ,"Uyumsoft","","222","rue de la Pomme","31000","Toulouse","folder sftp");
            sitesPage.WaitPageLoading();        
            int count_after = sitesPage.GetLinesFiscalEntitiesCount();
            Assert.AreEqual(count_before + 1, count_after, "L'entité fiscale n'a pas pu être créée");
            int index = count_after;
            bool created = sitesPage.CheckFiscalEntityExists("A name", "A code", "Uyumsoft", index);
            if (created) 
            {
                sitesPage.DeleteFiscalEntityAt(index);
                int last_count = sitesPage.GetLinesFiscalEntitiesCount();
                Assert.AreEqual(count_before, last_count, "L'entité fiscale n'a pas pu être supprimée");
            }
            else Assert.Fail("L'entité fiscale n'a pas pu être supprimée");
        }
    }
}
