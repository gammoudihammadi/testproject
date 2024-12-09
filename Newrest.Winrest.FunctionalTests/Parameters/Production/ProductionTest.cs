using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects;
using Newrest.Winrest.FunctionalTests.PageObjects.Parameters.Production;

namespace Newrest.Winrest.FunctionalTests.Parameters.Production
{
    [TestClass]
    public class ProductionTest : TestBase
    {
        private const int Timeout = 600000;

        [TestMethod]
        [Timeout(Timeout)]
        public void SETT_PRO_Settings_Suppression_Keywords()
        {
            string keyword = "TEST_KEYWORD";
            string OkToDelete = "Keyword successfully deleted";
            HomePage homePage = LogInAsAdmin();
            ParametersProduction parametersProduction = homePage.GoToParameters_ProductionPage();
            parametersProduction.GoToTab_Keyword();
            parametersProduction.AddKeyword(keyword);
            bool KeywordDeleted = parametersProduction.DeleteKeyword(keyword, OkToDelete);
            Assert.IsTrue(KeywordDeleted, "Keyword cannot be deleted.");
        }
    }
}
