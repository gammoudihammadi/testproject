using Newrest.Winrest.FunctionalTests.PageObjects;
using Newrest.Winrest.FunctionalTests.PageObjects.Parameters.Accounting;
using Newrest.Winrest.FunctionalTests.PageObjects.Parameters.Sites;

namespace Newrest.Winrest.FunctionalTests.Parameters
{
    public abstract class RedistributiveTaxTestBase : TestBase
    {
        public void CreateFiscalEntityUyumsoft()
        {
            HomePage homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            ParametersSites sitesPage = homePage.GoToParameters_Sites();
            sitesPage.ClickOnFiscalEntitiesTab();

            ParametersSitesModalPage modalCreate = sitesPage.ClickOnNewFiscalEntityModal();
            modalCreate.CreateNewFiscalEntity("test name", "test code", "Uyumsoft", "...", "222", "rue de la Pomme", "31000", "Toulouse", "folder sftp");

            sitesPage.WaitPageLoading();
        }

        public void DeleteFiscalEntityUyumsoft()
        {
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            ParametersSites sitesPage = homePage.GoToParameters_Sites();
            sitesPage.ClickOnFiscalEntitiesTab();

            sitesPage.DeleteFiscalEntityByValues("test code", "test name", "Uyumsoft");

            sitesPage.WaitPageLoading();
        }

        public void CreateRedistributiveCustomerTax(string code)
        {
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            var accountingPage = homePage.GoToParameters_AccountingPage();
            accountingPage.GoToTab_RedistributiveCustomerTax();

            var modalCreate = accountingPage.ClickOnNewRedistributiveCustomerTax();
            modalCreate.CreateNewRedistributiveCustomerTax("2-General", "ACE - ACE", "AIR CPU SL", "BEBIDAS CALIENTES", 35, "CJ_" + code, "AC_" + code);
            accountingPage.WaitPageLoading();
        }

        public int CheckRedistributiveCustomerTaxCount()
        {
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            var accountingPage = homePage.GoToParameters_AccountingPage();
            accountingPage.GoToTab_RedistributiveCustomerTax();

            return accountingPage.RedistributiveCustomerTaxCount("2-General", "ACE", "AIR CPU SL", "BEBIDAS CALIENTES", 35);
        }


        public void CreateRedistributiveSupplierTax(string code)
        {
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            ParametersAccounting accountingPage = homePage.GoToParameters_AccountingPage();
            accountingPage.GoToTab_RedistributiveSupplierTax();

            var modalCreate = accountingPage.ClickOnNewRedistributiveSupplierTax();
            modalCreate.CreateNewRedistributiveSupplierTax("2-General", "ACE - ACE", "ADPAN EUROPA, S.L.", "CONSERVAS VEGETALES", 35, "CJ_" + code, "AC_" + code);

            accountingPage.WaitPageLoading();

        }

        public int CheckRedistributiveSupplierTaxCount()
        {
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            var accountingPage = homePage.GoToParameters_AccountingPage();
            accountingPage.GoToTab_RedistributiveSupplierTax();

            return accountingPage.RedistributiveSupplierTaxCount("2-General", "ACE", "ADPAN EUROPA, S.L.", "CONSERVAS VEGETALES", 35);
        }


    }
}
