using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Purchasing.Suppliers
{
    public class SupplierEDITab : PageBase
    {
        public SupplierEDITab(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {

        }

        private const string ID_EDI = "IdEdi";
        private const string ORDER_EDI = "purchasing-supplier-detailsorderEDI";
        private const string ORDER_EDI_PATCH = "purchasing-supplier-detailsorderEDI";
        private const string SITES_FILTER = "SelectedSites_ms";
        private const string SEARCH_SITE = "/html/body/div[12]/div/div/label/input"; 
        private const string SITES_SELECT_ALL = "/html/body/div[12]/div/ul/li[1]/a/span[2]";
        private const string SITE_UNSELECT_ALL = "/html/body/div[12]/div/ul/li[2]/a/span[2]";
        private const string EDI_GNL_CODE = "DeliveriesSites_1__CodeSupplierEdiGLN";
        private const string ITEM_TAB = "hrefTabContentArticles";
        private const string CONFIRM_BUTTON = "//*[@id=\"modal-1\"]/div[2]/div/form/div[2]/button[2]"; 







        [FindsBy(How = How.Id, Using = ID_EDI)]
        private IWebElement _idEdi;
        [FindsBy(How = How.Id, Using = ORDER_EDI)]
        private IWebElement _orderEdi;
        [FindsBy(How = How.Id, Using = ORDER_EDI_PATCH)]
        private IWebElement _orderEdiPatch;
        [FindsBy(How = How.Id, Using = SITES_FILTER)]
        private IWebElement _SitesFilter;
        [FindsBy(How = How.XPath, Using = SEARCH_SITE)]
        private IWebElement _searchSite;
        [FindsBy(How = How.XPath, Using = SITES_SELECT_ALL)]
        private IWebElement _siteSelectAll;
        [FindsBy(How = How.XPath, Using = SITE_UNSELECT_ALL)]
        private IWebElement _siteUnselectAll;
        [FindsBy(How = How.Id, Using = EDI_GNL_CODE)]
        private IWebElement _ediGnlCode;
        [FindsBy(How = How.Id, Using = ITEM_TAB)]
        private IWebElement _itemTabBtn;


        public void SetIdEDI( string value)
        {
            WaitForLoad();
            _idEdi = WaitForElementIsVisible(By.Id(ID_EDI));
            _idEdi.SetValue(ControlType.TextBox, value);
            WaitForLoad();

        }

        public void SetEdiGNLCode(string value)
        {
            _ediGnlCode = WaitForElementIsVisible(By.Id(EDI_GNL_CODE));
            _ediGnlCode.SetValue(ControlType.TextBox, value);
            WaitForLoad();

        }

        public void SelectASiteInOrderEDI(string value)
        {
            if (IsDev())
            {
                _orderEdi = WaitForElementIsVisible(By.Id(ORDER_EDI));
                _orderEdi.Click();
                WaitForLoad();

            }
            else
            {
                _orderEdiPatch = WaitForElementIsVisible(By.Id(ORDER_EDI_PATCH));
                _orderEdiPatch.Click();
                WaitForLoad();

            }
            if (isElementVisible(By.Id(SITES_FILTER)))
            {
                //_SitesFilter = WaitForElementIsVisible(By.Id(SITES_FILTER));
                //_SitesFilter.Click();

                //    _siteUnselectAll = WaitForElementIsVisible(By.XPath(SITE_UNSELECT_ALL));
                //    _siteUnselectAll.Click();

                //    _searchSite = WaitForElementIsVisible(By.XPath(SEARCH_SITE));
                //    _searchSite.SetValue(ControlType.TextBox, value);

                //     var valueToCheckCustomersType = WaitForElementIsVisible(By.XPath("//span[text()='" + value + "']"));
                //     valueToCheckCustomersType.SetValue(ControlType.CheckBox, true);

                //_SitesFilter.Click();
                ComboBoxSelectById(new ComboBoxOptions(SITES_FILTER, (string)value,false));
            }
            var button  = WaitForElementIsVisible(By.XPath(CONFIRM_BUTTON));
            button.Click();
        }

        public SupplierItem ClickOnItemsTab()
        {
            _itemTabBtn = WaitForElementIsVisible(By.Id(ITEM_TAB));
            _itemTabBtn.Click();
            WaitForLoad();
            return new SupplierItem(_webDriver, _testContext);
        }

        public void SetPurchaseOrderFileFormatToXML()
        {
            WaitLoading();
            _ediGnlCode = WaitForElementIsVisible(By.Id("EdiPOFileFormat"));
            _ediGnlCode.SetValue(ControlType.DropDownList, "XML");
            WaitLoading();
        }

    }
}
