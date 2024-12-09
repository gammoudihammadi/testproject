using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Security.Policy;
using System.Threading;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Purchasing
{
    public class BestPriceModal : PageBase
    {
        public BestPriceModal(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        private const string SUPPLIER = "BestPriceSelectedSuppliers_ms";
        private const string SUPPLIER_CHECK_ALL = "/html/body/div[17]/div/ul/li[1]/a";
        private const string GROUP = "BestPriceSelectedGroups_ms";
        private const string GROUP_CHECK_ALL = "/html/body/div[18]/div/ul/li[1]/a";
        private const string SELECT_FIRST_LINE = "//*[@id='item_IsSelected']";
        private const string SITE_TABLE = "//*[@id=\"tableBestPrice\"]/tbody/tr[{0}]/td[2]";
        private const string SUPPLIER_TABLE= "//*[@id=\"tableBestPrice\"]/tbody/tr[{0}]/td[3]";
        private const string ITEM_NAME_TABLE = "//*[@id=\"tableBestPrice\"]/tbody/tr[{0}]/td[4]";
        private const string PACK_TABLE = "//*[@id=\"tableBestPrice\"]/tbody/tr[{0}]/td[6]";
        private const string SUPPLIER_CHECK_ALL_ONE = "/html/body/div[12]/div/ul/li[1]/a";
        private const string GROUP_CHECK_ALL_ONE = "/html/body/div[13]/div/ul/li[1]/a/span[2]";
        private const string OK_UPDATE_BUTTON = "//*[@id=\"modal-1\"]/div[3]/button";
        private const string BEST_PRICE_CHECKBOX = ".//td[1]//input[@type='checkbox']";

        [FindsBy(How = How.Id, Using = SUPPLIER)]
        private IWebElement _supplier;

        [FindsBy(How = How.XPath, Using = SUPPLIER_CHECK_ALL)]
        private IWebElement _supplierCheckAll;

        [FindsBy(How = How.Id, Using = GROUP)]
        private IWebElement _group;

        [FindsBy(How = How.XPath, Using = GROUP_CHECK_ALL)]
        private IWebElement _groupCheckAll;

        [FindsBy(How = How.XPath, Using = SELECT_FIRST_LINE)]
        private IWebElement _resultTable;
        [FindsBy(How = How.XPath, Using = OK_UPDATE_BUTTON)]
        private IWebElement _OKUpdateButton;

        public void Fill(string site, string supplier = null, string group = null)
        {
            // combo box site
            ComboBoxSelectById(new ComboBoxOptions("BestPriceSelectedSites_ms", site + " - " + site, false));
            // combo box supplier (All)
            _supplier = WaitForElementIsVisible(By.Id(SUPPLIER));
            if (supplier == null)
            {
                _supplier.Click();

                _supplierCheckAll =  isElementExists(By.XPath(SUPPLIER_CHECK_ALL)) ? WaitForElementIsVisible(By.XPath(SUPPLIER_CHECK_ALL)) : WaitForElementIsVisible(By.XPath(SUPPLIER_CHECK_ALL_ONE));
                _supplierCheckAll.Click();
                // on referme
                _supplier.Click();
            }
            else
            {
                ComboBoxSelectById(new ComboBoxOptions("BestPriceSelectedSuppliers_ms", supplier, false));
            }
            // combo box group (All)
            _group = WaitForElementIsVisible(By.Id(GROUP));
            if (group == null)
            {
                _group.Click();

                _groupCheckAll = isElementExists(By.XPath(GROUP_CHECK_ALL)) ? WaitForElementIsVisible(By.XPath(GROUP_CHECK_ALL)) : WaitForElementIsVisible(By.XPath(GROUP_CHECK_ALL_ONE));
                _groupCheckAll.Click();
                // on referme
                _group.Click();
            }
            else
            {
                ComboBoxSelectById(new ComboBoxOptions("BestPriceSelectedGroups_ms", group, false));
            }
            WaitForLoad();
        }

        public void ClickSearch()
        {
            //SEARCH
            var searchAction = WaitForElementIsVisible(By.Id("SearchBestPriceBtn"));
            searchAction.Click();
            WaitPageLoading();
            WaitPageLoading();
        }

        public void SelectFirstLine()
        {
            var resultTable = WaitForElementIsVisible(By.XPath(SELECT_FIRST_LINE));
            resultTable.Click();
        }
        public void ClickUpdate()
        {
            var update = WaitForElementIsVisible(By.Id("updateBestPriceBtn"));
            update.Click();
            WaitPageLoading();
            WaitPageLoading();
            Assert.IsTrue(isElementVisible(By.XPath("//*/h4[text()='Import done!']")),"Import not done");
        }

        public List<bool> PackagingIsMain()
        {
            List<bool> isMains = new List<bool>();
            var isMainElements = _webDriver.FindElements(By.XPath("//*[@id='detailsItemContainer']/div/table/tbody/tr[*]/td[2]/div/input"));
            for (var e = 0; e < isMainElements.Count; e++)
            {
                isMains.Add(isMainElements[e].Selected);
            }
            return isMains;
        }

        public void RestoreIsMain(List<bool> isMainsFlat)
        {
            for (var e = 0; e < isMainsFlat.Count; e++)
            {
                if (isMainsFlat[e] == true)
                {
                    int offset = e + 1;
                    // selection
                    var IsMainScanLine = WaitForElementExists(By.XPath("//*[@id='detailsItemContainer']/div/table/tbody/tr[" + offset + "]/td[1]"));
                    new Actions(_webDriver).MoveToElement(IsMainScanLine).Click().Perform();
                    // action
                    var IsMainScan = WaitForElementExists(By.XPath("//*[@id='detailsItemContainer']/div/table/tbody/tr[" + offset + "]/td[2]/div/input"));
                    new Actions(_webDriver).MoveToElement(IsMainScan).Click().Perform();
                    // affiche la disquette
                    //save-state btn-icon btn-icon-status-inprogress
                    //var dskOn = WaitForElementExists(By.XPath("//*[@id='detailsItemContainer']/div/table/tbody/tr[" + offset + "]/td[17]/span[1]/span[contains(@class,'btn-icon-status-inprogress')]"));
                    //var dskOff = WaitForElementExists(By.XPath("//*[@id='detailsItemContainer']/div/table/tbody/tr[" + offset + "]/td[17]/span[1]/span[not(contains(@class,'btn-icon-status-inprogress'))]"));
                    Thread.Sleep(4000);
                    WaitForLoad();
                }
            }

        }

        public string GetSiteFromBestPriceTable(int line)
        {
            var siteName = WaitForElementIsVisible(By.XPath(string.Format(SITE_TABLE, line)));
            return siteName.Text.Trim();
        }
        public string GetSupplierNameFromBestPriceTable(int line)
        {
            var supplierName = WaitForElementIsVisible(By.XPath(string.Format(SUPPLIER_TABLE, line)));
            return supplierName.Text.Trim();
        }
        public string GetItemNameFromBestPriceTable(int line)
        {
            var itemName = WaitForElementIsVisible(By.XPath(string.Format(ITEM_NAME_TABLE, line)));
            return itemName.Text.Trim();
        }
        public string GetPackFromBestPriceTable(int line)
        {
            var packName = WaitForElementIsVisible(By.XPath(string.Format(PACK_TABLE, line)));
            return packName.Text.Trim();
        }
        public void ClickOnOKUpdate()
        {
            WaitPageLoading();
            _OKUpdateButton = WaitForElementIsVisible(By.XPath(OK_UPDATE_BUTTON));
            _OKUpdateButton.Click();
           
        }
        public bool SelectCheckBoxByItemName(string itemName)
        {
            var rows = _webDriver.FindElements(By.XPath($"//*[@id='tableBestPrice']/tbody/tr[td[4][contains(text(), '{itemName}')]]"));

            if (rows.Count > 0)
            {
                var row = rows[0];
                var checkbox = row.FindElement(By.XPath(BEST_PRICE_CHECKBOX));

                if (checkbox != null && !checkbox.Selected)
                {
                    checkbox.Click();
                    WaitPageLoading();
                    return true;  
                }
                else
                {
                    return false;
                }
            }
            else
            {
                return false;
            }
        }




    }
}