using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Purchasing.Item
{
    public class ItemLastReceiptNotesPage : PageBase
    {
        public ItemLastReceiptNotesPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        private const string ITEM_SUPPLIER_FIRST_RN = "//*[@id=\"list-item-with-action\"]/div[2]/div[2]/div/div";
        //private const string WHRN_SUPPLIER_FIRST_RN = "//*[@id=\"div-body\"]/div/div[2]/h3/span[2]";
        //private const string WHRN_RN_PENCIL_LINK = "//*[@id=\"list-item-with-action\"]/div[2]/div[2]/div[1]/div/div[8]/div/a";
        //private const string RN_TITLE = "//*[@id=\"div-body\"]/div/div[1]/h1";

        //[FindsBy(How = How.XPath, Using = ITEM_SUPPLIER_FIRST_RN)]
        //private IWebElement _itemSupplierFirstRN;

        //[FindsBy(How = How.XPath, Using = WHRN_SUPPLIER_FIRST_RN)]
        //private IWebElement _whrnSupplierFirstRN;

        //[FindsBy(How = How.Id, Using = WHRN_RN_PENCIL_LINK)]
        //private IWebElement _whrn_RN_pencilLink;

        public bool IsLinkToRNOK()
        {
            if(isElementVisible(By.XPath(ITEM_SUPPLIER_FIRST_RN)))
            {
                WaitForElementIsVisible(By.XPath(ITEM_SUPPLIER_FIRST_RN));//*[@id="list-item-with-action"]/div[2]/div[2]/div/div
            }
            else
            { 
                return false; 
            }
            return true;
        }
        public bool IsExpireDateColumnVisible()
        {
            try
            {
                WaitPageLoading();
                // Localiser l'élément de la colonne "Expire date"
                var expireDateColumn = _webDriver.FindElement(By.XPath("//*[@id=\"list-item-with-action\"]/div[2]/div[1]/div[9]"));
                return expireDateColumn.Displayed;
            }
            catch (NoSuchElementException)
            {
                return false; 
            }
        }
        public string GetExpirationDate()
        {
            try
            {
                WaitPageLoading();
                var expirationDateElement = _webDriver.FindElement(By.XPath("//*[@id=\"list-item-with-action\"]/div[2]/div[1]/div[9]"));
                return expirationDateElement.Text;
            }
            catch (NoSuchElementException)
            {
                throw new Exception("Impossible de trouver la date d'expiration dans la colonne Expire date.");
            }
        }
        public bool IsExpirationDateTruncated()
        {
            try
            {
                var expirationDateElement = _webDriver.FindElement(By.XPath("//*[@id=\"list-item-with-action\"]/div[2]/div[1]/div[9]"));

                // Récupérer la largeur de la cellule
                var width = expirationDateElement.Size.Width;
                var scrollWidth = (long)((IJavaScriptExecutor)_webDriver).ExecuteScript("return arguments[0].scrollWidth;", expirationDateElement);

                // Si scrollWidth est supérieur à la largeur réelle de l'élément, cela signifie que le texte est tronqué
                return scrollWidth > width;
            }
            catch (NoSuchElementException)
            {
                throw new Exception("Impossible de vérifier si la date d'expiration est tronquée.");
            }
        }


        //public string GetOnReceiptNoteNumber()
        //{
        //    _whrn_RN_pencilLink = WaitForElementIsVisible(By.XPath(WHRN_RN_PENCIL_LINK));
        //    _whrn_RN_pencilLink.Click();

        //    var wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(60));
        //    wait.Until((driver) => driver.WindowHandles.Count > 1);
        //    var tab = _webDriver.SwitchTo().Window(_webDriver.WindowHandles[1]);

        //    var numberRn = WaitForElementIsVisible(By.XPath(RN_TITLE));

        //    string resultString = Regex.Match(numberRn.Text, "(\\d+)").Value;

        //    tab.Close();

        //    return resultString;
        //}
    }
}
