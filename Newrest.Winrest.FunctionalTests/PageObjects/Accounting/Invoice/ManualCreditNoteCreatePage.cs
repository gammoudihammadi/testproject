using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Accounting.Invoice
{
    public class ManualCreditNoteCreatePage : PageBase
    {
        private const string SITE_SELECT = "dropdown-site";
        private const string CUSTOMER_INPUT = "//*[@id='createFormVipInvoice']/div[4]/div/span[1]/input[2]";
        private const string CUSTOMER_INPUT_SELECT = "//*[@id='createFormVipInvoice']/div[4]/div/span[1]/div/div/div";
        private const string DETAILS_ADD = "addVipInvoiceDetailBtn";
        private const string DETAILS_FREEPRICE = "//*[contains(@id,'ServiceName')]";
        private const string FREEPRICE_NAME = "Name";
        private const string FREEPRICE_QTY = "Quantity";
        private const string FREEPRICE_PRICE = "SellingPrice";
        private const string PRINT_MENU = "btn-printreportpopup";
        private const string PRINT = "btn-print-submit";
        private const string DROP_BTN = "//*[@id='div-body']/div/div[1]/div/div[1]/button/parent::div";
        
        public ManualCreditNoteCreatePage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        public void FillField_CreateCreditNote(string site, string customer)
        {
            var siteSelect = WaitForElementIsVisible(By.Id(SITE_SELECT));
            siteSelect.SetValue(PageBase.ControlType.DropDownList, site);
            var customerInput = WaitForElementIsVisible(By.XPath(CUSTOMER_INPUT));
            customerInput.Clear();
            customerInput.SendKeys(customer);
            var customerInputSelect = WaitForElementIsVisible(By.XPath(CUSTOMER_INPUT_SELECT));
            customerInputSelect.Click();
        }

        public InvoiceDetailsPage Create()
        {
            var customerInput = WaitForElementIsVisible(By.XPath(CUSTOMER_INPUT));
            customerInput.Submit();
            return new InvoiceDetailsPage(_webDriver, _testContext);
        }

        public void AddDetailsLine(string bananas, string qty, string price)
        {
            var detailsAdd = WaitForElementIsVisible(By.Id(DETAILS_ADD));
            detailsAdd.Click();
            var freePrice = WaitForElementIsVisible(By.XPath(DETAILS_FREEPRICE)); ;

            // ouverture du formulaire
            freePrice.Clear();
            freePrice.SendKeys("[Free");
            WaitPageLoading();
            WaitForLoad();
            Actions a = new Actions(_webDriver);
            a.MoveToElement(freePrice);
            a.MoveByOffset(0, 19);
            a.Click().Perform();

            //animation
            WaitPageLoading();
            WaitForLoad();

            // remplir le formulaire
            var nameInput = WaitForElementIsVisible(By.Id(FREEPRICE_NAME));
            nameInput.SendKeys(bananas);
            var qtyInput = WaitForElementIsVisible(By.Id(FREEPRICE_QTY));
            qtyInput.SendKeys(qty);
            var priceInput = WaitForElementIsVisible(By.Id(FREEPRICE_PRICE));
            priceInput.SendKeys(price);
            priceInput.Submit();
            WaitForLoad();
        }

        public void ClickPrint()
        {
            WaitForLoad();
            var printButton = WaitForElementIsVisible(By.Id(PRINT_MENU));
            printButton.Click();
            var print = WaitForElementIsVisible(By.Id(PRINT));
            print.Click();
            WaitForLoad();
            _webDriver.Navigate().Refresh();
            //Thread.Sleep(10000);
            IsFileLoaded(By.CssSelector("[target='_blank'][class='far fa-file-pdf']"));
            //Results are opened in a new tab, switch the driver to the newly created one
            var wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(60));
            wait.Until((driver) => driver.WindowHandles.Count > 1);
            _webDriver.SwitchTo().Window(_webDriver.WindowHandles[1]);
            Close();

            WaitForLoad();
        }

        public void dropBtnClick()
        {
            Actions a = new Actions(_webDriver);
            var dropBtn = WaitForElementIsVisible(By.XPath(DROP_BTN));
            a.MoveToElement(dropBtn).Perform();
        }

    }
}
