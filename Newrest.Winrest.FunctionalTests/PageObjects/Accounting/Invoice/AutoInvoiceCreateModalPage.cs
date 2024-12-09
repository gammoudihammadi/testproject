using iText.StyledXmlParser.Jsoup.Helper;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Customers.Reinvoice;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium.Support.UI;
using System;
using System.Security.Policy;
using System.Text.RegularExpressions;
using System.Threading;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Accounting.Invoice
{
    public class AutoInvoiceCreateModalPage : PageBase
    {
        public AutoInvoiceCreateModalPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        // _________________________________________ Constantes _______________________________________________

        private const string SITE = "dropdown-site";
        private const string SUBMIT = "//*/button[@onclick='GoToNextStep();']";
        private const string CUSTOMER_PICK = "drop-down-customer-pick-method";
        private const string CUSTOMER_TO_SELECT = "//*[@id=\"SourceCustomers\"]/option[contains(text(),'{0}')]";
        private const string SELECT_ALL = "//*[@id=\"href-select-all-deliveries\"]";
        private const string SELECT_ORDER = "SelectedEntities";
        private const string NEW_INVOICE_POPUP_CLOSE_BUTTON = "//button[contains(text(), 'Close')]";
        private const string CLICK_CREATE = "btnCreateInvoice";
        private const string SELECTED_FLIGHTS = "//*[@id=\"SelectedEntities\"]/option";
        private const string MESSAGE_ERROR = "//*[@id=\"dataAlertModal\"]/div/div/div[2]/div/p/b";
        private const string STRAT_DATE = "StartDate";
        private const string INVOICE_NUMBER = "//*[@id=\"div-body\"]/div/div[1]/h1";
        private const string VALIDATE = "//*[@id=\"IsInvoiceValidated\"]";


        // _________________________________________ Variables _________________________________________________
        [FindsBy(How = How.Id, Using = VALIDATE)]
        private IWebElement _validate;

        [FindsBy(How = How.Id, Using = SITE)]
        private IWebElement _site;
        [FindsBy(How = How.Id, Using = STRAT_DATE)]
        private IWebElement _startdate;

        [FindsBy(How = How.Id, Using = CUSTOMER_PICK)]
        private IWebElement _customerPickMethod;

        [FindsBy(How = How.Id, Using = CUSTOMER_TO_SELECT)]
        private IWebElement _availableCustomerToSelect;

        [FindsBy(How = How.Id, Using = SELECT_ORDER)]
        private IWebElement _selectOrder;

        [FindsBy(How = How.XPath, Using = SUBMIT)]
        private IWebElement _submit;

        [FindsBy(How = How.XPath, Using = NEW_INVOICE_POPUP_CLOSE_BUTTON)]
        private IWebElement _close;

        [FindsBy(How = How.XPath, Using = MESSAGE_ERROR)]
        private IWebElement _messageError;

        [FindsBy(How = How.XPath, Using = INVOICE_NUMBER)]
        private IWebElement _invoiceNumber;

        // _________________________________________ Methodes _________________________________________________

        public void FillField_CreateNewAutoInvoice(string customer, string site, string customerPickMethod, DateTime? dateFrom = null)
        {
            CreateNewAutoInvoiceFirstPart(customer, site, customerPickMethod, dateFrom);

            _availableCustomerToSelect = WaitForElementIsVisible
            (By.XPath(string.Format(CUSTOMER_TO_SELECT, customer)));
            _availableCustomerToSelect.Click();
            WaitForLoad();

            IWebElement addOneCustomerButton = WaitForElementIsVisible(By.XPath("//a[contains(@class,'btn right')][2]"));
            addOneCustomerButton.Click();
            WaitForLoad();

            // le premier busy
            //WaitPageLoading();
            // parasite <div id="div-categories-loading" class="busy-indicator">&nbsp;</div> // deuxième busy infini
            // on a trois busy à la fois
            Submit(true);
            WaitForLoad();
        }

        public void CreateNewAutoInvoiceFirstPart(string customer, string site, string customerPickMethod, DateTime? dateFrom = null)
        {
            _site = WaitForElementIsVisible(By.Id(SITE));
            _site.SetValue(ControlType.DropDownList, site);
            WaitForLoad();

            _customerPickMethod = WaitForElementIsVisible(By.Id(CUSTOMER_PICK));
            _customerPickMethod.SetValue(ControlType.DropDownList, customerPickMethod);
            WaitForLoad();

            IWebElement From = WaitForElementIsVisible(By.Id(STRAT_DATE));
            if (dateFrom != null)
            {
                From.SetValue(ControlType.DateTime, dateFrom);
            }
            else
            {
                From.SetValue(ControlType.DateTime, DateUtils.Now.AddDays(-1));
            }
            WaitForLoad();

            Submit(true);
            WaitForLoad();
        }

        /**
         * A lancer après FillField_CreateNewAutoInvoice
         */
        //Select All 'orders'
        public InvoiceDetailsPage FillFieldSelectAll()
        {
            var selectAll = WaitForElementIsVisible(By.XPath(SELECT_ALL));
            selectAll.Click();
            WaitPageLoading();
            WaitForLoad();
            return Submit();
        }

        public InvoiceDetailsPage FillFieldSelectSomes(int nb)
        {

            WaitPageLoading();
            var select = WaitForElementIsVisible(By.Id("SelectedEntities"));
            SelectElement selectElements = new SelectElement(select);
            for (int i = selectElements.Options.Count - 1; selectElements.Options.Count - nb <= i && i >= 0; i--)
            {
                selectElements.SelectByIndex(i);
                WaitPageLoading();
            }
            return Submit();

        }

        public bool FillFieldSelectOneFlight(string flightNb)
        {
            var flightAllLines = _webDriver.FindElements(By.XPath("//*[@id='SelectedEntities']/option"));
            foreach (var line in flightAllLines)
            {
                line.Click();
            }
            if (isElementExists(By.XPath(string.Format("//*[@id='SelectedEntities']/option[contains(text(), '{0}')]", flightNb))))
            {
                var flightLine = WaitForElementIsVisible(By.XPath(string.Format("//*[@id='SelectedEntities']/option[contains(text(), '{0}')]", flightNb)));
                flightLine.Click();
                return true;
            }
            else
            {
                return false;
            }
        }

        public bool IsSelectOrdersEmpty()
        {
            if (isElementVisible(By.XPath(SELECTED_FLIGHTS)))
            {
                return false;
            }
            return true;
        }

        public bool IsSelectCustomersEmpty(string customer)
        {
            if (isElementVisible(By.XPath(string.Format(CUSTOMER_TO_SELECT, customer))))
            {
                return false;
            }
            return true;
        }

        public InvoiceDetailsPage Submit(bool nextStep = false)
        {
            // button CREATE disabled
            WaitForLoad();
            if (!nextStep)
            {
                // exception Délai d'attente dépassé pour le chargement de la page.
                //WaitPageLoading();
                Thread.Sleep(2000);
                WaitForLoad();
            }
            // exception [20/04/2023 12:58:52] [WaitForElementExists] Element does not exists By.XPath: //*/button[@onclick='GoToNextStep();']
            if (isElementVisible(By.XPath(SUBMIT)))
            {
                _submit = WaitForElementExists(By.XPath(SUBMIT));
            }
            else
            {
                _submit = null;
            }

            int counter = 0;
            while ((_submit == null || !_submit.Displayed) && counter < 30)
            {
                // le second busy
                Thread.Sleep(2000);
                WaitForLoad();
                Console.WriteLine(DateTime.Now.ToString("HH:mm:ss") + " not Displayed " + counter);
                if (isElementVisible(By.XPath(SUBMIT)))
                {
                    _submit = WaitForElementExists(By.XPath(SUBMIT));
                }
                else
                {
                    _submit = null;
                }

                counter++;
            }
            _submit = WaitForElementIsVisible(By.XPath(SUBMIT));
            _submit.Click();
            WaitForLoad();
            WaitLoading();
            return new InvoiceDetailsPage(_webDriver, _testContext);
        }

        public void CheckBoxSeparatedInvoices()
        {
            var checkbox = WaitForElementIsVisible(By.Id("SplitSeparatedInvoices"));
            new Actions(_webDriver).MoveToElement(checkbox).Perform();
            checkbox.Click();
            WaitForLoad();
        }
        public void CloseNewInvoicePopup()
        {
            _close = WaitForElementIsVisible(By.XPath(NEW_INVOICE_POPUP_CLOSE_BUTTON), nameof(NEW_INVOICE_POPUP_CLOSE_BUTTON));
            _close.Click();
            _webDriver.Navigate().Refresh();
            WaitForLoad();
        }
        public int GetSelectedFlightsNumber()
        {
            var flights = _webDriver.FindElements(By.XPath(SELECTED_FLIGHTS));
            return flights.Count;
        }
        public void CloseWarningInvoicePopup()
        {
            if (isElementVisible(By.XPath("//*[@id=\"modal-1\"]/div[1]")))
            {
                var closeButton = WaitForElementExists(By.XPath("//*[@id=\"modal-1\"]/div[3]/button"));
                closeButton.Click();
            }
        }

        public bool TestMsgAfficher(string MsgError)
        {
            var Message = WaitForElementIsVisible(By.XPath(MESSAGE_ERROR));

            if (Message.Text == MsgError)
            {
                return true;
            }
            else return false;
        }
        public string GetInvoiceNumber()
        {
            WaitPageLoading();
            _invoiceNumber = WaitForElementExists(By.XPath(INVOICE_NUMBER));
            return Regex.Match(_invoiceNumber.Text, @"\d+").Value;
        }
        public void Validate()
        {
            ShowValidationMenu();

            while (!isElementVisible(By.XPath(VALIDATE)))
            {
                WaitLoading();
            }
            _validate = WaitForElementIsVisible(By.XPath(VALIDATE));
            _validate.Click();
            //Carl : Attend de validation en requete
            WaitPageLoading();
        }
    }
}
