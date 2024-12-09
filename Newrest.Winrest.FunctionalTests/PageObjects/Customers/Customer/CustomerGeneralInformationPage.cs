using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium.Support.UI;
using System;
using System.IO;
using System.Threading;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Customers.Customer
{
    public class CustomerGeneralInformationPage : PageBase
    {

        public CustomerGeneralInformationPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        //____________________________________ Constantes _____________________________________

        // General
        private const string BACK_TO_LIST = "/html/body/div[2]/a/span[2]";
        public const string IS_ACTIVE = "Customer_IsActive";
        public const string IS_ACTIVE_SPAN = "//*[@id=\"customer-filter-form\"]/div[1]/div/div[1]/div[2]/div[1]/div/div/div/span[1]";

        // Onglets
        private const string CONTACTS_TAB = "//*[@id=\"customerDetailsTab\"]/li[*]/a[text()='Contacts']";
        private const string PRICE_LIST_TAB = "//*[@id=\"customerDetailsTab\"]/li[*]/a[text()='Price list']";
        private const string DISCOUNT_TAB = "//*[@id=\"customerDetailsTab\"]/li[*]/a[text()='Discount']";
        private const string CARDEX_TAB = "//*[@id=\"customerDetailsTab\"]/li[*]/a[text()='Cardex']";
        private const string AGREEMENT_TAB = "//*[@id=\"customerDetailsTab\"]/li[*]/a[text()='Agreement Number']";
        private const string REINVOICE_TAB = "//*[@id=\"customerDetailsTab\"]/li[*]/a[text()='Reinvoice from site to site']";
        private const string BUYONBOARD_TAB = "//*[@id=\"customerDetailsTab\"]/li[*]/a[text()='Buy on Board']";
        private const string GENERAL_INFORMATION_TAB = "tabGeneralInformation";



        // Tableau infos
        private const string CURRENCY = "Customer_CurrencyId";
        private const string PAYMENT_TERM = "Customer_PaymentTypeId";
        private const string CURRENCY_SELECTED = "//*[@id=\"Customer_CurrencyId\"]/option[@selected = 'selected']";
        private const string CONFIRM_DEVISE = "dataAlertCancel";
        private const string CUSTOMER_MAIL = "Customer_ContactMail";
        private const string INVOICE_COMMENT = "Customer_InvoiceComment";
        private const string ACCOUNTING_ID = "Customer_IdAccounting";
        private const string THIRD_ID = "Customer_IdAnalytical";
        private const string ICAO = "Customer_Code";
        private const string CUSTOMER_DELIVERY_NOTE_COMMENT = "Customer_DeliveryNoteComment";
        private const string CATEGORY = "Customer_CustomerTypeId";
        private const string CUSTOMER_NAME = "Customer_Name";
        private const string CUSTOMER_TAXOFFICE = "Customer_TaxOffice";
        private const string CUSTOMER_PROCESSING_DURATION = "processing-duration";
        private const string CUSTOMER_REVISION_PRICE = "datepicker-revision-date";

        private const string AIRPORT_TAX_BTN = "//*[@id=\"customer-filter-form\"]/div[1]/div/div[1]/div[2]/div[2]/div/div";
        private const string BUYONBOARD_BTN = "//*[@id=\"customer-filter-form\"]/div[1]/div/div[1]/div[2]/div[8]/div/div";

        private const string AIRVISION = "//*[@id=\"customer-filter-form\"]/div[1]/div/div[1]/div[2]/div[7]/div/div/div/span[1]";
        private const string BUY_ON_BOARD = "//*[@id=\"customer-filter-form\"]/div[1]/div/div[1]/div[2]/div[8]/div/div";
        private const string CUSTOMER_SIRET = "Customer_SIRET";
        private const string PRO_FORMA_ADDRESS_YES = "//*[@id=\"customer-filter-form\"]/div[1]/div/div[1]/div[2]/div[5]/div/div/div/span[1]";
        private const string PRO_FORMA_ADDRESS_NO = "//*[@id=\"customer-filter-form\"]/div[1]/div/div[1]/div[2]/div[5]/div/div/div/span[3]";
        private const string PRO_FORMA_ADDRESS = "Customer_IsProFormaAddress";



        private const string VATFREE_YES = "//*[@id=\"customer-filter-form\"]/div[1]/div/div[1]/div[2]/div[2]/div/div/div/span[1]";
        private const string VATFREE_NO = "//*[@id=\"customer-filter-form\"]/div[1]/div/div[1]/div[2]/div[2]/div/div/div/span[3]";
        private const string VAT_FREE_SWITCH = "Customer_IsVATFree";
        private const string PRO_FORMA_ADRESSE_SWITCH = "Customer_IsProFormaAddress";
        private const string IMAGE = "logo";
        private const string DELIVERY_NOTE_VALORIZATION = "Customer_IsValorizationDeliveryNote";
        private const string ENGAGEMENT_NO = "Customer_EngagementNumber";
        private const string EXTERNAL_IDENTIFIER = "Customer_ExternalIdentifier";
        private const string FILIGRAN = "Customer_IsFiligran";
        private const string IATA = "Customer_CodeIATA";
        private const string LAST_UPDATE = "last-update";


        //_______________________________________________________Variables_____________________________________________________________

        // General

        [FindsBy(How = How.XPath, Using = AIRVISION)]
        private IWebElement _airvision;

        [FindsBy(How = How.Id, Using = CUSTOMER_NAME)]
        private IWebElement _customername;

        [FindsBy(How = How.XPath, Using = BACK_TO_LIST)]
        private IWebElement _backToList;

        [FindsBy(How = How.Id, Using = IS_ACTIVE)]
        private IWebElement _isActive;

        // Onglets
        [FindsBy(How = How.XPath, Using = CONTACTS_TAB)]
        private IWebElement _contactsTab;

        [FindsBy(How = How.XPath, Using = PRICE_LIST_TAB)]
        private IWebElement _priceListTab;

        [FindsBy(How = How.XPath, Using = DISCOUNT_TAB)]
        private IWebElement _discountTab;

        [FindsBy(How = How.XPath, Using = CARDEX_TAB)]
        private IWebElement _cardexTab;

        [FindsBy(How = How.XPath, Using = AGREEMENT_TAB)]
        private IWebElement _agreementTab;

        [FindsBy(How = How.XPath, Using = REINVOICE_TAB)]
        private IWebElement _reinvoiceTab;

        [FindsBy(How = How.XPath, Using = BUYONBOARD_TAB)]
        private IWebElement _buyOnBoardTab;

        [FindsBy(How = How.XPath, Using = GENERAL_INFORMATION_TAB)]
        private IWebElement _generalInformationTab;

        // Tableau informations
        [FindsBy(How = How.Id, Using = CURRENCY)]
        private IWebElement _currency;
        [FindsBy(How = How.Id, Using = PAYMENT_TERM)]
        private IWebElement _paymentTerm;

        [FindsBy(How = How.Id, Using = CONFIRM_DEVISE)]
        private IWebElement _confirmDevise;

        [FindsBy(How = How.Id, Using = CUSTOMER_MAIL)]
        private IWebElement _customerMail;

        [FindsBy(How = How.Id, Using = INVOICE_COMMENT)]
        private IWebElement _invoiceComment;

        [FindsBy(How = How.Id, Using = ACCOUNTING_ID)]
        private IWebElement _accountingId;

        [FindsBy(How = How.Id, Using = THIRD_ID)]
        private IWebElement _thirdId;

        [FindsBy(How = How.Id, Using = ICAO)]
        private IWebElement _icao;

        [FindsBy(How = How.Id, Using = CUSTOMER_DELIVERY_NOTE_COMMENT)]
        private IWebElement _deliveryNoteComment;

        [FindsBy(How = How.XPath, Using = AIRPORT_TAX_BTN)]
        private IWebElement _airportTaxBtn;

        [FindsBy(How = How.XPath, Using = BUYONBOARD_BTN)]
        private IWebElement _buyOnBoardBtn;

        [FindsBy(How = How.Id, Using = CUSTOMER_TAXOFFICE)]
        private IWebElement _taxOffice;

        [FindsBy(How = How.Id, Using = CUSTOMER_SIRET)]
        private IWebElement _siret;

        [FindsBy(How = How.Id, Using = CUSTOMER_PROCESSING_DURATION)]
        private IWebElement _processingDuration;

        [FindsBy(How = How.Id, Using = CUSTOMER_REVISION_PRICE)]
        private IWebElement _revisionPrice;

        [FindsBy(How = How.XPath, Using = PRO_FORMA_ADDRESS_YES)]
        private IWebElement _proFormaAddressNo;

        [FindsBy(How = How.XPath, Using = PRO_FORMA_ADDRESS_NO)]
        private IWebElement _proFormaAddressYes;

        [FindsBy(How = How.XPath, Using = VATFREE_YES)]
        private IWebElement _vatFreeYes;

        [FindsBy(How = How.XPath, Using = VATFREE_NO)]
        private IWebElement _vatFreeNo;

        [FindsBy(How = How.XPath, Using = DELIVERY_NOTE_VALORIZATION)]
        private IWebElement _dNoteValorization;

        [FindsBy(How = How.XPath, Using = ENGAGEMENT_NO)]
        private IWebElement _engagementNo;

        [FindsBy(How = How.XPath, Using = EXTERNAL_IDENTIFIER)]
        private IWebElement _externalId;

        [FindsBy(How = How.XPath, Using = FILIGRAN)]
        private IWebElement _filigran;

        [FindsBy(How = How.XPath, Using = IATA)]
        private IWebElement _iata;
        [FindsBy(How = How.Id, Using = LAST_UPDATE)]
        private IWebElement _lastUpdate;


        //____________________________________ Méthodes_________________________________________

        // General
        public CustomerPage BackToList()
        {
            _backToList = WaitForElementIsVisible(By.XPath(BACK_TO_LIST));
            _backToList.Click();
            WaitForLoad();

            return new CustomerPage(_webDriver, _testContext);
        }

        // Onglets
        public CustomerPriceListPage GoToGeneralInformation()
        {
            _generalInformationTab = WaitForElementIsVisible(By.Id(GENERAL_INFORMATION_TAB));
            _generalInformationTab.Click();
            WaitPageLoading();

            return new CustomerPriceListPage(_webDriver, _testContext);
        }
        public CustomerContactsPage GoToContactsPage()
        {
            _contactsTab = WaitForElementIsVisible(By.XPath(CONTACTS_TAB), nameof(CONTACTS_TAB));
            _contactsTab.Click();
            WaitForLoad();

            return new CustomerContactsPage(_webDriver, _testContext);
        }

        public CustomerPriceListPage GoToPriceListPage()
        {
            _priceListTab = WaitForElementIsVisible(By.XPath(PRICE_LIST_TAB));
            _priceListTab.Click();
            WaitPageLoading();

            return new CustomerPriceListPage(_webDriver, _testContext);
        }

        public CustomerDiscountPage GoToDiscountPage()
        {
            _discountTab = WaitForElementIsVisible(By.XPath(DISCOUNT_TAB));
            _discountTab.Click();
            WaitForLoad();

            return new CustomerDiscountPage(_webDriver, _testContext);
        }
        public bool isBuyOnBoard()
        {
            var bobChecked = false;

            if (isElementVisible(By.XPath("//*[@id=\"customer-filter-form\"]/div[1]/div/div[1]/div[2]/div[8]/div/div[contains(@class,  'bootstrap-switch-on ')]")))
            {
                bobChecked = true;
            }
            else
            {
                bobChecked = false;
            }
            return bobChecked;
        }
        public CustomerCardexNotifPage GoToCardexPage()
        {
            _cardexTab = WaitForElementIsVisible(By.XPath(CARDEX_TAB));
            _cardexTab.Click();
            WaitForLoad();

            return new CustomerCardexNotifPage(_webDriver, _testContext);
        }

        public CustomerAgreementPage GoToAgreementPage()
        {
            _agreementTab = WaitForElementIsVisible(By.XPath(AGREEMENT_TAB));
            _agreementTab.Click();
            WaitForLoad();

            return new CustomerAgreementPage(_webDriver, _testContext);
        }

        public CustomerReinvoicePage GoToReinvoicePage()
        {
            _reinvoiceTab = WaitForElementIsVisible(By.XPath(REINVOICE_TAB));
            _reinvoiceTab.Click();
            WaitForLoad();

            return new CustomerReinvoicePage(_webDriver, _testContext);
        }

        public CustomerBuyOnBoardPage ClickBobTab()
        {
            _buyOnBoardTab = WaitForElementIsVisible(By.XPath(BUYONBOARD_TAB));
            _buyOnBoardTab.Click();
            WaitPageLoading();

            return new CustomerBuyOnBoardPage(_webDriver, _testContext);
        }

        public CustomerPricePage ClickPriceListTab()
        {
            var _priceTab = WaitForElementIsVisible(By.Id("tabPriceList"));
            _priceTab.Click();
            WaitForLoad();

            return new CustomerPricePage(_webDriver, _testContext);
        }

        // Tableau informations
        public void ModifyInformations(string comment, string mail)
        {
            // Modification du mail du client
            _customerMail = WaitForElementIsVisible(By.Id(CUSTOMER_MAIL));
            _customerMail.SetValue(ControlType.TextBox, mail);

            // Modification du commentaire
            _invoiceComment = WaitForElementIsVisible(By.Id(INVOICE_COMMENT));
            _invoiceComment.SetValue(ControlType.TextBox, comment);

            // Temps d'enregistrement des données
            WaitPageLoading();
        }

        public void SetCurrency(string devise)
        {
            _currency = WaitForElementIsVisible(By.Id(CURRENCY));
            _currency.SetValue(ControlType.DropDownList, devise);

            _confirmDevise = WaitForElementIsVisible(By.Id(CONFIRM_DEVISE));
            _confirmDevise.Click();
            WaitForLoad();
        }
        public void SetPaymentTerm(string paymentTerm)
        {
            WaitLoading();
            _paymentTerm = WaitForElementIsVisible(By.Id(PAYMENT_TERM));
            _paymentTerm.SetValue(ControlType.DropDownList, paymentTerm);
            WaitLoading();
            WaitPageLoading();
            
        }
        public string GetPaymentTerm()
        {
            var paymentTermElement = WaitForElementIsVisible(By.Id(PAYMENT_TERM));
            var selectElement = new SelectElement(paymentTermElement);
            return selectElement.SelectedOption.Text.Trim();
        }

        public void SetCategory(string category)
        {
            var categoryElement = WaitForElementIsVisible(By.Id(CATEGORY));
            categoryElement.SetValue(ControlType.DropDownList, category);
            WaitForLoad();
        }
        public string GetDevise()
        {
            var currencySelected = WaitForElementIsVisible(By.XPath(CURRENCY_SELECTED));
            return currencySelected.Text;
        }

        public string GetIcao()
        {
            _icao = WaitForElementIsVisible(By.Id(ICAO));
            return _icao.GetAttribute("value");
        }

        public string GetCustomerMail()
        {
            _customerMail = WaitForElementIsVisible(By.Id(CUSTOMER_MAIL));
            return _customerMail.GetAttribute("value");
        }

        public string GetComment()
        {
            _invoiceComment = WaitForElementExists(By.Id(INVOICE_COMMENT));
            return _invoiceComment.GetAttribute("value");
        }

        public string GetCustomerName()
        {
            _customername = WaitForElementExists(By.Id(CUSTOMER_NAME));
            return _customername.GetAttribute("value");
        }
        public void SetCustomerName(string newCustomerName)
        {
            var customerNameElement = WaitForElementExists(By.Id(CUSTOMER_NAME));
                customerNameElement.SetValue(ControlType.TextBox, newCustomerName);
            WaitLoading();
            WaitPageLoading();
        }


        public void SetAccountingId(string accountingId)
        {
            Actions action = new Actions(_webDriver);

            _accountingId = WaitForElementExists(By.Id(ACCOUNTING_ID));
            action.MoveToElement(_accountingId).Perform();
            _accountingId.SetValue(ControlType.TextBox, accountingId);
            WaitPageLoading();
        }

        public void SetThirdId(string thirdId)
        {
            Actions action = new Actions(_webDriver);

            _thirdId = WaitForElementExists(By.Id(THIRD_ID));
            action.MoveToElement(_thirdId).Perform();
            _thirdId.SetValue(ControlType.TextBox, thirdId);
            WaitPageLoading();
        }


        public void ActivateAirportTax()
        {
            _airportTaxBtn = WaitForElementIsVisible(By.XPath(AIRPORT_TAX_BTN));

            if (!_airportTaxBtn.GetAttribute("class").Contains("bootstrap-switch-on"))
            {
                _airportTaxBtn.Click();
                WaitPageLoading();
            }

        }

        public void ActiveBuyOnBoard()
        {
            _buyOnBoardBtn = WaitForElementIsVisible(By.XPath(BUY_ON_BOARD));
            _buyOnBoardBtn.Click();
            WaitPageLoading();
        }

        public bool IsBoBVisible()
        {
            if (isElementVisible(By.XPath(BUYONBOARD_TAB)))
            {
                WaitForElementIsVisible(By.XPath(BUYONBOARD_TAB));
                return true;
            }
            else
            {
                return false;
            }
        }

        public void SetCustomerActive()
        {
            try
            {
                var noCheck = WaitForElementIsVisible(By.XPath("//*[@id=\"customer-filter-form\"]/div[1]/div/div[1]/div[2]/div[1]/div/div/div/span[text()=\"No\"]"));
                noCheck.Click();
                WaitPageLoading();
            }
            catch
            {
                // déjà à Active
            }

            WaitForLoad();
        }

        public void SetCustomerInactive()
        {
            try
            {
                var yesCheck = WaitForElementExists(By.XPath("//*[@id=\"customer-filter-form\"]/div[1]/div/div[1]/div[2]/div[1]/div/div/div/span[text()=\"Yes\"]"));
                yesCheck.Click();
                WaitPageLoading();
            }
            catch
            {
                // déjà à No Active
            }
            WaitForLoad();
        }

        public string GetCustomerDeliveryNoteComment()
        {
            _deliveryNoteComment = WaitForElementIsVisible(By.Id(CUSTOMER_DELIVERY_NOTE_COMMENT));
            return _deliveryNoteComment.Text;
        }

        public void SetCustomerDeliveryNoteComment(string deliveryLabelComment)
        {
            _deliveryNoteComment = WaitForElementIsVisible(By.Id(CUSTOMER_DELIVERY_NOTE_COMMENT));
            _deliveryNoteComment.SetValue(ControlType.TextBox, deliveryLabelComment);
            WaitForLoad();
        }

        public void UploadLogo(FileInfo fiUpload)
        {
            Assert.IsTrue(fiUpload.Exists, "Fichier d'entrée non trouve");

            var buttonFile = WaitForElementIsVisible(By.Id("fileSent"));
            buttonFile.SendKeys(fiUpload.FullName);
            //chargement du preview
            Thread.Sleep(2000);
            WaitForLoad();
        }

        /**
         * format : "Portrait" ou "Square"
         */
        public void SetPrintLabelFormat(string format)
        {
            var labelFormatSelect = WaitForElementExists(By.Id("Customer_PrintLabelFormat"));
            new Actions(_webDriver).MoveToElement(labelFormatSelect);
            labelFormatSelect.SetValue(ControlType.DropDownList, format);
            WaitForLoad();
        }
        public string GetCustomerAirVision()
        {
            _airvision = WaitForElementIsVisible(By.XPath(AIRVISION));
            return _airvision.Text;
            WaitForLoad();
        }
       public void SetVATfree(bool switchTo)
        {
            var val = _webDriver.FindElement(By.Id(VAT_FREE_SWITCH));
            bool isChecked = Convert.ToBoolean(val.GetAttribute("checked") ?? "false");
            var parent = val.FindElement(By.XPath(".."));
            if (switchTo ^ isChecked)
            {
                parent.Click();
            } 
            WaitPageLoading();
        }

        public void SetTaxOffice(string taxOffice)
        {
            _taxOffice = WaitForElementIsVisible(By.Id(CUSTOMER_TAXOFFICE));
            _taxOffice.Clear();
            _taxOffice.SendKeys(taxOffice);
            //chargement du preview
            WaitPageLoading();
            WaitLoading();
        }

        public string GetTaxOffice()
        {
            _taxOffice = WaitForElementIsVisible(By.Id(CUSTOMER_TAXOFFICE));
            return _taxOffice.GetAttribute("value");
        }

        public void SetProcessingDuration(string processingDuration)
        {
            _processingDuration = WaitForElementIsVisible(By.Id(CUSTOMER_PROCESSING_DURATION));
            _processingDuration.Clear();
            _processingDuration.SendKeys(processingDuration);
            //chargement du preview
            WaitPageLoading();
        }

        public string GetProcessingDuration()
        {
            _processingDuration = WaitForElementIsVisible(By.Id(CUSTOMER_PROCESSING_DURATION));
            return _processingDuration.GetAttribute("value");
        }

        public void SetRevisionPrice(DateTime revisionPrice)
        {
            _revisionPrice.SetValue(ControlType.DateTime, revisionPrice);
            //chargement du preview
            WaitLoading();
        }

        public string GetRevisionPrice()
        {
            var revisionPrice = WaitForElementIsVisible(By.Id(CUSTOMER_REVISION_PRICE));
            return revisionPrice.GetAttribute("value");
        }

        public string GetThirdId()
        {
            _thirdId = WaitForElementIsVisible(By.Id(THIRD_ID));
            return _thirdId.GetAttribute("value");
        }
        public void SetSiret(string siret)
        {
            _siret.SetValue(ControlType.TextBox,siret);
            //chargement du preview
            WaitLoading();
        }
        public bool IsVATFREE()
        {
            WaitPageLoading();
            var checkbox = _webDriver.FindElement(By.Id("Customer_IsVATFree"));
            if (!checkbox.Selected)
                return false;
            return true;
        }

        public string GetSiret()
        {
            _siret = WaitForElementIsVisible(By.Id(CUSTOMER_SIRET));
            return _siret.GetAttribute("value");
        }

        public bool IsProFormaAddress()
        {
            WaitPageLoading();
            var checkbox = _webDriver.FindElement(By.Id(PRO_FORMA_ADDRESS));
            if (!checkbox.Selected)
                return false;
            return true;
        }
        public void SetProFormaAddress(bool switchTo)
        {
            var el = _webDriver.FindElement(By.Id(PRO_FORMA_ADRESSE_SWITCH));
            bool isChecked = Convert.ToBoolean(el.GetAttribute("checked") ?? "false");
            var parent = el.FindElement(By.XPath(".."));
            if (switchTo ^ isChecked)
            {
                parent.Click();
            }
            WaitPageLoading();
        }

        public void UpToMenueTabs()
        {
            var html = _webDriver.FindElement(By.TagName("html"));
            html.SendKeys(Keys.PageUp);
            html.SendKeys(Keys.PageUp);
            html.SendKeys(Keys.PageUp);
            html.SendKeys(Keys.PageUp);
            WaitPageLoading();
        }
        public void SetComment(string comment)
        {
            Actions action = new Actions(_webDriver);
            _invoiceComment = WaitForElementExists(By.Id(INVOICE_COMMENT));
            action.MoveToElement(_invoiceComment).Perform();
            _invoiceComment.SetValue(ControlType.TextBox, comment);
            WaitPageLoading();
        }

        public bool IsDeliveryNoteValorization()
        {
            WaitPageLoading();
            _dNoteValorization = _webDriver.FindElement(By.Id(DELIVERY_NOTE_VALORIZATION));
            if (!_dNoteValorization.Selected)
                return false;
            return true;
        }

        public void SetDeliveryNoteValorization(bool switchTo)
        {
            _dNoteValorization = _webDriver.FindElement(By.Id(DELIVERY_NOTE_VALORIZATION));
            bool isChecked = Convert.ToBoolean(_dNoteValorization.GetAttribute("checked") ?? "false");
            var parent = _dNoteValorization.FindElement(By.XPath(".."));
            if (switchTo ^ isChecked)
            {
                parent.Click();
            }
            WaitPageLoading();
        }
        public void ModifyICAO(string icao)
        {
            Actions action = new Actions(_webDriver);
            _icao = WaitForElementExists(By.Id(ICAO));
            action.MoveToElement(_icao).Perform();
            _icao.SetValue(ControlType.TextBox, icao);
            WaitPageLoading();
        }
        public bool IsImageExist()
        {
            bool _image = isElementVisible(By.Id(IMAGE));
            return _image;
        }


        public string GetEngagementNo()
        {
            _engagementNo = WaitForElementIsVisible(By.Id(ENGAGEMENT_NO));
            return _engagementNo.GetAttribute("value");
        }

        public void SetEngagementNo(string engagementNo)
        {
            _engagementNo = WaitForElementIsVisible(By.Id(ENGAGEMENT_NO));
            _engagementNo.SetValue(ControlType.TextBox, engagementNo);
            WaitPageLoading();
        }

        public string GetExternalId()
        {
            _externalId = WaitForElementIsVisible(By.Id(EXTERNAL_IDENTIFIER));
            return _externalId.GetAttribute("value");
        }

        public void SetExternalid(string externalId)
        {
            _externalId = WaitForElementIsVisible(By.Id(EXTERNAL_IDENTIFIER));
            _externalId.SetValue(ControlType.TextBox, externalId);
            WaitPageLoading();
        }

        public bool IsFiligran()
        {
            WaitPageLoading();
            _filigran = _webDriver.FindElement(By.Id(FILIGRAN));
            if (!_filigran.Selected)
                return false;
            return true;
        }

        public void SetFiligran(bool switchTo)
        {
            _filigran = _webDriver.FindElement(By.Id(FILIGRAN));
            bool isChecked = Convert.ToBoolean(_filigran.GetAttribute("checked") ?? "false");
            var parent = _filigran.FindElement(By.XPath(".."));
            if (switchTo ^ isChecked)
            {
                parent.Click();
            }
            WaitPageLoading();
        }

        public string GetIata()
        {
            _iata = WaitForElementIsVisible(By.Id(IATA));
            return _iata.GetAttribute("value").Trim();
        }

        public void SetIata(string iata)
        {
            _iata = WaitForElementIsVisible(By.Id(IATA));
            _iata.SetValue(ControlType.TextBox, iata);
            WaitPageLoading();
        }
        public string GetLasUpdate()
        {
            _lastUpdate = WaitForElementIsVisible(By.Id(LAST_UPDATE));
            return _lastUpdate.GetAttribute("value").Trim();
        }
        public void SetLasUpdate(string lastupdate)
        {
            _lastUpdate = WaitForElementIsVisible(By.Id(LAST_UPDATE));
            _lastUpdate.SetValue(ControlType.TextBox, lastupdate);
            WaitPageLoading();
        }
    }

}
