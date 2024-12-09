using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using UglyToad.PdfPig.Content;
using UglyToad.PdfPig.Graphics;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Accounting.Invoice
{
    public class InvoiceDetailsPage : PageBase
    {

        public InvoiceDetailsPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        //____________________________________________Constantes____________________________________________________

        // Général
        private const string BACK_TO_LIST = "BackToList";
        private const string VALIDATE = "//*[@id=\"IsInvoiceValidated\"]";

        private const string INVOICE_HEADER = "//h1[contains(text(), 'Invoice : ')]";

        private const string EXTENDED_BUTTON = "//*[@id=\"div-body\"]/div/div[1]/div/div[1]/button";
        private const string EXPORT = "btn-export-excel";
        private const string PRINT = "btn-printreportpopup";
        private const string CONFIRM_PRINT = "btn-print-submit";
        private const string UPDATE_PRICES = "btn-UpdateElements";
        private const string SEND_BY_MAIL = "//*[@id=\"div-body\"]/div/div[1]/div/div[1]/div/a[text()='Send by email']";
        private const string EMAIL_RECEIVER = "ToAddresses";
        private const string CHOOSE_ATTACHMENT = "add-attachment-input";
        private const string ATTACHMENT = "//span[text()='pieceJointe.jpg']";
        private const string CONFIRM_SEND_EMAIL = "btn-init-async-send-mail";
        private const string EXPORT_SAGE = "btn-export-for-sage";
        private const string SET_INTEGRATION_DATE = "ShowIntegrationDate";
        private const string SET_DATE = "datapicker-integration-date";
        private const string VALIDATE_EXPORT_SAGE = "btn-validate-export";
        private const string ENABLE_SAGE = "btn-enable-export-for-sage";

        private const string INVOICE_NAVIGATE_LINK = "//*[@id=\"modal-1\"]//p[1]/a";
        private const string INVOICE_NUMBER = "//*[@id=\"div-body\"]/div/div[1]/h1";
        private const string FIRST_INVOICE_NUMBER = "//*[@id=\"tableListMenu\"]/tbody/tr/td[4]/a/b";
        private const string PDF = "//*[@id=\"modal-1\"]/div/div/div/form/div[2]/div/div[4]/div/div/div/div/div";
        private const string PDFNAME = "//*[@id=\"modal-1\"]/div/div/div/form/div[2]/div/div[4]/div/div/div/div/div/span";

        // Onglets
        private const string GENERAL_INFORMATION_TAB = "hrefTabContentInformations";
        private const string FOOTER_TAB = "hrefTabInvoiceFooter";
        private const string ACCOUNTING_TAB = "hrefTabContentExportSageWriting";

        // Tableau
        private const string INVOICE_FLIGHT_DETAILS = "//*[@id=\"invoiceInfoTable\"]/tbody/tr";
        private const string ADD_SERVICE = "addVipInvoiceDetailBtn";
        private const string SERVICE_NAME = "//*[contains(@id, 'ServiceName')]";
        private const string FIRST_SERVICE_NAME_VALUE = "//*[@id=\"invoiceDetailTable\"]/tbody/tr[2]/td[2]";
        private const string SERVICE_VALUE = "//*[contains(@id, 'DeliveryQuantity')]";
        private const string FREE_PRICE = "[Free price]";
        private const string FIRST_SERVICE = "//*[@id=\"invoiceDetailTable\"]/tbody/tr[2]";
        private const string DELETE_FIRST_SERVICE = "/html/body/div[3]/div/div[3]/div/div/div/div/div/div[2]/div[1]/div/div/div/table/tbody/tr[2]/td[11]/div/a[3]";
        private const string COMMENT_FIRST_SERVICE = "/html/body/div[3]/div/div[3]/div/div/div/div/div/div[2]/div[1]/div/div/div/table/tbody/tr[2]/td[11]/div/a[2]";
        private const string COMMENT = "Comment";
        private const string VALIDATE_COMMENT = "//button[text()='Validate']";
        private const string COMMENT_ICON = "green-text";
        private const string VAR_RAT = "/html/body/div[3]/div/div[3]/div[1]/div/div/table[1]/tbody/tr[2]/td[3]";

        // Fill
        private const string FILL_SERVICE_CATEGORY = "//*/td[3]/div[2]/select";
        private const string FILL_QUANTITY = "//*/td[4]/div[2]/input";
        private const string FILL_UNIT_PRICE = "//*/td[6]/div[2]/div/div/input";
        private const string FILL_DISCOUNT_SIGNE_INPUT = "//*/td[7]/div[2]/div/div[1]/div/div/input";
        private const string FILL_DISCOUNT_SIGNE_BUTTON = "//*/td[7]/div[2]/div/div[1]/div/div/button[1]";
        private const string FILL_DISCOUNT = "//*/td[7]/div[2]/div/div[2]/div/div/div[2]/input";
        private const string FILL_TOTAL = "//*/td[8]/div[2]/div/div[2]/input";
        private const string SELECT_FIRST_LINE = "//*/table/tbody/tr[2]";
        //Print
        private const string GROUPED_BY_DELIVERY = "//*[@id=\"drop-down-report-type\"]/option[2]";

        private const string SERVICE_CATEGORY = "//*[@id=\"invoiceDetailTable\"]/tbody/tr[2]/td[3]";
        private const string QUANTITY = "//*[@id=\"invoiceDetailTable\"]/tbody/tr[2]/td[4]";
        private const string UNIT_PRICE = "//*[@id=\"invoiceDetailTable\"]/tbody/tr[2]/td[6]";
        private const string NB_FLIGHT = "//*[@id=\"dvInvoiceInfoTable\"]/table/tbody/tr[*]";
        private const string VALIDATE_BTN = "//*[@id=\"div-body\"]/div/div[1]/div/div[2]/button";
        private const string SENDCUSTOMERINVOICE = "//*[@id=\"div-body\"]/div/div[1]/div/div[1]/div/a[3]";
        private const string NEXTBUTTON = "btn-init-async-send-mail";
        private const string VERIFPASSED = "//*[@id=\"div-body\"]/div[1]/div/h1";
        private const string TOTAL_PRICE = "//*[@id=\"dataAlertModal\"]/div/div/div[2]/div/div[3]/div/div/p/b";
        private const string OK_BTN = "//*[@id=\"dataAlertCancel\"]";
        private const string INVOICE_NUM = "//*[@id=\"div-body\"]/div/div[1]/h1";
        private const string SEND_MAIL_POPUP = "//*[@id=\"modal-1\"]";
        private const string SUBJECT = "//*[@id='Subject']";

        private const string PUPUP_TITLE = "//*[@id=\"dataAlertLabel\"]";
        private const string VAT_NAME = "/html/body/div[3]/div/div[3]/div[1]/div/div/div/div/div[2]/div[1]/div/div/div/table/tbody/tr[2]/td[9]/div[2]/select";
        

        //____________________________________________Variables_____________________________________________________

        // General
        [FindsBy(How = How.XPath, Using = UNIT_PRICE)]
        private IWebElement _unit_price_invoice;
        [FindsBy(How = How.XPath, Using = VERIFPASSED)]
        private IWebElement _verifpassed;
        [FindsBy(How = How.Id, Using = NEXTBUTTON)]
        private IWebElement _nextbutton;
        [FindsBy(How = How.XPath, Using = SENDCUSTOMERINVOICE)]
        private IWebElement _sendcustomerinvoice;
        [FindsBy(How = How.XPath, Using = QUANTITY)]
        private IWebElement _quantity_invoice;

        [FindsBy(How = How.XPath, Using = SERVICE_CATEGORY)]
        private IWebElement _servicecategory;

        [FindsBy(How = How.XPath, Using = GROUPED_BY_DELIVERY)]
        private IWebElement _groupedbydelivery;

        [FindsBy(How = How.Id, Using = BACK_TO_LIST)]
        private IWebElement _backToList;

        [FindsBy(How = How.Id, Using = VALIDATE)]
        private IWebElement _validate;

        [FindsBy(How = How.XPath, Using = EXTENDED_BUTTON)]
        private IWebElement _extendedButton;

        [FindsBy(How = How.Id, Using = EXPORT)]
        private IWebElement _export;

        [FindsBy(How = How.Id, Using = PRINT)]
        private IWebElement _print;

        [FindsBy(How = How.Id, Using = CONFIRM_PRINT)]
        private IWebElement _confirmPrint;

        [FindsBy(How = How.XPath, Using = SEND_BY_MAIL)]
        private IWebElement _sendByEmail;

        [FindsBy(How = How.Id, Using = EMAIL_RECEIVER)]
        private IWebElement _emailReceiver;

        [FindsBy(How = How.Id, Using = CONFIRM_SEND_EMAIL)]
        private IWebElement _confirmSendMail;

        [FindsBy(How = How.Id, Using = CHOOSE_ATTACHMENT)]
        private IWebElement _chooseAttachment;

        [FindsBy(How = How.Id, Using = EXPORT_SAGE)]
        private IWebElement _exportSage;

        [FindsBy(How = How.Id, Using = SET_INTEGRATION_DATE)]
        private IWebElement _setIntegrationDate;

        [FindsBy(How = How.Id, Using = SET_DATE)]
        private IWebElement _setDate;

        [FindsBy(How = How.Id, Using = VALIDATE_EXPORT_SAGE)]
        private IWebElement _validateSage;

        [FindsBy(How = How.Id, Using = ENABLE_SAGE)]
        private IWebElement _enableSage;

        [FindsBy(How = How.XPath, Using = INVOICE_NAVIGATE_LINK)]
        private IWebElement _invoiceNavigateLink;

        [FindsBy(How = How.XPath, Using = INVOICE_NUMBER)]
        private IWebElement _invoiceNumber;

        // Onglets
        [FindsBy(How = How.Id, Using = GENERAL_INFORMATION_TAB)]
        private IWebElement _generalInformationTab;

        [FindsBy(How = How.Id, Using = FOOTER_TAB)]
        private IWebElement _footerTab;

        [FindsBy(How = How.Id, Using = ACCOUNTING_TAB)]
        private IWebElement _accountingTab;

        // Tableau
        [FindsBy(How = How.XPath, Using = INVOICE_FLIGHT_DETAILS)]
        private IWebElement _învoiceFlightDetails;

        [FindsBy(How = How.Id, Using = ADD_SERVICE)]
        private IWebElement _addServiceButton;

        [FindsBy(How = How.XPath, Using = SERVICE_NAME)]
        private IWebElement _serviceName;

        [FindsBy(How = How.XPath, Using = FIRST_SERVICE_NAME_VALUE)]
        private IWebElement _firstServiceName;

        [FindsBy(How = How.XPath, Using = SERVICE_VALUE)]
        private IWebElement _serviceValue;

        [FindsBy(How = How.XPath, Using = FIRST_SERVICE)]
        private IWebElement _firstService;

        [FindsBy(How = How.XPath, Using = DELETE_FIRST_SERVICE)]
        private IWebElement _deleteFirstService;

        [FindsBy(How = How.XPath, Using = COMMENT_FIRST_SERVICE)]
        private IWebElement _commentFirstService;

        [FindsBy(How = How.Id, Using = COMMENT)]
        private IWebElement _comment;

        [FindsBy(How = How.XPath, Using = VALIDATE_COMMENT)]
        private IWebElement _validateComment;

        [FindsBy(How = How.ClassName, Using = COMMENT_ICON)]
        private IWebElement _commentIcon;

        // Fill
        [FindsBy(How = How.XPath, Using = FILL_SERVICE_CATEGORY)]
        private IWebElement _serviceCategory;

        [FindsBy(How = How.XPath, Using = FILL_QUANTITY)]
        private IWebElement _quantity;

        [FindsBy(How = How.XPath, Using = FILL_UNIT_PRICE)]
        private IWebElement _unitPrice;

        [FindsBy(How = How.XPath, Using = FILL_DISCOUNT_SIGNE_INPUT)]
        private IWebElement _discount_signe;
        [FindsBy(How = How.XPath, Using = FILL_DISCOUNT)]
        private IWebElement _discount;

        [FindsBy(How = How.XPath, Using = FILL_TOTAL)]
        private IWebElement _total;

        //____________________________________________Pages_____________________________________________________

        // General
        public InvoicesPage BackToList()
        {
            // parfois on est dans index, parfois on est dans invoice après la création automatique d'un invoice
            if (isElementVisible(By.Id(BACK_TO_LIST)))
            {
                _backToList = WaitForElementToBeClickable(By.Id(BACK_TO_LIST));
                WaitPageLoading();
                _backToList.Click();
            }
            WaitForLoad();

            return new InvoicesPage(_webDriver, _testContext);
        }
        public bool IsDisplayed()
        {
            var elements = _webDriver.FindElements(By.XPath(PUPUP_TITLE));
            return elements.Count > 0 && elements[0].Displayed;
        }
        public bool AreElementsNotOverlapping()
        {
            var popUpTitle = _webDriver.FindElement(By.XPath(PUPUP_TITLE));
            var errorMessage = _webDriver.FindElement(By.XPath(TOTAL_PRICE));
            var okButton = _webDriver.FindElement(By.XPath(OK_BTN));

            var titleLocation = popUpTitle.Location;
            var errorMessageLocation = errorMessage.Location;
            var okButtonLocation = okButton.Location;

            var titleSize = popUpTitle.Size;
            var errorMessageSize = errorMessage.Size;
            var okButtonSize = okButton.Size;

            bool isErrorMessageBelowTitle = errorMessageLocation.Y > (titleLocation.Y + titleSize.Height);
            bool isOkButtonBelowErrorMessage = okButtonLocation.Y > (errorMessageLocation.Y + errorMessageSize.Height);

            return isErrorMessageBelowTitle && isOkButtonBelowErrorMessage;
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
        public void NextButton()
        {
            _nextbutton = WaitForElementIsVisible(By.Id(NEXTBUTTON));
            _nextbutton.Click();
            WaitPageLoading();
            WaitForLoad();
        }
        public void SendCustomerInvoice()
        {
            ShowExtendedMenu();

            _sendcustomerinvoice = WaitForElementIsVisible(By.XPath(SENDCUSTOMERINVOICE));
            _sendcustomerinvoice.Click();
            WaitPageLoading();
            WaitForLoad();
        }


        public override void ShowExtendedMenu()
        {
            var actions = new Actions(_webDriver);
            _extendedButton = WaitForElementIsVisible(By.XPath(EXTENDED_BUTTON));
            actions.MoveToElement(_extendedButton).Perform();
            WaitForLoad();
        }

        public void ExportExcelFile()
        {
            ShowExtendedMenu();

            _export = WaitForElementIsVisible(By.Id(EXPORT));
            _export.Click();
            WaitForLoad();

            IsFileLoaded(By.CssSelector("[target='_blank'][class='far fa-file-excel']"));
            ClickPrintButton();
            WaitForDownload();
            Close();
        }

        public FileInfo GetExportExcelFile(FileInfo[] taskFiles)
        {
            List<FileInfo> correctDownloadFiles = new List<FileInfo>();

            foreach (var file in taskFiles)
            {
                //  Test REGEX
                if (IsExcelFileCorrect(file.Name))
                {
                    correctDownloadFiles.Add(file);
                }
            }

            if (correctDownloadFiles.Count <= 0)
            {
                return null;
            }

            var time = correctDownloadFiles[0].LastWriteTimeUtc;
            var correctFile = correctDownloadFiles[0];

            correctDownloadFiles.ForEach(file =>
            {
                if (time < file.LastWriteTimeUtc)
                {
                    time = file.LastWriteTimeUtc;
                    correctFile = file;
                }
            });

            return correctFile;
        }

        public bool IsExcelFileCorrect(string filePath)
        {
            // "export-invoice-2019-12-05 13-54-38.xlsx"

            string mois = "(?:0[1-9]|1[0-2])";         // mois
            string space = "(\\s)";                    // Espace
            string annee = "\\d{4}";                   // annee YYYY
            string jour = "[0-3]\\d";                  // jour
            string heure = "(?:0[0-9]|1[0-9]|2[0-3])"; // heure
            string minutes = "[0-5]\\d";               // minutes
            string secondes = "[0-5]\\d";              // secondes

            Regex r = new Regex("^export-invoice-" + annee + "-" + mois + "-" + jour + space + heure + "-" + minutes + "-" + secondes + ".xlsx$", RegexOptions.IgnoreCase | RegexOptions.Singleline);
            Match m = r.Match(filePath);

            return m.Success;
        }

        public PrintReportPage PrintInvoiceResults(bool GroupedByDelivery = false)
        {
            ShowExtendedMenu();

            _print = WaitForElementIsVisible(By.Id(PRINT));
            _print.Click();
            WaitForLoad();
            if (GroupedByDelivery)
            {
                _groupedbydelivery = WaitForElementIsVisible(By.XPath(GROUPED_BY_DELIVERY));
                _groupedbydelivery.Click();
                WaitForLoad();
            }

            _confirmPrint = WaitForElementIsVisible(By.Id(CONFIRM_PRINT));
            _confirmPrint.Click();
            WaitForLoad();

            IsFileLoaded(By.CssSelector("[target='_blank'][class='far fa-file-pdf']"));

            //Results are opened in a new tab, switch the driver to the newly created one
            var wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(60));
            wait.Until((driver) => driver.WindowHandles.Count > 1);
            _webDriver.SwitchTo().Window(_webDriver.WindowHandles[1]);
            WaitPageLoading();
            WaitPageLoading();
            return new PrintReportPage(_webDriver, _testContext);
        }

        public void SendByMail(string mail, string pathAttachment = null)
        {
            ShowExtendedMenu();

            // Click sur le bouton Send By Mail
            if (IsDev())
            {
                _sendByEmail = WaitForElementIsVisible(By.XPath("//*[@id=\"div-body\"]/div/div[1]/div/div[1]/div/a[4]"));
            }
            else
            {
                _sendByEmail = WaitForElementIsVisible(By.XPath("//*[@id=\"div-body\"]/div/div[1]/div/div[1]/div/a[3]"));
            }
            _sendByEmail.Click();
            WaitForLoad();

            _emailReceiver = WaitForElementIsVisible(By.Id(EMAIL_RECEIVER));
            _emailReceiver.SetValue(ControlType.TextBox, mail);

            if (pathAttachment != null)
            {
                bool isAttachment = JoinAttachment(pathAttachment);
                Assert.IsTrue(isAttachment, "La pièce jointe n'a pas été ajoutée.");
                WaitPageLoading();
                WaitForLoad();
            }

            _confirmSendMail = WaitForElementIsVisible(By.Id(CONFIRM_SEND_EMAIL));
            _confirmSendMail.Click();
            if (pathAttachment != null)
            {
                waitUntilMailWithAttachmentIsSent();
            }
            else
            {
                waitUntilMailIsSent();
            }

        }
        private bool JoinAttachment(string pathAttachment)
        {
            _chooseAttachment = WaitForElementIsVisible(By.Id(CHOOSE_ATTACHMENT));
            _chooseAttachment.SendKeys(pathAttachment);
            WaitForLoad();

            bool pathAttached;
            // On attend que la pièce jointe soit chargée
            pathAttached = isElementVisible(By.XPath(ATTACHMENT));
            WaitForLoad();
            return pathAttached;
        }

        public string ExportSageError(bool isMessage)
        {
            string errorMessage = "";
            ShowExtendedMenu();

            _exportSage = WaitForElementIsVisible(By.Id(EXPORT_SAGE));
            _exportSage.Click();
            WaitForLoad();

            _validateSage = WaitForElementIsVisible(By.Id(VALIDATE_EXPORT_SAGE));
            _validateSage.Click();
            WaitForLoad();

            errorMessage = IsFileInError(By.CssSelector("[class='fa fa-info-circle']"), isMessage);
            ClickPrintButton();

            return errorMessage;
        }

        public void ExportSage(bool integrationDate = false)
        {
            ShowExtendedMenu();
            WaitForLoad();
            _exportSage = WaitForElementIsVisible(By.Id(EXPORT_SAGE));
            _exportSage.Click();

            if (integrationDate)
            {
                var firstDayOftheMonth = new DateTime(DateUtils.Now.Year, DateUtils.Now.Month, 1);

                _setIntegrationDate = WaitForElementExists(By.Id(SET_INTEGRATION_DATE));
                _setIntegrationDate.SetValue(ControlType.CheckBox, true);
                WaitForLoad();

                _setDate = WaitForElementExists(By.Id(SET_DATE));
                _setDate.SetValue(ControlType.DateTime, firstDayOftheMonth);
                WaitForLoad();
            }
            _validateSage = WaitForElementIsVisible(By.Id(VALIDATE_EXPORT_SAGE));

            _validateSage.Click();
            WaitForLoad();

            IsFileLoaded(By.CssSelector("[target='_blank'][class='far fa-file-alt']"));
            ClickPrintButton();

            WaitForDownload();
            Close();
            ClosePrintButton();
        }

        public FileInfo GetExportSAGEFile(FileInfo[] taskFiles, string numberInvoice)
        {
            List<FileInfo> correctDownloadFiles = new List<FileInfo>();

            foreach (var file in taskFiles)
            {
                //  Test REGEX
                if (IsSAGEFileCorrect(file.Name, numberInvoice))
                {
                    correctDownloadFiles.Add(file);
                }
            }

            if (correctDownloadFiles.Count <= 0)
            {
                return null;
            }

            var time = correctDownloadFiles[0].LastWriteTimeUtc;
            var correctFile = correctDownloadFiles[0];

            correctDownloadFiles.ForEach(file =>
            {
                if (time < file.LastWriteTimeUtc)
                {
                    time = file.LastWriteTimeUtc;
                    correctFile = file;
                }
            });

            return correctFile;
        }

        public bool IsSAGEFileCorrect(string filePath, string numberInvoice)
        {
            // "invoice 106335 2019-12-05 15-55-09.txt"

            string mois = "(?:0[1-9]|1[0-2])";         // mois
            string space = "(\\s)";                    // Espace
            string annee = "\\d{4}";                   // annee YYYY
            string jour = "[0-3]\\d";                  // jour
            string heure = "(?:0[0-9]|1[0-9]|2[0-3])"; // heure
            string minutes = "[0-5]\\d";               // minutes
            string secondes = "[0-5]\\d";              // secondes

            Regex r = new Regex("^invoice" + space + numberInvoice + space + annee + "-" + mois + "-" + jour + space + heure + "-" + minutes + "-" + secondes + ".txt$", RegexOptions.IgnoreCase | RegexOptions.Singleline);
            Match m = r.Match(filePath);

            return m.Success;
        }

        public bool IsExportSageEnabled()
        {
            ShowExtendedMenu();
            _exportSage = WaitForElementIsVisible(By.Id(EXPORT_SAGE));

            if (_exportSage.GetAttribute("disabled") != null)
                return false;

            return true;
        }

        public void EnableExportForSage()
        {
            ShowExtendedMenu();

            _enableSage = WaitForElementIsVisible(By.Id(ENABLE_SAGE));
            _enableSage.Click();
            WaitForLoad();
        }

        public bool IsEnableExportForSageEnabled()
        {
            ShowExtendedMenu();
            _enableSage = WaitForElementIsVisible(By.Id(ENABLE_SAGE));

            if (_enableSage.GetAttribute("disabled") != null)
                return false;

            return true;
        }

        public string GetInvoiceNumber()
        {
            WaitPageLoading();
            _invoiceNumber = WaitForElementExists(By.XPath(INVOICE_NUMBER));
            return Regex.Match(_invoiceNumber.Text, @"\d+").Value;
        }
        public string GetInvoiceNumber_details()
        {
            _invoiceNumber = WaitForElementExists(By.XPath("//*[@id=\"div-body\"]/div/div[1]/h1"));
            return Regex.Match(_invoiceNumber.Text, @"\d+").Value;
        }

        public bool IsOnInvoiceDetailsPage()
        {
            if (isElementVisible(By.XPath(INVOICE_HEADER)))
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        public string GetFirstInvoiceNumber()
        {
            var firstInvoiceNumber = WaitForElementIsVisible(By.XPath(FIRST_INVOICE_NUMBER));
            return firstInvoiceNumber.Text;
        }
        public void NavigateToCreatedInvoice()
        {
            _invoiceNavigateLink = WaitForElementExists(By.XPath(INVOICE_NAVIGATE_LINK));
            _invoiceNavigateLink.Click();
            WaitForLoad();

            //Results are opened in a new tab, switch the driver to the newly created one
            var wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(60));
            wait.Until((driver) => driver.WindowHandles.Count > 1);
            _webDriver.SwitchTo().Window(_webDriver.WindowHandles[1]);
        }

        // Onglets
        public InvoiceGeneralInformations ClickOnGeneralInformation()
        {
            _generalInformationTab = WaitForElementIsVisible(By.Id(GENERAL_INFORMATION_TAB));
            _generalInformationTab.Click();
            WaitForLoad();

            return new InvoiceGeneralInformations(_webDriver, _testContext);
        }

        public InvoiceFooterPage ClickOnInvoiceFooter()
        {
            _footerTab = WaitForElementIsVisible(By.Id(FOOTER_TAB));

            _footerTab.Click();
            WaitForLoad();

            return new InvoiceFooterPage(_webDriver, _testContext);
        }

        public InvoiceAccountingPage ClickOnAccounting()
        {
            _accountingTab = WaitForElementIsVisible(By.Id(ACCOUNTING_TAB));
            _accountingTab.Click();
            WaitForLoad();

            return new InvoiceAccountingPage(_webDriver, _testContext);
        }

        // Tableau

        public bool AddService(string serviceName, string quantity = "0")
        {
            try
            {
                // Click sur le bouton +
                _addServiceButton = WaitForElementIsVisible(By.Id(ADD_SERVICE));
                _addServiceButton.Click();
                WaitForLoad();

                if (IsDev())
                {
                    _serviceName = WaitForElementIsVisible(By.XPath("//*[contains(@id, 'accounting-invoice-detail-service-name-')]"));
                }
                else
                {
                    _serviceName = WaitForElementIsVisible(By.XPath(SERVICE_NAME));
                }
                _serviceName.SetValue(ControlType.TextBox, serviceName);
                WaitForLoad();

                // Selection du premier élément de la liste (le prix à retirer au bout)
                var serviceSelected = WaitForElementIsVisible(By.XPath("//div[contains(text(),'" + serviceName + "')]"));
                serviceSelected.Click();
                WaitForLoad();

                // Temps d'attente obligatoire pour les informations du service créé se chargent
                _serviceValue = WaitForElementIsVisible(By.XPath(SERVICE_VALUE));
                _serviceValue.SetValue(ControlType.TextBox, quantity);
                WaitForLoad();

                // Temps d'attente obligatoire pour la prise en compte du service : ne pas enlever
                WaitPageLoading();
            }
            catch
            {
                return false;
            }

            return true;
        }

        public CreateFreePriceModalPage AddFreePrice()
        {
            _addServiceButton = WaitForElementIsVisible(By.Id(ADD_SERVICE));
            _addServiceButton.Click();

            // Entrer "Free price" dans la textBox 
            _serviceName = WaitForElementIsVisible(By.XPath(SERVICE_NAME));
            _serviceName.Click();
            _serviceName.SendKeys(FREE_PRICE);
            WaitForLoad();

            // Selection du premier élément de la liste
            var serviceSelected = WaitForElementIsVisible(By.XPath("//div[text()='" + FREE_PRICE + "']"));
            serviceSelected.Click();
            WaitForLoad();

            return new CreateFreePriceModalPage(_webDriver, _testContext);
        }

        public bool VerifyFreePrice(string serviceName, string quantity)
        {
            _firstService = WaitForElementIsVisible(By.XPath(FIRST_SERVICE));
            _firstService.Click();

            _serviceName = WaitForElementIsVisible(By.XPath(SERVICE_NAME));
            _serviceValue = WaitForElementIsVisible(By.XPath(SERVICE_VALUE));

            return _serviceName.GetAttribute("value").Equals(serviceName) && _serviceValue.GetAttribute("value").Equals(quantity);
        }

        public bool DeleteService()
        {
            try
            {
                _firstService = WaitForElementIsVisible(By.XPath(FIRST_SERVICE));
                _firstService.Click();

                _deleteFirstService = WaitForElementIsVisible(By.XPath(DELETE_FIRST_SERVICE));
                _deleteFirstService.Click();
                WaitForLoad();

                var nbServices = _webDriver.FindElements(By.XPath(FIRST_SERVICE)).Count;

                if (nbServices != 0)
                {
                    return false;
                }
            }
            catch
            {
                return false;
            }

            return true;
        }

        public void CommentService(string comment)
        {
            _firstService = WaitForElementIsVisible(By.XPath(FIRST_SERVICE));
            _firstService.Click();

            _commentFirstService = WaitForElementIsVisible(By.XPath(COMMENT_FIRST_SERVICE));
            _commentFirstService.Click();
            WaitForLoad();

            _comment = WaitForElementIsVisible(By.Id(COMMENT));
            _comment.SetValue(ControlType.TextBox, comment);

            _validateComment = WaitForElementIsVisible(By.XPath(VALIDATE_COMMENT));
            _validateComment.Click();
            WaitForLoad();
        }

        public bool VerifyComment(string comment)
        {
            // Première partie du test : vérifier que l'icône comment est verte
            _commentIcon = WaitForElementIsVisible(By.ClassName(COMMENT_ICON));
            bool isIconDisplayed = _commentIcon.Displayed;

            // Seconde partie du test : vérifier que le commentaire a été pris en compte
            _firstService = WaitForElementIsVisible(By.XPath(FIRST_SERVICE));
            _firstService.Click();

            _commentFirstService = WaitForElementIsVisible(By.XPath(COMMENT_FIRST_SERVICE));
            _commentFirstService.Click();
            WaitForLoad();

            _comment = WaitForElementIsVisible(By.Id(COMMENT));

            return isIconDisplayed && _comment.Text.Equals(comment);
        }

        public string GetInvoiceFirstServiceName()
        {
            _firstServiceName = WaitForElementIsVisible(By.XPath(FIRST_SERVICE_NAME_VALUE));
            return _firstServiceName.Text;
        }

        public string GetInvoiceFlightDetails()
        {
            _învoiceFlightDetails = WaitForElementIsVisible(By.XPath(INVOICE_FLIGHT_DETAILS));
            return _învoiceFlightDetails.Text;
        }

        public void Fill(string serviceCategory, long? quantity, double? unitPrice, double? discount, double? total)
        {
            WaitForLoad();

            if (serviceCategory != null)
            {
                _serviceCategory = WaitForElementIsVisible(By.XPath(FILL_SERVICE_CATEGORY));
                _serviceCategory.SetValue(ControlType.DropDownList, serviceCategory);
                WaitForLoad();
            }

            if (quantity != null)
            {
                _quantity = WaitForElementIsVisible(By.XPath(FILL_QUANTITY));
                _quantity.SetValue(ControlType.TextBox, string.Format("{0}", quantity));
                WaitForLoad();
            }

            if (total == null && unitPrice != null)
            {
                _unitPrice = WaitForElementIsVisible(By.XPath(FILL_UNIT_PRICE));
                _unitPrice.SetValue(ControlType.TextBox, string.Format("{0}", unitPrice));
                WaitForLoad();
            }

            if (discount != null)
            {
                _discount_signe = WaitForElementIsVisible(By.XPath(FILL_DISCOUNT_SIGNE_INPUT));
                string sign = _discount_signe.GetAttribute("value");
                if ((discount < 0 && sign == "+") || (discount >= 0 && sign == "-"))
                {
                    _discount_signe = WaitForElementIsVisible(By.XPath(FILL_DISCOUNT_SIGNE_BUTTON));
                    _discount_signe.Click();
                }
                _discount = WaitForElementIsVisible(By.XPath(FILL_DISCOUNT));
                _discount.SetValue(ControlType.TextBox, Math.Abs((double)discount).ToString());
                WaitForLoad();
            }

            if (total != null)
            {
                //parfois en lecture
                _total = WaitForElementIsVisible(By.XPath(FILL_TOTAL));
                _total.SetValue(ControlType.TextBox, string.Format("{0}", total));
            }
            WaitDisquette();
        }
        public void WaitDisquette()
        {

            // On attend que le disquette apparaisse dans la page

            int compteur = 1;
            bool isLoading = false;
            while (compteur <= 1000)
            {
                if (isElementVisible(By.XPath("//span[contains(@class, 'save')]")))
                {
                    _webDriver.FindElement(By.XPath("//span[contains(@class, 'save')]"));
                    isLoading = true;
                    break;
                }
                else
                {
                    compteur++;
                }
            }


            // Si le sablier n'est pas détecté, on arrête le traitement
            if (isLoading)
            {
                // On attend que le sablier disparaisse pour le chargement de la page
                bool vueSablier = true;
                compteur = 1;

                while (compteur <= 600 && vueSablier)
                {
                    try
                    {
                        var wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(1));
                        wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//span[contains(@class, 'save')]")));
                        compteur++;
                        // Ophélie : ajout d'un sleep pour augmenter le temps d'attente (équivalent à 1 minute max au total)
                        Thread.Sleep(1500);
                    }
                    catch
                    {
                        vueSablier = false;
                    }
                }

                if (vueSablier)
                {
                    Console.WriteLine($"[{DateTime.UtcNow.ToShortDateString()} {DateTime.UtcNow.ToLongTimeString()}] [LoadingPage] Infinite Loading");
                    throw new Exception("Délai d'attente dépassé pour le chargement de la disquette.");
                }
            }
        }

        public double GetTotal()
        {
            WaitPageLoading();
            WaitForLoad();
            _total = WaitForElementExists(By.XPath(FILL_TOTAL));
            return double.Parse(_total.GetAttribute("value").Replace(",", "."), CultureInfo.InvariantCulture);
        }

        public double GetQuantity()
        {
            var quantity = WaitForElementExists(By.XPath(FILL_QUANTITY));
            return double.Parse(quantity.GetAttribute("value").Replace(",", "."), CultureInfo.InvariantCulture);
        }
        public double GetUnitPrice()
        {
            var unitPrice = WaitForElementExists(By.XPath(FILL_UNIT_PRICE));
            return double.Parse(unitPrice.GetAttribute("value").Replace(",", "."), CultureInfo.InvariantCulture);
        }

        public void SelectFirstLine()
        {
            var _firstLine = WaitForElementIsVisible(By.XPath(SELECT_FIRST_LINE));
            _firstLine.Click();
        }
        public bool VerifPassedSendCustomerIvoice()
        {
            if (isElementVisible(By.XPath(VERIFPASSED)))
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        public void FillServiceCategorie(string serviceCategory)
        {
            WaitForLoad();
            var _category = WaitForElementIsVisible(By.XPath(FILL_SERVICE_CATEGORY));
            _category.SetValue(ControlType.DropDownList, serviceCategory);
            _category.SendKeys(Keys.Tab);
            WaitDisquette();

        }

        public void UpdatePrices()
        {
            ShowExtendedMenu();
            var _updatePrices = WaitForElementIsVisible(By.Id(UPDATE_PRICES));
            _updatePrices.Click();
            WaitForLoad();
            SelectFirstLine();
        }

        public void ClickOnSendByEmail()
        {

            ShowExtendedMenu();

            var sendByEmail = WaitForElementIsVisible(By.XPath(SEND_BY_MAIL));
            sendByEmail.Click();
            WaitForLoad();
            WaitPageLoading();

        }
        public void Submit()
        {
            var sendEmail = WaitForElementIsVisible(By.Id(CONFIRM_SEND_EMAIL));
            sendEmail.Click();
            WaitPageLoading();
            WaitForLoad();
            WebDriverWait wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(60));
            wait.Until(emailPopUp => !isElementVisible(By.XPath(SEND_MAIL_POPUP)));


        }
        public bool IsPdfAttachmentVisible()
        {
            try
            {
                WaitPageLoading();
                WaitForLoad();
                var pdfAttachment = WaitForElementIsVisible(By.XPath(PDF));
                WaitForLoad();

                return pdfAttachment.Displayed;
            }
            catch (NoSuchElementException)
            {
                return false;
            }
        }
        public string ExtractNumbersFromElement()
        {
            try
            {
                string elementText = _webDriver.FindElement(By.XPath(PDFNAME)).Text;

                return new string(elementText.Where(char.IsDigit).ToArray());
            }
            catch (NoSuchElementException)
            {
                return string.Empty;
            }
        }
        public string GetSubject()
        {
            WaitLoading();
            var subjectText = _webDriver.FindElement(By.XPath(SUBJECT));
            subjectText.Click();
            var subjectValue = subjectText.GetAttribute("value"); // Get the value of the input field
            return subjectValue;
        }



        public string GetServiceCategory()
        {
            _servicecategory = WaitForElementIsVisible(By.XPath(SERVICE_CATEGORY));
            return _servicecategory.Text;
        }
        public int Get_Quantity()
        {
            _quantity_invoice = WaitForElementIsVisible(By.XPath(QUANTITY));
            return int.Parse(_quantity_invoice.Text);
        }
        public double Get_Unit_Price()
        {
            _unit_price_invoice = WaitForElementIsVisible(By.XPath(UNIT_PRICE));
            string unitprice = new string(_unit_price_invoice.Text.Where(char.IsDigit).ToArray()); // Filtre uniquement les chiffres
            return double.Parse(unitprice.Replace(",", "."));
        }
        public int GetTotal_Flight()
        {
            var nbflight = _webDriver.FindElements(By.XPath(NB_FLIGHT));
            return nbflight.Count;
        }

        public override void ShowValidationMenu()
        {
            var _validateButton = WaitForElementExists(By.XPath(VALIDATE_BTN));
            var actions = new Actions(_webDriver);
            actions.MoveToElement(_validateButton).Perform();
            WaitPageLoading();
        }

        public void ConfirmValidation()
        {
            WaitPageLoading();
            var _confirmValidateButton = WaitForElementExists(By.Id("dataConfirmOK"));
            _confirmValidateButton.Click();
            WaitPageLoading();
        }
        public bool IsValidate()
        {
            ShowValidationMenu();

            _validate = WaitForElementIsVisible(By.XPath(VALIDATE));
            var disabled = _validate.GetAttribute("disabled");
            if (disabled.Equals("true"))
            {
                return true;
            }
            return false;
        }
        public string GetVATRATFromFooterInvoice()
        {
            var taxeFree = WaitForElementIsVisible(By.XPath(VAR_RAT));
            return taxeFree.Text;
        }
        public void waitUntilMailIsSent()
        {
            int compteur = 1;
            while (compteur < 10)
            {
                WaitPageLoading();

                if (!popUpIsVisible())
                {
                    break;
                }
                compteur++;
            }
        }

        public bool popUpIsVisible()
        {
            var showServiceName = _webDriver.FindElements(By.XPath("//*[@id=\"modal-1\"]/div/div/div/form/div[1]/h4"));
            return showServiceName.Count != 0;
        }
        public void waitUntilMailWithAttachmentIsSent()
        {
            int compteur = 1;
            while (compteur < 10)
            {
                WaitPageLoading();

                if (!popUpIsVisibleWithAttachment())
                {
                    break;
                }
                compteur++;
            }
        }

        public bool popUpIsVisibleWithAttachment()
        {
            var showServiceName = _webDriver.FindElements(By.XPath("//*[@id=\"modal-1\"]/div/div/div/div/div/div/form/div[1]/h4"));
            return showServiceName.Count != 0;
        }

        public int ServiceTotalNumber()
        {
            return _webDriver.FindElements(By.XPath("//*[@id=\"invoiceDetailTable\"]/tbody/tr[*]/td[2]")).Count;
        }

        public string GetVatName()
        {
            var vatName = WaitForElementIsVisible(By.XPath(VAT_NAME));
            var vat = new SelectElement(vatName);
            return vat.SelectedOption.Text;
        }
        public bool VerifVatTotal(List<string> mots, string VATRAT, string montant)
        {
            return (mots.Any(a => a.Contains(VATRAT)) && (mots.Any(a => a.Contains(montant))));

        }
        public string GetFirstServiceDiscount()
        {
            var discount = WaitForElementExists(By.XPath("/html/body/div[3]/div/div[3]/div[1]/div/div/div/div/div[2]/div[1]/div/div/div/table/tbody/tr[2]/td[7]/div[2]/div/div[2]/div/div/div[2]/input"));
            return discount.GetAttribute("value");
        }

    }
}
