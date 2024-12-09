using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;
using System.Threading;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Warehouse.ReceiptNotes
{
    public class ReceiptNotesQualityChecks : PageBase
    {
        public ReceiptNotesQualityChecks(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        //___________________________ Constantes ______________________________________

        // Onglets
        private const string ITEMS_TAB = "hrefTabContentItems";
        private const string SECURITY_CHECKS_TAB = "form-update-security-check";
        private const string SECURITY_CHECKS = "//*[starts-with(@id,\"securityCheckStatus-\")][@value='{0}']";
        private const string RADIO_BTN_QUALITY_CHECK = "//html/body/div[3]/div/div[3]/div/div/div[1]/div/div/div/form/div/div[*]/div[*]/div/input[1]";


        // Tableau
        private const string FROZEN_TEMPERATURE = "//*[@id=\"form-update-quality-check\"]/div/div[1]/div[2]/div/div/input";
        private const string DELIVERY_ACCEPTED = "accepted";
        private const string DELIVERY_PARTIAL = "partially";
        private const string CONFIRM_DELIVERY = "dataConfirmOK";
        private const string VERIFIED_BY = "VerifiedBy";
        private const string VALIDATE_BTN = "btn-validate-receipt-note";
        private const string VALIDATE_BTN_MODAL = "btn-popup-validate";
        private const string NOT_APPLICABLE_RADIO = "//*[contains(@id, 'Not-Applicable')]\r\n";

        //___________________________ Variables ______________________________________

        // Onglets
        [FindsBy(How = How.Id, Using = ITEMS_TAB)]
        private IWebElement _itemsTab;

        [FindsBy(How = How.Id, Using = SECURITY_CHECKS_TAB)]
        private IWebElement _securityChecksTab;

        // Tableau
        [FindsBy(How = How.XPath, Using = FROZEN_TEMPERATURE)]
        private IWebElement _frozenTemperature;

        [FindsBy(How = How.Id, Using = DELIVERY_ACCEPTED)]
        private IWebElement _deliveryAccepted;

        [FindsBy(How = How.Id, Using = DELIVERY_PARTIAL)]
        private IWebElement _deliveryPartial;

        [FindsBy(How = How.Id, Using = CONFIRM_DELIVERY)]
        private IWebElement _confirmDelivery;

        [FindsBy(How = How.Id, Using = VERIFIED_BY)]
        private IWebElement _verifiedBy;   
        
        [FindsBy(How = How.Id, Using = VALIDATE_BTN)]
        private IWebElement _validateBtn;   
        
        [FindsBy(How = How.Id, Using = VALIDATE_BTN_MODAL)]
        private IWebElement _validateBtnModal;

        // Valeurs
        [FindsBy(How = How.Id, Using = SECURITY_CHECKS)]
        private IWebElement _securityChecks;

        //__________________________ Méthodes ________________________________________

        // Onglets
        public ReceiptNotesItem ClickOnReceiptNoteItemTab()
        {
            _itemsTab = WaitForElementIsVisibleNew(By.Id(ITEMS_TAB));
            _itemsTab.Click();
            WaitForLoad();

            return new ReceiptNotesItem(_webDriver, _testContext);
        }

        public bool CanClickOnSecurityChecks()
        {
            if (isElementVisible(By.Id(SECURITY_CHECKS_TAB)))
            {
                return true;
            }
            else
            {
                return false;
            }
        }


        public void SetQualityChecks()
        {
                var Elements = _webDriver.FindElements(By.XPath("//html/body/div[3]/div/div[3]/div/div/div[1]/div/div/div/form/div/div[*]/div[*]/div/div[1]/input[@type='radio'][1]"));
                foreach (var elm in Elements)
                {
                    elm.Click();
                }

            //temps de sauvegarde waitForLoad non suffisant et waitPageLoading inadapté (pas de loading bar)
            WaitPageLoading();
        }

        public void SetSecurityChecks(string status)
        {

            IJavaScriptExecutor js = (IJavaScriptExecutor)(IWebDriver)_webDriver;
            js.ExecuteScript("window.scrollTo(0,0)");

            _securityChecks = WaitForElementToBeClickable(By.XPath(string.Format(SECURITY_CHECKS, status)));
            _securityChecks.SetValue(ControlType.RadioButton, true);
            WaitForLoad();
        }

 

        // Tableau
        public void SetFrozenTemperature(string temp)
        {
            _frozenTemperature = WaitForElementExists(By.XPath(FROZEN_TEMPERATURE));
            _frozenTemperature.SetValue(ControlType.TextBox, temp);
        }
        public void SetRefrigeratedVehicleTemperature(string temp)
        {
            var element = WaitForElementExists(By.XPath("/html/body/div[3]/div/div[3]/div[1]/div/div[1]/div/div/div/form/div/div[1]/div[1]/div/div[1]/input"));
            element.SetValue(ControlType.TextBox, temp);
        }
        public string GetVerifiedBy()
        {
            _verifiedBy = WaitForElementExists(By.Id(VERIFIED_BY));
            return _verifiedBy.GetAttribute("value");
        }
       
        public void DeliveryAccepted()
        {
            _deliveryAccepted = WaitForElementIsVisibleNew(By.Id(DELIVERY_ACCEPTED));
            _deliveryAccepted.Click();

            // Ophélie : Temps de prise en compte de la donnée
            Thread.Sleep(2000);
        }

        public void DeliveryPartial()
        {
            _deliveryPartial = WaitForElementIsVisibleNew(By.Id(DELIVERY_PARTIAL));
            _deliveryPartial.Click();

            if (isElementVisible(By.Id(CONFIRM_DELIVERY)))
            {
                _confirmDelivery = WaitForElementIsVisibleNew(By.Id(CONFIRM_DELIVERY));
                _confirmDelivery.Click();
                WaitForLoad();
            }
            else
            {
            }

            // Ophélie : Temps de prise en compte de la donnée
            Thread.Sleep(2000);
        }

        public void ValidateQualityChecks()
        {
            this.ShowValidationMenu();
            _validateBtn = WaitForElementIsVisible(By.Id(VALIDATE_BTN));
            _validateBtn.Click();

            WaitForLoad();

            _validateBtnModal = WaitForElementIsVisible(By.Id(VALIDATE_BTN_MODAL));
            _validateBtnModal.Click();
            WaitPageLoading();
        }

        public void SetNotApplicable()
        {
            var Elements = _webDriver.FindElements(By.XPath(NOT_APPLICABLE_RADIO));
            foreach (var elm in Elements)
            {
                elm.Click();
            }

            WaitPageLoading();
        }
        public void SetQAcceptance()
        {
            var Element = _webDriver.FindElement(By.XPath("/html/body/div[3]/div/div[3]/div[1]/div/div[4]/div/div/div/form/div/div/div[1]/div/div/input[1]"));

            Element.Click();
            WaitLoading();




        }
    }
}
