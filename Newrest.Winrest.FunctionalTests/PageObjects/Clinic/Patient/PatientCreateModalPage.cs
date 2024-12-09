using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;


namespace Newrest.Winrest.FunctionalTests.PageObjects.Clinic.Patient
{
    public class PatientCreateModalPage : PageBase
    {
        public PatientCreateModalPage(IWebDriver _webDriver, TestContext _testContext) : base(_webDriver, _testContext)
        {
        }

        //_______________________________________________________Constantes____________________________________________________________

        private const string PATIENT_FIRST_NAME = "Patient_FirstName";
        private const string PATIENT_LAST_NAME = "Patient_LastName";
        private const string PATIENT_IPP = "Patient_Ipp";
        private const string PATIENT_VISIT_NUMBER = "Patient_VisitNumber";
        private const string PATIENT_SITE = "SelectSite";
        private const string PATIENT_DELIVERY = "SelectDelivery";
        private const string PATIENT_ROOM_NUMBER = "Patient_RoomNumber";
        private const string PATIENT_BED_NUMBER = "Patient_BedNumber";
        private const string PATIENT_DIET_MONITO = "Patient_DieticianMonitoring";
        private const string NEW_PATIENT_SAVE_BUTTON = "btn_save";
        //_______________________________________________________Variables_____________________________________________________________

        [FindsBy(How = How.Id, Using = PATIENT_FIRST_NAME)]
        private IWebElement _patientFirstNameInput;

        [FindsBy(How = How.Id, Using = PATIENT_LAST_NAME)]
        private IWebElement _patientLastNameInput;

        [FindsBy(How = How.Id, Using = PATIENT_IPP)]
        private IWebElement _patientIppInput;

        [FindsBy(How = How.Id, Using = PATIENT_VISIT_NUMBER)]
        private IWebElement _patientVisitNumberInput;

        [FindsBy(How = How.Id, Using = PATIENT_SITE)]
        private IWebElement _patientSiteSelect;

        [FindsBy(How = How.Id, Using = PATIENT_DELIVERY)]
        private IWebElement _patientDeliverySelect;

        [FindsBy(How = How.Id, Using = PATIENT_ROOM_NUMBER)]
        private IWebElement _patientRoomNumberInput;

        [FindsBy(How = How.Id, Using = PATIENT_BED_NUMBER)]
        private IWebElement _patientBedNumberInput;

        [FindsBy(How = How.Id, Using = PATIENT_DIET_MONITO)]
        private IWebElement _patientDietMonito;

        [FindsBy(How = How.Id, Using = NEW_PATIENT_SAVE_BUTTON)]
        private IWebElement _newPatientSaveButton;

        //_______________________________________________________Pages_________________________________________________________________

        public void FillFields_CreatePatientModalPage(string patientFirstName, string patientLastName, string rndNumber, string site, string delivery, string diet = null)
        {
            // Renseigner le prénom
            _patientFirstNameInput = WaitForElementIsVisible(By.Id(PATIENT_FIRST_NAME));
            _patientFirstNameInput.SetValue(ControlType.TextBox, patientFirstName);
            WaitForLoad();

            // Renseigner le nom
            _patientLastNameInput = WaitForElementIsVisible(By.Id(PATIENT_LAST_NAME));
            _patientLastNameInput.SetValue(ControlType.TextBox, patientLastName);
            WaitForLoad();

            // Renseigner l'IPP
            _patientIppInput = WaitForElementIsVisible(By.Id(PATIENT_IPP));
            _patientIppInput.SetValue(ControlType.TextBox, rndNumber);
            WaitForLoad();

            // Renseigner le numéro de visite
            _patientVisitNumberInput = WaitForElementIsVisible(By.Id(PATIENT_VISIT_NUMBER));
            _patientVisitNumberInput.SetValue(ControlType.TextBox, rndNumber);
            WaitForLoad();

            // Renseigner le site
            _patientSiteSelect = WaitForElementIsVisible(By.Id(PATIENT_SITE));
            _patientSiteSelect.Click();
            ScrollUntilElementIsInView(By.XPath("//*[@id='modal-1']//select[@id='SelectSite']/option[text()='" + site + "']"));
            var selectElement = WaitForElementIsVisible(By.XPath("//*[@id='modal-1']//select[@id='SelectSite']/option[text()='"+site+"']"));
            selectElement.Click();
            _patientSiteSelect.Click();
            //_patientSiteSelect.SetValue(ControlType.DropDownList, site);
            //_patientSiteSelect.SendKeys(Keys.Tab);
            WaitForLoad();

            // Renseigner le delivery
            _patientDeliverySelect = WaitForElementIsVisible(By.Id(PATIENT_DELIVERY));
            WaitForLoad();
            _patientDeliverySelect.SetValue(ControlType.DropDownList, delivery);
            _patientDeliverySelect.SendKeys(Keys.Tab);

            // Renseigner la chambre et le lit
            _patientRoomNumberInput = WaitForElementIsVisible(By.Id(PATIENT_ROOM_NUMBER));
            _patientRoomNumberInput.SetValue(ControlType.TextBox, rndNumber);
            WaitForLoad();

            _patientBedNumberInput = WaitForElementIsVisible(By.Id(PATIENT_BED_NUMBER));
            _patientBedNumberInput.SetValue(ControlType.TextBox, rndNumber);
            WaitForLoad();

            if (diet != null)
            {
                _patientDietMonito = WaitForElementIsVisible(By.Id(PATIENT_DIET_MONITO));
                _patientDietMonito.SetValue(ControlType.DropDownList, diet);
                _patientDietMonito.SendKeys(Keys.Tab);
                WaitForLoad();
            }

            _newPatientSaveButton = WaitForElementIsVisible(By.Id(NEW_PATIENT_SAVE_BUTTON));
            _newPatientSaveButton.Click();
            WaitForLoad();
        }
    }
}
