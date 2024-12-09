using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Menus.Datasheet
{
    public class DatasheetGeneralInformationPage : PageBase
    {

        public DatasheetGeneralInformationPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        // ________________________________________________ Constantes ______________________________________________

        private const string NAME = "datasheet-name";
        private const string COMMERCIAL_NAME = "datasheet-commercialName";
        private const string COMMERCIAL_NAME_2 = "datasheet-commercialName2";
        private const string CUSTOMER_CODE = "Datasheet_CustomerCode";
        private const string ACTIVATE = "Datasheet_IsActive";
        private const string DELETE_DATASHEET_Dev = "//*/span[contains(@class,'fas fa-trash-alt')]";
        private const string DELETE_DATASHEET_Patch = "//*/span[contains(@class,'glyphicon-trash')]";
        private const string GUEST_TYPE = "Datasheet_GuestTypeId";
        private const string GUEST_TYPE_SELECTED = "//*[@id=\"Datasheet_GuestTypeId\"]/option[@selected='selected']";

        private const string BACK_TO_LIST = "/html/body/div[2]/a/span[2]";
        private const string DETAILS_TAB = "//*[@id=\"div-body\"]/div/div[1]/table/tbody/tr[2]";

        // ________________________________________________ Variables _______________________________________________

        [FindsBy(How = How.Id, Using = NAME)]
        private IWebElement _name;

        [FindsBy(How = How.Id, Using = COMMERCIAL_NAME)]
        private IWebElement _commercialName;

        [FindsBy(How = How.Id, Using = COMMERCIAL_NAME_2)]
        private IWebElement _commercialName2;

        [FindsBy(How = How.Id, Using = ACTIVATE)]
        private IWebElement _activate;

        [FindsBy(How = How.Id, Using = CUSTOMER_CODE)]
        private IWebElement _customerCode;

        [FindsBy(How = How.Id, Using = GUEST_TYPE)]
        private IWebElement _guestType;


        [FindsBy(How = How.XPath, Using = DETAILS_TAB)]
        private IWebElement _detailsTab;

        // ________________________________________________ Méthodes ________________________________________________

        public string GetName()
        {
            _name = WaitForElementIsVisible(By.Id(NAME));
            return _name.GetAttribute("value");
        }

        public void SetName(string name)
        {
            _name = WaitForElementIsVisible(By.Id(NAME));
            _name.SetValue(ControlType.TextBox, name);
            WaitPageLoading();
        }

        public string GetCommercialName()
        {
            _commercialName = WaitForElementIsVisible(By.Id(COMMERCIAL_NAME));
            return _commercialName.GetAttribute("value");
        }

        public void SetCommercialName(string commercialName)
        {
            _commercialName = WaitForElementIsVisible(By.Id(COMMERCIAL_NAME));
            _commercialName.SetValue(ControlType.TextBox, commercialName);
            WaitPageLoading();
        }

        public string GetCommercialName2()
        {
            _commercialName2 = WaitForElementIsVisible(By.Id(COMMERCIAL_NAME_2));
            return _commercialName2.GetAttribute("value");
        }

        public void SetCommercialName2(string commercialName2)
        {
            _commercialName2 = WaitForElementIsVisible(By.Id(COMMERCIAL_NAME_2));
            _commercialName2.SetValue(ControlType.TextBox, commercialName2);
            WaitPageLoading();
        }

        public void SetActive(bool isActive)
        {
            _activate = WaitForElementExists(By.Id(ACTIVATE));
            _activate.SetValue(ControlType.CheckBox, isActive);
            WaitPageLoading();
        }

        public string GetCustomerCode()
        {
            _customerCode = WaitForElementIsVisible(By.Id(CUSTOMER_CODE));
            return _customerCode.GetAttribute("value");
        }

        public void SetCustomerCode(string customerCode)
        {
            _customerCode = WaitForElementIsVisible(By.Id(CUSTOMER_CODE));
            _customerCode.SetValue(ControlType.TextBox, customerCode);
            WaitPageLoading();
        }

        public string GetGuestType()
        {
            var guestTypeSelected = WaitForElementIsVisible(By.XPath(GUEST_TYPE_SELECTED));
            return guestTypeSelected.Text;
        }

        public void SetGuestType(string guestType)
        {
            _guestType = WaitForElementIsVisible(By.Id(GUEST_TYPE));
            _guestType.SetValue(ControlType.DropDownList, guestType);
            WaitPageLoading();
        }

        public DatasheetDetailsPage ClickOnDetailsTab()
        {
            _detailsTab = WaitForElementIsVisible(By.XPath(DETAILS_TAB));
            _detailsTab.Click();
            WaitForLoad();

            return new DatasheetDetailsPage(_webDriver, _testContext);
        }

        public void DeleteDatasheet()
        {
            var _trash = WaitForElementIsVisible(By.XPath(DELETE_DATASHEET_Dev));
            _trash.Click();
            var _confirmTrash = WaitForElementIsVisible(By.Id("dataConfirmOK"));
            _confirmTrash.Click();
        }

        public DatasheetPage BackToList()
        {
            var _backToList = WaitForElementIsVisible(By.XPath(BACK_TO_LIST));
            _backToList.Click();
            WaitForLoad();

            return new DatasheetPage(_webDriver, _testContext);
        }
    }
}
