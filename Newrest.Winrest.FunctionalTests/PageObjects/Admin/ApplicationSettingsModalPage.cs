using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Admin
{
    public class ApplicationSettingsModalPage : PageBase
    {
        public ApplicationSettingsModalPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        //____________________________________________________Constantes_____________________________________________________________

        // New version Flight
        private const string FLIGHT_V0 = "//*[@id=\"v0\"]";
        private const string FLIGHT_V1 = "//*[@id=\"v1\"]";
        private const string FLIGHT_V2 = "//*[@id=\"v2\"]";

        // CustomerPortalURI
        private const string CUSTOMER_PORTAL_URI_VALUE = "StringValue";

        // WinrestTLCountryCode
        private const string WINREST_TL_COUNTRY_CODE_VALUE = "StringValue";

        // WinrestExportTLSageDbOverload
        private const string WINREST_EXPORT_TL_SAGE_DB_OVERLOAD_VALUE = "StringValue";

        // WinrestExportTLSageCountryCodeOverload
        private const string WINREST_EXPORT_TL_SAGE_COUNTRY_CODE_OVERLOAD_VALUE = "StringValue";

        private const string BACKDATING_VALUE = "IntValue";

        private const string WHOMUSTAPPROVEPO_VALUE = "collapseStringListValueFilter";

        private const string SAVE = "//*[@id=\"last\"]";

        //____________________________________________________Variables______________________________________________________________

        // New version Flight
        [FindsBy(How = How.XPath, Using = FLIGHT_V0)]
        private IWebElement _newVersionFlightV0;

        [FindsBy(How = How.XPath, Using = FLIGHT_V1)]
        private IWebElement _newVersionFlightV1;

        [FindsBy(How = How.XPath, Using = FLIGHT_V2)]
        private IWebElement _newVersionFlightV2;

        // CustomerPortalURI
        [FindsBy(How = How.Id, Using = CUSTOMER_PORTAL_URI_VALUE)]
        private IWebElement _customerPortalURI;

        // WinrestTLCountryCode
        [FindsBy(How = How.Id, Using = WINREST_TL_COUNTRY_CODE_VALUE)]
        private IWebElement _winrestTLCountryCode;

        // WinrestExportTLSageDbOverload
        [FindsBy(How = How.Id, Using = WINREST_EXPORT_TL_SAGE_DB_OVERLOAD_VALUE)]
        private IWebElement _winrestExportTLSageDbOverload;

        // WinrestExportTLSageCountryCodeOverload
        [FindsBy(How = How.Id, Using = WINREST_EXPORT_TL_SAGE_COUNTRY_CODE_OVERLOAD_VALUE)]
        private IWebElement _winrestExportTLSageCountryCodeOverload;

        [FindsBy(How = How.Id, Using = BACKDATING_VALUE)]
        private IWebElement _backDating;

        [FindsBy(How = How.Id, Using = WHOMUSTAPPROVEPO_VALUE)]
        private IWebElement _whomustApprovePO;

        [FindsBy(How = How.XPath, Using = SAVE)]
        private IWebElement _save;

        //____________________________________________________Méthodes_______________________________________________________________

        // New version Flight
        public void SetNewVersionFlightValue(int version)
        {
            WaitForElementToBeClickable(By.XPath(SAVE));

            switch (version)
            {
                case 0:
                    _newVersionFlightV0 = WaitForElementExists(By.XPath(FLIGHT_V0));
                    _newVersionFlightV0.SetValue(ControlType.RadioButton, true);
                    break;
                case 1:
                    _newVersionFlightV1 = WaitForElementExists(By.XPath(FLIGHT_V1));
                    _newVersionFlightV1.SetValue(ControlType.RadioButton, true);
                    break;
                case 2:
                    _newVersionFlightV2 = WaitForElementExists(By.XPath(FLIGHT_V2));
                    _newVersionFlightV2.SetValue(ControlType.RadioButton, true);
                    break;
                default:
                    break;
            }
        }

        // CustomerPortalURI
        public void SetCustomerPortalURI(string uri)
        {
            WaitForElementToBeClickable(By.XPath(SAVE));

            _customerPortalURI = WaitForElementIsVisible(By.Id(CUSTOMER_PORTAL_URI_VALUE));
            _customerPortalURI.SetValue(ControlType.TextBox, uri);
        }

        public string GetCustomerPortalURI()
        {
            _customerPortalURI = WaitForElementIsVisible(By.Id(CUSTOMER_PORTAL_URI_VALUE));
            return _customerPortalURI.GetAttribute("value");
        }

        // WinrestTLCountryCode
        public void SetWinrestTLCountryCode(string countryCode)
        {
            WaitForElementToBeClickable(By.XPath(SAVE));

            _winrestTLCountryCode = WaitForElementIsVisible(By.Id(WINREST_TL_COUNTRY_CODE_VALUE));
            _winrestTLCountryCode.SetValue(ControlType.TextBox, countryCode);
        }

        // WinrestExportTLSageDbOverload
        public void SetWinrestExportTLSageDbOverload(string dbName)
        {
            WaitForElementToBeClickable(By.XPath(SAVE));

            _winrestExportTLSageDbOverload = WaitForElementIsVisible(By.Id(WINREST_EXPORT_TL_SAGE_DB_OVERLOAD_VALUE));
            _winrestExportTLSageDbOverload.SetValue(ControlType.TextBox, dbName);
        }

        // WinrestExportTLSageCountryCodeOverload
        public void SetWinrestExportTLSageCountryCodeOverload(string countryCode)
        {
            WaitForElementToBeClickable(By.XPath(SAVE));

            _winrestExportTLSageCountryCodeOverload = WaitForElementIsVisible(By.Id(WINREST_EXPORT_TL_SAGE_COUNTRY_CODE_OVERLOAD_VALUE));
            _winrestExportTLSageCountryCodeOverload.SetValue(ControlType.TextBox, countryCode);
        }
        public string GetBackDating()
        {
            WaitForElementToBeClickable(By.XPath(SAVE));

            _backDating = WaitForElementIsVisible(By.Id(BACKDATING_VALUE));
            return _backDating.GetAttribute("value");
        }

        public void SetBackDating(string backDating)
        {
            WaitForElementToBeClickable(By.XPath(SAVE));

            _backDating = WaitForElementIsVisible(By.Id(BACKDATING_VALUE));
            _backDating.SetValue(ControlType.DropDownList, backDating);
        }

        public void SetWhoMustApprovePO(string role, bool check)
        {
            WaitForElementToBeClickable(By.XPath(SAVE));

            var options = new ComboBoxOptions(WHOMUSTAPPROVEPO_VALUE, role);
            if (check) { options.ClickCheckAllAfterSelection = true; }
            else { options.ClickUncheckAllAfterSelection = true; }
            ComboBoxSelectById(options);
        }

        public ApplicationSettingsPage Save()
        {
            _save = WaitForElementIsVisible(By.XPath(SAVE));
            _save.Click();

            WaitForLoad();

            return new ApplicationSettingsPage(_webDriver, _testContext);
        }
    }
}
