using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.PageObjects;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Admin
{
    public class ApplicationSettingsPage : PageBase
    {

        public ApplicationSettingsPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        // ______________________________________________ Constantes ________________________________________________

        private const string APP_SETTINGS_HACCP_DOC_SETTINGS_BUTTON = "//a[contains(text(),'Haccp document settings')]";
        private const string APP_SETTINGS_CUSTOMIZABLE_COLUMNS_BUTTON = "//a[contains(text(),'customizable columns')]";

        private const string NEW_VERSION_FLIGHT_EDIT = "//*[@id=\"settingsTable\"]/tbody/tr[*]/td[1][normalize-space(text())=\"NewVersionFlight\"]/parent::*/td[4]/a/span";
        private const string NEW_VERSION_FLIGHT_LINE = "//*[@id=\"settingsTable\"]/tbody/tr[*]/td[1][normalize-space(text())=\"NewVersionFlight\"]";

        private const string CUSTOMER_PORTAL_URI_EDIT = "//*[@id=\"settingsTable\"]/tbody/tr[*]/td[1][normalize-space(text())=\"CustomerPortalUri\"]/parent::*/td[4]/a/span";
        private const string CUSTOMER_PORTAL_URI_LINE = "//*[@id=\"settingsTable\"]/tbody/tr[*]/td[1][normalize-space(text())=\"CustomerPortalUri\"]";

        private const string WINREST_TL_COUNTRY_CODE_EDIT = "//*[@id=\"settingsTable\"]/tbody/tr[*]/td[1][normalize-space(text())=\"WinrestTLCountryCode\"]/parent::*/td[4]/a/span";
        private const string WINREST_TL_COUNTRY_CODE_LINE = "//*[@id=\"settingsTable\"]/tbody/tr[*]/td[1][normalize-space(text())=\"WinrestTLCountryCode\"]";

        private const string WINREST_EXPORT_TL_SAGE_DB_OVERLOAD_EDIT = "//*[@id=\"settingsTable\"]/tbody/tr[*]/td[1][normalize-space(text())=\"WinrestExportTLSageDBOverload\"]/parent::*/td[4]/a/span";
        private const string WINREST_EXPORT_TL_SAGE_DB_OVERLOAD_LINE = "//*[@id=\"settingsTable\"]/tbody/tr[*]/td[1][normalize-space(text())=\"WinrestExportTLSageDBOverload\"]";

        private const string WINREST_EXPORT_TL_SAGE_COUNTRY_CODE_OVERLOAD_EDIT = "//*[@id=\"settingsTable\"]/tbody/tr[*]/td[1][normalize-space(text())=\"WinrestExportTLSageCountryCodeOverload\"]/parent::*/td[4]/a/span";
        private const string WINREST_EXPORT_TL_SAGE_COUNTRY_CODE_OVERLOAD_LINE = "//*[@id=\"settingsTable\"]/tbody/tr[*]/td[1][normalize-space(text())=\"WinrestExportTLSageCountryCodeOverload\"]";

        private const string BACKDATING_EDIT = "//*[@id=\"settingsTable\"]/tbody/tr[*]/td[1][normalize-space(text())=\"Backdating\"]/parent::*/td[4]/a/span";
        private const string BACKDATING_LINE = "//*[@id=\"settingsTable\"]/tbody/tr[*]/td[1][normalize-space(text())=\"Backdating\"]";

        private const string WHOMUSTAPPROVEPO_EDIT = "//*[@id=\"settingsTable\"]/tbody/tr[*]/td[1][normalize-space(text())=\"WhoMustApprovePO\"]/parent::*/td[4]/a/span";
        private const string WHOMUSTAPPROVEPO_LINE = "//*[@id=\"settingsTable\"]/tbody/tr[*]/td[1][normalize-space(text())=\"WhoMustApprovePO\"]";

        private const string DOCUMENTS_DISPLAY_EDIT = "//*[@id=\"settingsTable\"]/tbody/tr[*]/td[1][normalize-space(text())=\"DocumentsDisplay\"]/parent::*/td[4]/a/span";

        private const string DOCUMENTS_DISPLAY_VALUE_SELECT = "StringListValue_ms";

        private const string DOCUMENTS_DISPLAY_VALUE_SELECT_CHECK_ALL = "/html/body/div[11]/div/ul/li[1]/a";

        private const string DOCUMENTS_DISPLAY_SAVE = "last";
        private const string APP_SETTINGS_NAME_LIST = "//*[@id=\"settingsTable\"]/tbody/tr[*]/td[1]";
        private const string APP_SETTINGS_VALUE = "//*[@id=\"settingsTable\"]/tbody/tr[{0}]/td[2]";

        // ______________________________________________ Variables _________________________________________________

        [FindsBy(How = How.XPath, Using = APP_SETTINGS_HACCP_DOC_SETTINGS_BUTTON)]
        private IWebElement _HACCPDocumentSettings;

        [FindsBy(How = How.XPath, Using = APP_SETTINGS_CUSTOMIZABLE_COLUMNS_BUTTON)]
        private IWebElement _customizableColumns;

        [FindsBy(How = How.XPath, Using = NEW_VERSION_FLIGHT_EDIT)]
        private IWebElement _newVersionFlightEdit;

        [FindsBy(How = How.XPath, Using = NEW_VERSION_FLIGHT_LINE)]
        private IWebElement _newVersionFlightLine;

        [FindsBy(How = How.XPath, Using = CUSTOMER_PORTAL_URI_EDIT)]
        private IWebElement _customerPortalUriEdit;

        [FindsBy(How = How.XPath, Using = CUSTOMER_PORTAL_URI_LINE)]
        private IWebElement _customerPortalUriLine;

        [FindsBy(How = How.XPath, Using = WINREST_TL_COUNTRY_CODE_EDIT)]
        private IWebElement _winrestTLCountryCodeEdit;

        [FindsBy(How = How.XPath, Using = WINREST_TL_COUNTRY_CODE_LINE)]
        private IWebElement _winrestTLCountryCodeLine;

        [FindsBy(How = How.XPath, Using = WINREST_EXPORT_TL_SAGE_DB_OVERLOAD_EDIT)]
        private IWebElement _winrestExportTLSageDBOverloadEdit;

        [FindsBy(How = How.XPath, Using = WINREST_EXPORT_TL_SAGE_DB_OVERLOAD_LINE)]
        private IWebElement _winrestExportTLSageDBOverloadLine;

        [FindsBy(How = How.XPath, Using = WINREST_EXPORT_TL_SAGE_COUNTRY_CODE_OVERLOAD_EDIT)]
        private IWebElement _winrestExportTLSageCountryCodeOverloadEdit;

        [FindsBy(How = How.XPath, Using = WINREST_EXPORT_TL_SAGE_COUNTRY_CODE_OVERLOAD_LINE)]
        private IWebElement _winrestExportTLSageCountryCodeOverloadLine;

        [FindsBy(How = How.XPath, Using = BACKDATING_EDIT)]
        private IWebElement _backDatingEdit;

        [FindsBy(How = How.XPath, Using = BACKDATING_LINE)]
        private IWebElement _backDatingLine;

        [FindsBy(How = How.XPath, Using = WHOMUSTAPPROVEPO_EDIT)]
        private IWebElement _whoMustApprovePOEdit;

        [FindsBy(How = How.XPath, Using = WHOMUSTAPPROVEPO_LINE)]
        private IWebElement _whoMustApprovePOLine;

        [FindsBy(How = How.XPath, Using = DOCUMENTS_DISPLAY_EDIT)]
        private IWebElement _documentsDisplayEdit;

        [FindsBy(How = How.Id, Using = DOCUMENTS_DISPLAY_VALUE_SELECT)]
        private IWebElement _documentsDisplayValueSelect;

        [FindsBy(How = How.XPath, Using = DOCUMENTS_DISPLAY_VALUE_SELECT_CHECK_ALL)]
        private IWebElement _documentsDisplayValueSelectCheckAll;

        [FindsBy(How = How.XPath, Using = DOCUMENTS_DISPLAY_SAVE)]
        private IWebElement _documentsDisplaySave;

        // ______________________________________________ Méthodes __________________________________________________

        public HACCPDocumentSettingsPage GoToHACCPDocumentSettings()
        {
            _HACCPDocumentSettings = WaitForElementExists(By.XPath(APP_SETTINGS_HACCP_DOC_SETTINGS_BUTTON));
            _HACCPDocumentSettings.Click();
            WaitForLoad();

            return new HACCPDocumentSettingsPage(_webDriver, _testContext);
        }

        public CustomizableColumnsPage GoToCustomizableColumns()
        {
            _customizableColumns = WaitForElementExists(By.XPath(APP_SETTINGS_CUSTOMIZABLE_COLUMNS_BUTTON));
            _customizableColumns.Click();
            WaitForLoad();

            return new CustomizableColumnsPage(_webDriver, _testContext);
        }

        public ApplicationSettingsModalPage GetNewFlightGlobalPage()
        {
            Actions action = new Actions(_webDriver);
            WaitForElementExists(By.XPath(NEW_VERSION_FLIGHT_LINE));
            action.MoveToElement(_newVersionFlightLine).Perform();
            _newVersionFlightLine.Click();

            _newVersionFlightEdit = WaitForElementIsVisible(By.XPath(NEW_VERSION_FLIGHT_EDIT));
            _newVersionFlightEdit.Click();

            WaitForLoad();

            return new ApplicationSettingsModalPage(_webDriver, _testContext);
        }

        public ApplicationSettingsModalPage GetCustomerPortalURIPage()
        {
            Actions action = new Actions(_webDriver);

            WaitForElementExists(By.XPath(CUSTOMER_PORTAL_URI_LINE));
            action.MoveToElement(_customerPortalUriLine).Perform();
            _customerPortalUriLine.Click();

            _customerPortalUriEdit = WaitForElementIsVisible(By.XPath(CUSTOMER_PORTAL_URI_EDIT));
            _customerPortalUriEdit.Click();

            WaitForLoad();

            return new ApplicationSettingsModalPage(_webDriver, _testContext);
        }

        public ApplicationSettingsModalPage GetWinrestTLCountryCodePage()
        {
            Actions action = new Actions(_webDriver);

            WaitForElementExists(By.XPath(WINREST_TL_COUNTRY_CODE_LINE));
            action.MoveToElement(_winrestTLCountryCodeLine).Perform();
            _winrestTLCountryCodeLine.Click();

            _winrestTLCountryCodeEdit = WaitForElementIsVisible(By.XPath(WINREST_TL_COUNTRY_CODE_EDIT));
            _winrestTLCountryCodeEdit.Click();
            WaitForLoad();

            return new ApplicationSettingsModalPage(_webDriver, _testContext);
        }

        public ApplicationSettingsModalPage GetWinrestExportTLSageDbOverloadPage()
        {
            Actions action = new Actions(_webDriver);

            WaitForElementExists(By.XPath(WINREST_EXPORT_TL_SAGE_DB_OVERLOAD_LINE));
            action.MoveToElement(_winrestExportTLSageDBOverloadLine).Perform();
            _winrestExportTLSageDBOverloadLine.Click();

            _winrestExportTLSageDBOverloadEdit = WaitForElementIsVisible(By.XPath(WINREST_EXPORT_TL_SAGE_DB_OVERLOAD_EDIT));
            _winrestExportTLSageDBOverloadEdit.Click();
            WaitForLoad();

            return new ApplicationSettingsModalPage(_webDriver, _testContext);
        }

        public ApplicationSettingsModalPage GetWinrestExportTLSageCountryCodeOverloadPage()
        {
            Actions action = new Actions(_webDriver);

            WaitForElementExists(By.XPath(WINREST_EXPORT_TL_SAGE_COUNTRY_CODE_OVERLOAD_LINE));
            action.MoveToElement(_winrestExportTLSageCountryCodeOverloadLine).Perform();
            _winrestExportTLSageCountryCodeOverloadLine.Click();

            _winrestExportTLSageCountryCodeOverloadEdit = WaitForElementIsVisible(By.XPath(WINREST_EXPORT_TL_SAGE_COUNTRY_CODE_OVERLOAD_EDIT));
            _winrestExportTLSageCountryCodeOverloadEdit.Click();
            WaitForLoad();

            return new ApplicationSettingsModalPage(_webDriver, _testContext);
        }

        public ApplicationSettingsModalPage GetBackDatingPage()
        {
            Actions action = new Actions(_webDriver);

            WaitForElementExists(By.XPath(BACKDATING_LINE));
            action.MoveToElement(_backDatingLine).Perform();
            _backDatingLine.Click();

            _backDatingEdit = WaitForElementIsVisible(By.XPath(BACKDATING_EDIT));
            _backDatingEdit.Click();
            WaitForLoad();

            return new ApplicationSettingsModalPage(_webDriver, _testContext);
        }

        public ApplicationSettingsModalPage GetWhoMustApprovePOPage()
        {
            Actions action = new Actions(_webDriver);

            WaitForElementExists(By.XPath(WHOMUSTAPPROVEPO_LINE));
            action.MoveToElement(_whoMustApprovePOLine).Perform();
            _whoMustApprovePOLine.Click();

            _whoMustApprovePOEdit = WaitForElementIsVisible(By.XPath(WHOMUSTAPPROVEPO_EDIT));
            _whoMustApprovePOEdit.Click();
            WaitForLoad();

            return new ApplicationSettingsModalPage(_webDriver, _testContext);
        }

        public ApplicationSettingsModalPage CheckAllDocumentsDisplay()
        {
            Actions action = new Actions(_webDriver);

            _documentsDisplayEdit = WaitForElementIsVisible(By.XPath(DOCUMENTS_DISPLAY_EDIT));
            _documentsDisplayEdit.Click();
            WaitForLoad();

            _documentsDisplayValueSelect = WaitForElementIsVisible(By.Id(DOCUMENTS_DISPLAY_VALUE_SELECT));
            _documentsDisplayValueSelect.Click();
            WaitForLoad();

            //CheckAll
            _documentsDisplayValueSelectCheckAll = WaitForElementIsVisible(By.XPath(DOCUMENTS_DISPLAY_VALUE_SELECT_CHECK_ALL));
            _documentsDisplayValueSelectCheckAll.Click();
            WaitForLoad();

            _documentsDisplaySave = WaitForElementIsVisible(By.Id(DOCUMENTS_DISPLAY_SAVE));
            _documentsDisplaySave.Click();
            WaitForLoad();

            return new ApplicationSettingsModalPage(_webDriver, _testContext);
        }

        public bool IsWMSInventoryExportExist(string expectedResult)
        {
            var appSettingNameList = _webDriver.FindElements(By.XPath(APP_SETTINGS_NAME_LIST));

            for (int i = 0; i < appSettingNameList.Count; i++)
            {
                if (appSettingNameList[i].Text.Trim() == expectedResult)
                {
                    var appSettingValue = WaitForElementExists(By.XPath(string.Format(APP_SETTINGS_VALUE, i + 1)));

                    if (!string.IsNullOrEmpty(appSettingValue.Text))
                    {
                        return true;
                    }
                }
            }
            return false;
        }

    }
}
