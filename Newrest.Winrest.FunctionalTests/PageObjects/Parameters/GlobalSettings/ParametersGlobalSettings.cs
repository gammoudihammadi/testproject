using DocumentFormat.OpenXml.Bibliography;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Accounting.Edi;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Parameters.GlobalSettings
{
    public class ParametersGlobalSettings : PageBase
    {

        public ParametersGlobalSettings(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {

        }

        //____________________________________________________Constantes_____________________________________________________________

        // DateFormatPicker
        private const string DATE_FORMAT_PICKER_EDIT = "//*[@id=\"tabContentParameters\"]/table/tbody/tr[*]/td[3][normalize-space(text())=\"DateFormatPicker\"]/../td[6]/a/span";
        private const string DATE_FORMAT_PICKER_LINE = "//*[@id=\"tabContentParameters\"]/table/tbody/tr[*]/td[3][normalize-space(text())=\"DateFormatPicker\"]";

        // Decimal separator
        private const string DECIMAL_SEPARATOR_EDIT = "//*[@id=\"tabContentParameters\"]/table/tbody/tr[*]/td[3][normalize-space(text())=\"DecimalSeparator\"]/../td[6]/a/span";
        private const string DECIMAL_SEPARATOR_LINE = "//*[@id=\"tabContentParameters\"]/table/tbody/tr[*]/td[3][normalize-space(text())=\"DecimalSeparator\"]";

        // Show Interim
        //private const string SHOW_INTERIM_EDIT = "//*[@id=\"tabContentParameters\"]/table/tbody/tr[*]/td[3][normalize-space(text())=\"ShowInterim\"]/../td[6]/a/span";
        //private const string SHOW_INTERIM_LINE = "//*[@id=\"tabContentParameters\"]/table/tbody/tr[*]/td[3][normalize-space(text())=\"ShowInterim\"]";

        // Show WarehouseclaimV3
        private const string SHOW_CLAIM_V3_EDIT = "//*[@id=\"tabContentParameters\"]/table/tbody/tr[*]/td[3][normalize-space(text())=\"WarehouseClaimV3\"]/../td[6]/a/span";
        private const string SHOW_CLAIM_V3_LINE = "//*[@id=\"tabContentParameters\"]/table/tbody/tr[*]/td[3][normalize-space(text())=\"WarehouseClaimV3\"]";

        // Is Sage Auto Models Enabled
        private const string IS_SAGE_AUTO_ENABLED_EDIT = "//*[@id=\"tabContentParameters\"]/table/tbody/tr[*]/td[3][normalize-space(text())=\"IsSageOrCegidAutoModeIsEnabled\"]/../td[6]/a/span";
        private const string IS_SAGE_AUTO_ENABLED_LINE = "//*[@id=\"tabContentParameters\"]/table/tbody/tr[*]/td[3][normalize-space(text())=\"IsSageOrCegidAutoModeIsEnabled\"]";

        // New version Keyword
        private const string NEW_VERSION_KEYWORD_EDIT = "//*[@id=\"tabContentParameters\"]/table/tbody/tr[*]/td[3][normalize-space(text())=\"IsNewKeywordSearch\"]/../td[6]/a/span";
        private const string NEW_VERSION_KEYWORD_LINE = "//*[@id=\"tabContentParameters\"]/table/tbody/tr[*]/td[3][normalize-space(text())=\"IsNewKeywordSearch\"]";

        // New search mode
        private const string NEW_SEARCH_MODE_EDIT = "//*[@id=\"tabContentParameters\"]/table/tbody/tr[*]/td[3][normalize-space(text())=\"IsNewSearchModeActive\"]/../td[6]/a/span";
        private const string NEW_SEARCH_MODE_LINE = "//*[@id=\"tabContentParameters\"]/table/tbody/tr[*]/td[3][normalize-space(text())=\"IsNewSearchModeActive\"]";

        // New claim type
        private const string NEW_CLAIM_TYPE_EDIT = "//*[@id=\"tabContentParameters\"]/table/tbody/tr[*]/td[3][normalize-space(text())=\"IsNewClaimTypeActive\"]/../td[6]/a/span";
        private const string NEW_CLAIM_TYPE_LINE = "//*[@id=\"tabContentParameters\"]/table/tbody/tr[*]/td[3][normalize-space(text())=\"IsNewClaimTypeActive\"]";

        // New group display
        private const string NEW_GROUP_DISPLAY_EDIT = "//*[@id=\"tabContentParameters\"]/table/tbody/tr[*]/td[3][normalize-space(text())=\"IsNewGroupDisplayActive\"]/../td[6]/a/span";
        private const string NEW_GROUP_DISPLAY_LINE = "//*[@id=\"tabContentParameters\"]/table/tbody/tr[*]/td[3][normalize-space(text())=\"IsNewGroupDisplayActive\"]";

        // Sub group
        private const string SUB_GROUP_EDIT = "//*[@id=\"tabContentParameters\"]/table/tbody/tr[*]/td[3][normalize-space(text())=\"IsSubGroupFunctionActive\"]/../td[6]/a/span";
        private const string SUB_GROUP_LINE = "//*[@id=\"tabContentParameters\"]/table/tbody/tr[*]/td[3][normalize-space(text())=\"IsSubGroupFunctionActive\"]";

        // New version item details
        private const string NEW_VERSION_ITEM_DETAILS_EDIT = "//*[@id=\"tabContentParameters\"]/table/tbody/tr[*]/td[3][normalize-space(text())=\"NewVersionItemDetail\"]/../td[6]/a/span";
        private const string NEW_VERSION_ITEM_DETAILS_LINE = "//*[@id=\"tabContentParameters\"]/table/tbody/tr[*]/td[3][normalize-space(text())=\"NewVersionItemDetail\"]";

        // New supplier invoice
        private const string NEW_SUPPLIER_INVOICE_EDIT = "//*[@id=\"tabContentParameters\"]/table/tbody/tr[*]/td[3][normalize-space(text())=\"NewSupplierInvoiceVersion\"]/parent::*/td[6]/a/span";
        private const string NEW_SUPPLIER_INVOICE_LINE = "//*[@id=\"tabContentParameters\"]/table/tbody/tr[*]/td[3][normalize-space(text())=\"NewSupplierInvoiceVersion\"]";

        // Compare
        private const string IS_COMPARE_ACTIVE_EDIT = "//*[@id=\"tabContentParameters\"]/table/tbody/tr[*]/td[3][normalize-space(text())=\"IsCompareActive\"]/parent::*/td[6]/a/span";
        private const string IS_COMPARE_ACTIVE_LINE = "//*[@id=\"tabContentParameters\"]/table/tbody/tr[*]/td[3][normalize-space(text())=\"IsCompareActive\"]";

        // Number of items to compare
        private const string NUMBER_ITEMS_COMPARE_EDIT = "//*[@id=\"tabContentParameters\"]/table/tbody/tr[*]/td[3][normalize-space(text())=\"NumberOfItemsToCompare\"]/parent::*/td[6]/a/span";
        private const string NUMBER_ITEMS_COMPARE_LINE = "//*[@id=\"tabContentParameters\"]/table/tbody/tr[*]/td[3][normalize-space(text())=\"NumberOfItemsToCompare\"]";

        // Recipe Expiry Date
        private const string RECIPE_EXPIRY_DATE_LINE = "//*[@id=\"tabContentParameters\"]/table/tbody/tr[*]/td[3][normalize-space(text())=\"RecipeExpiryDate\"]";
        private const string RECIPE_EXPIRY_DATE_EDIT = "//*[@id=\"tabContentParameters\"]/table/tbody/tr[*]/td[3][normalize-space(text())=\"RecipeExpiryDate\"]/parent::*/td[6]/a/span";

        // Show Interim
        private const string SHOW_INTERIM_DETAILS_LINE = "//*[@id=\"tabContentParameters\"]/table/tbody/tr[*]/td[3][normalize-space(text())=\"ShowInterim\"]";

        // Back Date
        private const string BACKDATING_LINE = "//*[@id=\"tabContentParameters\"]/table/tbody/tr[*]/td[3][normalize-space(text())=\"Backdating\"]";
        private const string BACKDATING_EDIT = "//*[@id=\"tabContentParameters\"]/table/tbody/tr[*]/td[3][normalize-space(text())=\"Backdating\"]/parent::*/td[6]/a/span";

        private const string TRANSLATIONTAB = "//*[@id=\"paramGlobalSettingTab\"]/li[6]/a";

        private const string CULTURE = "SelectedCulture";
        private const string MODULE = "SelectedModule";
        private const string SEARCH = "SearchPattern";
        private const string TRANSLATIONLIST = "//*[@id=\"translation-list-content\"]/div[1]/div/ul/li[*]/div/div/div/form/div/div[2]/div[2]/span";
        private const string SECTIONLIST = "/html/body/div[2]/div/div/div/table/tbody/tr[*]/td[2]";
        //____________________________________________________Variables_______________________________________________________________
        // DateFormatPicker
        [FindsBy(How = How.XPath, Using = DATE_FORMAT_PICKER_EDIT)]
        private IWebElement _dateFormatPickerEdit;

        [FindsBy(How = How.XPath, Using = DATE_FORMAT_PICKER_LINE)]
        private IWebElement _dateFormatPickerLine;

        // Decimal separator
        [FindsBy(How = How.XPath, Using = DECIMAL_SEPARATOR_EDIT)]
        private IWebElement _decimalSeparatorEdit;

        [FindsBy(How = How.XPath, Using = DECIMAL_SEPARATOR_LINE)]
        private IWebElement _decimalSeparatorLine;

        // Show Interim
        //[FindsBy(How = How.XPath, Using = SHOW_INTERIM_EDIT)]
        //private IWebElement _showInterimEdit;

        //[FindsBy(How = How.XPath, Using = SHOW_INTERIM_LINE)]
        //private IWebElement _showInterimLine;

        // Is Sage Auto Models Enabled
        [FindsBy(How = How.XPath, Using = IS_SAGE_AUTO_ENABLED_EDIT)]
        private IWebElement _isSageAutoEnabledEdit;

        [FindsBy(How = How.XPath, Using = IS_SAGE_AUTO_ENABLED_LINE)]
        private IWebElement _isSageAutoEnabledLine;

        // New version Keyword
        [FindsBy(How = How.XPath, Using = NEW_VERSION_KEYWORD_EDIT)]
        private IWebElement _newVersionKeywordEdit;

        [FindsBy(How = How.XPath, Using = NEW_VERSION_KEYWORD_LINE)]
        private IWebElement _newVersionKeywordLine;

        // New search mode
        [FindsBy(How = How.XPath, Using = NEW_SEARCH_MODE_EDIT)]
        private IWebElement _newSearchModeEdit;

        [FindsBy(How = How.XPath, Using = NEW_SEARCH_MODE_LINE)]
        private IWebElement _newSearchModeLine;

        // New claim type
        [FindsBy(How = How.XPath, Using = NEW_CLAIM_TYPE_EDIT)]
        private IWebElement _newClaimTypeEdit;

        [FindsBy(How = How.XPath, Using = NEW_CLAIM_TYPE_LINE)]
        private IWebElement _newClaimTypeLine;

        // New group display
        [FindsBy(How = How.XPath, Using = NEW_GROUP_DISPLAY_EDIT)]
        private IWebElement _newGroupDisplayEdit;

        [FindsBy(How = How.XPath, Using = NEW_GROUP_DISPLAY_LINE)]
        private IWebElement _newGroupDisplayLine;

        // Sub group
        [FindsBy(How = How.XPath, Using = SUB_GROUP_EDIT)]
        private IWebElement _subGroupEdit;

        [FindsBy(How = How.XPath, Using = SUB_GROUP_LINE)]
        private IWebElement _subGroupLine;

        // New version item details
        [FindsBy(How = How.XPath, Using = NEW_VERSION_ITEM_DETAILS_EDIT)]
        private IWebElement _newVersionItemDetailsEdit;

        [FindsBy(How = How.XPath, Using = NEW_VERSION_ITEM_DETAILS_LINE)]
        private IWebElement _newVersionItemDetailsLine;

        // New supplier invoice
        [FindsBy(How = How.XPath, Using = NEW_SUPPLIER_INVOICE_EDIT)]
        private IWebElement _newSupplierInvoiceEdit;

        [FindsBy(How = How.XPath, Using = NEW_SUPPLIER_INVOICE_LINE)]
        private IWebElement _newSupplierInvoiceLine;

        // Compare
        [FindsBy(How = How.XPath, Using = IS_COMPARE_ACTIVE_EDIT)]
        private IWebElement _isCompareActiveEdit;

        [FindsBy(How = How.XPath, Using = IS_COMPARE_ACTIVE_LINE)]
        private IWebElement _isCompareActiveLine;

        // Number of items to compare
        [FindsBy(How = How.XPath, Using = NUMBER_ITEMS_COMPARE_EDIT)]
        private IWebElement _numberOfItemsEdit;

        [FindsBy(How = How.XPath, Using = NUMBER_ITEMS_COMPARE_LINE)]
        private IWebElement _numberOfItemsLine;

        [FindsBy(How = How.XPath, Using = RECIPE_EXPIRY_DATE_EDIT)]
        private IWebElement _recipeExpiryDateEdit;

        [FindsBy(How = How.XPath, Using = RECIPE_EXPIRY_DATE_LINE)]
        private IWebElement _recipeExpiryDateLine;

        [FindsBy(How = How.XPath, Using = BACKDATING_EDIT)]
        private IWebElement _backDatingEdit;

        [FindsBy(How = How.XPath, Using = BACKDATING_LINE)]
        private IWebElement _backDatingLine;

        [FindsBy(How = How.XPath, Using = TRANSLATIONTAB)]
        private IWebElement _translationTab;

        [FindsBy(How = How.Id, Using = CULTURE)]
        private IWebElement _culture;

        [FindsBy(How = How.Id, Using = MODULE)]
        private IWebElement _module;

        [FindsBy(How = How.Id, Using = SEARCH)]
        private IWebElement _search;

        [FindsBy(How = How.XPath, Using = TRANSLATIONLIST)]
        private IWebElement _translationList;
        // _______________________________________________ METHODES ______________________________________________________

        public ParametersGlobalSettingsModalPage GetDateFormatPickerPage()
        {
            Actions action = new Actions(_webDriver);

            action.MoveToElement(_dateFormatPickerLine).Perform();
            _dateFormatPickerLine.Click();

            _dateFormatPickerEdit = WaitForElementIsVisible(By.XPath(DATE_FORMAT_PICKER_EDIT));
            _dateFormatPickerEdit.Click();
            WaitForLoad();

            return new ParametersGlobalSettingsModalPage(_webDriver, _testContext);
        }

        public ParametersGlobalSettingsModalPage GetDecimalSeparatorPage()
        {
            Actions action = new Actions(_webDriver);

            action.MoveToElement(_decimalSeparatorLine).Perform();
            _decimalSeparatorLine.Click();

            _decimalSeparatorEdit = WaitForElementIsVisible(By.XPath(DECIMAL_SEPARATOR_EDIT));
            _decimalSeparatorEdit.Click();
            WaitForLoad();

            return new ParametersGlobalSettingsModalPage(_webDriver, _testContext);
        }

        public ParametersGlobalSettingsModalPage IsSageAutoEnabled()
        {
            Actions action = new Actions(_webDriver);

            action.MoveToElement(_isSageAutoEnabledLine).Perform();
            _isSageAutoEnabledLine.Click();
            WaitForLoad();

            _isSageAutoEnabledEdit = WaitForElementIsVisible(By.XPath(IS_SAGE_AUTO_ENABLED_EDIT));
            _isSageAutoEnabledEdit.Click();
            WaitForLoad();

            return new ParametersGlobalSettingsModalPage(_webDriver, _testContext);
        }

        public ParametersGlobalSettingsModalPage GetNewKeywordGlobalPage()
        {
            Actions action = new Actions(_webDriver);

            action.MoveToElement(_newVersionKeywordLine).Perform();
            _newVersionKeywordLine.Click();

            _newVersionKeywordEdit = WaitForElementIsVisible(By.XPath(NEW_VERSION_KEYWORD_EDIT));
            _newVersionKeywordEdit.Click();
            WaitForLoad();

            return new ParametersGlobalSettingsModalPage(_webDriver, _testContext);
        }

        public ParametersGlobalSettingsModalPage GetIsNewSearchModePage()
        {
            Actions action = new Actions(_webDriver);

            action.MoveToElement(_newSearchModeLine).Perform();
            _newSearchModeLine.Click();

            _newSearchModeEdit = WaitForElementIsVisible(By.XPath(NEW_SEARCH_MODE_EDIT));
            _newSearchModeEdit.Click();
            WaitForLoad();

            return new ParametersGlobalSettingsModalPage(_webDriver, _testContext);
        }

        public ParametersGlobalSettingsModalPage GetIsNewClaimTypePage()
        {
            Actions action = new Actions(_webDriver);

            action.MoveToElement(_newClaimTypeLine).Perform();
            _newClaimTypeLine.Click();

            _newClaimTypeEdit = WaitForElementIsVisible(By.XPath(NEW_CLAIM_TYPE_EDIT));
            _newClaimTypeEdit.Click();
            WaitForLoad();

            return new ParametersGlobalSettingsModalPage(_webDriver, _testContext);
        }

        public ParametersGlobalSettingsModalPage GetIsNewGroupDisplayPage()
        {
            Actions action = new Actions(_webDriver);

            action.MoveToElement(_newGroupDisplayLine).Perform();
            _newGroupDisplayLine.Click();

            _newGroupDisplayEdit = WaitForElementIsVisible(By.XPath(NEW_GROUP_DISPLAY_EDIT));
            _newGroupDisplayEdit.Click();
            WaitForLoad();

            return new ParametersGlobalSettingsModalPage(_webDriver, _testContext);
        }

        public ParametersGlobalSettingsModalPage GetSubGroupGlobalPage()
        {
            Actions action = new Actions(_webDriver);

            action.MoveToElement(_subGroupLine).Perform();
            _subGroupLine.Click();

            _subGroupEdit = WaitForElementIsVisible(By.XPath(SUB_GROUP_EDIT));
            _subGroupEdit.Click();
            WaitForLoad();

            return new ParametersGlobalSettingsModalPage(_webDriver, _testContext);
        }

        public ParametersGlobalSettingsModalPage GetNewVersionItemDetailsPage()
        {
            Actions action = new Actions(_webDriver);

            action.MoveToElement(_newVersionItemDetailsLine).Perform();
            _newVersionItemDetailsLine.Click();

            _newVersionItemDetailsEdit = WaitForElementIsVisible(By.XPath(NEW_VERSION_ITEM_DETAILS_EDIT));
            _newVersionItemDetailsEdit.Click();
            WaitForLoad();

            return new ParametersGlobalSettingsModalPage(_webDriver, _testContext);
        }

        public ParametersGlobalSettingsModalPage GetNewSupplierInvoiceVersionPage()
        {
            Actions action = new Actions(_webDriver);

            action.MoveToElement(_newSupplierInvoiceLine).Perform();
            _newSupplierInvoiceLine.Click();

            _newSupplierInvoiceEdit = WaitForElementIsVisible(By.XPath(NEW_SUPPLIER_INVOICE_EDIT));
            _newSupplierInvoiceEdit.Click();
            WaitForLoad();

            return new ParametersGlobalSettingsModalPage(_webDriver, _testContext);
        }

        public ParametersGlobalSettingsModalPage GetIsCompareActiveGlobalPage()
        {
            Actions action = new Actions(_webDriver);

            action.MoveToElement(_isCompareActiveLine).Perform();
            _isCompareActiveLine.Click();

            _isCompareActiveEdit = WaitForElementIsVisible(By.XPath(IS_COMPARE_ACTIVE_EDIT));
            _isCompareActiveEdit.Click();
            WaitForLoad();

            return new ParametersGlobalSettingsModalPage(_webDriver, _testContext);
        }

        public ParametersGlobalSettingsModalPage GetNumberOfItemsToComparePage()
        {
            Actions action = new Actions(_webDriver);

            action.MoveToElement(_numberOfItemsLine).Perform();
            _numberOfItemsLine.Click();

            _numberOfItemsEdit = WaitForElementIsVisible(By.XPath(NUMBER_ITEMS_COMPARE_EDIT));
            _numberOfItemsEdit.Click();
            WaitForLoad();

            return new ParametersGlobalSettingsModalPage(_webDriver, _testContext);
        }

        public ParametersGlobalSettingsModalPage GetRecipeExpiryDatePage()
        {
            Actions action = new Actions(_webDriver);

            action.MoveToElement(_recipeExpiryDateLine).Perform();
            _recipeExpiryDateLine.Click();

            _recipeExpiryDateEdit = WaitForElementIsVisible(By.XPath(RECIPE_EXPIRY_DATE_EDIT));
            _recipeExpiryDateEdit.Click();
            WaitForLoad();

            return new ParametersGlobalSettingsModalPage(_webDriver, _testContext);
        }

        public ParametersGlobalSettingsModalPage GetBackDatingDatePage()
        {
            Actions action = new Actions(_webDriver);

            action.MoveToElement(_backDatingLine).Perform();
            _backDatingLine.Click();

            _backDatingEdit = WaitForElementIsVisible(By.XPath(BACKDATING_EDIT));
            _backDatingEdit.Click();
            WaitForLoad();

            return new ParametersGlobalSettingsModalPage(_webDriver, _testContext);
        }
        
          public ParametersGlobalSettingsModalPage ClickOnTranslationTab()
          {
            var tab = WaitForElementIsVisible(By.XPath(TRANSLATIONTAB));
            tab.Click();
            WaitForLoad();
            return new ParametersGlobalSettingsModalPage(_webDriver, _testContext);

            
          }


        public enum FilterEdi
        {
           culture, 
           module, 
           search
        }
        public void Filter(FilterEdi FilterEdi, object value)
        {
            switch (FilterEdi)
            {
                case FilterEdi.culture:
                    _culture = WaitForElementIsVisible(By.Id(CULTURE));
                    _culture.SetValue(ControlType.TextBox, value);
                    break;
                case FilterEdi.module:
                    _module = WaitForElementIsVisible(By.Id(MODULE));
                    _module.SetValue(ControlType.DropDownList, value);
                    break;
                case FilterEdi.search:
                    _search = WaitForElementIsVisible(By.Id(SEARCH));
                    _search.SetValue(ControlType.DropDownList, value);
                    break;
            }
            WaitForLoad();
        }


        public bool IsListeAffiche()
        {
            var status = _webDriver.FindElements(By.XPath(TRANSLATIONLIST));

            if (status.Count == 0)
                return false;

            return status.Any(c => !string.IsNullOrEmpty(c.Text));
        }

        // Version print
        //public ParametersGlobalSettingsModalPage GetNewPrintGlobalPage()
        //{
        //    //Actions action = new Actions(_webDriver);

        //    //action.MoveToElement(_newVersionPrintLine).Perform();
        //    //_newVersionPrintLine.Click();

        //    //_newVersionPrintEdit = WaitForElementIsVisible(By.XPath(NEW_VERSION_PRINT_EDIT));
        //    //_newVersionPrintEdit.Click();
        //    //WaitForLoad();

        //    return new ParametersGlobalSettingsModalPage(_webDriver, _testContext);
        //}

        /*public ParametersGlobalSettingsModalPage GetShowInterimDetailsPage()
        {

            Actions action = new Actions(_webDriver);
            _showInterimLine = WaitForElementExists(By.XPath(SHOW_INTERIM_LINE));
            action.MoveToElement(_showInterimLine).Click().Perform();

            _showInterimEdit = WaitForElementIsVisible(By.XPath(SHOW_INTERIM_EDIT));
            _showInterimEdit.Click();
            WaitForLoad();

            return new ParametersGlobalSettingsModalPage(_webDriver, _testContext);
        }*/

        public bool DetectClosingPeriod(string sectionToCheck)
        {
            var sectionElements = _webDriver.FindElements(By.XPath(SECTIONLIST));

            bool sectionExist = sectionElements.Select(c=>c.Text.ToLower().Trim()).Contains(sectionToCheck.ToLower().Trim());
            
            return sectionExist;
            
           
        }

    }
}
