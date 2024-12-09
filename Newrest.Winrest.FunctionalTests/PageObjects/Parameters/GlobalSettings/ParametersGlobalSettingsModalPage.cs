using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Parameters.GlobalSettings
{
    public class ParametersGlobalSettingsModalPage : PageBase
    {

        public ParametersGlobalSettingsModalPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        //____________________________________________________Constantes_____________________________________________________________

        private const string VALUE_DECIMAL_SEPARATOR = "SelectedSpecificValue";
        private const string VALUE_DATE_FORMAT_PICKER = "Value";

        // sage auto enabled
        private const string SAGE_AUTO_FILTER = "SelectedItems_ms";
        private const string SAGE_AUTO_CHECK_ALL = "/html/body/div[11]/div/ul/li[1]/a";
        private const string SAGE_AUTO_UNCHECK_ALL = "/html/body/div[11]/div/ul/li[2]/a";
        private const string SAGE_AUTO_INPUT = "/html/body/div[11]/div/div/label/input";


        // New search mode
        private const string NEW_SEARCH_MODE_FILTER = "SelectedItems_ms";
        private const string NEW_SEARCH_MODE_CHECK_ALL = "/html/body/div[11]/div/ul/li[1]/a";
        private const string NEW_SEARCH_MODE_UNCHECK_ALL = "/html/body/div[11]/div/ul/li[2]/a";

        private const string NUMBER_ITEMS_COMPARE = "Value";
        private const string RECIPE_EXPIRY_DATE_VALUE = "Value";
        private const string BACK_DATING = "SelectedSpecificValue";

        private const string CHECK_BOX = "BooleanValue";
        private const string SAVE = "last";

        //____________________________________________________Variables_______________________________________________________________

        // DateFormatPicker value
        [FindsBy(How = How.Id, Using = VALUE_DECIMAL_SEPARATOR)]
        private IWebElement _dateFormatPickerValue;

        // Decimal value
        [FindsBy(How = How.Id, Using = VALUE_DECIMAL_SEPARATOR)]
        private IWebElement _valueDecimalSeparator;

        // Winrest 4.0
        [FindsBy(How = How.Id, Using = CHECK_BOX)]
        private IWebElement _valueShowInterim;

        [FindsBy(How = How.Id, Using = CHECK_BOX)]
        private IWebElement _valueKeywordSearch;

        // Winrest 4.0
        [FindsBy(How = How.Id, Using = CHECK_BOX)]
        private IWebElement _valueClaimV3;



        // sage auto enabled
        [FindsBy(How = How.Id, Using = SAGE_AUTO_FILTER)]
        private IWebElement _sageAutoEnabledFilter;

        [FindsBy(How = How.XPath, Using = SAGE_AUTO_CHECK_ALL)]
        private IWebElement _sageAutoEnabledCheckAll;

        [FindsBy(How = How.XPath, Using = SAGE_AUTO_UNCHECK_ALL)]
        private IWebElement _sageAutoEnabledUncheckAll;

        [FindsBy(How = How.XPath, Using = SAGE_AUTO_INPUT)]
        private IWebElement _sageAutoEnabledInput;

        // new search mode
        [FindsBy(How = How.Id, Using = NEW_SEARCH_MODE_FILTER)]
        private IWebElement _newSearchModeFilter;

        [FindsBy(How = How.XPath, Using = NEW_SEARCH_MODE_CHECK_ALL)]
        private IWebElement _newSearchModeCheckAll;

        [FindsBy(How = How.XPath, Using = NEW_SEARCH_MODE_UNCHECK_ALL)]
        private IWebElement _newSearchModeUncheckAll;

        [FindsBy(How = How.Id, Using = CHECK_BOX)]
        private IWebElement _valueClaimType;

        [FindsBy(How = How.Id, Using = CHECK_BOX)]
        private IWebElement _valueGroupDisplay;

        [FindsBy(How = How.Id, Using = CHECK_BOX)]
        private IWebElement _valueSubGroup;

        [FindsBy(How = How.Id, Using = CHECK_BOX)]
        private IWebElement _valueVersionItemDetails;

        [FindsBy(How = How.Id, Using = CHECK_BOX)]
        private IWebElement _valueNewSupplierInvoice;

        [FindsBy(How = How.Id, Using = CHECK_BOX)]
        private IWebElement _valueIsCompareActive;

        [FindsBy(How = How.Id, Using = NUMBER_ITEMS_COMPARE)]
        private IWebElement _valueNumberOfItems;
        
        [FindsBy(How = How.Id, Using = RECIPE_EXPIRY_DATE_VALUE)]
        private IWebElement _recipeExpiryDateValue;

        [FindsBy(How = How.Id, Using = NUMBER_ITEMS_COMPARE)]
        private IWebElement _valueBackDating;

        // Commun
        [FindsBy(How = How.Id, Using = SAVE)]
        private IWebElement _saveBtn;


        //____________________________________________________ Méthodes ____________________________________________________________________

        // DateFormatPicker
        public string GetDateFormatPickerValue()
        {
            _dateFormatPickerValue = WaitForElementExists(By.Id(VALUE_DATE_FORMAT_PICKER));

            return _dateFormatPickerValue.GetAttribute("value");
        }

        // Decimal separator
        public string GetDecimalValue()
        {
            _valueDecimalSeparator = WaitForElementExists(By.Id(VALUE_DECIMAL_SEPARATOR));

            if (_valueDecimalSeparator.GetAttribute("value") == ",")
                return ",";
            else
                return ".";
        }

        // Winrest 4.0
        public void SetShowInterimAreaValue(bool value = true)
        {
            WaitForElementToBeClickable(By.Id(SAVE));

            _valueShowInterim = WaitForElementExists(By.Id(CHECK_BOX));
            _valueShowInterim.SetValue(ControlType.CheckBox, value);
        }

        // Warehouse Claim V3
        public void SetClaimV3AreaValue(bool value = true)
        {
            WaitForElementToBeClickable(By.Id(SAVE));

            _valueClaimV3 = WaitForElementExists(By.Id(CHECK_BOX));
            _valueClaimV3.SetValue(ControlType.CheckBox, value);
        }

        public void SetSageAutoEnabledValue(bool value = true, string area = null)
        {
            if (value && area == null)
            {
                // ou ComboBoxSelectById(SAGE_AUTO_FILTER, area, false,true); ?
                _sageAutoEnabledFilter = WaitForElementIsVisible(By.Id(SAGE_AUTO_FILTER));
                _sageAutoEnabledFilter.Click();
                WaitForLoad();
                
                _sageAutoEnabledCheckAll = WaitForElementIsVisible(By.XPath(SAGE_AUTO_CHECK_ALL));
                _sageAutoEnabledCheckAll.Click();
                WaitForLoad();

                _sageAutoEnabledFilter = WaitForElementIsVisible(By.Id(SAGE_AUTO_FILTER));
                _sageAutoEnabledFilter.Click();
                WaitForLoad();
            }
            else if (value && area != "1")
            {
                ComboBoxSelectById(new ComboBoxOptions(SAGE_AUTO_FILTER, area, false));
            }
            else if (!value && area == null)
            {
                _sageAutoEnabledFilter = WaitForElementIsVisible(By.Id(SAGE_AUTO_FILTER));
                _sageAutoEnabledFilter.Click();
                WaitForLoad();

                _sageAutoEnabledUncheckAll = WaitForElementIsVisible(By.XPath(SAGE_AUTO_UNCHECK_ALL));
                _sageAutoEnabledUncheckAll.Click();
                WaitForLoad();

                _sageAutoEnabledFilter = WaitForElementIsVisible(By.Id(SAGE_AUTO_FILTER));
                _sageAutoEnabledFilter.Click();
                WaitForLoad();
            }
        }

        public void SetKeywordVersionValue(bool value = true)
        {
            WaitForElementToBeClickable(By.Id(SAVE));
            _valueKeywordSearch = WaitForElementExists(By.Id(CHECK_BOX));
            _valueKeywordSearch.SetValue(ControlType.CheckBox, value);
        }

        public void SetNewSearchModeValue(bool value = true)
        {
            WaitForElementToBeClickable(By.Id(SAVE));
            _newSearchModeFilter = WaitForElementIsVisible(By.Id(NEW_SEARCH_MODE_FILTER));
            _newSearchModeFilter.Click();

            if (value)
            {
                _newSearchModeCheckAll = WaitForElementIsVisible(By.XPath(NEW_SEARCH_MODE_CHECK_ALL));
                _newSearchModeCheckAll.Click();
            }
            else
            {
                _newSearchModeUncheckAll = WaitForElementIsVisible(By.XPath(NEW_SEARCH_MODE_UNCHECK_ALL));
                _newSearchModeUncheckAll.Click();
            }

            _newSearchModeFilter.Click();
        }

        public void SetClaimTypeValue(bool value = true)
        {
            WaitForElementToBeClickable(By.Id(SAVE));
            _valueClaimType = WaitForElementExists(By.Id(CHECK_BOX));
            _valueClaimType.SetValue(ControlType.CheckBox, value);
        }

        public void SetGroupDisplayValue(bool value = true)
        {
            WaitForElementToBeClickable(By.Id(SAVE));
            _valueGroupDisplay = WaitForElementExists(By.Id(CHECK_BOX));
            _valueGroupDisplay.SetValue(ControlType.CheckBox, value);
        }

        public void SetSubGroupValue(bool value = true)
        {
            WaitForElementToBeClickable(By.Id(SAVE));
            _valueSubGroup = WaitForElementExists(By.Id(CHECK_BOX));
            _valueSubGroup.SetValue(ControlType.CheckBox, value);
        }

        public bool GetSubGroupValue()
        {
            _valueSubGroup = WaitForElementExists(By.Id(CHECK_BOX));

            if (_valueSubGroup.GetAttribute("checked") == "true")
                return true;
            else
                return false;
        }

        public void SetNewVersionItemDetailValue(bool value = true)
        {
            WaitForElementToBeClickable(By.Id(SAVE));
            _valueVersionItemDetails = WaitForElementExists(By.Id(CHECK_BOX));
            _valueVersionItemDetails.SetValue(ControlType.CheckBox, value);
        }

        public void SetNewSupplierInvoiceValue(bool value = true)
        {
            WaitForElementToBeClickable(By.Id(SAVE));
            _valueNewSupplierInvoice = WaitForElementExists(By.Id(CHECK_BOX));
            _valueNewSupplierInvoice.SetValue(ControlType.CheckBox, value);
        }

        public void SetCompareActiveValue(bool value = true)
        {
            WaitForElementToBeClickable(By.Id(SAVE));
            _valueIsCompareActive = WaitForElementExists(By.Id(CHECK_BOX));
            _valueIsCompareActive.SetValue(ControlType.CheckBox, value);
        }

        public void SetNumberItemCompareValue(string value)
        {
            WaitForElementToBeClickable(By.Id(SAVE));
            _valueNumberOfItems = WaitForElementIsVisible(By.Id(NUMBER_ITEMS_COMPARE));
            _valueNumberOfItems.SetValue(ControlType.TextBox, value);
        }

        public string GetNumberItemCompareValue()
        {
            WaitForElementToBeClickable(By.Id(SAVE));
            _valueNumberOfItems = WaitForElementIsVisible(By.Id(NUMBER_ITEMS_COMPARE));
            return _valueNumberOfItems.GetAttribute("value");
        }

        public void SetRecipeExpiryDateValue(string value)
        {
            WaitForElementToBeClickable(By.Id(SAVE));
            _recipeExpiryDateValue = WaitForElementIsVisible(By.Id(RECIPE_EXPIRY_DATE_VALUE));
            _recipeExpiryDateValue.SetValue(ControlType.TextBox, value);
        }

        public void SetBackDatingValue(string value)
        {
            WaitForElementToBeClickable(By.Id(SAVE));
            _valueBackDating = WaitForElementIsVisible(By.Id(BACK_DATING));
            _valueBackDating.SetValue(ControlType.DropDownList, value);
        }

        public string GetBackDatingValue()
        {
            WaitForElementToBeClickable(By.Id(SAVE));
            _valueBackDating = WaitForElementIsVisible(By.Id(BACK_DATING));
            return _valueBackDating.GetAttribute("value");
        }

        // Commun
        public ParametersGlobalSettings Save()
        {
            _saveBtn = WaitForElementIsVisible(By.Id(SAVE));
            _saveBtn.Click();
            WaitForLoad();
            return new ParametersGlobalSettings(_webDriver, _testContext);
        }

        // Version Print
        //public void SetPrintValue(bool value = true)
        //{
        //    WaitForElementToBeClickable(By.Id(SAVE));

        //    _valueVersionPrint = WaitForElementExists(By.Id(CHECK_BOX));
        //    _valueVersionPrint.SetValue(ControlType.CheckBox, value);
        //}

    }


}
