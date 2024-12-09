using System;
using System.Security.Cryptography.Xml;
using System.Threading;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium.Support.UI;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Purchasing
{
    public class ItemCreateModalPage : PageBase
    {
        public ItemCreateModalPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        // ____________________________________ Constantes ______________________________________________

        private const string ITEM_NAME = "Name";
        private const string WEIGHT = "ItemUnitProductionWeight";
        private const string ITEM_COMMERCIALNAME1 = "CommercialName1";
        private const string ITEM_COMMERCIALNAME2 = "CommercialName2";
        private const string GROUP_NAME = "ddl_group";
        private const string SUBGROUP_NAME = "ddl_subgroup";
        private const string WORKSHOP = "WorkshopId";
        private const string TAX_TYPE = "TaxTypeId";
        private const string PROD_UNIT = "ItemUnitProductionTypeId";
        private const string SAVE = "//*[@id=\"item-filter-form\"]/div/div[2]/div/button[2]";
        private const string ITEM_REFERENCE = "Reference";
        private const string IS_SANITIZATION_CHECKBOX = "IsSanitization";
        private const string IS_THAWING_CHECKBOX = "IsThawing"; 
        private const string CheckBox_NoHACCPRecord = "IsHACCPRecordExcluded"; 
        private const string ERROR_MESSAGE_NAME = "//*[@id=\"item-filter-form\"]/div/div[1]/div[1]/div[1]/div[1]/div/span[1]";
        private const string ERROR_MESSAGE_DOUBLE = "//*[@id=\"item-filter-form\"]/div/div[1]/div[1]/div[1]/div[1]/div/span[2]";
        private const string ERROR_MESSAGE_DOUBLE_NAME = "//*[@id=\"item-filter-form\"]/div[1]/div[1]/div[1]/div[1]/div/span[2]";
        private const string WEIGHTONGENERALINFORMATION = "//*[@id=\"ItemUnitProductionWeight\"]";
        private const string CHECKBOX_ISVEGETABLE = "IsVegetable";
        private const string CHECKBOX_IsCSR = "IsCsr";



        // ____________________________________ Variables _______________________________________________

        [FindsBy(How = How.Id, Using = ITEM_NAME)]
        private IWebElement _itemName;
        [FindsBy(How = How.XPath, Using = WEIGHTONGENERALINFORMATION)]
        private IWebElement _weightongeneralinformation;
        [FindsBy(How = How.Id, Using = WEIGHT)]
        private IWebElement _weight;
        [FindsBy(How = How.Id, Using = ITEM_COMMERCIALNAME1)]
        private IWebElement _commercialName1;

        [FindsBy(How = How.Id, Using = ITEM_COMMERCIALNAME2)]
        private IWebElement _commercialName2;

        [FindsBy(How = How.Id, Using = ITEM_REFERENCE)]
        private IWebElement _reference;

        [FindsBy(How = How.Id, Using = GROUP_NAME)]
        private IWebElement _groupName;

        [FindsBy(How = How.Id, Using = SUBGROUP_NAME)]
        private IWebElement _subGroupName;

        [FindsBy(How = How.Id, Using = WORKSHOP)]
        private IWebElement _workshop;

        [FindsBy(How = How.Id, Using = TAX_TYPE)]
        private IWebElement _taxType;

        [FindsBy(How = How.Id, Using = PROD_UNIT)]
        private IWebElement _prodUnit; 

        [FindsBy(How = How.Id, Using = IS_SANITIZATION_CHECKBOX)]
        private IWebElement _checkboxIsSantization; 

        [FindsBy(How = How.Id, Using = IS_THAWING_CHECKBOX)]
        private IWebElement _checkboxIsThawing;

        [FindsBy(How = How.Id, Using = CheckBox_NoHACCPRecord)]
        private IWebElement _checkboxNoHACCPRecord; 

        [FindsBy(How = How.Id, Using = CHECKBOX_ISVEGETABLE)]
        private IWebElement _checkboxIsVegetable;

        [FindsBy(How = How.Id, Using = CHECKBOX_IsCSR)]
        private IWebElement _checkboxIsCSR;

        [FindsBy(How = How.XPath, Using = SAVE)]
        private IWebElement _save;

        [FindsBy(How = How.XPath, Using = ERROR_MESSAGE_NAME)]
        private IWebElement _errorMessageName;

        [FindsBy(How = How.XPath, Using = ERROR_MESSAGE_DOUBLE)]
        private IWebElement _errorMessageDouble;

        [FindsBy(How = How.XPath, Using = ERROR_MESSAGE_DOUBLE_NAME)]
        private IWebElement _errorMessageDoubleName;

        // ___________________________________________ Méthodes ___________________________________________

        public ItemGeneralInformationPage FillField_CreateNewItem(string name, string group, string workshop, string taxType, string prodUnit, string weight = null, string subgroup = null, string commercialName = null, string commercialName2 = null, string reference = null, bool isSanitization = false, bool IsThawing = false, bool NoHACCPRecord = false, bool IsVegetable = false, bool IsCSR = false)
        {
            // Définition du nom
            _itemName = WaitForElementIsVisible(By.Id(ITEM_NAME));
            _itemName.SetValue(ControlType.TextBox, name);

            //Define CommercialName1
            if (!string.IsNullOrEmpty(commercialName))
            {
                _commercialName1 = WaitForElementIsVisible(By.Id(ITEM_COMMERCIALNAME1));
                _commercialName1.SetValue(ControlType.TextBox, commercialName);
            }

            //Define CommercialName2
            if (!string.IsNullOrEmpty(commercialName2))
            {
                _commercialName2 = WaitForElementIsVisible(By.Id(ITEM_COMMERCIALNAME2));
                _commercialName2.SetValue(ControlType.TextBox, commercialName2);
            }

            // Définition du groupe
            _groupName = WaitForElementIsVisible(By.Id(GROUP_NAME));
            _groupName.SetValue(ControlType.DropDownList, group);

            // Attachement du SubGroup
            if (subgroup != null)
            {
                //cascade
                Thread.Sleep(2000);
                _subGroupName = WaitForElementIsVisible(By.Id(SUBGROUP_NAME));
                _subGroupName.SetValue(ControlType.DropDownList, subgroup);
            }

            // Définition du workshop
            _workshop = WaitForElementIsVisible(By.Id(WORKSHOP));
            _workshop.SetValue(ControlType.DropDownList, workshop);

            // Définition du reference
            if (!string.IsNullOrEmpty(reference))
            {
                _reference = WaitForElementIsVisible(By.Id(ITEM_REFERENCE));
                _reference.SetValue(ControlType.TextBox, reference);
            }

            // Définition de type de taxe
            _taxType = WaitForElementIsVisible(By.Id(TAX_TYPE));
            _taxType.SetValue(ControlType.DropDownList, taxType);

            // Définition de Prod Unit
            _prodUnit = WaitForElementIsVisible(By.Id(PROD_UNIT));
            _prodUnit.SetValue(ControlType.DropDownList, prodUnit);

            // Définition du weight
            if (weight != null)
            {
                _weight = WaitForElementIsVisible(By.Id(WEIGHT));
                _weight.SetValue(ControlType.TextBox, weight);
            }
            // Définition du Sanitization
            if (isSanitization)
            {
                _checkboxIsSantization = WaitForElementExists(By.Id(IS_SANITIZATION_CHECKBOX));
                _checkboxIsSantization.SetValue(PageBase.ControlType.CheckBox, isSanitization);
                WaitForLoad();
            }
            // Définition du IsThawing
            if (IsThawing)
            {
                _checkboxIsThawing = WaitForElementExists(By.Id(IS_THAWING_CHECKBOX));
                _checkboxIsThawing.SetValue(PageBase.ControlType.CheckBox, IsThawing);
                WaitForLoad();
            }
            // Définition du NoHACCPRecord
            if (NoHACCPRecord)
            {
                _checkboxNoHACCPRecord = WaitForElementExists(By.Id(CheckBox_NoHACCPRecord));
                _checkboxNoHACCPRecord.SetValue(PageBase.ControlType.CheckBox, NoHACCPRecord);
                WaitForLoad();
            }
            
            // Définition du IsVegetable
            if (IsVegetable)
            {
                _checkboxIsVegetable = WaitForElementExists(By.Id(CHECKBOX_ISVEGETABLE));
                _checkboxIsVegetable.SetValue(PageBase.ControlType.CheckBox, IsVegetable);
                WaitForLoad();
            }

            // Définition du IsCSR
            if (IsCSR)
            {
                _checkboxIsCSR = WaitForElementExists(By.Id(CHECKBOX_IsCSR));
                _checkboxIsCSR.SetValue(PageBase.ControlType.CheckBox, IsCSR);
                WaitForLoad();
            }
            // Click sur le bouton "Create"
            Save();

            return new ItemGeneralInformationPage(_webDriver, _testContext);
        }

        public void Save()
        {
            _save = WaitForElementToBeClickable(By.XPath("//*[@id='item-filter-form']/div[2]/div/button[2]"));
            _save.Click();
            WaitForLoad();
        }

        public bool ErrorMessageNameRequired()
        {
                
            _errorMessageName = WaitForElementIsVisible(By.XPath("//*[@id=\"item-filter-form\"]/div[1]/div[1]/div[1]/div[1]/div/span[1]"));

            string expectedMessage = "The Name field is required.";

            if (_errorMessageName.Text == expectedMessage)
            {
                return true;
            }

            return false;
        }

        public bool ErrorMessageNameAlreadyExists()
        {
            _errorMessageDouble = WaitForElementIsVisible(By.XPath("//*[@id=\"item-filter-form\"]/div[1]/div[1]/div[1]/div[1]/div/span[2]"));

            string expectedMessage = "An item with the same name already exists in database. Please chooose another name.";

            if (_errorMessageDouble.Text == expectedMessage)
            {
                return true;
            }

            return false;
        }

        public ItemGeneralInformationPage FillField_Name(string name)

        {
            // Définition du nom
            _itemName = WaitForElementIsVisibleNew(By.Id(ITEM_NAME));
            _itemName.SetValue(ControlType.TextBox, name);


            Save();

            return new ItemGeneralInformationPage(_webDriver, _testContext);
        }
        public bool ErrorMessageNameDouble()
        {

            _errorMessageDoubleName = WaitForElementIsVisibleNew(By.XPath(ERROR_MESSAGE_DOUBLE_NAME));

            string expectedMessage = "An item with the same name already exists in database. Please chooose another name.";

            if (_errorMessageDoubleName.Text == expectedMessage)
            {
                return true;
            }

            return false;
        }
        public string GetWeightInGram()
        {
            _weightongeneralinformation = WaitForElementIsVisible(By.XPath(WEIGHTONGENERALINFORMATION));
             return _weightongeneralinformation.GetAttribute("value");

        }
    }

}
