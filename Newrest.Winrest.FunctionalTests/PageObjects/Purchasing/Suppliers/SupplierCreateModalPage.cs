using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Purchasing.Suppliers;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium.Support.UI;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Purchasing
{
    public class SupplierCreateModalPage : PageBase
    {    
        public SupplierCreateModalPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {

        }

        //__________________________________Constantes_______________________________________

        private const string NAME_TEXT_FIELD = "Name";
        private const string ACTIVATED_CHECKBOX = "IsActive";
        private const string AUDITED_CHECKBOX = "IsAudited";
        private const string CREATE_BTN = "//*[@id=\"modal-1\"]/div[2]/div/form/div[2]/button[2]";
        private const string SUPPLIER_NUMBER = "NextSupplierId";

        //__________________________________Variables________________________________________

        [FindsBy(How = How.Id, Using = NAME_TEXT_FIELD)]
        private IWebElement _nameTextField;

        [FindsBy(How = How.Id, Using = ACTIVATED_CHECKBOX)]
        private IWebElement _activatedCheckbox;

        [FindsBy(How = How.Id, Using = AUDITED_CHECKBOX)]
        private IWebElement _auditedCheckbox;

        [FindsBy(How = How.XPath, Using = CREATE_BTN)]
        private IWebElement _createBtn;

        [FindsBy(How = How.Id, Using = SUPPLIER_NUMBER)]
        private IWebElement _supplierNumber;

        //__________________________________Utilitaire______________________________________

        public SupplierGeneralInfoTab FillField_CreatNewSupplier(string name, bool isActivated = true, bool isAudited = false)
        {           
            _nameTextField = WaitForElementIsVisibleNew(By.Id(NAME_TEXT_FIELD));
            _nameTextField.SetValue(ControlType.TextBox, name);

            _activatedCheckbox = WaitForElementExists(By.Id(ACTIVATED_CHECKBOX));
            _activatedCheckbox.SetValue(ControlType.CheckBox, isActivated);

            _auditedCheckbox = WaitForElementExists(By.Id(AUDITED_CHECKBOX));
            _auditedCheckbox.SetValue(ControlType.CheckBox, isAudited);
            return new SupplierGeneralInfoTab(_webDriver, _testContext);
        }


        // __________________________________________ METHODES ___________________________________________

        public SupplierItem Submit()
        {
                _createBtn = WaitForElementIsVisibleNew(By.XPath("//*/button[@value='Create']"));
            
            _createBtn.Click();
            WaitPageLoading();

            return new SupplierItem(_webDriver, _testContext);
        }

        public SupplierGeneralInfoTab SubmitToGeneralInfo()
        {
            _createBtn = WaitForElementIsVisible(By.XPath("//*/button[@value='Create']"));

            _createBtn.Click();
            WaitPageLoading();

            return new SupplierGeneralInfoTab(_webDriver, _testContext);
        }

        public string CollectSupplierNumber()
        {
            _supplierNumber = WaitForElementExists(By.Id(SUPPLIER_NUMBER));
            return _supplierNumber.Text;
        }
    }
}
