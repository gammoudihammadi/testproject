using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Parameters.Customer;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.PageObjects;
using System;

namespace Newrest.Winrest.FunctionalTests.PageObject.Parameters.Customer
{
    public class ParametersCustomer : PageBase
    {

        private const string CATEGORY = "//*[@id=\"paramCustomerTab\"]/li[2]/a";
        private const string ORDER_TYPE_TAB = "//*[@id=\"paramCustomerTab\"]/li[3]/a";

        private const string ADD_SERVICE = "//*[@id=\"tabContentParameters\"]/div[1]/a";

        private const string CUSTOMER_TYPE = "//td[contains(text(), '{0}')]";
        private const string ADD_NEW = "/html/body/div[2]/div/div/div/div[1]/a";
        private const string FAMILLE = "FLIGHT";

        private const string SERVICE_NAME = "//tr[{0}]/td[1]";
        private const string SERVICE_EDIT = "//tr[{0}]//span[contains(@class, 'pencil')]";
        private const string SERVICE_EDIT_PATCH = "//td[contains(text(), 'HOT MEAL YC UAE')]/..//span[@class='glyphicon glyphicon-pencil']";

        //category
        private const string GROUP_FAMILIES_NAME = "//*[@id=\"tabContentParameters\"]/table/tbody/tr/td[contains(text(), '{0}')]";
        private const string ISPRODUCT_CATEGORY = "//*[@id=\"tabContentParameters\"]/div[2]/div/table/tbody/tr[*]/td[1][contains(text(),'{0}')]/../td[7]/div/input[1]";
        private const string ISFOOD_CATEGORY = "//*[@id=\"tabContentParameters\"]/div[2]/div/table/tbody/tr[*]/td[1][contains(text(),'{0}')]/../td[2]/div/input[1]";
        private const string EDIT_CATEGORY = "//*[@id=\"tabContentParameters\"]/div[2]/div/table/tbody/tr[*]/td[1][contains(text(),'{0}')]/../td[11]/a[contains(@href, 'Edit')]";
        private const string ADD_NEW_CATEGORY = "//*[@id=\"tabContentParameters\"]/div[1]/a";
        private const string CATEGORY_NEW_NAME = "first";

        //create order type
        private const string ADD_ORDER_TYPE = "//*[@id=\"tabContentParameters\"]/div[1]/a";
        private const string ORDER_TYPE_NAME = "//*[@id=\"first\"]";
        private const string ORDER_TYPE_SAVE_BUTTON = "//*[@id=\"last\"]";

        //create type of customer modaltext
        private const string CUST_NEW_TYPE_NAME = "first";
        private const string CUST_NEW_TYPE_SAVE_BUTTON = "last";

        //service category families tab
        private const string SERVICE_CATEGORY_FAMILIES_TAB = "//a[text()='Service category families']";
        private const string CATEGORY_NAME = "//td[contains(text(), '{0}')]";
        private const string CHECK_ISPRODUCT = "IsProduct";
        private const string CHECK_ISFOOD = "IsFood";
        private const string CATEGORY_SAVE_BUTTON = "last";


        [FindsBy(How = How.XPath, Using = CATEGORY)]
        private IWebElement _category_tab;

        [FindsBy(How = How.XPath, Using = ADD_NEW_CATEGORY)]
        private IWebElement _addNewCategory;


        [FindsBy(How = How.XPath, Using = ORDER_TYPE_TAB)]
        private IWebElement _orderType_tab;

        [FindsBy(How = How.XPath, Using = SERVICE_CATEGORY_FAMILIES_TAB)]
        private IWebElement _servCatFam_tab;

        [FindsBy(How = How.XPath, Using = ADD_NEW)]
        private IWebElement _addNew;

        [FindsBy(How = How.Id, Using = CUST_NEW_TYPE_NAME)]
        private IWebElement _inputTypeCustomer;

        [FindsBy(How = How.Id, Using = CUST_NEW_TYPE_SAVE_BUTTON)]
        private IWebElement _saveNewBtn;

        [FindsBy(How = How.XPath, Using = ADD_SERVICE)]
        private IWebElement _addService;

        [FindsBy(How = How.XPath, Using = ADD_ORDER_TYPE)]
        private IWebElement _addOrderType;

        [FindsBy(How = How.XPath, Using = ORDER_TYPE_NAME)]
        private IWebElement _inputOrderType;

        [FindsBy(How = How.Id, Using = CATEGORY_NEW_NAME)]
        private IWebElement _CategoryName;
      
        [FindsBy(How = How.Id, Using = CHECK_ISPRODUCT)]
        private IWebElement _isFood;
        [FindsBy(How = How.Id, Using = CATEGORY_SAVE_BUTTON)]
        private IWebElement _saveButton;

        public ParametersCustomer(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {

        }

        public void GoToTab_Category()
        {
            _category_tab = WaitForElementIsVisible(By.XPath(CATEGORY));
            _category_tab.Click();
            WaitForLoad();
        }

        public void GoToTab_ServiceCategoryFamilies()
        {
            _servCatFam_tab = WaitForElementIsVisible(By.XPath(SERVICE_CATEGORY_FAMILIES_TAB));
            _servCatFam_tab.Click();
            WaitForLoad();
        }

        public bool isCategoryExist(string serviceCategory)
        {
            bool isExists;
            if (isElementVisible(By.XPath(string.Format(CATEGORY_NAME, serviceCategory))))
            {
                isExists = true;
            }
            else
            {
                isExists = false;
            }
            return isExists;
        }

        public void SetServiceWithAirportTax(string name)
        {
            bool isServiceFound = false;

            // On regarde s'il y a des sites déjà définis
            var nbLignes = _webDriver.FindElements(By.TagName("tr")).Count;

            if (nbLignes > 1)
            {
                for (int i = 1; i < nbLignes; i++)
                {
                    var xpath = String.Format(SERVICE_NAME, i + 1);
                    var serviceName = WaitForElementIsVisible(By.XPath(xpath));
                    //scroll to each service
                    Actions a = new Actions(_webDriver);
                    a.MoveToElement(serviceName).Perform();
                    if (serviceName.Text.Equals(name))
                    {
                        // On ouvre la page d'édition du site
                        var editServicePage = EditService(i);
                        editServicePage.SetAirportTax();
                        editServicePage.Save();
                        
                        isServiceFound = true;
                        break;
                    }
                }
            }

            if (!isServiceFound)
            {
                isServiceFound = AddService(name);
            }

            Assert.IsTrue(isServiceFound, "La taxe Airport Tax n'a pas pu être assignée au service " + name + ".");
        }

        public ParametersCustomerCreateServiceModalPage EditService(int i)
        {
                var edit = String.Format(SERVICE_EDIT, i + 1);
                var editService = WaitForElementIsVisible(By.XPath(edit));
                editService.Click();
            
            WaitForLoad();

            return new ParametersCustomerCreateServiceModalPage(_webDriver, _testContext);
        }

        public ParametersCustomerCreateServiceModalPage AddNewService()
        {
            _addService = WaitForElementIsVisible(By.XPath(ADD_SERVICE));
            _addService.Click();
            WaitForLoad();

            return new ParametersCustomerCreateServiceModalPage(_webDriver, _testContext);
        }

        public bool AddService(string serviceName)
        {
            bool isServiceAdded = true;

            try
            {
                // Click sur le bouton New
                var createNewServicePage = AddNewService();

                // Ajout du site, du customer et de la value
                createNewServicePage.SetServiceName(serviceName);
                createNewServicePage.SetAirportTax();
                createNewServicePage.SetFamily(FAMILLE);
                createNewServicePage.AddNew();
            }
            catch
            {
                isServiceAdded = false;
            }

            return isServiceAdded;
        }

        public bool isTypeOfCustomerExist(string customerType)
        {
            bool isExists;
            if(isElementVisible(By.XPath(string.Format(CUSTOMER_TYPE, customerType))))
            {
                WaitForElementIsVisible(By.XPath(string.Format(CUSTOMER_TYPE, customerType)));
                isExists = true;
            }
            else
            {
                isExists = false;
            }
            return isExists;
        }

        public void AddNewTypeOfCustomer(string value)
        {
            _addNew = WaitForElementIsVisible(By.XPath(ADD_NEW));
            _addNew.Click();
            WaitForLoad();

            _inputTypeCustomer = WaitForElementIsVisible(By.Id(CUST_NEW_TYPE_NAME));
            _inputTypeCustomer.SetValue(ControlType.TextBox, value);

            _saveNewBtn = WaitForElementIsVisible(By.Id(CUST_NEW_TYPE_SAVE_BUTTON), nameof(CUST_NEW_TYPE_SAVE_BUTTON));
            _saveNewBtn.Click();
            WaitForLoad();
        }

        public void ClickOrderTypeTab()
        {
            _orderType_tab = WaitForElementIsVisible(By.XPath(ORDER_TYPE_TAB));
            _orderType_tab.Click();
            WaitForLoad();
        }
        public void AddNewOrderType(string value)
        {
            _addOrderType = WaitForElementIsVisible(By.XPath(ADD_ORDER_TYPE));
            _addOrderType.Click();

            _inputOrderType = WaitForElementIsVisible(By.XPath(ORDER_TYPE_NAME));
            _inputOrderType.SetValue(ControlType.TextBox, value);

            _saveNewBtn = WaitForElementIsVisible(By.XPath(ORDER_TYPE_SAVE_BUTTON));
            _saveNewBtn.Click();
        }

        public bool isServiceCategoryFamiliesExist(string groupFamilies)
        {
            bool isExists;
            if (isElementVisible(By.XPath(string.Format(GROUP_FAMILIES_NAME, groupFamilies))))
            {
                isExists = true;
            }
            else
            {
                isExists = false;
            }
            return isExists;
        }

        public void AddNewServiceCategoryFamilies(string name)
        {
            _addNew = WaitForElementIsVisible(By.XPath(ADD_NEW));
            _addNew.Click();
            WaitForLoad();

            _inputTypeCustomer = WaitForElementIsVisible(By.Id(CUST_NEW_TYPE_NAME));
            _inputTypeCustomer.SetValue(ControlType.TextBox, name);

            _saveNewBtn = WaitForElementIsVisible(By.Id(CUST_NEW_TYPE_SAVE_BUTTON), nameof(CUST_NEW_TYPE_SAVE_BUTTON));
            _saveNewBtn.Click();
            WaitForLoad();
        }

        public bool isCategoryProduct(string categoryName)
        {
            IWebElement isproductGroup;
            isproductGroup = WaitForElementExists(By.XPath(string.Format(ISPRODUCT_CATEGORY, categoryName)));
            return isproductGroup.Selected;
        }
        public bool isCategoryFood(string categoryName)
        {
            var isproductFood = WaitForElementExists(By.XPath(string.Format(ISFOOD_CATEGORY, categoryName)));
            return isproductFood.Selected;
        }
        public void EditCategory(string categoryName, bool isProduct, bool isFood)
        {
            var editGroup = WaitForElementExists(By.XPath(string.Format(EDIT_CATEGORY, categoryName)));
            editGroup.Click();
            WaitForLoad();
            
            if (isProduct)
            {
                var isProductCheck = WaitForElementExists(By.XPath("//*[@id=\"IsProduct\"]"));
                isProductCheck.Click();
                WaitForLoad();
            }
            if (isFood)
            {
                var isFoodCheck = WaitForElementExists(By.XPath("//*[@id=\"IsFood\"]"));
                isFoodCheck.Click();
                WaitForLoad();
            }
        }
        public void AddNewCategory(string name, bool isFood, bool isProduct)
        {
            _addNewCategory = WaitForElementIsVisible(By.XPath(ADD_NEW_CATEGORY));
            _addNewCategory.Click();
            WaitForLoad();

            _CategoryName = WaitForElementIsVisible(By.Id(CATEGORY_NEW_NAME));
            _CategoryName.SetValue(ControlType.TextBox, name);

            if (isFood)
            {
                _isFood = WaitForElementExists(By.Id(CHECK_ISFOOD));
                _isFood.Click();
                WaitForLoad();
            }
            if (isProduct)
            {
                var _isProduct = WaitForElementExists(By.Id(CHECK_ISPRODUCT));
                _isProduct.Click();
                WaitForLoad();
            }
            _saveButton = WaitForElementIsVisible(By.Id(CATEGORY_SAVE_BUTTON), nameof(CATEGORY_SAVE_BUTTON));
            _saveButton.Click();
            WaitForLoad();
        }

    }
}
