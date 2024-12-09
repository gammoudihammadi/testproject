using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Purchasing;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.PageObjects;
using System.Threading;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Customers.Delivery
{
    public class DeliveryLoadingPage : PageBase
    {

        public DeliveryLoadingPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        //__________________________________________ Constantes ____________________________________________________________

        // General
        private const string BACK_TO_LIST = "/html/body/div[2]/a/span[2]";
        private const string DELIVERY_NAME = "//*[@id=\"div-body\"]/div/h1";

        // Onglets
        private const string GENERAL_INFORMATONS_TAB = "hrefTabContentInformation";

        // Ajout service
        private const string ADD_SERVICE = "ServiceName";
        private const string SERVICE = "//*[@id=\"service-list-result\"]/table/tbody/tr[*]/td[2]/span[text()='{0}']";
        private const string DELETE_SERVICE = "customers-flightdelivery-packagingList-deleteBtn-1";
        private const string CONFIRM_DELETE_SERVICE = "dataConfirmOK";

        // Tableau jours
        private const string MONDAY_QTY = "item_Monday";
        private const string TUESDAY_QTY = "item_Tuesday";
        private const string WEDNESDAY_QTY = "item_Wednesday";
        private const string THURSDAY_QTY = "item_Thursday";
        private const string FRIDAY_QTY = "item_Friday";
        private const string SATURDAY_QTY = "item_Saturday";
        private const string SUNDAY_QTY = "item_Sunday";

        private const string MONDAY_QTY_WITH_SERVICE = "//*[starts-with(@id,'deliveryToService')]/div/div/form/div/div[1]/div[2][contains(text(),'{0}')]/../../div[2]//*[@id='item_Monday']";
        private const string TUESDAY_QTY_WITH_SERVICE = "//*[starts-with(@id,'deliveryToService')]/div/div/form/div/div[1]/div[2][contains(text(),'{0}')]/../../div[3]//*[@id='item_Tuesday']";
        private const string WEDNESDAY_QTY_WITH_SERVICE = "//*[starts-with(@id,'deliveryToService')]/div/div/form/div/div[1]/div[2][contains(text(),'{0}')]/../../div[4]//*[@id='item_Wednesday']";
        private const string THURSDAY_QTY_WITH_SERVICE = "//*[starts-with(@id,'deliveryToService')]/div/div/form/div/div[1]/div[2][contains(text(),'{0}')]/../../div[5]//*[@id='item_Thursday']";
        private const string FRIDAY_QTY_WITH_SERVICE = "//*[starts-with(@id,'deliveryToService')]/div/div/form/div/div[1]/div[2][contains(text(),'{0}')]/../../div[6]//*[@id='item_Friday']";
        private const string SATURDAY_QTY_WITH_SERVICE = "//*[starts-with(@id,'deliveryToService')]/div/div/form/div/div[1]/div[2][contains(text(),'{0}')]/../../div[7]//*[@id='item_Saturday']";
        private const string SUNDAY_QTY_WITH_SERVICE = "//*[starts-with(@id,'deliveryToService')]/div/div/form/div/div[1]/div[2][contains(text(),'{0}')]/../../div[8]//*[@id='item_Sunday']";
        private const string LIST_SERVICE_NAME = "//*[starts-with(@id,'deliveryToService')]/div/div/form/div/div[1]/div[contains(text(),'{0}')]";

        private const string ITEM_SERVICE = "item-service-row";

        // Packaging
        private const string ADD_PACKAGING_PATCH = "//*[@id=\"recipe-variant-detail-list\"]/div[contains(@class,'item-tab-body')]/ul/li/div/div/form/div/div[17]/div/a[1]";
        private const string ADD_PACKAGING_DEV = "//*[@id=\"recipe-variant-detail-list\"]/div[*]/ul/li/div/div/form/div/div[18]/div/a[1]";
        private const string EDIT_BTN_DEV = "//*[@id=\"modal-1\"]/div[2]/div/div/div[2]/div/div/div/div/form/div[2]/div[6]/div/div[2]/a[1]";
        private const string EDIT_BTN_PATCH = "customers-flightdelivery-editpackaging-display-1";
        private const string ADD_PACKAGING_BTN = "//*[@id=\"modal-1\"]/div/div/div[3]/a";

        private const string PACKAGING_METHOD = "drop-down-methods";
        private const string PACKAGING_METHOD_SELECTED = "customers-flightdelivery-packaging-method-locked";
        private const string PACKAGING_NAME = "first";
        private const string PACKAGING_PAX_MIN = "PaxMin";
        private const string PACKAGING_PAX_MAX = "PaxMax";
        private const string PACKAGING_PORTIONS_NUMBER = "NumberOfPortion";
        private const string INPUT_PORTION = "NumberOfPortion";
        private const string ALL_RECIPE_TYPES_CHECKBOX = "/html/body/div[5]/div/div/div/div/form/div[2]/div[6]/div/div[1]/div/div";
        private const string RECIPE_TYPE_CHECKBOX = "//span[contains(text(),'{0}')]/../div";
        private const string RECIPE_TYPE_CHECKBOX_ITALIC = "//i[contains(text(),'{0}')]/../../div";
        private const string SAVE_PACKAGING = "/html/body/div[5]/div/div/div/div/form/div[3]/button[2]";

        private const string FIRST_PACKAGING_NAME = "name_0";
        private const string FIRST_PACKAGING_METHOD = "MethodName_0";
                  
        private const string DUPLICATE_PACKAGING_BTN_DEV = "customers-flightdelivery-duplicate-packaging-btn";
        private const string DUPLICATE_PACKAGING_BTN_PATCH = "customers-flightdelivery-duplicate-packaging-btn";
        private const string INPUT_DELIVERY_PACKAGING = "SearchVM_FlightDelivery";
        private const string ITEM_PACK = "//*[@id=\"form\"]/table/tbody/tr[2]";
        private const string DUPLICATE_BTN = "btn-duplicate";
        private const string DELETE_PACKAGING = "//*[@id=\"modal-1\"]/div/div/div[2]/div/div/div[2]/div/div/div/div/form/div[2]/div[6]/div/div[2]/a[2]/span";
        private const string CONFIRM_DELETE_PACKAGING = "dataConfirmOK";
        private const string DELETE_PACK = "//*[@id=\"customers-flightdelivery-deletepackaging-display-1\"]/span";
        private const string DELETE_SERVICE_PATH = "//*/span[contains(@class,'trash')]/..";
        private const string CLOSE_PACKAGING_MODAL = "last";
        private const string VERIFYADDEDSERVICE = "/html/body/div[3]/div/div[2]/div[2]/div/div/div[1]/div/div/div[2]/ul/li[1]";
        private const string SERVICE_NAME = "//*[starts-with(@id,\"deliveryToService\")]/div/div/form/div/div[1]/div[2]";

        //_________________________________________ Variables _____________________________________________________________

        // General
        [FindsBy(How = How.XPath, Using = BACK_TO_LIST)]
        private IWebElement _backToList;
        [FindsBy(How = How.XPath, Using = VERIFYADDEDSERVICE)]
        private IWebElement _verifyaddedservice;

        [FindsBy(How = How.XPath, Using = DELIVERY_NAME)]
        private IWebElement _deliveryName;

        // Onglets
        [FindsBy(How = How.Id, Using = GENERAL_INFORMATONS_TAB)]
        private IWebElement _generalInfoTab;

        // Ajout service
        [FindsBy(How = How.Id, Using = ADD_SERVICE)]
        private IWebElement _addService;

        [FindsBy(How = How.XPath, Using = SERVICE)]
        private IWebElement _serviceToAdd;


        [FindsBy(How = How.Id, Using = DELETE_SERVICE)]
        private IWebElement _deleteService;

        [FindsBy(How = How.Id, Using = CONFIRM_DELETE_SERVICE)]
        private IWebElement _confirmDeleteService;


        // Tableau jours
        [FindsBy(How = How.Id, Using = MONDAY_QTY)]
        private IWebElement _monday;

        [FindsBy(How = How.Id, Using = TUESDAY_QTY)]
        private IWebElement _tuesday;

        [FindsBy(How = How.Id, Using = WEDNESDAY_QTY)]
        private IWebElement _wednesday;

        [FindsBy(How = How.Id, Using = THURSDAY_QTY)]
        private IWebElement _thursday;

        [FindsBy(How = How.Id, Using = FRIDAY_QTY)]
        private IWebElement _friday;

        [FindsBy(How = How.Id, Using = SATURDAY_QTY)]
        private IWebElement _saturday;

        [FindsBy(How = How.Id, Using = SUNDAY_QTY)]
        private IWebElement _sunday;

        [FindsBy(How = How.ClassName, Using = ITEM_SERVICE)]
        private IWebElement _itemService;

        // Packaging
        [FindsBy(How = How.XPath, Using = ADD_PACKAGING_DEV)]
        private IWebElement _addPackagingDev;

        [FindsBy(How = How.XPath, Using = ADD_PACKAGING_PATCH)]
        private IWebElement _addPackagingPatch;

        [FindsBy(How = How.XPath, Using = ADD_PACKAGING_BTN)]
        private IWebElement _addPackagingBtn;

        [FindsBy(How = How.Id, Using = PACKAGING_METHOD)]
        private IWebElement _packagingMethod;

        [FindsBy(How = How.Id, Using = PACKAGING_METHOD_SELECTED)]
        private IWebElement _packagingMethodSelected;

        [FindsBy(How = How.Id, Using = PACKAGING_NAME)]
        private IWebElement _packagingName;

        [FindsBy(How = How.XPath, Using = PACKAGING_PAX_MIN)]
        private IWebElement _packagingPaxMin;

        [FindsBy(How = How.XPath, Using = PACKAGING_PAX_MAX)]
        private IWebElement _packagingPaxMax;

        [FindsBy(How = How.XPath, Using = PACKAGING_PORTIONS_NUMBER)]
        private IWebElement _packagingPortionsNumber; 

        [FindsBy(How = How.Id, Using = INPUT_PORTION)]
        private IWebElement _inputPortion;

        [FindsBy(How = How.XPath, Using = ALL_RECIPE_TYPES_CHECKBOX)]
        private IWebElement _all_recipe_types_checkbox;

        [FindsBy(How = How.XPath, Using = RECIPE_TYPE_CHECKBOX)]
        private IWebElement _recipe_type_checkbox;

        [FindsBy(How = How.XPath, Using = SAVE_PACKAGING)]
        private IWebElement _savePackaging;

        [FindsBy(How = How.Id, Using = FIRST_PACKAGING_NAME)]
        private IWebElement _firstPackagingName;

        [FindsBy(How = How.XPath, Using = FIRST_PACKAGING_METHOD)]
        private IWebElement _firstPackagingMethod;

        [FindsBy(How = How.XPath, Using = EDIT_BTN_DEV)]
        private IWebElement _editPackagingDev;

        [FindsBy(How = How.Id, Using = EDIT_BTN_PATCH)]
        private IWebElement _editPackagingPatch;

        [FindsBy(How = How.Id, Using = DUPLICATE_PACKAGING_BTN_DEV)]
        private IWebElement _duplicatePackagingDev;

        [FindsBy(How = How.Id, Using = DUPLICATE_PACKAGING_BTN_PATCH)]
        private IWebElement _duplicatePackagingPatch;

        [FindsBy(How = How.Id, Using = INPUT_DELIVERY_PACKAGING)]
        private IWebElement _inputDeliveryPackaging;

        [FindsBy(How = How.XPath, Using = ITEM_PACK)]
        private IWebElement _checkDelivery;

        [FindsBy(How = How.Id, Using = DUPLICATE_BTN)]
        private IWebElement _duplicateBtn;

        [FindsBy(How = How.XPath, Using = DELETE_PACKAGING)]
        private IWebElement _deletePackaging;

        [FindsBy(How = How.Id, Using = CONFIRM_DELETE_PACKAGING)]
        private IWebElement _confirmDeletePackaging;
        
        [FindsBy(How = How.Id, Using = CLOSE_PACKAGING_MODAL)]
        private IWebElement _closePackagingModal;
     

        //_______________________________________  Methodes ___________________________________________

        // General
        public DeliveryPage BackToList()
        {
            _backToList = WaitForElementIsVisible(By.XPath(BACK_TO_LIST));
            _backToList.Click();
            WaitForLoad();

            return new DeliveryPage(_webDriver, _testContext);
        }

        public string GetDeliveryName()
        {
            _deliveryName = WaitForElementIsVisible(By.XPath(DELIVERY_NAME));
            return _deliveryName.Text.Substring(_deliveryName.Text.IndexOf(":") + 1).Trim();
        }

        // Onglets
        public DeliveryGeneralInformationPage ClickOnGeneralInformation()
        {
            _generalInfoTab = WaitForElementIsVisible(By.Id(GENERAL_INFORMATONS_TAB));
            _generalInfoTab.Click();
            WaitForLoad();

            return new DeliveryGeneralInformationPage(_webDriver, _testContext);
        }

        // Ajout service
        public void AddService(string service)
        {
            _addService = WaitForElementIsVisible(By.Id(ADD_SERVICE));
            _addService.SetValue(ControlType.TextBox, service);
            WaitPageLoading();

            _serviceToAdd = WaitForElementIsVisible(By.XPath(string.Format(SERVICE, service)));
            _serviceToAdd.Click();
            WaitPageLoading();
        }

        public bool IsServiceVisible()
        {
            WaitForLoad();
            if (isElementVisible(By.ClassName(ITEM_SERVICE)))
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        public void DeleteService()
        {
            Actions action = new Actions(_webDriver);
            _itemService = WaitForElementIsVisible(By.ClassName(ITEM_SERVICE));
            action.MoveToElement(_itemService).Perform();

            if (IsDev())
            {
                _deleteService = WaitForElementIsVisible(By.Id(DELETE_SERVICE));
            }
            else
            {
                _deleteService = WaitForElementIsVisible(By.XPath(DELETE_SERVICE_PATH));
            }
            _deleteService.Click();
            WaitForLoad();

            _confirmDeleteService = WaitForElementIsVisible(By.Id(CONFIRM_DELETE_SERVICE));
            _confirmDeleteService.Click();
            WaitForLoad();

            // Temps de chargement de la page
            WaitPageLoading();
        }

        // Tableau jours
        public void AddQuantity(string qty)
        {
            Actions action = new Actions(_webDriver);
            _itemService = WaitForElementIsVisible(By.ClassName(ITEM_SERVICE));
            action.MoveToElement(_itemService).Perform();

            _monday = WaitForElementIsVisible(By.Id(MONDAY_QTY));
            _monday.SetValue(ControlType.TextBox, qty);

            _tuesday = WaitForElementIsVisible(By.Id(TUESDAY_QTY));
            _tuesday.SetValue(ControlType.TextBox, qty);

            _wednesday = WaitForElementIsVisible(By.Id(WEDNESDAY_QTY));
            _wednesday.SetValue(ControlType.TextBox, qty);

            _thursday = WaitForElementIsVisible(By.Id(THURSDAY_QTY));
            _thursday.SetValue(ControlType.TextBox, qty);

            _friday = WaitForElementIsVisible(By.Id(FRIDAY_QTY));
            _friday.SetValue(ControlType.TextBox, qty);

            _saturday = WaitForElementIsVisible(By.Id(SATURDAY_QTY));
            _saturday.SetValue(ControlType.TextBox, qty);

            _sunday = WaitForElementIsVisible(By.Id(SUNDAY_QTY));
            _sunday.SetValue(ControlType.TextBox, qty);

            Thread.Sleep(2000);
        }
        public void AddQuantityForService(string qty ,string servicName)
        {
            Actions action = new Actions(_webDriver);
           var _itemService = WaitForElementIsVisible(By.XPath(string.Format(LIST_SERVICE_NAME, servicName)));
            action.MoveToElement(_itemService).Perform();

            var _monday = WaitForElementIsVisible(By.XPath(string.Format(MONDAY_QTY_WITH_SERVICE, servicName)));
            _monday.SetValue(ControlType.TextBox, qty);

            var _tuesday = WaitForElementIsVisible(By.XPath(string.Format(TUESDAY_QTY_WITH_SERVICE, servicName)));
            _tuesday.SetValue(ControlType.TextBox, qty);

            var _wednesday = WaitForElementIsVisible(By.XPath(string.Format(WEDNESDAY_QTY_WITH_SERVICE, servicName)));
            _wednesday.SetValue(ControlType.TextBox, qty);

           var  _thursday = WaitForElementIsVisible(By.XPath(string.Format(THURSDAY_QTY_WITH_SERVICE, servicName)));
            _thursday.SetValue(ControlType.TextBox, qty);

           var  _friday = WaitForElementIsVisible(By.XPath(string.Format(FRIDAY_QTY_WITH_SERVICE, servicName)));
            _friday.SetValue(ControlType.TextBox, qty);

           var  _saturday = WaitForElementIsVisible(By.XPath(string.Format(SATURDAY_QTY_WITH_SERVICE, servicName)));
            _saturday.SetValue(ControlType.TextBox, qty);

            var _sunday = WaitForElementIsVisible(By.XPath(string.Format(SUNDAY_QTY_WITH_SERVICE, servicName)));
            _sunday.SetValue(ControlType.TextBox, qty);

            WaitPageLoading();
        }

        // Packaging
        public void AddPackaging()
        {
            Actions action = new Actions(_webDriver);
            action.MoveToElement(_itemService).Perform();

            _addPackagingDev = WaitForElementIsVisible(By.XPath(ADD_PACKAGING_DEV));
            _addPackagingDev.Click();
           
            WaitForLoad(); 
        }

        public void FillField_FoodPackaging(string packaging, string name, bool isPaxPerPackaging, string paxMin = null, string paxMax = null, string portionsNumber = null, string recipeType = null)
        {
            WaitForLoad();
            _addPackagingBtn = WaitForElementIsVisible(By.XPath("//*[@id=\"modal-1\"]/div[3]/a"));
            _addPackagingBtn.Click();
            WaitForLoad();

            if (isElementVisible(By.Id(PACKAGING_METHOD)))
            {
                _packagingMethod = WaitForElementIsVisible(By.Id(PACKAGING_METHOD));
                _packagingMethod.SetValue(ControlType.DropDownList, packaging);
            }
            else
            {
                _packagingMethodSelected = WaitForElementIsVisible(By.Id(PACKAGING_METHOD_SELECTED));
                Assert.IsTrue(_packagingMethodSelected.GetAttribute("value").Equals(packaging));
            }

            _packagingName = WaitForElementIsVisible(By.XPath("/html/body/div[5]/div/div/div/div/form/div[2]/div[2]/div/div/select"));
            _packagingName.SetValue(ControlType.DropDownList, name);

            if (paxMin != null)
            {
                _packagingPaxMin = WaitForElementIsVisible(By.Id(PACKAGING_PAX_MIN));
                _packagingPaxMin.SetValue(ControlType.TextBox, paxMin);
            }

            if (paxMax != null)
            {
                _packagingPaxMax = WaitForElementIsVisible(By.Id(PACKAGING_PAX_MAX));
                _packagingPaxMax.SetValue(ControlType.TextBox, paxMax);
            }

            if (portionsNumber != null)
            {
                _packagingPortionsNumber = WaitForElementIsVisible(By.Id(PACKAGING_PORTIONS_NUMBER));
                _packagingPortionsNumber.SetValue(ControlType.TextBox, portionsNumber);
            }

            if (isPaxPerPackaging)
            {
                _inputPortion = WaitForElementIsVisible(By.Id(INPUT_PORTION));
                _inputPortion.SetValue(ControlType.TextBox, "10");
            }

            if (recipeType != null)
            {
                //first uncheck all recipe types checkbox
                _all_recipe_types_checkbox = WaitForElementIsVisible(By.XPath(ALL_RECIPE_TYPES_CHECKBOX));
                var actions = new Actions(_webDriver);
                actions.MoveToElement(_all_recipe_types_checkbox).Click();
                WaitForLoad();

                //then check recipe type checkbox for test
                if (isElementVisible(By.XPath(string.Format(RECIPE_TYPE_CHECKBOX, recipeType))))
                {
                    _recipe_type_checkbox = WaitForElementIsVisible(By.XPath(string.Format(RECIPE_TYPE_CHECKBOX, recipeType)));
                }
                else
                {
                    _recipe_type_checkbox = WaitForElementIsVisible(By.XPath(string.Format(RECIPE_TYPE_CHECKBOX_ITALIC, recipeType)));
                }
                actions.MoveToElement(_recipe_type_checkbox).Click().Build().Perform();
                WaitForLoad();
            }

            _savePackaging = WaitForElementToBeClickable(By.XPath(SAVE_PACKAGING));
            _savePackaging.Click();
            WaitForLoad();
        }

        public void FillFields_FoodPackaging( string name, bool isPaxPerPackaging, string paxMin, string paxMax , string portionsNumber , string recipeType = null)
        {
            WaitForLoad();
            _addPackagingBtn = WaitForElementIsVisible(By.XPath("//*[@id=\"modal-1\"]/div[3]/a"));
            _addPackagingBtn.Click();
            WaitForLoad();

            _packagingName = WaitForElementIsVisible(By.Id(PACKAGING_NAME));
            _packagingName.SetValue(ControlType.DropDownList, name);

                _packagingPaxMin = WaitForElementIsVisible(By.Id(PACKAGING_PAX_MIN));
                _packagingPaxMin.SetValue(ControlType.TextBox, paxMin);


                _packagingPaxMax = WaitForElementIsVisible(By.Id(PACKAGING_PAX_MAX));
                _packagingPaxMax.SetValue(ControlType.TextBox, paxMax);


                _packagingPortionsNumber = WaitForElementIsVisible(By.Id(PACKAGING_PORTIONS_NUMBER));
                _packagingPortionsNumber.SetValue(ControlType.TextBox, portionsNumber);


            if (isPaxPerPackaging)
            {
                _inputPortion = WaitForElementIsVisible(By.Id(INPUT_PORTION));
                _inputPortion.SetValue(ControlType.TextBox, "10");
            }

            if (recipeType != null)
            {
                //first uncheck all recipe types checkbox
                _all_recipe_types_checkbox = WaitForElementIsVisible(By.XPath(ALL_RECIPE_TYPES_CHECKBOX));
                var actions = new Actions(_webDriver);
                actions.MoveToElement(_all_recipe_types_checkbox).Click();
                WaitForLoad();

                //then check recipe type checkbox for test
                if (isElementVisible(By.XPath(string.Format(RECIPE_TYPE_CHECKBOX, recipeType))))
                {
                    _recipe_type_checkbox = WaitForElementIsVisible(By.XPath(string.Format(RECIPE_TYPE_CHECKBOX, recipeType)));
                }
                else
                {
                    _recipe_type_checkbox = WaitForElementIsVisible(By.XPath(string.Format(RECIPE_TYPE_CHECKBOX_ITALIC, recipeType)));
                }
                actions.MoveToElement(_recipe_type_checkbox).Click().Build().Perform();
                WaitForLoad();
            }

            _savePackaging = WaitForElementToBeClickable(By.XPath(SAVE_PACKAGING));
            _savePackaging.Click();
            WaitForLoad();
        }

        public void CloseFoodPackagingModal()
        {
            _closePackagingModal = WaitForElementIsVisible(By.Id(CLOSE_PACKAGING_MODAL));
            _closePackagingModal.Click();
        }

        public string GetFirstPackagingName()
        {
            _firstPackagingName = WaitForElementIsVisible(By.Id(FIRST_PACKAGING_NAME));
            return _firstPackagingName.Text;
        }

        public string GetFirstPackagingMethod()
        {
            _firstPackagingMethod = WaitForElementIsVisible(By.Id(FIRST_PACKAGING_METHOD));
            return _firstPackagingMethod.Text;
        }

        public void UpdatePackaging(string packaging, string name, bool isPaxPerPackaging)
        {
            if (IsDev())
            {
                _editPackagingDev = WaitForElementIsVisible(By.XPath(EDIT_BTN_DEV));
                _editPackagingDev.Click();
                WaitForLoad();
            }
            else
            {
                _editPackagingDev = WaitForElementIsVisible(By.Id(EDIT_BTN_PATCH));
                _editPackagingDev.Click();
                WaitForLoad();
            }

            _packagingMethod = WaitForElementIsVisible(By.Id(PACKAGING_METHOD));
            _packagingMethod.SetValue(ControlType.DropDownList, packaging);

            _packagingName = WaitForElementIsVisible(By.XPath("/html/body/div[5]/div/div/div/div/form/div[2]/div[2]/div/div/select"));
            _packagingName.SetValue(ControlType.DropDownList, name);

            if (isPaxPerPackaging)
            {
                _inputPortion = WaitForElementIsVisible(By.Id(INPUT_PORTION));
                _inputPortion.SetValue(ControlType.TextBox, "10");
            }

            _savePackaging = WaitForElementToBeClickable(By.XPath(SAVE_PACKAGING));
            _savePackaging.Click();
            WaitForLoad();
        }


        public void UpdatePackagingFood(string name, bool isPaxPerPackaging)
        {
            WaitForLoad();

            if (IsDev())
            {
                _editPackagingDev = WaitForElementIsVisible(By.XPath(EDIT_BTN_DEV));
                _editPackagingDev.Click();
                WaitForLoad();
            }
            else
            {
                _editPackagingDev = WaitForElementIsVisible(By.Id(EDIT_BTN_PATCH));
                _editPackagingDev.Click();
                WaitForLoad();
            }
            _packagingName = WaitForElementIsVisible(By.Id(PACKAGING_NAME));
            _packagingName.SetValue(ControlType.DropDownList, name);

            if (isPaxPerPackaging)
            {
                _inputPortion = WaitForElementIsVisible(By.Id(INPUT_PORTION));
                _inputPortion.SetValue(ControlType.TextBox, "10");
            }

            _savePackaging = WaitForElementToBeClickable(By.XPath(SAVE_PACKAGING));
            _savePackaging.Click();
            WaitForLoad();
        }

        public void DuplicatePackaging(string value)
        {
            if (IsDev())
            {
                _duplicatePackagingDev = WaitForElementIsVisible(By.Id(DUPLICATE_PACKAGING_BTN_DEV));
                _duplicatePackagingDev.Click();
                WaitForLoad();
            }
            else
            {
                _duplicatePackagingDev = WaitForElementIsVisible(By.Id(DUPLICATE_PACKAGING_BTN_PATCH));
                _duplicatePackagingDev.Click();
                WaitForLoad();

            }

            _inputDeliveryPackaging = WaitForElementIsVisible(By.Id(INPUT_DELIVERY_PACKAGING));
            _inputDeliveryPackaging.SetValue(ControlType.TextBox, value);
            WaitForLoad();

            _checkDelivery = WaitForElementIsVisible(By.XPath(ITEM_PACK));
            _checkDelivery.Click();
            WaitForLoad();

            _duplicateBtn = WaitForElementIsVisible(By.Id(DUPLICATE_BTN));
            _duplicateBtn.Click();
            WaitForLoad();
        }

        public void DeletePackaging()
        {
            if (IsDev())
            {
                _deletePackaging = WaitForElementIsVisible(By.XPath("//*[@id=\"customers-flightdelivery-deletepackaging-display-1\"]/span"));
                _deletePackaging.Click();
                WaitForLoad();
            }
            else
            {
                _deletePackaging = WaitForElementIsVisible(By.XPath(DELETE_PACK));
                _deletePackaging.Click();
                WaitForLoad();
            }
            _confirmDeletePackaging = WaitForElementIsVisible(By.Id(CONFIRM_DELETE_PACKAGING));
            _confirmDeletePackaging.Click();
            WaitForLoad();

            // Temps de chargement de la page
            Thread.Sleep(2000);
        }

        public bool IsPackagingVisible()
        {
            WaitForLoad();
            if (isElementVisible(By.XPath(ITEM_PACK)))
            {
                return true;
            }
            else
            {
                return false;
            }
        }
        public string GetFirstStyleService()
        {
            var firstStyleService = WaitForElementIsVisible(By.XPath("/html/body/div[3]/div/div[2]/div[2]/div/div/div[1]/div/div/div[2]/ul/li/div/div/form/div/div[1]/div[2]"));
            IJavaScriptExecutor js = (IJavaScriptExecutor)_webDriver;
            string color = (string)js.ExecuteScript("return window.getComputedStyle(arguments[0], null).getPropertyValue('color');", firstStyleService);
            return color;
        }

        public string GetServiceName()
        {
            var ServiceName = WaitForElementIsVisible(By.XPath(SERVICE_NAME));
            return ServiceName.Text;
        }
        public bool VerifyAddedService()
        {
            if (isElementExists(By.XPath(VERIFYADDEDSERVICE)))
            {
                return true;
            }
            else {
                return false;
            }
        }
    }
}
