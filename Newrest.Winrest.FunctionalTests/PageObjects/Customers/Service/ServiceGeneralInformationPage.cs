using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium.Support.UI;
using System.Linq;
using System.Threading;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Customers.Service
{
    public class ServiceGeneralInformationPage : PageBase
    {
        public ServiceGeneralInformationPage(IWebDriver _webDriver, TestContext _testContext) : base(_webDriver, _testContext)
        {
        }

        //____________________________________________ Constantes ____________________________________________________________

        // Général
        private const string BACK_TO_LIST = "/html/body/div[2]/a/span[2]";

        // Onglets
        private const string PRICE_TAB = "hrefTabContentPrice";

        // Informations
        private const string SERVICE_NAME = "first";
        private const string SERVICE_CODE = "Code";
        private const string PRODUCTION_NAME = "ProductionName";
        private const string CATEGORY = "CategoryId";
        private const string GUEST_TYPE = "GuestTypeId";
        private const string SERVICE_TYPE = "ServiceTypeId";
        private const string IS_ACTIVATED = "IsActive";
        private const string PRODUCED = "//*[@id=\"IsProduced\"][@value='true']";
        private const string NOTPRODUCED = "//*[@id=\"IsProduced\"][@value='false']";
        private const string GENERIC_BTN = "/html/body/div[3]/div/div/div[2]/div/div/div/form/div/div[1]/div[1]/div[2]/div[4]/div/input[1]";

        private const string NOISINVOISED = "//*[@id=\"IsInvoiced\"][@value='false']";
        private const string ISINVOISED = "//*[@id=\"IsInvoiced\"][@value='true']";
        private const string ISSPML = "//*[@id=\"IsSPML\"][@value='true']";
        private const string ISNOSPML = "//*[@id=\"IsSPML\"][@value='false']";
        private const string GENERAL_INFORMATION = "//*[@id=\"serviceDetailsTab\"]/li[1]";
        private const string PRICE = "//*[@id=\"serviceDetailsTab\"]/li[2]";
        private const string MESSAGE = "//*[@id=\"Serive-IsSPML-withoutSPML\"]/div[2]/div/div/div/p";
        private const string PRODUCED_VALIDATOR_MESSAGE = "//*[@id=\"service-filter-form\"]/div/div[1]/div[1]/div[2]/div[2]/div/div/span]";
        private const string EXTERNAL_IDENTIFIER = "ExternalCode";
        private const string TEMPLATE_ID = "TemplateId";


    //_____________________________________________ Variables _____________________________________________________________

    // Général
    [FindsBy(How = How.XPath, Using = ISSPML)]
        private IWebElement _isspml;

        [FindsBy(How = How.XPath, Using = ISNOSPML)]
        private IWebElement _isnospml;


        [FindsBy(How = How.XPath, Using = NOISINVOISED)]
        private IWebElement _noisinvoised;


        [FindsBy(How = How.XPath, Using = ISINVOISED)]
        private IWebElement _isinvoised;

        [FindsBy(How = How.XPath, Using = BACK_TO_LIST)]
        private IWebElement _backToList;

        // Onglets
        [FindsBy(How = How.Id, Using = PRICE_TAB)]
        private IWebElement _priceTab;

        // Informations
        [FindsBy(How = How.Id, Using = SERVICE_NAME)]
        private IWebElement _serviceName;

        [FindsBy(How = How.Id, Using = SERVICE_CODE)]
        private IWebElement _serviceCode;

        [FindsBy(How = How.Id, Using = PRODUCTION_NAME)]
        private IWebElement _productionName;

        [FindsBy(How = How.Id, Using = CATEGORY)]
        private IWebElement _category;

        [FindsBy(How = How.Id, Using = GUEST_TYPE)]
        private IWebElement _guestType;

        [FindsBy(How = How.Id, Using = SERVICE_TYPE)]
        private IWebElement _serviceType;

        [FindsBy(How = How.Id, Using = IS_ACTIVATED)]
        private IWebElement _isActive;

        [FindsBy(How = How.XPath, Using = PRODUCED)]
        private IWebElement _produced;

        [FindsBy(How = How.XPath, Using = NOTPRODUCED)]
        private IWebElement _notProduced;

        [FindsBy(How = How.XPath, Using = GENERIC_BTN)]
        private IWebElement _genericBtn;

        [FindsBy(How = How.XPath, Using = EXTERNAL_IDENTIFIER)]
        private IWebElement _externalId;

        [FindsBy(How = How.XPath, Using = TEMPLATE_ID)]
        private IWebElement _templateId;


        //___________________________________________  Méthodes __________________________________________________

        // Général
        public ServicePage BackToList()
        {
            _backToList = WaitForElementIsVisible(By.XPath(BACK_TO_LIST));
            _backToList.Click();
            WaitForLoad();

            return new ServicePage(_webDriver, _testContext);
        }

        // Onglets
        public ServicePricePage GoToPricePage()
        {
            _priceTab = WaitForElementToBeClickable(By.Id(PRICE_TAB));
            _priceTab.Click();
            WaitPageLoading();
            WaitForLoad();
            return new ServicePricePage(_webDriver, _testContext);
        }

        // Informations
        public void Update_General_Informations(string serviceName, string category, string guestType, string serviceType)
        {
            _serviceName = WaitForElementIsVisible(By.Id(SERVICE_NAME));
            _serviceName.SetValue(ControlType.TextBox, serviceName);

            _serviceCode = WaitForElementIsVisible(By.Id(SERVICE_CODE));
            _serviceCode.SetValue(ControlType.TextBox, serviceName);

            _productionName = WaitForElementIsVisible(By.Id(PRODUCTION_NAME));
            _productionName.SetValue(ControlType.TextBox, serviceName);

            _category = WaitForElementIsVisible(By.Id(CATEGORY));
            _category.SetValue(ControlType.DropDownList, category);

            _guestType = WaitForElementIsVisible(By.Id(GUEST_TYPE));
            _guestType.SetValue(ControlType.DropDownList, guestType);

            _serviceType = WaitForElementIsVisible(By.Id(SERVICE_TYPE));
            _serviceType.SetValue(ControlType.DropDownList, serviceType);

            _isActive = WaitForElementIsVisible(By.Id(IS_ACTIVATED));
            _isActive.Click();
        }

        public string GetServiceName()
        {
            _serviceName = WaitForElementIsVisible(By.Id(SERVICE_NAME));
            return _serviceName.GetAttribute("value");
        }

        public string GetServiceCode()
        {
            _serviceCode = WaitForElementIsVisible(By.Id(SERVICE_CODE));
            return _serviceCode.GetAttribute("value");
        }

        public string GetProductionName()
        {
            _productionName = WaitForElementIsVisible(By.Id(PRODUCTION_NAME));
            return _productionName.GetAttribute("value");
        }

        public string GetCategory()
        {
            var category = new SelectElement(_webDriver.FindElement(By.Id(CATEGORY)));
            return category.AllSelectedOptions.FirstOrDefault().Text;
        }

        public void SetCategory(string category)
        {
            _category = WaitForElementIsVisible(By.Id(CATEGORY));
            _category.SetValue(ControlType.DropDownList, category);

            // Temps d'enregistrement de la valeur
            WaitPageLoading();
        }

        public bool IsGeneric()
        {
            WaitForLoad();
            _produced = WaitForElementIsVisible(By.XPath("//*[@id='IsGeneric'][@value='true']"));
            return _produced.Selected;
        }

        public string GetGuestType()
        {
            var guest = new SelectElement(_webDriver.FindElement(By.Id(GUEST_TYPE)));
            return guest.AllSelectedOptions.FirstOrDefault().Text;
        }

        public string GetServiceType()
        {
            SelectElement serviceType = new SelectElement(_webDriver.FindElement(By.Id(SERVICE_TYPE)));
            return serviceType.AllSelectedOptions.FirstOrDefault().Text;
        }

        public void SetActive(bool active)
        {
            _isActive = WaitForElementExists(By.Id(IS_ACTIVATED));
            _isActive.SetValue(ControlType.CheckBox, active);

            // Prise en compte de la modification
            Thread.Sleep(1500);
            WaitForLoad();
        }

        public void SetProduced(bool produced)
        {
            _produced = WaitForElementExists(By.XPath(PRODUCED));
            _produced.SetValue(ControlType.RadioButton, produced);

            // Prise en compte de la modification
            Thread.Sleep(1500);
            WaitForLoad();
        }

        public void SetGeneric()
        {
            _genericBtn = WaitForElementIsVisible(By.Id("IsGeneric"));
            _genericBtn.Click();

            WaitPageLoading();
        }

        public bool IsProduced()
        {
            _produced = WaitForElementIsVisible(By.XPath(PRODUCED));
            return _produced.GetAttribute("checked") != null;
        }

        //*****************************
        public bool IsInvoiced()
        {
            _produced = WaitForElementIsVisible(By.XPath("//*[@id=\"IsInvoiced\"][@value='true']"));
            return _produced.GetAttribute("checked") != null;
        }

        public bool IsSPML()
        {
            _produced = WaitForElementIsVisible(By.XPath("//*[@id=\"IsSPML\"][@value='true']"));
            return _produced.GetAttribute("checked") != null;
        }

        public bool IsCheckList()
        {
            _produced = WaitForElementIsVisible(By.XPath("//*[@id=\"CheckList\"][@value='true']"));
            return _produced.GetAttribute("checked") != null;
        }

        public void SetNotProduced()
        {
            _notProduced = WaitForElementExists(By.XPath(NOTPRODUCED));
            _notProduced.SetValue(ControlType.RadioButton, true);

            // Prise en compte de la modification
            Thread.Sleep(1500);
            WaitForLoad();
        }

        public void SetInvoised()
        {
            _isinvoised = WaitForElementExists(By.XPath(ISINVOISED));
            _isinvoised.SetValue(ControlType.RadioButton, true);

            // Prise en compte de la modification
            Thread.Sleep(1500);
            WaitForLoad();
        }

         public void SetNoInvoised()
         {
            _noisinvoised = WaitForElementExists(By.XPath(NOISINVOISED));
            _noisinvoised.SetValue(ControlType.RadioButton, true);

            // Prise en compte de la modification
            Thread.Sleep(1500);
            WaitForLoad();
         }

        public void SetProductionName(string productionname)
        {
            _productionName = WaitForElementIsVisible(By.Id(PRODUCTION_NAME));
            _productionName.SetValue(ControlType.TextBox, productionname);
            WaitPageLoading();
            WaitForLoad();
        }

        public void SetSPML()
        {
            _isspml = WaitForElementExists(By.XPath(ISSPML));
            _isspml.SetValue(ControlType.RadioButton, true);
            // Prise en compte de la modification
            Thread.Sleep(1500);
            WaitForLoad();
        }
        public void SetNoSPML()
        {
            _isnospml = WaitForElementExists(By.XPath(ISNOSPML));
            _isnospml.SetValue(ControlType.RadioButton, false);
            // Prise en compte de la modification
            Thread.Sleep(1500);
            WaitForLoad();
        }

        public void SpecialMeals()
        {
            _category = WaitForElementIsVisible(By.Id("SpecialMealId"));
            _category.SetValue(ControlType.DropDownList, "AVML");
            WaitForLoad();
        }
        public ServiceGeneralInformationPage SelectGeneralInformation()
        {
            WaitForLoad();
            var generalinformation = WaitForElementIsVisible(By.XPath(GENERAL_INFORMATION));
            generalinformation.Click();
            WaitPageLoading();
            return new ServiceGeneralInformationPage(_webDriver, _testContext);
        }
        public ServiceGeneralInformationPage SelectISSPML()
        {
            WaitForLoad();
            var isspml = WaitForElementIsVisible(By.XPath(ISSPML));
            isspml.Click();
            WaitPageLoading();
            return new ServiceGeneralInformationPage(_webDriver, _testContext);
        }
        public ServiceGeneralInformationPage SelectPrice()
        {
            WaitForLoad();
            var price = WaitForElementIsVisible(By.XPath(PRICE));
            price.Click();
            WaitPageLoading();
            return new ServiceGeneralInformationPage(_webDriver, _testContext);
        }
        public bool IsVisible()
        {
            return isElementVisible(By.XPath(MESSAGE));
          
        }

        public bool CheckValidator()
        {
            return isElementExists(By.XPath(PRODUCED_VALIDATOR_MESSAGE));
        }

        public void SetServiceCode(string serviceCode)
        {
            _serviceCode = WaitForElementIsVisible(By.Id(SERVICE_CODE));
            _serviceCode.SetValue(ControlType.TextBox, serviceCode);
            WaitPageLoading();
            WaitForLoad();
        }

        public void SetGuestType(string guestType)
        {
            _guestType = WaitForElementIsVisible(By.Id(GUEST_TYPE));
            _guestType.SetValue(ControlType.DropDownList, guestType);

            // Temps d'enregistrement de la valeur
            WaitPageLoading();
        }

        public void SetServiceName(string serviceName)
        {
            _serviceName = WaitForElementIsVisible(By.Id(SERVICE_NAME));
            _serviceName.SetValue(ControlType.TextBox, serviceName);
            WaitPageLoading();
            WaitForLoad();
        }

        public void SetServiceType(string serviceType)
        {
            _serviceType = WaitForElementIsVisible(By.Id(SERVICE_TYPE));
            _serviceType.SetValue(ControlType.DropDownList, serviceType);
            // Temps d'enregistrement de la valeur
            WaitPageLoading();
        }

        public string GetExternalId()
        {
            _externalId = WaitForElementIsVisible(By.Id(EXTERNAL_IDENTIFIER));
            return _externalId.GetAttribute("value");
        }

        public void SetExternalId(string externalId)
        {
            _externalId = WaitForElementIsVisible(By.Id(EXTERNAL_IDENTIFIER));
            _externalId.SetValue(ControlType.TextBox, externalId);
            WaitPageLoading();
            WaitForLoad();
        }

        public string GetTemplateId()
        {
            _templateId = WaitForElementIsVisible(By.Id(TEMPLATE_ID));
            return _templateId.GetAttribute("value");
        }

        public void SetTemplateId(string templateId)
        {
            _templateId = WaitForElementIsVisible(By.Id(TEMPLATE_ID));
            _templateId.SetValue(ControlType.TextBox, templateId);
            WaitPageLoading();
            WaitForLoad();
        }
    }
}
