using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Parameters.Clinic
{
    public class ParametersClinic : PageBase
    {

        public ParametersClinic(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {

        }

        // __________________________________________ Constantes _________________________________________

        // TABS
        private const string CL_PAT_TEXTURE_TAB = "//li/a[contains(text(), 'Texture')]";
        private const string CL_PAT_OFFER_TAB = "//li/a[contains(text(), 'Offer')]";
        private const string CL_PAT_PACKAGE_TAB = "//li/a[contains(text(), 'Package')]";

        //TEXTURE
        private const string CL_PAT_TEXTURE = "//*[@id=\"tabContentParameters\"]/div[2]/div/table/tbody/tr[2]/td[1][contains(text(), '{0}')]";
        private const string CL_PAT_NEW_TEXTURE_BUTTON = "//a[contains(text(), 'New')]";
        private const string CL_PAT_NEW_TEXTURE_NAME_INPUT = "Name";
        private const string CL_PAT_NEW_TEXTURE_ORDER_INPUT = "Order";
        private const string CL_PAT_NEW_TEXTURE_SAVE_BUTTON = "last";

        //OFFER
        private const string CL_PAT_OFFER = "//*[@id=\"tabContentParameters\"]/div[2]/div/table/tbody/tr[2]/td[1][contains(text(), '{0}')]";
        private const string CL_PAT_NEW_OFFER_BUTTON = "//a[contains(text(), 'New')]";
        private const string CL_PAT_NEW_OFFER_NAME_INPUT = "Name";
        private const string CL_PAT_NEW_OFFER_ORDER_INPUT = "Order";
        private const string CL_PAT_NEW_OFFER_SAVE_BUTTON = "last";

        //PACKAGE
        private const string CL_PAT_PACKAGE = "//*[@id=\"tabContentParameters\"]/div[2]/div/table/tbody/tr[2]/td[1][contains(text(), '{0}')]";
        private const string CL_PAT_NEW_PACKAGE_BUTTON = "//a[contains(text(), 'New')]";
        private const string CL_PAT_NEW_PACKAGE_NAME_INPUT = "Name";
        private const string CL_PAT_NEW_PACKAGE_ORDER_INPUT = "Order";
        private const string CL_PAT_NEW_PACKAGE_SAVE_BUTTON = "last";

        // __________________________________________ Variables __________________________________________
        // TABS
        [FindsBy(How = How.XPath, Using = CL_PAT_TEXTURE_TAB)]
        private IWebElement _textureTab;

        [FindsBy(How = How.XPath, Using = CL_PAT_OFFER_TAB)]
        private IWebElement _offerTab;

        [FindsBy(How = How.XPath, Using = CL_PAT_PACKAGE_TAB)]
        private IWebElement _packageTab;

        //TEXTURE
        [FindsBy(How = How.XPath, Using = CL_PAT_NEW_TEXTURE_BUTTON)]
        private IWebElement _newTextureButton;

        [FindsBy(How = How.Name, Using = CL_PAT_NEW_TEXTURE_NAME_INPUT)]
        private IWebElement _newTextureNameInput;

        [FindsBy(How = How.Name, Using = CL_PAT_NEW_TEXTURE_ORDER_INPUT)]
        private IWebElement _newTextureOrderInput;

        [FindsBy(How = How.Id, Using = CL_PAT_NEW_TEXTURE_SAVE_BUTTON)]
        private IWebElement _newTextureSaveButton;

        //OFFER
        [FindsBy(How = How.XPath, Using = CL_PAT_NEW_OFFER_BUTTON)]
        private IWebElement _newOfferButton;

        [FindsBy(How = How.Name, Using = CL_PAT_NEW_OFFER_NAME_INPUT)]
        private IWebElement _newOfferNameInput;

        [FindsBy(How = How.Name, Using = CL_PAT_NEW_OFFER_ORDER_INPUT)]
        private IWebElement _newOfferOrderInput;

        [FindsBy(How = How.Id, Using = CL_PAT_NEW_OFFER_SAVE_BUTTON)]
        private IWebElement _newOfferSaveButton;

        //PACKAGE
        [FindsBy(How = How.XPath, Using = CL_PAT_NEW_PACKAGE_BUTTON)]
        private IWebElement _newPackageButton;

        [FindsBy(How = How.Name, Using = CL_PAT_NEW_PACKAGE_NAME_INPUT)]
        private IWebElement _newPackageNameInput;

        [FindsBy(How = How.Name, Using = CL_PAT_NEW_PACKAGE_ORDER_INPUT)]
        private IWebElement _newPackageOrderInput;

        [FindsBy(How = How.Id, Using = CL_PAT_NEW_PACKAGE_SAVE_BUTTON)]
        private IWebElement _newPackageSaveButton;

        // _________________________________________ Méthodes ______________________________________________
        // TABS
        public void ClickToTexture()
        {
            _textureTab = WaitForElementIsVisible(By.XPath(CL_PAT_TEXTURE_TAB), nameof(CL_PAT_TEXTURE_TAB));
            _textureTab.Click();
            WaitForLoad();
        }
        public void ClickToOffer()
        {
            _offerTab = WaitForElementIsVisible(By.XPath(CL_PAT_OFFER_TAB), nameof(CL_PAT_OFFER_TAB));
            _offerTab.Click();
            WaitForLoad();
        }
        public void ClickToPackage()
        {
            _packageTab = WaitForElementIsVisible(By.XPath(CL_PAT_PACKAGE_TAB), nameof(CL_PAT_PACKAGE_TAB));
            _packageTab.Click();
            WaitForLoad();
        }

        //TEXTURE
        public bool isTextureExist(string textureName)
        {
            bool isExists;
            try
            {
                WaitForElementIsVisible(By.XPath(string.Format(CL_PAT_TEXTURE, textureName)));
                isExists = true;
            }
            catch
            {
                isExists = false;
            }
            return isExists;
        }

        public void AddNewTexture(string offerName, string offerOrder)
        {
            _newTextureButton = WaitForElementIsVisible(By.XPath(CL_PAT_NEW_TEXTURE_BUTTON), nameof(CL_PAT_NEW_TEXTURE_BUTTON));
            _newTextureButton.Click();
            WaitForLoad();

            // Renseigner le name
            _newTextureNameInput = WaitForElementIsVisible(By.Name(CL_PAT_NEW_TEXTURE_NAME_INPUT), nameof(CL_PAT_NEW_TEXTURE_NAME_INPUT));
            _newTextureNameInput.SetValue(ControlType.TextBox, offerName);

            // Renseigner l'ordre
            _newTextureOrderInput = WaitForElementIsVisible(By.Name(CL_PAT_NEW_TEXTURE_ORDER_INPUT), nameof(CL_PAT_NEW_TEXTURE_ORDER_INPUT));
            _newTextureOrderInput.SetValue(ControlType.TextBox, offerOrder);

            //Save
            _newTextureSaveButton = WaitForElementIsVisible(By.Id(CL_PAT_NEW_TEXTURE_SAVE_BUTTON), nameof(CL_PAT_NEW_TEXTURE_SAVE_BUTTON));
            _newTextureSaveButton.Click();
            WaitForLoad();
        }

        //OFFER
        public bool isOfferExist(string offerName)
        {
            bool isExists;
            try
            {
                WaitForElementIsVisible(By.XPath(string.Format(CL_PAT_OFFER, offerName)));
                isExists = true;
            }
            catch
            {
                isExists = false;
            }
            return isExists;
        }

        public void AddNewOffer(string offerName, string offerOrder)
        {
            _newOfferButton = WaitForElementIsVisible(By.XPath(CL_PAT_NEW_OFFER_BUTTON), nameof(CL_PAT_NEW_OFFER_BUTTON));
            _newOfferButton.Click();
            WaitForLoad();

            // Renseigner le name
            _newOfferNameInput = WaitForElementIsVisible(By.Name(CL_PAT_NEW_OFFER_NAME_INPUT), nameof(CL_PAT_NEW_OFFER_NAME_INPUT));
            _newOfferNameInput.SetValue(ControlType.TextBox, offerName);

            // Renseigner l'ordre
            _newOfferOrderInput = WaitForElementIsVisible(By.Name(CL_PAT_NEW_OFFER_ORDER_INPUT), nameof(CL_PAT_NEW_OFFER_ORDER_INPUT));
            _newOfferOrderInput.SetValue(ControlType.TextBox, offerOrder);
            
            //Save
            _newOfferSaveButton = WaitForElementIsVisible(By.Id(CL_PAT_NEW_OFFER_SAVE_BUTTON), nameof(CL_PAT_NEW_OFFER_SAVE_BUTTON));
            _newOfferSaveButton.Click();
            WaitForLoad();
        }

        //PACKAGE
        public bool isPackageExist(string packageName)
        {
            bool isExists;
            try
            {
                WaitForElementIsVisible(By.XPath(string.Format(CL_PAT_PACKAGE, packageName)));
                isExists = true;
            }
            catch
            {
                isExists = false;
            }
            return isExists;
        }

        public void AddNewPackage(string offerName, string offerOrder)
        {
            _newPackageButton = WaitForElementIsVisible(By.XPath(CL_PAT_NEW_PACKAGE_BUTTON), nameof(CL_PAT_NEW_PACKAGE_BUTTON));
            _newPackageButton.Click();
            WaitForLoad();

            // Renseigner le name
            _newPackageNameInput = WaitForElementIsVisible(By.Name(CL_PAT_NEW_PACKAGE_NAME_INPUT), nameof(CL_PAT_NEW_PACKAGE_NAME_INPUT));
            _newPackageNameInput.SetValue(ControlType.TextBox, offerName);

            // Renseigner l'ordre
            _newPackageOrderInput = WaitForElementIsVisible(By.Name(CL_PAT_NEW_PACKAGE_ORDER_INPUT), nameof(CL_PAT_NEW_PACKAGE_ORDER_INPUT));
            _newPackageOrderInput.SetValue(ControlType.TextBox, offerOrder);

            //Save
            _newPackageSaveButton = WaitForElementIsVisible(By.Id(CL_PAT_NEW_PACKAGE_SAVE_BUTTON), nameof(CL_PAT_NEW_PACKAGE_SAVE_BUTTON));
            _newPackageSaveButton.Click();
            WaitForLoad();
        }
    }
}
