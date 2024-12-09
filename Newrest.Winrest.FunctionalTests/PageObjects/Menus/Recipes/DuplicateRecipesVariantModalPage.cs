using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.PageObjects;
using System;
using System.Threading;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Menus.Recipes
{
    public class DuplicateRecipesVariantModalPage : PageBase
    {
        // ________________________________________________ Constantes ___________________________________________

        private const string SITE_SOURCE = "duplicate-source-site";
        private const string SITE_DESTINATION = "duplicate-destination-site";
        private const string VARIANT_SOURCE = "duplicate-source-variant";
        private const string VARIANT_DESTINATION = "duplicate-destination-variant";

        private const string NEXT = "duplicate-recipe-next-btn";
        private const string UNFOLD = "//*[@id=\"duplicate-recipes-variant-container\"]/div[2]/div[1]/div[1]/div/div[1]/span";
        private const string FIRST_VARIANT_SUCCEDED = "//*[@id='content_succeeded']/div/table/tbody[2]/tr[1]/td[2]";
        private const string FIRST_VARIANT_ERRORS = "//*[@id='content_errors']/div/table/tbody[2]/tr[1]/td[2]";
        private const string DUPLICATE = "btn-modal-duplicate";
        private const string PROGRESS_BAR = "//*[@id=\"duplicate-recipes-variant-form\"]/div[1]";
        private const string CLOSE = "//*/button[text()='Close']";
        private const string BUTTON_CREATE_DUPLICATE = "btn-create-new-row";


        // ________________________________________________ Variables ____________________________________________

        [FindsBy(How = How.Id, Using = SITE_SOURCE)]
        private IWebElement _siteSource;

        [FindsBy(How = How.Id, Using = BUTTON_CREATE_DUPLICATE)]
        private IWebElement _btnCreateDuplicate;

        [FindsBy(How = How.Id, Using = SITE_DESTINATION)]
        private IWebElement _siteDestination;

        [FindsBy(How = How.Id, Using = VARIANT_SOURCE)]
        private IWebElement _variantSource;

        [FindsBy(How = How.Id, Using = VARIANT_DESTINATION)]
        private IWebElement _variantDestination;

        [FindsBy(How = How.Id, Using = NEXT)]
        private IWebElement _nextBtn;

        [FindsBy(How = How.XPath, Using = UNFOLD)]
        private IWebElement _unfold;

        [FindsBy(How = How.XPath, Using = FIRST_VARIANT_SUCCEDED)]
        private IWebElement _firstVariantSucceded;

        [FindsBy(How = How.XPath, Using = FIRST_VARIANT_ERRORS)]
        private IWebElement _firstVariantErrors;

        [FindsBy(How = How.Id, Using = DUPLICATE)]
        private IWebElement _duplicate;

        [FindsBy(How = How.XPath, Using = CLOSE)]
        private IWebElement _close;

        // ________________________________________________ Méthodes _____________________________________________

        public DuplicateRecipesVariantModalPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        public void FillField_DuplicateRecipesVariant(
            string siteSource, string siteDestination, string variantSource, string variantDestination)
        {
           
            //var addNewRecipeDestination = WaitForElementIsVisible(By.Id(BUTTON_CREATE_DUPLICATE));
            //addNewRecipeDestination.Click();
            //WaitForLoad();
            _siteSource = WaitForElementIsVisible(By.Id(SITE_SOURCE));
            _siteSource.SetValue(ControlType.DropDownList, siteSource);
            WaitForLoad();

            _siteDestination = WaitForElementIsVisible(By.Id(SITE_DESTINATION));
            _siteDestination.SetValue(ControlType.DropDownList, siteDestination);
            WaitForLoad();

            _variantSource = WaitForElementIsVisible(By.Id(VARIANT_SOURCE));
            _variantSource.SetValue(ControlType.DropDownList, variantSource);
            WaitForLoad();

            _variantDestination = WaitForElementIsVisible(By.Id(VARIANT_DESTINATION));
            _variantDestination.SetValue(ControlType.DropDownList, variantDestination);
            WaitForLoad();
            
            _nextBtn = WaitForElementToBeClickable(By.Id(NEXT));
            _nextBtn.Click();
            //progress bar
            WaitPageLoading();
            WaitForLoad();

            //_firstVariantErrors = WaitForElementIsVisible(By.XPath(FIRST_VARIANT_ERRORS));
            //var firstVariantText = _firstVariantErrors.Text;

            _duplicate = WaitForElementExists(By.Id(DUPLICATE));
            new Actions(_webDriver).MoveToElement(_duplicate).Perform();
            _duplicate = WaitForElementIsVisible(By.Id(DUPLICATE));
            _duplicate.Click();
            //var javaScriptExecutor = _webDriver as IJavaScriptExecutor;
            //javaScriptExecutor.ExecuteScript("StartOperation()");
            //javaScriptExecutor.ExecuteScript("arguments[0].scrollIntoView(true);", _duplicate);
            //_duplicate.Click();
            WaitPageLoading();
            WaitForLoad();

            DuplicationLoading();

            _close = WaitForElementIsVisible(By.XPath(CLOSE));
            _close.Click();
            WaitForLoad();

                //return firstVariantText;
        }

        public bool FillField_DuplicateRecipesVariantError(string siteDestination)
        {
            
                var addNewRecipeDestination = WaitForElementIsVisible(By.Id("btn-create-new-row"));
                addNewRecipeDestination.Click();
                WaitForLoad();

                _siteDestination = WaitForElementIsVisible(By.Id("duplicate-destination-site"));
           
            _siteDestination.SetValue(ControlType.DropDownList, siteDestination);
            WaitPageLoading();

            _nextBtn = _webDriver.FindElement(By.Id(NEXT));
            return _nextBtn.Displayed;

        }

        public void DuplicationLoading()
        {
            bool isLoading = true;
            int cpt = 1;

            while (isLoading && cpt <= 1000)
            {
                try
                {
                   _webDriver.FindElement(By.Id("RecipeProgress"));
                    
                    Thread.Sleep(5000);
                    cpt++;
                }
                catch
                {
                    isLoading = false;
                }
            }
            WaitForLoad();

            if (isLoading)
            {
                throw new Exception("La duplication n'a pas réussi à être chargée.");
            }
            
        }
    }
}
