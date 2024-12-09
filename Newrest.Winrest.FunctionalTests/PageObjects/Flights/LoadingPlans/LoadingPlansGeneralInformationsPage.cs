using DocumentFormat.OpenXml.Presentation;
using Limilabs.Client.IMAP;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Flights.LoadingPlans
{
    public class LoadingPlansGeneralInformationsPage : PageBase
    {

        public LoadingPlansGeneralInformationsPage(IWebDriver webDriver, TestContext testContext) : base(webDriver,
            testContext)
        {
        }

        //_________________________________________Constantes___________________________________________________

        private const string ACTION_BTN = "//*[@id=\"div-body\"]/div/div[1]/div/div/button";
        private const string BACK_TO_LIST = "/html/body/div[2]/a/span[2]";
        private const string LINKED_LP = "hrefTabContentLinkedLoadingPlans";
        private const string FLIGHTS = "hrefTabContentFlights";

        private const string DUPLICATE = "duplicate";
        private const string DELETE = "delete";
        private const string SPLIT = "split";
        private const string HISTORY = "history";
        private const string PROPAGATE = "propagate";

        private const string END_DATE = "popup-end-date-picker";
        private const string START_DATE = "popup-start-date-picker";
        private const string SAVE = "//*[@id=\"div-body\"]/div/div[1]/div/a/span";
        private const string CONFIRM_SAVE = "confirm-loading-plan-create";
        private const string CONFIRM_DELETE = "//*[@id=\"form-delete-loading-plan\"]/div/div[3]/button[2]";
        private const string LOADING_PLAN_TITLE_NAME = "/html/body/div[3]/div/div[1]/h1";
        private const string LOADING_PLAN_NAME = "Name";
        private const string SAVE_BTN = "//*[@id=\"div-body\"]/div/div[1]/div/a";
        private const string SAVE_CONFIRM = "confirm-loading-plan-create";
        private const string TYPE = "cbSortBy";
        private const string LP_CART = "dropdown-lpcart";
        //__________________________________________Variables___________________________________________________

        [FindsBy(How = How.XPath, Using = BACK_TO_LIST)]
        private IWebElement _backToList;

        [FindsBy(How = How.Id, Using = LINKED_LP)]
        private IWebElement _linkedLP;
        [FindsBy(How = How.Id, Using = TYPE)]
        private IWebElement _cbSortBy;

        [FindsBy(How = How.Id, Using = FLIGHTS)]
        private IWebElement _flightBtn;

        [FindsBy(How = How.XPath, Using = ACTION_BTN)]
        private IWebElement _actionBtn;

        [FindsBy(How = How.Id, Using = END_DATE)]
        private IWebElement _endDate;

        [FindsBy(How = How.Id, Using = START_DATE)]
        private IWebElement _startDate;

        [FindsBy(How = How.XPath, Using = SAVE)]
        private IWebElement _save;

        [FindsBy(How = How.Id, Using = CONFIRM_SAVE)]
        private IWebElement _confirmSave;

        [FindsBy(How = How.Id, Using = DUPLICATE)]
        private IWebElement _duplicateLoadingPlan;

        [FindsBy(How = How.Id, Using = SPLIT)]
        private IWebElement _splitLoadingPlan;

        [FindsBy(How = How.Id, Using = DELETE)]
        private IWebElement _deleteLP;

        [FindsBy(How = How.XPath, Using = CONFIRM_DELETE)]
        private IWebElement _confirmDelete;

        [FindsBy(How = How.Id, Using = HISTORY)]
        private IWebElement _historyLoadingPlan;

        [FindsBy(How = How.Id, Using = PROPAGATE)]
        private IWebElement _propagateLoadingPlan;

        [FindsBy(How = How.Id, Using = LOADING_PLAN_NAME)]
        private IWebElement _loadingPlanName;
        [FindsBy(How = How.XPath, Using = SAVE_BTN)]
        private IWebElement _saveBtn;
        [FindsBy(How = How.Id, Using = SAVE_CONFIRM)]
        private IWebElement save_confirm;

        [FindsBy(How = How.Id, Using = LP_CART)]
        private IWebElement _lpcart_Combox;

        //_______________________________________Pages___________________________________________________________

        public LoadingPlansDuplicationModalPage DuplicateLoadingPlanPage()
        {
            ShowActionMenu();

            _duplicateLoadingPlan = WaitForElementIsVisible(By.Id(DUPLICATE));
            _duplicateLoadingPlan.Click();

            return new LoadingPlansDuplicationModalPage(_webDriver, _testContext);
        }

        public LoadingPlansSplitModalPage SplitLoadingPlanPage()
        {
            ShowActionMenu();

            _splitLoadingPlan = WaitForElementIsVisible(By.Id(SPLIT));
            _splitLoadingPlan.Click();

            return new LoadingPlansSplitModalPage(_webDriver, _testContext);
        }

        public LoadingPlansFlightPage ClickOnFlightBtn()
        {
            _flightBtn = WaitForElementIsVisible(By.Id(FLIGHTS));
            _flightBtn.Click();
            WaitForLoad();

            return new LoadingPlansFlightPage(_webDriver, _testContext);
        }

        public LinkedLoadingPlansPage RedirectToLinkedLoadingPlansPage()
        {
            _linkedLP = WaitForElementIsVisible(By.Id(LINKED_LP));
            _linkedLP.Click();
            WaitForLoad();

            return new LinkedLoadingPlansPage(_webDriver, _testContext);
        }

        public LoadingPlansDetailsPage GoToLoadingPlanDetailPage()
        {
            var detailPage = WaitForElementToBeClickable(By.Id("hrefTabContentDetails"));
            detailPage.Click();
            WaitPageLoading();
            WaitForLoad();

            return new LoadingPlansDetailsPage(_webDriver, _testContext);
        }

        //________________________________________Utilitaires_______________________________________________

        public void ShowActionMenu()
        {
            Actions actions = new Actions(_webDriver);

            _actionBtn = WaitForElementIsVisible(By.XPath(ACTION_BTN));
            actions.MoveToElement(_actionBtn).Perform();
        }

        public LoadingPlansPage BackToList()
        {
            WaitLoading();
            _backToList = WaitForElementToBeClickable(By.XPath(BACK_TO_LIST));
            _backToList.Click();
            WaitPageLoading();
            return new LoadingPlansPage(_webDriver, _testContext);
        }

        public void EditLoadingPlanInformations(DateTime endDate, DateTime? startDate = null)
        {
         

            if (startDate != null)
            {
                _startDate = WaitForElementIsVisible(By.Id(START_DATE));
                _startDate.SetValue(ControlType.DateTime, startDate);
                _startDate.SendKeys(Keys.Enter);
                WaitForLoad();
            }
            // On modifie le champ End Date
            _endDate = WaitForElementIsVisible(By.Id(END_DATE));
            _endDate.SetValue(ControlType.DateTime, endDate);
            _endDate.SendKeys(Keys.Enter);
            WaitForLoad();

            // Sauvegarde
            _save = WaitForElementIsVisible(By.XPath(SAVE));
            _save.Click();
            WaitForLoad();

            _confirmSave = WaitForElementIsVisible(By.Id(CONFIRM_SAVE));
            _confirmSave.Click();
            // Temps de fermeture de la fenêtre de confirmation
            WaitPageLoading();
            WaitPageLoading();
        }

        public void SelectLpCart(string lpCart)
        {

            // Définition du site
            _lpcart_Combox = WaitForElementIsVisible(By.Id(LP_CART));
            _lpcart_Combox.SetValue(ControlType.DropDownList, lpCart);           
            WaitForLoad();

            // Sauvegarde
            _save = WaitForElementIsVisible(By.XPath(SAVE));
            _save.Click();
            WaitPageLoading();
            WaitForLoad();

            _confirmSave = WaitForElementIsVisible(By.Id(CONFIRM_SAVE));
            _confirmSave.Click();
            WaitPageLoading();
            WaitForLoad();

            // Temps de fermeture de la fenêtre de confirmation (la barre de progression n'est pas buzy)
            Thread.Sleep(4000);
        }

        public void ClickLoadingPlanLPCartEditLabels(string position)
        {
            var editLabel = WaitForElementIsVisible(By.Id("btn-lpcart-editlabels"));
            editLabel.Click();

            WaitForLoad();


            // Définition de la position
            var positionInput = WaitForElementIsVisible(By.Id("Labels_0__Position"));
            positionInput.SetValue(ControlType.TextBox, position);


            // Définition de la qty
            var qty = WaitForElementIsVisible(By.Id("Labels_0__Quantity"));
            qty.SetValue(ControlType.TextBox, "10");

            // Sauvegarde
            _save = WaitForElementIsVisible(By.XPath("//*[@id=\"form-savelabels\"]/div[2]/button[2]"));
            _save.Click();
            WaitForLoad();


            // Temps de fermeture de la fenêtre de confirmation
            Thread.Sleep(1000);
        }

        public void DeleteLoadingPlan()
        {
            ShowActionMenu();

            _deleteLP = WaitForElementIsVisible(By.Id(DELETE));
            _deleteLP.Click();
            WaitLoading();
            _confirmDelete = WaitForElementIsVisible(By.XPath(CONFIRM_DELETE));
            _confirmDelete.Click();
            WaitForLoad();
        }

        public string GetLpCartName()
        {
            var newSite = new SelectElement(_webDriver.FindElement(By.Id("dropdown-lpcart")));
            return newSite.AllSelectedOptions.FirstOrDefault()?.Text;
        }

        public LoadingPlansHistoryModalPage CheckLoadingPlanHistory()
        {
            ShowActionMenu();

            _historyLoadingPlan = WaitForElementIsVisible(By.Id(HISTORY));
            _historyLoadingPlan.Click();
            return new LoadingPlansHistoryModalPage(_webDriver, _testContext);
        }

        public LoadingPlansPropagateModalPage PropagateLoadingPlan()
        {
            ShowActionMenu();

            _propagateLoadingPlan = WaitForElementIsVisible(By.Id(PROPAGATE));
            _propagateLoadingPlan.Click();
            return new LoadingPlansPropagateModalPage(_webDriver, _testContext);
        }

        public string GetLoadingPlanTitleName()
        {
            var title = WaitForElementExists(By.XPath(LOADING_PLAN_TITLE_NAME));
            return title.Text;
        }
        public bool isLoadingPlanDetailPage()
        {
            if (GetLoadingPlanTitleName().Contains("LOADING PLAN :"))
                return true;

            return false;
        }
        public string GetFlightLoadingPlanName()
        {
            _loadingPlanName = WaitForElementExists(By.Id(LOADING_PLAN_NAME));
            return _loadingPlanName.GetAttribute("value");
        }
        public LoadingPlansGeneralInformationsPage EditLoadingPlanInformationsType()
        {
            _cbSortBy = WaitForElementIsVisible(By.Id(TYPE));
            var selectedOption = _cbSortBy.FindElement(By.CssSelector("option[selected='selected']"));
            string selectedValue = selectedOption.GetAttribute("value");
            var options = new List<string> { "Meal", "Loading", "Other", "BuyOnBoard" };
            options.Remove(selectedValue);
            var newOptionValue = options.First();
            // Change the value
            _cbSortBy.SetValue(ControlType.DropDownList, newOptionValue);
            _cbSortBy.SendKeys(Keys.Enter);
            WaitForLoad();
            return new LoadingPlansGeneralInformationsPage(_webDriver, _testContext);
        }
        public string GetTypeCheck()
        {
            _cbSortBy = WaitForElementIsVisible(By.Id(TYPE));
            var selectedOption = _cbSortBy.FindElement(By.CssSelector("option[selected='selected']"));
            return selectedOption.GetAttribute("value");
        }

        public void SaveConfirm()
        {
            _saveBtn = WaitForElementIsVisible(By.XPath(SAVE_BTN));
            _saveBtn.Click();
            WaitForLoad();
            save_confirm = WaitForElementIsVisible(By.Id(SAVE_CONFIRM));
            save_confirm.Click();
            WaitForLoad();

        }
        public string GetLoadingPlanInfo()
        {
            var element = WaitForElementIsVisible(By.XPath("//*[@id='div-body']/div/div[1]/h1"));
            string fullText = element.Text;
            return fullText.Replace("LOADING PLAN : ", "").Trim();
        }




    }
}
