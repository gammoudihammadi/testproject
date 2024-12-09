using DocumentFormat.OpenXml;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Customers.Service;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Flights.LoadingPlans
{
    public class LoadingPlansDetailsPage : PageBase
    {

        public LoadingPlansDetailsPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        // _____________________________________ Constantes _________________________________________

        private const string BACK_TO_LIST = "/html/body/div[2]/a/span[2]";
        private const string ADD_GUEST = "//*/a[@href='/Flights/LoadingPlan/AddGuestTypeDetails']";
        private const string YC_GUEST = "//*[@id=\"leg-guest-type4\"]/a";
        private const string BOB_GUEST = "//*[@id=\"leg-guest-type14\"]/a";
        private const string GUEST_TYPE = "GuestTypeId";
        private const string CREATE_GUEST = "/html/body/div[5]/div/div/div[2]/form/div[2]/button[2]";
        private const string UNFOLD_YC_GUEST = "leg-guest-type4";

        private const string ADD_SERVICE = "add-detail-link";
        private const string DELETE_SERVICE = "//*[@id=\"loading-plan-table\"]/tbody/tr[2]/td[11]/a";
        private const string SERVICE_NAME = "LoadingPlanDetails_{0}__ServiceId-selectized";
        private const string SERVICE_VALUE = "//*[@id=\"ServiceIdContainer_0\"]/div/div[1]/div";
        private const string FIRST_SERVICE = "//*[@id=\"ServiceIdContainer_{0}\"]/div/div[2]/div";

        private const string SAVE = "//*[@id=\"div-body\"]/div/div[1]/div/a/span";
        private const string CONFIRM = "confirm-loading-plan-create";
        private const string DETAILS_PAGE = "hrefTabContentDetails";

        private const string GENERAL_INFORMATION = "hrefTabContentGeneral";
        private const string FIRSTGUEST = "//*[@id=\"leg-guest-type2\"]/a";
        private const string CLICKVALUE = "//*[@id=\"loading-plan-table\"]/tbody/tr[2]/td[9]";
        private const string CONFIRMSAVE = "//*[@id=\"confirm-loading-plan-create\"]";
        private const string CLICKVALUE2 = "//*[@id=\"loading-plan-table\"]/tbody/tr[3]/td[9]";
        private const string FIRSTLINETYPE = "//*[@id=\"loading-plan-table\"]/tbody/tr[2]/td[6]/div[1]";
        private const string SECONDLINETYPE = "//*[@id=\"loading-plan-table\"]/tbody/tr[3]/td[6]/div[1]";




        //______________________________________Variables____________________________________________

        [FindsBy(How = How.XPath, Using = BACK_TO_LIST)]
        private IWebElement _backToList;




        [FindsBy(How = How.XPath, Using = CONFIRMSAVE)]
        private IWebElement _confirmsave;
        [FindsBy(How = How.XPath, Using = FIRSTLINETYPE)]
        private IWebElement _firstlinetype;
        [FindsBy(How = How.XPath, Using = SECONDLINETYPE)]
        private IWebElement _secondlinetype;
        [FindsBy(How = How.XPath, Using = CLICKVALUE)]
        private IWebElement _clickvalue;
        [FindsBy(How = How.XPath, Using = FIRSTGUEST)]
        private IWebElement _firstguest;

        [FindsBy(How = How.XPath, Using = CLICKVALUE2)]
        private IWebElement _clickvalue2;


        [FindsBy(How = How.XPath, Using = ADD_GUEST)]
        private IWebElement _addGuest;

        [FindsBy(How = How.XPath, Using = YC_GUEST)]
        private IWebElement _ycGuestLink;

        [FindsBy(How = How.XPath, Using = BOB_GUEST)]
        private IWebElement _bobGuestLink;

        [FindsBy(How = How.Id, Using = GUEST_TYPE)]
        private IWebElement _guestType;

        [FindsBy(How = How.XPath, Using = CREATE_GUEST)]
        private IWebElement _createGuest;

        [FindsBy(How = How.Id, Using = UNFOLD_YC_GUEST)]
        private IWebElement _unfoldYcGuest;

        [FindsBy(How = How.Id, Using = ADD_SERVICE)]
        private IWebElement _addService;

        [FindsBy(How = How.XPath, Using = DELETE_SERVICE)]
        private IWebElement _deleteService;

        [FindsBy(How = How.Id, Using = SERVICE_NAME)]
        private IWebElement _serviceName;

        [FindsBy(How = How.XPath, Using = SERVICE_VALUE)]
        private IWebElement _serviceValue;

        [FindsBy(How = How.XPath, Using = FIRST_SERVICE)]
        private IWebElement _firstService;

        [FindsBy(How = How.XPath, Using = SAVE)]
        private IWebElement _save;

        [FindsBy(How = How.Id, Using = CONFIRM)]
        private IWebElement _confirmSave;

        [FindsBy(How = How.Id, Using = GENERAL_INFORMATION)]
        private IWebElement _generalInformationPage;

        [FindsBy(How = How.Id, Using = DETAILS_PAGE)]
        private IWebElement _detailPage;

        //_________________________________________________UTILITAIRE_________________________________________________

        public LoadingPlansPage BackToList()
        {
            _backToList = WaitForElementIsVisible(By.XPath(BACK_TO_LIST), nameof(BACK_TO_LIST));
            _backToList.Click();
            WaitPageLoading();
            WaitForLoad();

            return new LoadingPlansPage(_webDriver, _testContext);
        }

        public void ClickCreateGuestBtn()
        {
            _createGuest = WaitForElementIsVisible(By.XPath(CREATE_GUEST));
            _createGuest.Click();
            WaitPageLoading();
            WaitForLoad();
        }

        //_________________________________________________GUEST_________________________________________________

        public void ClickAddGuestBtn()
        {
            _addGuest = WaitForElementIsVisible(By.XPath(ADD_GUEST));
            _addGuest.Click();
            WaitLoading();
        }
        public void ClickGuestBtn()
        {
            _ycGuestLink = WaitForElementExists(By.XPath(YC_GUEST));
            _ycGuestLink.Click();

            WaitPageLoading();
            WaitLoading();
        }

        public void ClickGuestBOBBtn()
        {
            _bobGuestLink = WaitForElementExists(By.XPath(BOB_GUEST));
            _bobGuestLink.Click();

            WaitLoading();
        }


        public void ClickGuestBtnBOB(string guest)
        {
            _ycGuestLink = WaitForElementExists(By.XPath("//*[contains(text(), '" + guest + "')]"));
            _ycGuestLink.Click();

            WaitLoading();
        }



        public void SelectGuest(string guestName)
        {
            _guestType = WaitForElementIsVisible(By.Id(GUEST_TYPE));
            var dropdown = new SelectElement(_guestType);
            dropdown.SelectByText(guestName);
            WaitForLoad();
        }

        public Boolean IsGuestAdded()
        {
            bool valueBool = false;

            _unfoldYcGuest = _webDriver.FindElement(By.Id(UNFOLD_YC_GUEST));
            if (_unfoldYcGuest.GetAttribute("class") == "leg-guest-type-name label-dark-gray")
                valueBool = true;

            return valueBool;
        }

        //_________________________________________________SERVICE_________________________________________________

        public void AddServiceBtn()
        {
            _addService = WaitForElementIsVisible(By.Id(ADD_SERVICE));
            _addService.Click();
            //WaitPageLoading();
            WaitForLoad();
        }

        public void AddNewService(string serviceName)
        {
            // Définition du Service Name);
            for (int i = 0; i < 5; i++) // limite à l'ajout de 5 services
            {
                if (isElementVisible(By.Id(string.Format(SERVICE_NAME, i))))
                {
                    _serviceName = WaitForElementIsVisible(By.Id(string.Format(SERVICE_NAME, i)));
                    _serviceName.ClearElement();
                    _serviceName.SetValue(ControlType.TextBox, serviceName);
                    _serviceName.SendKeys(Keys.Tab);
                    WaitForLoad();

                    Save();

                    break;
                }
            }

            if (isElementVisible(By.Id("dataAlertCancel")))
            {
                var button = WaitForElementIsVisible(By.Id("dataAlertCancel"));
                button.Click();
                WaitForLoad();
            }
            WaitPageLoading();
            WaitForLoad();
        }

        public void AddNewServiceWithoutsave(string serviceName)
        {
            // Définition du Service Name);
            for (int i = 0; i < 5; i++) // limite à l'ajout de 5 services
            {
                if (isElementVisible(By.Id(string.Format(SERVICE_NAME, i))))
                {
                    _serviceName = WaitForElementIsVisible(By.Id(string.Format(SERVICE_NAME, i)));
                    _serviceName.ClearElement();
                    _serviceName.SetValue(ControlType.TextBox, serviceName);
                    _serviceName.SendKeys(Keys.Tab);
                    WaitForLoad();


                    break;


                }
            }

            if (isElementVisible(By.Id("dataAlertCancel")))
            {
                var button = WaitForElementIsVisible(By.Id("dataAlertCancel"));
                button.Click();
                WaitForLoad();
            }
        }



        public void AddNewServiceWithValueAndName(string value, string servicename)
        {
            // Définition du Service Name);
            for (int i = 0; i < 5; i++) // limite à l'ajout de 5 services
            {
                if (isElementVisible(By.Id(string.Format(SERVICE_NAME, i))))
                {
                    _serviceName = WaitForElementIsVisible(By.Id(string.Format(SERVICE_NAME, i)));
                    _serviceName.ClearElement();
                    _serviceName.SetValue(ControlType.TextBox, servicename);
                    _serviceName.SendKeys(Keys.Tab);

                    _serviceValue = WaitForElementIsVisible(By.XPath(string.Format(CLICKVALUE2, i)));
                    //_serviceName.ClearElement();
                    _serviceValue.SetValue(ControlType.TextBox, value);
                    WaitForLoad();

                    Save();

                    break;


                }
            }

            if (isElementVisible(By.Id("dataAlertCancel")))
            {
                var button = WaitForElementIsVisible(By.Id("dataAlertCancel"));
                button.Click();
                WaitForLoad();
            }
        }
        public bool IsServiceAdded(string serviceName)
        {
            WaitForLoad();
            if (!isElementExists(By.XPath(string.Format("//*/tr[contains(@class,'editline')]/td[3]/div[contains(text(),'{0}')]", serviceName))))
                return false;
            return true;
        }

        public Boolean IsServiceDeleted()
        {
            bool valueBool;

            try
            {
                _ycGuestLink = WaitForElementExists(By.XPath(YC_GUEST));
                var periodLocationDisplayed = _ycGuestLink.Displayed;

                if (!periodLocationDisplayed)
                    valueBool = true;
                else
                    valueBool = false;
            }
            catch
            {
                valueBool = false;
            }
            return valueBool;
        }

        public string GetServiceName()
        {
            _serviceValue = WaitForElementExists(By.XPath(SERVICE_VALUE));
            return _serviceValue.GetAttribute("innerText");
        }

        public void DeleteService()
        {
            // Définition du Service Name
            _deleteService = WaitForElementIsVisible(By.XPath(DELETE_SERVICE));
            _deleteService.Click();
            WaitPageLoading();
            WaitForLoad();

            // enregistre et va sur la page general info
            Save();
        }



        public void SaveConfirm()
        {
            // Définition du Service Name
            _deleteService = WaitForElementIsVisible(By.XPath(CONFIRMSAVE));
            _deleteService.Click();
            WaitForLoad();

        }

        public void Save()
        {
            _save = WaitForElementIsVisible(By.XPath(SAVE));
            _save.Click();
            WaitPageLoading();
            WaitForLoad();

            // Confirmation sauvegarde
            _confirmSave = WaitForElementIsVisible(By.Id(CONFIRM));
            _confirmSave.Click();
            // Récupérer le nouveau nom du service et valider que le changement est bien pris en compte
            // barre de progression non buzy (c'est pas un rond qui tourne)
            WaitPageLoading();
            WaitForLoad();
        }

        public LoadingPlansGeneralInformationsPage ClickOnGeneralInformation()
        {
            _generalInformationPage = WaitForElementIsVisible(By.Id(GENERAL_INFORMATION));
            _generalInformationPage.Click();
            WaitPageLoading();
            WaitLoading();

            return new LoadingPlansGeneralInformationsPage(_webDriver, _testContext);
        }

        public void WaiForLoad()
        {
            base.WaitForLoad();
        }
        public LoadingPlansDetailsPage ClickOnDetailsPage()
        {
            _detailPage = WaitForElementIsVisible(By.Id(DETAILS_PAGE));
            // FIXME page vide
            _detailPage.Click();
            _detailPage.Click();
            // changement d'onglet
            WaitPageLoading();
            WaitLoading();
            // chargement du tableau dans l'onglet
            WaitPageLoading();
            WaitLoading();

            return new LoadingPlansDetailsPage(_webDriver, _testContext);
        }

        public LoadingPlansDetailsPage GoToLoadingPlanDetailPage()
        {
            var detailPage = WaitForElementToBeClickable(By.Id("hrefTabContentDetails"));
            detailPage.Click();

            return new LoadingPlansDetailsPage(_webDriver, _testContext);
        }
        public void ClickFirstGuest()
        {
            var _guestclick = WaitForElementIsVisible(By.XPath(FIRSTGUEST));
            _guestclick.Click();
            WaitForLoad();
        }


        public void Editvalue(string value)
        {


            _serviceName = WaitForElementIsVisible(By.XPath(string.Format(CLICKVALUE)));
            _serviceName.ClearElement();
            _serviceName.SetValue(ControlType.TextBox, value);
            WaitForLoad();

            Save();
        }

        public string GetFirstLineType()
        {
            _firstlinetype = WaitForElementExists(By.XPath(FIRSTLINETYPE));
            return _firstlinetype.GetAttribute("innerText");
        }
        public string GetSecondLineType()
        {
            _secondlinetype = WaitForElementExists(By.XPath(SECONDLINETYPE));
            return _secondlinetype.GetAttribute("innerText");
        }

        public bool GetServicesNames(string serviceName, string newServiceName)
        {
            var servicesNames = new List<string>();
            var rows = _webDriver.FindElements(By.XPath("//*[@id=\"loading-plan-table\"]/tbody/tr[*]"));


            for (int i = 1; i < rows.Count; i++)
            {
                Console.WriteLine(rows[i].Text);
                string[] lines = rows[i].Text.Split(new[] { "\r\n" }, StringSplitOptions.None);
                servicesNames.Add(lines[0]);

            }
            var serviceExist = servicesNames.Any(s => s.Contains(serviceName));
            var newServiceExist = servicesNames.Any(s => s.Contains(newServiceName));

            return serviceExist && newServiceExist;


        }

        public int GetWidthIconDelete()
        {
            var _sizeIconDelete = WaitForElementExists(By.XPath("//*[@id=\"loading-plan-table\"]/tbody/tr[2]/td[11]/a"));
            return _sizeIconDelete.Size.Width;
        }
        public int GetHeightIconDelete()
        {
            var _sizeIconDelete = WaitForElementExists(By.XPath("//*[@id=\"loading-plan-table\"]/tbody/tr[2]/td[11]/a"));
            return _sizeIconDelete.Size.Height;
        }

        public int GetWidthTDDelete()
        {
            var _sizeTdDelete = WaitForElementExists(By.XPath("//*[@id=\"loading-plan-table\"]/tbody/tr[2]/td[11]"));
            return _sizeTdDelete.Size.Width;
        }

        public int GetHeightTDDelete()
        {
            var _sizeTdDelete = WaitForElementExists(By.XPath("//*[@id=\"loading-plan-table\"]/tbody/tr[2]/td[11]"));
            return _sizeTdDelete.Size.Height;
        }

        public bool IsIconWithinBounds(int widthIcon, int heightIcon, int widthTd, int heightTd)
        {
            const int standardWidth = 20;
            const int standardHeight = 20;

            // Check if widthIcon is between 20 and widthTd, and if heightIcon is between 20 and heightTd
            bool isWidthValid = (standardWidth < widthIcon) && (widthIcon < widthTd);
            bool isHeightValid = (standardHeight < heightIcon) && (heightIcon < heightTd);

            // Return true if both conditions are true, otherwise return false
            return isWidthValid && isHeightValid;
        }
    }
}
         
   
 