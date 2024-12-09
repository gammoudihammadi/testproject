using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Customers.Service;
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

namespace Newrest.Winrest.FunctionalTests.PageObjects.Production.Dispatch
{
    public class PrevisionalQtyPage : PageBase
    {

        public PrevisionalQtyPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        //__________________________________Constantes_____________________________________

        private const string FIRST_ITEM = "//*[@id=\"dispatchTable\"]/tbody/tr[2]";
        private const string ALERT_ICON = "//*[@id=\"dispatchTable\"]/tbody/tr[2]/td[1]/span";
        private const string DELIVERY = "//*[@id=\"dispatchTable\"]/tbody/tr[2]/td[3]";
        private const string SERVICE = "//*[@id=\"dispatchTable\"]/tbody/tr[2]/td[4]";
        private const string SERVICE_LIST = "//*[@id=\"dispatchTable\"]/tbody/tr[*]/td[4]";
        private const string FIRST_SERVICE_BUTTON = "production-dispatch-service-detail-1-1";
        private const string QTY_INPUT = "/html/body/div[2]/div/div[2]/div[3]/div/div/div/table/tbody/tr[2]/td[12]/span[2]/input";

        private const string ALLCUSTOMER = "//*[@id=\"dispatchTable\"]/tbody/tr[{0}]/td[2]";
        private const string ALLCUSTOMER1 = "//*[@id=\"dispatchTable\"]/tbody/tr[*]/td[2]";
        private const string ALLDELIVERY = "//*[@id=\"dispatchTable\"]/tbody/tr[{0}]/td[3]";
        private const string ALLDELIVERY1 = "//*[@id='dispatchTable']/tbody/tr[contains(@class,'row-dispatch-line')]/td[3]";
        private const string ALLSERVICE = "//*[@id=\"dispatchTable\"]/tbody/tr[{0}]/td[4]";
        private const string ALLSERVICE1 = "//*[@id=\"dispatchTable\"]/tbody/tr[contains(@class,'row-dispatch-line')]/td[4]";
        
        private const string BTNVALIDATE = "0_ChkIsValidated";
        private const string SUNDAY_COLOR_VALIDATION = "//*[@id=\"thSun\"]/span[2]/span[1]";

        private const string WEEK_QTY_INPUT = "/html/body/div[2]/div/div[2]/div[3]/div/div/div/table/tbody/tr[2]/td[{0}]/span[2]/input";
        private const string DAY_COLOR_VALIDATION = "/html/body/div[2]/div/div[2]/div[3]/div/div/div/table/tbody/tr[1]/th[{0}]/span[2]/span[1]";       

        //__________________________________ Variables ______________________________________

        [FindsBy(How = How.XPath, Using = DELIVERY)]
        private IWebElement _delivery;

        [FindsBy(How = How.XPath, Using = SERVICE)]
        private IWebElement _service;

        [FindsBy(How = How.Id, Using = BTNVALIDATE)]
        private IWebElement _btnValidate;

        [FindsBy(How = How.XPath, Using = FIRST_ITEM)]
        private IWebElement _firstItem;

        [FindsBy(How = How.XPath, Using = QTY_INPUT)]
        private IWebElement _inputQty;

        [FindsBy(How = How.XPath, Using = SUNDAY_COLOR_VALIDATION)]
        private IWebElement _sundayValidation;

        //___________________________________ Methodes _________________________________________

        public string GetFirstDeliveryName()
        {
            _delivery = WaitForElementIsVisible(By.XPath(DELIVERY));
            return _delivery.Text;
        }

        public string GetFirstServiceName()
        {
            _service = WaitForElementIsVisible(By.XPath(SERVICE));
            return _service.Text;
        }

        public List<string> GetAllService()
        {

            List<string> names = new List<string>();
            var elements = _webDriver.FindElements(By.XPath(SERVICE_LIST));
            foreach (var element in elements)
            {
                names.Add(element.Text);
            }

            return names;
        }

        public ServicePricePage ClickFirstService()
        {
           var _fisrtService = WaitForElementIsVisible(By.Id(FIRST_SERVICE_BUTTON));
            _fisrtService.Click();
            var wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(60));
            wait.Until((driver) => driver.WindowHandles.Count > 1);
            _webDriver.SwitchTo().Window(_webDriver.WindowHandles[1]);
            return new ServicePricePage(_webDriver, _testContext);

        }


        public bool GetDeliveryName()
        {
            int tot;
            string beforeName = "";

            if (CheckTotalNumber() > 100)
            {
                tot = 100;
            }
            else
            {
                tot = CheckTotalNumber();
            }

            if (tot == 0)
                return false;

            //row-dispatch-line
            var elements = _webDriver.FindElements(By.XPath(ALLDELIVERY1));



            for (int i = 0; i < elements.Count(); i++)
            {

                try
                {
                    var name = elements[i].Text;
                    if (name.CompareTo(beforeName) < 0)
                        return false;

                    beforeName = name;

                }
                catch
                {
                    beforeName = "";
                }
            }
            return true;
        }




        public List<string> GetCustomer()
        {

            List<string> names = new List<string>();
            //int tot;

            //if (CheckTotalNumber() > 100)
            //{
            //    tot = 100;
            //}
            //else
            //{
            //    tot = CheckTotalNumber();
            //}


            var elements = _webDriver.FindElements(By.XPath(ALLCUSTOMER1));
            foreach(var element in elements)
            {
                names.Add(element.Text);
            }

            return names;
        }




        public bool VerifyCustomer(string value)
        {
            int tot;

            if (CheckTotalNumber() > 100)
            {
                tot = 100;
            }
            else
            {
                tot = CheckTotalNumber();
            }
            if (tot == 0)
                return false;

            IReadOnlyCollection<IWebElement> elements;
            elements = _webDriver.FindElements(By.XPath("//*[@id=\"dispatchTable\"]/tbody/tr[contains(@class,'row-dispatch-line')]/td[2]"));

            if (elements.Count == 0) return false;

            foreach (var element in elements)
            {
                if (element.Text != value)
                {
                    return false;
                }
            }

            return true;
        }

   


        public void CheckSortByServiceName()
        {
            // Attention, s'il y a plusieurs services pour un même couple delivery-customer,
            // ils sont groupés par delivery/customer par ordre croissant également
            // (côté test auto, ne prendre que le premier service dans ce cas là)
            List<string> service_avantTri = new List<string>();
            List<string> service_apresTri = new List<string>();

            var d = _webDriver.FindElements(By.XPath(ALLSERVICE1)).Select(x => x.Text).Distinct();
            foreach (var element in d)
            {
                var s = _webDriver.FindElements(By.XPath("//*[@id='dispatchTable']/tbody/tr[contains(@class,'row-dispatch-line')]/td[contains(text(),\"" + element+"\")]/../td[4]"));
                if (s.Count>=1)
                {
                    service_avantTri.Add(s[0].Text);
                    service_apresTri.Add(s[0].Text);
                }
            }

            service_apresTri.Sort();
            for (int i = 0; i < service_avantTri.Count; i++)
            {
                Assert.AreEqual(service_avantTri[i], service_apresTri[i]);
            }
        }


        public void ValidateFirstDispatch()
        {
            Actions action = new Actions(_webDriver);
            action.MoveToElement(_firstItem).Perform();

            _btnValidate = WaitForElementIsVisible(By.Id(BTNVALIDATE));
            _btnValidate.Click();
            WaitForLoad();
        }

        public void SetQuantity(string qty)
        {
            Actions action = new Actions(_webDriver);
            action.MoveToElement(_firstItem).Perform();

            _inputQty = WaitForElementIsVisible(By.XPath(QTY_INPUT));
            _inputQty.SetValue(ControlType.TextBox, qty);
            Thread.Sleep(2000); //long time to save value
        }

        public string GetQuantity()
        {
            Actions action = new Actions(_webDriver);
            action.MoveToElement(_firstItem).Perform();

            _inputQty = WaitForElementIsVisible(By.XPath(QTY_INPUT));
            return _inputQty.GetAttribute("value").Replace(" ", "");
        }

        public void UpdateQuantities(string newQty)
        {
            for (int i = 0; i < 7; i++)
            {
                Actions action = new Actions(_webDriver);
                action.MoveToElement(_firstItem).Perform();

                var element = _webDriver.FindElement(By.XPath(String.Format(WEEK_QTY_INPUT, i + 6)));
                element.SetValue(ControlType.TextBox, newQty);
                Thread.Sleep(1000);
            }
        }

        public bool IsQuantitiesUpdated(string qty)
        {
            bool isUpdated = true;

            for (int i = 0; i < 7; i++)
            {
                try
                {
                    Actions action = new Actions(_webDriver);
                    action.MoveToElement(_firstItem).Perform();

                    var element = _webDriver.FindElement(By.XPath(String.Format(WEEK_QTY_INPUT, i + 6)));

                    if (!element.GetAttribute("value").Equals(qty))
                    {
                        isUpdated = false;
                        break;
                    }
                }
                catch
                {
                    isUpdated = false;
                }
            }

            return isUpdated;
        }

        public bool IsValidatedByColorDay()
        {
            bool isOk = true;

            for (int i = 0; i < 7; i++)
            {
                try
                {
                    var element = _webDriver.FindElement(By.XPath(String.Format(DAY_COLOR_VALIDATION, i + 6)));

                    if (!element.GetAttribute("style").Equals("border-bottom-color: green;"))
                    {
                        isOk = false;
                        break;
                    }
                }
                catch
                {
                    isOk = false;
                }
            }

            return isOk;
        }

        public bool IsSundayValidatedByColorDay()
        {
            bool isOk = false;

            try
            {
                _sundayValidation = _webDriver.FindElement(By.XPath(String.Format(SUNDAY_COLOR_VALIDATION)));
                if (_sundayValidation.GetAttribute("style").Equals("border-bottom-color: green;"))
                    isOk = true;
            }          
            catch
            {
                isOk = false;
            }

            return isOk;
        }

        public bool CanUpdateQty()
        {
            bool isOk = false;

            for (int i = 0; i < 7; i++)
            {
                try
                {
                    Actions action = new Actions(_webDriver);
                    action.MoveToElement(_firstItem).Perform();

                    var element = _webDriver.FindElement(By.XPath(string.Format(WEEK_QTY_INPUT, i + 6)));
                    if(element.GetAttribute("disabled") == null)
                    {
                        isOk = true;
                        break;
                    }
                }
                catch
                {
                    isOk = false;
                }
            }

            return isOk;
        }

        public bool IsDispatchValidated()
        {
            _btnValidate = WaitForElementIsVisible(By.Id(BTNVALIDATE));
            if(_btnValidate.GetAttribute("data-isvalidated").Equals("true"))
            {
                return true;
            }
            else
            {
                return false;
            }

        }

        public bool IsNotLinked()
        {
            try
            {
                WaitForElementIsVisible(By.XPath(ALERT_ICON));
                return true;
            }
            catch
            {
                return false;
            }
        }

    }


}
