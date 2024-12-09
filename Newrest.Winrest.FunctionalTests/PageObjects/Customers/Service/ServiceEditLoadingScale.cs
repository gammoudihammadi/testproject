using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Customers.Service
{
    public class ServiceEditLoadingScale : PageBase
    {
        public ServiceEditLoadingScale(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {


        }

        private const string ADDSCALEMODE = "//html/body/div[5]/div/div/div/div/form/div[3]/a";
        private const string BTN_SAVE = "//html/body/div[5]/div/div/div/div/form/div[4]/button[2]";

        [FindsBy(How = How.XPath, Using = ADDSCALEMODE)]
        private IWebElement _addscalemode;

        [FindsBy(How = How.Id, Using = BTN_SAVE)]
        private IWebElement _btn_save;

        public void SetValueForFirstScaleMode(string value_nb_pax, string value_quantity)
        {
            var nb_pax = WaitForElementIsVisible(By.XPath("//*[@id=\"default-scales-table\"]/tbody/tr[2]/td[1]/input[3]"));
            nb_pax.SetValue(ControlType.TextBox, value_nb_pax);
            WaitForLoad();
            var quantity = WaitForElementIsVisible(By.XPath("//*[@id=\"default-scales-table\"]/tbody/tr[2]/td[2]/input[1]"));
            quantity.SetValue(ControlType.TextBox, value_quantity);
            WaitForLoad();
        }
        public void AddManyScaleMode(int nbligne, List<string> tab_value )
        {
            int indicerow = 3;
            int indicetab = 2;
            for (int i = 0; i < nbligne; i++)
            {
                _addscalemode = WaitForElementIsVisible(By.XPath(ADDSCALEMODE));
                _addscalemode.Click();
                WaitForLoad();

                var nb_pax = WaitForElementIsVisible(By.XPath("//*[@id=\"default-scales-table\"]/tbody/tr["+ indicerow + "]/td[1]/input[3]"));
                nb_pax.SetValue(ControlType.TextBox, tab_value[indicetab]);
                WaitForLoad();

                var quantity = WaitForElementIsVisible(By.XPath("//*[@id=\"default-scales-table\"]/tbody/tr[" + indicerow + "]/td[2]/input[1]"));
                quantity.SetValue(ControlType.TextBox, tab_value[indicetab + 1]);
                WaitForLoad();

                indicerow++;
                indicetab = indicetab+2;
            }
        }

        public ServiceCreatePriceModalPage Save_Scale_Mode()
        {
            _btn_save = WaitForElementIsVisible(By.XPath(BTN_SAVE));
            _btn_save.Click();
            WaitForLoad();
            return new ServiceCreatePriceModalPage(_webDriver, _testContext);
        }

        public List<string> GetAllValueScaleMode()
        {
            List<string> tab_value = new List<string>();
            var nb_row = _webDriver.FindElements(By.XPath("//*[@id=\"default-scales-table\"]/tbody/tr[*]"));

            if (nb_row.Count > 0)
            {
                int indicerow = 2;
            
                for (int i = 1; i < nb_row.Count; i++)
                {
                    //*[@id="default-scales-table"]/tbody/tr[*]
                    var nb_pax = WaitForElementIsVisible(By.XPath("//*[@id=\"default-scales-table\"]/tbody/tr[" + indicerow + "]/td[1]/input[3]"));
                    tab_value.Add(nb_pax.GetAttribute("value"));
                    var quantity = WaitForElementIsVisible(By.XPath("//*[@id=\"default-scales-table\"]/tbody/tr[" + indicerow + "]/td[2]/input[1]"));
                    tab_value.Add(quantity.GetAttribute("value"));
                    indicerow++;                 
                    WaitForLoad();
                }
            }

            return tab_value;
        }
    }
}
