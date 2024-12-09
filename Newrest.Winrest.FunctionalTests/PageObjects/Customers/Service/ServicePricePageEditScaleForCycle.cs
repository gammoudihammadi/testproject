using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.PageObjects;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static Newrest.Winrest.FunctionalTests.PageObjects.Shared.PageBase;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Customers.Service
{
    public class ServicePricePageEditScaleForCycle : PageBase
    {
        public ServicePricePageEditScaleForCycle(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
            
        }

        private const string FIRST_INPUT_PAX = "//*[@id=\"cycleScaleForm\"]/div[2]/table/tbody/tr[2]/td[1]/input[3]";
        private const string FIRST_INPUT_QTE = "//*[@id=\"cycleScaleForm\"]/div[2]/table/tbody/tr[2]/td[2]/input";
        private const string ADDSCALEMODE = "//*[@id=\"cycleScaleForm\"]/div[2]/a";
        private const string BTN_SAVE = "//html/body/div[5]/div/div/div/div/form/div[3]/button[2]";
        private const string BTN_CLOSE = "//*[@id=\"cycleScaleForm\"]/div[3]/button[1]";

        [FindsBy(How = How.XPath, Using = BTN_CLOSE)]
        private IWebElement _btn_close;

        [FindsBy(How = How.XPath, Using = ADDSCALEMODE)]
        private IWebElement _addscalemode;

        [FindsBy(How = How.Id, Using = BTN_SAVE)]
        private IWebElement _btn_save;

        [FindsBy(How = How.XPath, Using = FIRST_INPUT_PAX)]
        private IWebElement _first_input_pax;

        [FindsBy(How = How.XPath, Using = FIRST_INPUT_QTE)]
        private IWebElement _second_input_qte;

        public void SetValueForFirstScaleMode(string value_nb_pax, string value_quantity)
        {
            _first_input_pax = WaitForElementIsVisible(By.XPath(FIRST_INPUT_PAX));
            _first_input_pax.SetValue(ControlType.TextBox, value_nb_pax);
            WaitForLoad();
            _second_input_qte = WaitForElementIsVisible(By.XPath(FIRST_INPUT_QTE));
            _second_input_qte.SetValue(ControlType.TextBox, value_quantity);
            WaitForLoad();
        }
        public void AddManyScaleMode(int nbligne, List<int> tab_value,int coef)
        {
            int indicerow = 3;
            int indicetab = 2;
            for (int i = 0; i < nbligne; i++)
            {
                _addscalemode = WaitForElementIsVisible(By.XPath(ADDSCALEMODE));
                _addscalemode.Click();
                WaitForLoad();

                var nb_pax = WaitForElementIsVisible(By.XPath("//*[@id=\"cycleScaleForm\"]/div[2]/table/tbody/tr["+indicerow+"]/td[1]/input[3]"));
                nb_pax.SetValue(ControlType.TextBox, (tab_value[indicetab]*coef).ToString());
                WaitForLoad();

                var quantity = WaitForElementIsVisible(By.XPath("//*[@id=\"cycleScaleForm\"]/div[2]/table/tbody/tr["+indicerow+"]/td[2]/input"));
                quantity.SetValue(ControlType.TextBox, (tab_value[indicetab+1]*coef).ToString());
                WaitForLoad();
                indicerow++;
                indicetab = indicetab + 2;
            }
        }
        public ServiceCreatePriceModalPage Save_Scale_Mode()
        {
            _btn_save = WaitForElementIsVisible(By.XPath(BTN_SAVE));
            _btn_save.Click();
            WaitForLoad();
            return new ServiceCreatePriceModalPage(_webDriver, _testContext);
        }
        public void GetAllValueScaleMode( ref List<string> tab_value)
        {
            WaitForLoad();
            var nb_row_in_the_table = _webDriver.FindElements(By.XPath("//*[@id=\"cycleScaleForm\"]/div[2]/table/tbody/tr[*]"));
            if (nb_row_in_the_table.Count > 1)
            {
             int indicerow = 2;
             for (int j = 1; j < nb_row_in_the_table.Count;j++)
              {
                var nb_pax = WaitForElementIsVisible(By.XPath("//*[@id=\"cycleScaleForm\"]/div[2]/table/tbody/tr["+indicerow+"]/td[1]/input[3]"));
                tab_value.Add(nb_pax.GetAttribute("value"));
                var quantity = WaitForElementIsVisible(By.XPath("//*[@id=\"cycleScaleForm\"]/div[2]/table/tbody/tr["+indicerow+"]/td[2]/input"));
                tab_value.Add(quantity.GetAttribute("value"));
                indicerow++;
                WaitForLoad();
              }
            }           
        }
        public void CloseEditModeCycle()
        {
            _btn_close = WaitForElementIsVisible(By.XPath(BTN_CLOSE));
            _btn_close.Click();
            WaitForLoad();
        }
    }
}
