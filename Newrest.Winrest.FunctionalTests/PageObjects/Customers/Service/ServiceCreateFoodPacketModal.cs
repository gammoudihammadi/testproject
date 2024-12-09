using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Drawing.Charts;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Customers.Service
{
    public class ServiceCreateFoodPacketModal : PageBase
    {
        private const string SITE = "SiteId";
        private const string CUSTOMER = "//*[@id=\"dropdown-customer-selectized\"]";
        private const string FOOD_PACKET = "//*[@id=\"div-type-ahead-packets\"]/span/input[2]";
        private const string SUBMIT = "//*/button[text()='Save']";
        private const string PACKET_ITEM= "//*[@id=\"div-type-ahead-packets\"]/span/div/div/div[1]";
        private const string FIRST_ROW = "/html/body/div[3]/div/div/div[2]/div/div/div[2]/div/div";
        private const string DELETE_PACKET_ICON = "/html/body/div[5]/div/div/div/div/form/div[2]/div[4]/div/div[1]/a";

        public ServiceCreateFoodPacketModal(IWebDriver _webDriver, TestContext _testContext) : base(_webDriver, _testContext)
        {
        }

        public void FillServiceFoodPacket(string site, string foodPacket, string customer=null)
        {
            DropdownListSelectById(SITE, site);
            //DropdownListSelectById(CUSTOMER, customer);
            if (customer != null )
            {
                var s = WaitForElementIsVisible(By.XPath(CUSTOMER));
                ((IJavaScriptExecutor)_webDriver).ExecuteScript("arguments[0].value = '';", s);
                Thread.Sleep(1000);
                s.SetValue(ControlType.TextBox, customer);
                s.SendKeys(Keys.Enter);
                WaitForLoad();
            }
            var f = WaitForElementIsVisible(By.XPath(FOOD_PACKET));
            f.SetValue(ControlType.TextBox, "ATLANTIC");
            WaitPageLoading();
            WaitForLoad();
            if(isElementExists(By.XPath(PACKET_ITEM)))
            {
                var selectFirstfoodpacket = WaitForElementIsVisible(By.XPath(PACKET_ITEM));
                selectFirstfoodpacket.Click();
                WaitForLoad();
            }
            //WaitForLoad();
            var submit = WaitForElementIsVisible(By.XPath(SUBMIT));
            submit.Click();
            WaitPageLoading();
            WaitForLoad();
        }
        public bool IsVisible()
        {
            return isElementVisible(By.XPath(FIRST_ROW));
        }
        
        public void DeletePacket()
        {
            if (IsDev())
            {
                var modal = WaitForElementIsVisible(By.XPath(DELETE_PACKET_ICON));
                modal.Click();
            }
            else
            {
                var modal = WaitForElementIsVisible(By.XPath("//*/span[contains(text(),'ATLANTIC')]/../a/span"));
                modal.Click();
            }
            WaitForLoad();
        }
        public void FillPacketServiceFoodPacket( string foodPacket)
        {
            
            
            var f = WaitForElementIsVisible(By.XPath(FOOD_PACKET));
            f.SetValue(ControlType.TextBox, "ATLANTIC");
            WaitPageLoading();
            WaitForLoad();
            if (isElementExists(By.XPath(PACKET_ITEM)))
            {
                var selectFirstfoodpacket = WaitForElementIsVisible(By.XPath(PACKET_ITEM));
                selectFirstfoodpacket.Click();
                WaitForLoad();
            }
            //WaitForLoad();
            var submit = WaitForElementIsVisible(By.XPath(SUBMIT));
            submit.Click();
            WaitPageLoading();
            WaitForLoad();

        }
    }
}
