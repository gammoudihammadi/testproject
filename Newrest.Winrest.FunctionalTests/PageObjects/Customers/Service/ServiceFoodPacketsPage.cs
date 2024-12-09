using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Drawing;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newrest.Winrest.FunctionalTests.PageObjects.Customers.FoodPackets;
using System.Security.Policy;
using Newrest.Winrest.FunctionalTests.PageObjects.Purchasing;
using OpenQA.Selenium.Support.UI;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Customers.Service
{
    public class ServiceFoodPacketsPage : PageBase
    {

        private const string OPEN_FOODPACK_MODAL = "//*/a[@data-toggle=\"modal\" and contains(text(),'Create new packet association')]";
        private const string DELETE = "//*[@id=\"tabContentDetails\"]/div/div[2]/div/div/div[4]/div/div/a[2]/span";
        private const string OK = "//*[@id=\"dataConfirmOK\"]";
        private const string FOOD_PACKETS = "/html/body/div[3]/div/div/div[2]/div/div/div[2]/div[*]";
        private const string FIRST_PACKET_SITE = "//*[@id=\"tabContentDetails\"]/div/div[2]/div/div/div[1]/span";
        private const string FIRST_EDIT_ICON = "/html/body/div[3]/div/div/div[2]/div/div/div[2]/div/div/div[4]/div/div/a[1]/span";
        private const string FOOD_PACKET = "/html/body/div[3]/div/div/div[2]/div/div/div[2]/div[1]/div/div[3]";
        private const string PACKET_LIST = "//*[@id=\"tabContentDetails\"]/div/div[2]/div/div/div[3]/span";


        public ServiceFoodPacketsPage(IWebDriver _webDriver, TestContext _testContext) : base(_webDriver, _testContext)
        {
        }

        public ServiceCreateFoodPacketModal CreateFoodPackets()
        {
            var modal = WaitForElementIsVisible(By.XPath(OPEN_FOODPACK_MODAL));
            modal.Click();
            WaitForLoad();
            return new ServiceCreateFoodPacketModal(_webDriver, _testContext);
        }
        public void DeleteRow()
        {
            var modal = WaitForElementIsVisible(By.XPath(DELETE));
            modal.Click();
            WaitForLoad();
            var modal1 = WaitForElementIsVisible(By.XPath(OK));
            modal1.Click();
            WaitForLoad();
            WaitPageLoading();
        }
        public int GetNumberOfPacketFoodsInService()
        {
            var firstSite = _webDriver.FindElements(By.XPath(FIRST_PACKET_SITE));
            if (firstSite.Count != 0)
            {
                var packetFoods = _webDriver.FindElements(By.XPath(FOOD_PACKETS));
                return packetFoods.Count;
            }
            else
            {
                return 0;
            }
        }
        public void DeleteRows(int numberPackets)
        {
            if (numberPackets > 0)
            {

                int index = numberPackets;
                while (index > 0)
                {
                    DeleteRow();
                    index--;
                }

            }
        }
        public string GetPacket()
        {
            var packet = WaitForElementIsVisible(By.XPath(FOOD_PACKET));
            return packet.Text;
        }
        public ServiceCreateFoodPacketModal EditItem()
        {
            var modal = WaitForElementIsVisible(By.XPath(FIRST_EDIT_ICON));
            modal.Click();
            WaitForLoad();
            return new ServiceCreateFoodPacketModal(_webDriver, _testContext);
        }
        public int GetNumberOfPackets()
        {
            var listPackets = WaitForElementIsVisible(By.XPath(PACKET_LIST));
            var value = listPackets.Text;
            string[] packetsArray = value.Split(',');
            var index = 0;
            foreach (string packet in packetsArray)
            {
                index++;
            }
            return index;

        }
    }
}