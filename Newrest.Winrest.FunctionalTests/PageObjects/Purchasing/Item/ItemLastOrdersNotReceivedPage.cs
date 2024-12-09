using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Purchasing.Item
{
    public class ItemLastOrdersNotReceivedPage : PageBase
    {
        public ItemLastOrdersNotReceivedPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }
    }
}
