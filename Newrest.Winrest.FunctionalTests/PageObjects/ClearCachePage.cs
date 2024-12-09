using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Newrest.Winrest.FunctionalTests.PageObjects
{
    public class ClearCachePage
    {
        private readonly IWebDriver _webDriver;

        [FindsBy(How = How.XPath, Using = "/html/body/a")]
        private IWebElement _clearCacheLink;

        public ClearCachePage(IWebDriver webDriver)
        {
            _webDriver = webDriver;
            PageFactory.InitElements(_webDriver, this);
        }

        public void Clear()
        {
            _clearCacheLink.Click();
        }
    }

    
}
