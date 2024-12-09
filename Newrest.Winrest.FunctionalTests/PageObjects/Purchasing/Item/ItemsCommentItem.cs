using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using FluentAssertions;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium.Support.UI;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Purchasing
{
    public class ItemsCommentItem : PageBase
    {

        // ________________________________________ Constantes __________________________________________________

        private const string COMMENT = "Comment";
        private const string SAVE = "//*[@id=\"modal-1\"]/div/div/div[2]/div/form/div[2]/button[2]";

        // ________________________________________ Variables ___________________________________________________

        [FindsBy(How = How.Id, Using = COMMENT)]
        private IWebElement _comment;

        [FindsBy(How = How.XPath, Using = SAVE)]
        private IWebElement _save;  
        
        // ________________________________________ Méthodes ____________________________________________________

        public ItemsCommentItem(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        public void FillComment(string comment)
        {
            _comment = WaitForElementIsVisible(By.Id(COMMENT));
            _comment.SetValue(ControlType.TextBox, comment);   
        }

        public ItemGeneralInformationPage Save()
        {
                _save = WaitForElementIsVisible(By.XPath("//*/button[@value='Create']"));
            _save.Click();
            WaitForLoad();

            return new ItemGeneralInformationPage(_webDriver, _testContext);
        }

        public Boolean IsCommented(string comment)
        {
            _comment = WaitForElementIsVisible(By.Id(COMMENT));

            if (_comment.Text.Contains(comment))
            {
                return true;
            }
            else
            {
                return false;
            }
        }

    }
}
