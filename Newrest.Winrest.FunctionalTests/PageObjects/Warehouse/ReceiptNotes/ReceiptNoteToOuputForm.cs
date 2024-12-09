using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using Newrest.Winrest.FunctionalTests.PageObjects.Warehouse.OutputForm;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Warehouse.ReceiptNotes
{
    public class ReceiptNoteToOuputForm : PageBase
    {
        public ReceiptNoteToOuputForm(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        public void Fill(string placeFrom, string placeTo, bool select)
        {
            var from = WaitForElementIsVisible(By.Id("drop-down-places-from"));
            from.SetValue(ControlType.DropDownList, placeFrom);
            var to = WaitForElementIsVisible(By.Id("drop-down-places-to"));
            to.SetValue(ControlType.DropDownList, placeTo);
            // select is done via javascript
            WaitForLoad();
            var selectItem = WaitForElementExists(By.Id("item_IsSelected"));
            Assert.AreEqual("true", selectItem.GetAttribute("value"));

        }

        public OutputFormItem Create()
        {
            //var createButton = WaitForElementIsVisible(By.Id("btn-submit-create-output-form"));
            //createButton.Click();
            
            var createButton = WaitForElementIsVisible(By.Id("form-create-output-form"));
            createButton.Submit();
            WaitPageLoading();
            return new OutputFormItem(_webDriver, _testContext);
        }
    }
}
