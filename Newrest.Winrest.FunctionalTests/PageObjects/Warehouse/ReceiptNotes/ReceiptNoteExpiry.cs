using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Warehouse.ReceiptNotes
{
    public class ReceiptNoteExpiry : PageBase
    {

        private const string COUNT = "//*/input[starts-with(@id,'Quantity_')]";
        private const string ADD = "btn-create-new-row";
        private const string QUANTITY = "Quantity_{0}";
        private const string DATE = "datepicker-expiration_{0}";
        private const string DELETE = "//*[@id=\"rowExpiryDate_{0}\"]/div[5]/a";
        private const string SAVE = "btnSubmit";
        private const string CLOSE = "//*/span[text()='×']/parent::button";
        private const string GREEN_ICON = "(//*[contains(@name,'IconExpirationDate')])[1]";

        public ReceiptNoteExpiry(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        public int CountExpiryDate()
        {
            return _webDriver.FindElements(By.XPath(COUNT)).Count;
        }

        public void AddExpiryDate(string count, DateTime expiryDateTest)
        {
            WaitForLoad();
            var addLineButton = WaitForElementIsVisible(By.Id(ADD));

            if (isElementExists(By.XPath("//*[contains(@class, 'datepicker-switch')]")))
            { addLineButton.Click(); WaitForLoad(); }
            addLineButton.Click();
            WaitForLoad();

            int nbLignes = CountExpiryDate();

            var idInput = String.Format(QUANTITY, nbLignes - 1);
            var idDate = String.Format(DATE, nbLignes - 1);

            var input = WaitForElementIsVisible(By.Id(idInput));
            input.Clear();
            input.SendKeys(count);
            WaitForLoad();

            var date = WaitForElementIsVisible(By.Id(idDate));
            date.Clear();
            date.SendKeys(expiryDateTest.ToString("dd/MM/yyyy"));
            date.SendKeys(Keys.Tab);
            WaitForLoad();
        }

        public void ModifyFirstExpiryDate(string count, DateTime expiryDateTest)
        {
            var idInput = String.Format(QUANTITY, 0);
            var idDate = String.Format(DATE, 0);

            var input = WaitForElementIsVisible(By.Id(idInput));
            input.Clear();
            input.SendKeys(count);
            WaitForLoad();

            var date = WaitForElementIsVisible(By.Id(idDate));
            date.Clear();
            date.SendKeys(expiryDateTest.ToString("dd/MM/yyyy"));
            WaitForLoad();
        }

        public bool CheckGreenIcon()
        {
            var greenOn = WaitForElementIsVisible(By.XPath(GREEN_ICON));
            return greenOn.GetAttribute("class").Contains("green-text");
        }

        public ReceiptNotesItem Save()
        {
            var SaveButton = WaitForElementIsVisible(By.Id(SAVE));
            // bouton SAVE bleu > bouton SAVE vert
            new Actions(_webDriver).MoveToElement(SaveButton).Click().Perform();

            WaitForLoad();
            return new ReceiptNotesItem(_webDriver, _testContext);
        }


        public ReceiptNotesItem Cancel()
        {
            var closeButton = WaitForElementIsVisible(By.XPath(CLOSE));
            closeButton.Click();
            WaitForLoad();
            return new ReceiptNotesItem(_webDriver, _testContext);
        }

        public void DeleteAllExpiryDate()
        {
            int nbLines = CountExpiryDate();
            for (int i = nbLines - 1; i >= 0; i--)
            {
                // attendre le lien puis lancer le javascript
                string xPath = String.Format(DELETE, i);
                var lien = WaitForElementIsVisible(By.XPath(xPath));
                var javascript = lien.GetAttribute("onclick");
                IJavaScriptExecutor executor = (IJavaScriptExecutor)_webDriver;
                executor.ExecuteScript(javascript);
            }
        }
    }
}
