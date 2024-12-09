using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Warehouse.ReceiptNotes
{
    public class ReceiptNotesSecurityChecks : PageBase
    {

        public ReceiptNotesSecurityChecks(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        //_________________________________ Constantes ______________________________________

        // Onglet
        private const string ITEMS_TAB = "hrefTabContentItems";

        // Valeurs
        private const string SECURITY_CHECKS = "//*[starts-with(@id,\"securityCheckStatus-\")][@value='{0}']";

        //_________________________________ Variables _______________________________________

        // Onglets
        [FindsBy(How = How.Id, Using = ITEMS_TAB)]
        private IWebElement _receiptNoteItemsTab;

        // Valeurs
        [FindsBy(How = How.Id, Using = SECURITY_CHECKS)]
        private IWebElement _securityChecks;

        // ________________________________ Méthodes ________________________________________

        public ReceiptNotesItem ClickOnReceiptNoteItemTab()
        {
            _receiptNoteItemsTab = WaitForElementIsVisible(By.Id(ITEMS_TAB));
            _receiptNoteItemsTab.Click();
            WaitForLoad();

            return new ReceiptNotesItem(_webDriver, _testContext);
        }

        /// <summary>
        /// Permet d'activer les security checks pour une RN. 
        /// </summary>
        /// <param name="status">statut à renseigner : "Yes", "No", ou "NotApplicable" </param>
        public void SetSecurityChecks(string status)
        {
            _securityChecks = WaitForElementExists(By.XPath(string.Format(SECURITY_CHECKS, status)));
            _securityChecks.SetValue(ControlType.RadioButton, true);
            WaitForLoad();

            // Temps d'enregistrement de la donnée
            Thread.Sleep(2000);
        }

    }
}
