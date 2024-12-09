using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Warehouse.ReceiptNotes
{
    public class ReceiptNotesEditClaim : PageBase
    {
        private const string CLAIM_NON_COMPLIANT_RADIO_BTN = "//*/label[text()='Non compliant expiration date']";
        private const string CLAIM_NOT_RECEIVED_RADIO_BTN = "//*/label[text()='Not received quantity']";
        private const string CLAIM_WRONG_FEE_RADIO_BTN = "//*/label[text()='Wrong fee']";
        private const string CLAIM_SANCTION_AMOUNT = "//input[@id='SanctionAmount'][1]";

        private const string SAVE_CLAIM = "btn-valid-claim";

        [FindsBy(How = How.Id, Using = CLAIM_NON_COMPLIANT_RADIO_BTN)]
        private IWebElement _claimNonCompliantRadioBtn;

        [FindsBy(How = How.Id, Using = CLAIM_NOT_RECEIVED_RADIO_BTN)]
        private IWebElement _claimNotReceivedRadioBtn;

        [FindsBy(How = How.Id, Using = CLAIM_WRONG_FEE_RADIO_BTN)]
        private IWebElement _claimWrongFeeRadioBtn;

        [FindsBy(How = How.Id, Using = CLAIM_SANCTION_AMOUNT)]
        //private IWebElement _claimSanctionAmountInput;

        [FindsBy(How = How.Id, Using = SAVE_CLAIM)]
        private IWebElement _saveClaim;

        public ReceiptNotesEditClaim(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        public void SetDecreaseStock(string decr)
        {
            var checkBoxDecrease = WaitForElementIsVisible(By.Id("DecreaseStock"));
            if (!checkBoxDecrease.Selected)
            {
                checkBoxDecrease.Click();
            }
            var inputDecrease = WaitForElementIsVisible(By.Id("DecreaseQuantity"));
            inputDecrease.Clear();
            inputDecrease.SendKeys(decr);
            WaitForLoad();
        }

        public void UploadPicture(string img)
        {
            FileInfo fiUpload = new FileInfo(_testContext.TestDeploymentDir + "\\"+img);
            Assert.IsTrue(fiUpload.Exists, "Fichier d'entrée non trouve");

            var buttonFile = WaitForElementIsVisible(By.Id("inputFileSent"));
            buttonFile.SendKeys(fiUpload.FullName);
            //chargement du preview
            WaitForLoad();
        }

        public void SetComment(string comment)
        {
            var commentTextArea = WaitForElementIsVisible(By.Id("Comment"));
            commentTextArea.Clear();
            commentTextArea.SendKeys(comment);
            WaitForLoad();
        }

        public void SetSanctionAmount(string sanction)
        {
            var sanctionInput = WaitForElementIsVisible(By.Id("SanctionAmount"));
            sanctionInput.Clear();
            sanctionInput.SendKeys(sanction);
            WaitForLoad();
        }

        public void SetVAT(string vat)
        {
            var vatSelect = WaitForElementIsVisible(By.Id("TaxTypeId"));
            var selectElement = new SelectElement(vatSelect);
            selectElement.SelectByText(vat);
        }

        public void Fill()
        {
            _claimNonCompliantRadioBtn = WaitForElementIsVisible(By.XPath(CLAIM_NON_COMPLIANT_RADIO_BTN));
            _claimNonCompliantRadioBtn.SetValue(ControlType.CheckBox, true);
            _claimNotReceivedRadioBtn = WaitForElementIsVisible(By.XPath(CLAIM_NOT_RECEIVED_RADIO_BTN));
            _claimNotReceivedRadioBtn.SetValue(ControlType.CheckBox, true);
            _claimWrongFeeRadioBtn = WaitForElementIsVisible(By.XPath(CLAIM_WRONG_FEE_RADIO_BTN));
            _claimWrongFeeRadioBtn.SetValue(ControlType.CheckBox, true);
        }

        public ReceiptNotesClaims Save()
        {
            _saveClaim = WaitForElementIsVisible(By.Id(SAVE_CLAIM));
            _saveClaim.Click();
            WaitForLoad();
            return new ReceiptNotesClaims(_webDriver, _testContext);
        }
        public bool GetDecreaseStock()
        {
            var checkBoxDecrease = WaitForElementIsVisible(By.Id("DecreaseStock"));
            return checkBoxDecrease.Selected;
        }
    }
}
