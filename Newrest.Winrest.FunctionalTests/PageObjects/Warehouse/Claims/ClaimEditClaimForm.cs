using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.PageObjects.Warehouse.ReceiptNotes;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;
using System;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Warehouse.Claims
{
    public class ClaimEditClaimForm : PageBase
    {

        private const string UPLOAD_PICTURE = "inputFileSent";
        private const string PICTURE = "claimPhotoServer";
        private const string CLAIM_FORM_CHECKBOX = "//label[contains(text(), '{0}')]";
        private const string CLAIM_FORM_SAVE = "btn-valid-claim";
        private const string DECREASE_STOCK = "DecreaseStock";
        private const string COMMENT = "Comment";
        private const string SANCTION_AMOUNT = "SanctionAmount";


        [FindsBy(How = How.Id, Using = UPLOAD_PICTURE)]
        private IWebElement _uploadPicture;

        [FindsBy(How = How.XPath, Using = CLAIM_FORM_CHECKBOX)]
        private IWebElement _claimFormCheckbox;

        [FindsBy(How = How.Id, Using = CLAIM_FORM_SAVE)]
        private IWebElement _claimFormSave;

        public ClaimEditClaimForm(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }



        public void SetClaimForm()
        {
            var checkchevkClaim1 = WaitForElementIsVisible(By.Id("radioDelivery"));
            checkchevkClaim1.Click();

            WaitForLoad();

            var checkchevkClaim2 = WaitForElementIsVisible(By.Id("radioQuality"));
            checkchevkClaim2.Click();

            WaitForLoad();

            var checkchevkClaim3 = WaitForElementIsVisible(By.Id("radioNonCompliantTemperature"));
            checkchevkClaim3.Click();

            WaitForLoad();

            var vtnValid = WaitForElementIsVisible(By.Id("btn-valid-claim"));
            vtnValid.Click();

            WaitForLoad();
        }

        public void SetClaimFormForClaimed(string imagePath, string price)
        {
            var checkchevkClaim1 = WaitForElementIsVisible(By.Id("radioProduct"));
            checkchevkClaim1.Click();

            WaitForLoad();

            if(price != null)
            {
                var checkchevkClaim2 = WaitForElementIsVisible(By.Id("radioProductPricing"));
                checkchevkClaim2.Click();

                WaitForLoad(); 


                var checkchevkClaim3 = WaitForElementIsVisible(By.Id("radioErrorBoth"));
                checkchevkClaim3.Click();

                WaitForLoad();

                var ClaimPrice = WaitForElementIsVisible(By.Id("ClaimAmount_PriceError"));
                ClaimPrice.SetValue(ControlType.TextBox, price);

                WaitForLoad();


            }
            else
            {

                var checkchevkClaim2 = WaitForElementIsVisible(By.Id("radioProductQuality"));
                checkchevkClaim2.Click();

                WaitForLoad();

                var checkchevkClaim3 = WaitForElementIsVisible(By.Id("radioCompliantCalibration"));
                checkchevkClaim3.Click();

                WaitForLoad();


                var ClaimQty = WaitForElementIsVisible(By.Id("ClaimedReceivedQuantity"));
                ClaimQty.SetValue(ControlType.TextBox, "1");

                WaitForLoad();

                var incidentDescription = WaitForElementIsVisible(By.Id("PQ_IncidentElements_IncidentDescription"));
                incidentDescription.SetValue(ControlType.TextBox, "testIncident");

                WaitForLoad();

                if (imagePath != null)
                {

                }
            }


            WaitForLoad();

            var vtnValid = WaitForElementIsVisible(By.Id("btn-valid-claim"));
            vtnValid.Click();

            WaitForLoad();
        }


        public ReceiptNotesClaims SetClaimFormForClaimedV3(string imagePath = null)
        {
            var checkchevkClaim1 = WaitForElementIsVisible(By.XPath("//label[contains(text(), 'Non-compliant delivered ')]"));
            checkchevkClaim1.Click();

            WaitForLoad();
            if (imagePath != null)
            {
                _uploadPicture = WaitForElementIsVisible(By.Id(UPLOAD_PICTURE));
                _uploadPicture.SendKeys(imagePath);
            }

            WaitForLoad();

            var vtnValid = WaitForElementIsVisible(By.Id("btn-valid-claim"));
            vtnValid.Click();

            WaitForLoad();

            return new ReceiptNotesClaims(_webDriver, _testContext);
        }


        public bool IsPictureAdded()
        {
            try
            {
                _webDriver.FindElement(By.Id(PICTURE));
                return true;
            }
            catch
            {
                return false;
            }
        }

        public void FillClaimFormCheckbox(string texte)
        {
            // marche grâce à la magie du "<label for="
            _claimFormCheckbox = WaitForElementIsVisible(By.XPath(string.Format(CLAIM_FORM_CHECKBOX, texte)));
            _claimFormCheckbox.Click();
        }

        public ClaimsItem Save()
        {
            _claimFormSave = WaitForElementIsVisible(By.Id(CLAIM_FORM_SAVE));
            _claimFormSave.Click();
            WaitPageLoading();
            return new ClaimsItem(_webDriver, _testContext);
        }
        public void EditClaim(bool decreasestock,string claimtype,string newcomment,string sanctionamount)
        {
            var decreaseStockCheckBox = WaitForElementIsVisible(By.Id(DECREASE_STOCK));
            decreaseStockCheckBox.SetValue(ControlType.CheckBox, decreasestock);
            FillClaimFormCheckbox(claimtype);
            var commentInput = WaitForElementIsVisible(By.Id(COMMENT));
            commentInput.Clear();
            commentInput.SendKeys(newcomment);
            var sanctionAmountInput = WaitForElementIsVisible(By.Id(SANCTION_AMOUNT));
            sanctionAmountInput.Clear();
            sanctionAmountInput.SendKeys(sanctionamount);
            Save();
        }
    }
}
