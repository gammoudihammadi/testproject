using DocumentFormat.OpenXml.Bibliography;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using System;
using System.IO;
using System.Threading;

namespace Newrest.Winrest.FunctionalTests.PageObjects.TabletApp
{
    public class PlateauFlightTabletAppPage : PageBase
    {

        private const string UPLOAD_IMAGE = "//*/app-flight-detail-datasheet/datasheet-shared-view/div/div/div/div[3]/div/div/div/input";
        private const string VERIFY_IMAGE = "//*/div[contains(@class,'ngx-gallery-image') and contains(@style,'no-picture.svg')]";
        private const string QUANTITY_INPUT = "qty";
        private const string LIST_TYPED_PAX = "//*[@id=\"typed-pax\"]";
        private const string DELETE_IMAGE = "//*/button[text()='Delete image']";
        private const string UNFOLD_ALL_BTN = "/html/body/div[2]/div[2]/div/mat-dialog-container/div/div/app-flight-detail-datasheet/datasheet-shared-view/div/div/div/div[3]/div/button";
        private const string DETAIL_MENU = "//*[@id=\"flight-detail-datasheet-dialog\"]/div/div/app-flight-detail-datasheet/datasheet-shared-view/div/div/div/div[3]/mat-list[1]/details/div/div[2]/div/div/div[1]";
        private const string TAKE_A_PHOTO_BTN = "/html/body/div[2]/div[2]/div/mat-dialog-container/div/div/app-flight-detail-datasheet/datasheet-shared-view/div/div/div/div[3]/div/div/button[1]";
        private const string SEND_PHOTO_BTN = "/html/body/div[2]/div[4]/div/mat-dialog-container/div/div/webcampopup/mat-dialog-actions/button[2]/span[2]";
        public PlateauFlightTabletAppPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        public void UploadImage(FileInfo fiUpload)
        {
            Assert.IsTrue(fiUpload.Exists, "Fichier d'entrée non trouve");
            var input = WaitForElementExists(By.XPath(UPLOAD_IMAGE));
            input.SendKeys(fiUpload.FullName);
            //chargement du previews
            WaitPageLoading();
            WaitLoading();
        }

        public new void WaitForLoad()
        {
            IWait<IWebDriver> wait = new OpenQA.Selenium.Support.UI.WebDriverWait(_webDriver, TimeSpan.FromSeconds(30.00));
            wait.Until(driver1 => ((IJavaScriptExecutor)_webDriver).ExecuteScript("return document.readyState").Equals("complete"));
        }

        public bool VerifyImage()
        {
            WaitPageLoading();
            return !isElementVisible(By.XPath(VERIFY_IMAGE));
        }

        public void RemoveImage()
        {
            var button = WaitForElementIsVisible(By.XPath(DELETE_IMAGE));
            button.Click();
            //chargement du previews
            WaitPageLoading();
            WaitLoading();
        }
        public void UnfoldAll()
        {
            IWebElement unfoldAllBtn = WaitForElementIsVisible(By.XPath(UNFOLD_ALL_BTN));
            unfoldAllBtn.Click();   

        }
        public bool IsUnfoldAll()
        {
            if(isElementVisible(By.XPath(DETAIL_MENU)))
            {
                var detailMenu = WaitForElementExists(By.XPath(DETAIL_MENU));
                return true;
            }
            return false;
        }
        public void SetQuantity(string qty)
        {
            var qtyInput = WaitForElementIsVisible(By.Id(QUANTITY_INPUT));
            qtyInput.Clear();   
            qtyInput.SendKeys(qty);
            qtyInput.SendKeys(Keys.Enter);
            WaitForLoad();
        }
        public bool isRecalculate(string qty)
        {
            var listeTypedPax = _webDriver.FindElements(By.XPath(LIST_TYPED_PAX));
            bool result = false;
            foreach(var pax in listeTypedPax)
            {
                var text = pax.Text.Substring(1, pax.Text.Length - 6);
                if (text != qty)
                {
                    result = false;
                    break;
                }
                result = true;
            }
            return result;
        }

        public void TakeAPhoto()
        {
            var takeAPhotoBtn = WaitForElementIsVisible(By.XPath(TAKE_A_PHOTO_BTN));
            takeAPhotoBtn.Click();

            var sendPhotoBtn = WaitForElementIsVisible(By.XPath(SEND_PHOTO_BTN));
            sendPhotoBtn.Click();
            Thread.Sleep(1000);
        }
    }
}
