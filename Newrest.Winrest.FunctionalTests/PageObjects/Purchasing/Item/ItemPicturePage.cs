using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Purchasing.Item
{
    public class ItemPicturePage : PageBase
    {
        public ItemPicturePage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }


        private const string UPLOAD_PICTURE = "FileSent";
        private const string PICTURE = "img-logo";
        private const string DELETE_PICTURE = "delete-picture";
        private const string DELETE_PICTURE_MODAL_BTN = "dataConfirmOK";


        [FindsBy(How = How.Id, Using = UPLOAD_PICTURE)]
        private IWebElement _uploadPicture;

        [FindsBy(How = How.Id, Using = DELETE_PICTURE)]
        private IWebElement _deletePicture;

        [FindsBy(How = How.Id, Using = DELETE_PICTURE_MODAL_BTN)]
        private IWebElement _deletePictureModalBtn;

        public void AddPicture(string imagePath)
        {
            _uploadPicture = WaitForElementIsVisible(By.Id(UPLOAD_PICTURE));
            _uploadPicture.SendKeys(imagePath);
            WaitPageLoading();
            WaitForElementIsVisible(By.Id(DELETE_PICTURE));
            WaitForLoad();
        }

        public void DeletePicture() 
        {
            _deletePicture = WaitForElementIsVisible(By.Id(DELETE_PICTURE));
            _deletePicture.Click();
            WaitForLoad();
            _deletePictureModalBtn = WaitForElementIsVisible(By.Id(DELETE_PICTURE_MODAL_BTN));
            _deletePictureModalBtn.Click();
            WaitForLoad();
        }

        public bool IsPictureAdded()
        {
            WaitForLoad();
            try
            {
                WaitForElementIsVisible(By.Id(PICTURE));
                return true;
            }
            catch
            {
                return false;
            }
        }

        public bool IsPictureDeleted()
        {
            return !isElementVisible(By.Id(PICTURE));
        }
    }
}
