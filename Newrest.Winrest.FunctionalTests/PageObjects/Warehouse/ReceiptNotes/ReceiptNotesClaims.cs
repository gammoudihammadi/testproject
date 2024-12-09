using DocumentFormat.OpenXml.VariantTypes;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.PageObjects;
using System;
using System.Globalization;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Warehouse.ReceiptNotes
{
    public class ReceiptNotesClaims : PageBase
    {

        private const string ITEMS_TAB = "hrefTabContentItems";

        private const string ITEM_NAME = "//*[@class='item-detail-col-name-span']";
        private const string ITEM_PACKAGING = "//*[@name='ClaimDetailPackaging']";
        private const string ITEM_PACK_PRICE = "//*[@name='ClaimDetailPackagingPrice']";
        private const string ITEM_RECEIVED = "//*[@name='ClaimDetailReceivedQty']";
        private const string ITEM_DN_PRICE = "//*[@name='ClaimDetailDNPrice']";
        private const string ITEM_DN_QTY = "//*[@name='ClaimDetailDNQty']";
        private const string ITEM_DN_TOTAL = "//*[@name='ClaimDetailDNTotal']";
        private const string ITEM_SANCTION_AMOUNT = "//*[@name='ClaimDetailSanctionAmount']";
        private const string COMMENT_ICON= "//*[@name='ClaimHasComment']";
        private const string PICTURE_ICON = "//*[@name='ClaimHasPicture']";

        [FindsBy(How = How.Id, Using = ITEMS_TAB)]
        private IWebElement _itemsTab;

        [FindsBy(How = How.Id, Using = ITEM_NAME)]
        private IWebElement _itemName;

        [FindsBy(How = How.Id, Using = ITEM_PACKAGING)]
        private IWebElement _itemPackaging;

        [FindsBy(How = How.Id, Using = ITEM_PACK_PRICE)]
        private IWebElement _itemPackPrice;

        [FindsBy(How = How.Id, Using = ITEM_RECEIVED)]
        private IWebElement _itemReceived;

        [FindsBy(How = How.Id, Using = ITEM_DN_PRICE)]
        private IWebElement _priceDN;

        [FindsBy(How = How.Id, Using = ITEM_DN_QTY)]
        private IWebElement _qtyDN;

        [FindsBy(How = How.Id, Using = ITEM_DN_TOTAL)]
        private IWebElement _totalDN;

        [FindsBy(How = How.Id, Using = ITEM_SANCTION_AMOUNT)]
        private IWebElement _sanctionAmount;

        [FindsBy(How = How.Id, Using = COMMENT_ICON)]
        private IWebElement _comment_icon;

        [FindsBy(How = How.Id, Using = PICTURE_ICON)]
        private IWebElement _has_picture;


        public ReceiptNotesClaims(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        public ReceiptNotesItem ClickOnItemsTab()
        {
            WaitForLoad();
            _itemsTab = WaitForElementIsVisible(By.Id(ITEMS_TAB));
            _itemsTab.Click();
            WaitForLoad();

            return new ReceiptNotesItem(_webDriver, _testContext);
        }

        public string getItemName()
        {
            _itemName = WaitForElementIsVisible(By.XPath(ITEM_NAME));
            return _itemName.Text;
        }

        public string getItemPackaging()
        {
            _itemPackaging = WaitForElementIsVisible(By.XPath(ITEM_PACKAGING));
            return _itemPackaging.Text;
        }

        public double getItemPackPrice(string decimalSeparator)
        {
            _itemPackPrice = WaitForElementIsVisible(By.XPath(ITEM_PACK_PRICE));
            CultureInfo ci = decimalSeparator.Equals(",") ? new CultureInfo("fr-FR") : new CultureInfo("en-US");
            return double.Parse(_itemPackPrice.Text.Replace("€", "").Replace(" ", ""), ci);
        }

        public string getItemReceived(string decimalSeparator)
        {
            _itemReceived = WaitForElementIsVisible(By.XPath(ITEM_RECEIVED));
            return _itemReceived.Text;
        }


        public double getDNPrice(string decimalSeparator)
        {
            _priceDN = WaitForElementIsVisible(By.XPath(ITEM_DN_PRICE));
            CultureInfo ci = decimalSeparator.Equals(",") ? new CultureInfo("fr-FR") : new CultureInfo("en-US");
            return double.Parse(_priceDN.Text.Replace("€", "").Replace(" ", ""), ci);
        }

        public double getDNQty(string decimalSeparator)
        {
            _qtyDN = WaitForElementIsVisible(By.XPath(ITEM_DN_QTY));
            CultureInfo ci = decimalSeparator.Equals(",") ? new CultureInfo("fr-FR") : new CultureInfo("en-US");
            return double.Parse(_qtyDN.Text.Replace("€", "").Replace(" ", ""), ci);
        }

        public double getDNTotal(string decimalSeparator)
        {
            _totalDN = WaitForElementIsVisible(By.XPath(ITEM_DN_TOTAL));
            CultureInfo ci = decimalSeparator.Equals(",") ? new CultureInfo("fr-FR") : new CultureInfo("en-US");
            return double.Parse(_totalDN.Text.Replace("€", "").Replace(" ", ""), ci);
        }

        public bool GetDecrStock()
        {
            var _descrStockCheckBox = WaitForElementExists(By.XPath("//*[@id=\"tabContentServiceContainer\"]/div[1]/div[2]/div/div/div[10]/div/input"));
            return _descrStockCheckBox.Selected;
        }

        public string GetDecrQty()
        {
            var _descrQty = WaitForElementIsVisible(By.XPath("//*[@id=\"tabContentServiceContainer\"]/div[1]/div[2]/div/div/div[11]/span"));
            return _descrQty.Text;
        }

        public ReceiptNotesEditClaim EditClaim()
        {
            var editButton = WaitForElementIsVisible(By.XPath("//*/span[contains(@class,'fas fa-edit')]"));
            editButton.Click();
            return new ReceiptNotesEditClaim(_webDriver, _testContext);
        }

        public void DeleteClaim()
        {
            // /span
            var deleteCLaim = WaitForElementIsVisible(By.XPath("//*[@id=\"tabContentServiceContainer\"]/div[1]/div[2]/div/div/div[15]/a"));
            deleteCLaim.Click();
            WaitForLoad();
        }

        public int CheckTableSize()
        {
            return _webDriver.FindElements(By.XPath("//*[@class='item-detail-col-name-span']")).Count;
        }
        public double getSanctionAmount(string decimalSeparator , string currency)
        {
            WaitForLoad();
            CultureInfo ci = decimalSeparator.Equals(",") ? new CultureInfo("fr-FR") : new CultureInfo("en-US");

            _sanctionAmount = WaitForElementIsVisible(By.XPath(ITEM_SANCTION_AMOUNT));
            var sanctionAmountText = _sanctionAmount.Text;

            var format = (NumberFormatInfo)ci.NumberFormat.Clone();
            format.CurrencySymbol = currency;
            var mynumber = Decimal.Parse(sanctionAmountText, NumberStyles.Currency, format);

            return Convert.ToDouble(mynumber, ci);
            //return double.Parse(_sanctionAmount.Text.Replace("€", "").Replace(" ", ""), ci);
        }
        public bool CommentChecker()
        {
            _comment_icon = WaitForElementExists(By.XPath(COMMENT_ICON));
            return _comment_icon.Displayed;
        }
        public bool PictureChecker()
        {
            _has_picture = WaitForElementExists(By.XPath(PICTURE_ICON));
            return _has_picture.Displayed;
        }
    }
}
