using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Accounting.Invoice;
using Newrest.Winrest.FunctionalTests.PageObjects.Purchasing;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Production.CustomerOrder
{
    public class CustomerOrderItem : PageBase
    {
        public CustomerOrderItem(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        //__________________________________ Constantes _____________________________________

        // General
        private const string BACK_TO_LIST = "/html/body/div[2]/a/span[2]";
        private const string VALIDATION_BUTTON = "//*[@id=\"div-body\"]/div/div[1]/div/div[3]/button";
        private const string VALIDATE = "IsOrderValidated";
        private const string EXTENTED_MENU = "//*[@id=\"div-body\"]/div/div[1]/div/div[2]/button";
        private const string PRINT = "//*[@id=\"div-body\"]/div/div[1]/div/div[2]/div/a[1]";
        private const string PRINTPROD = "btn-print-production-order";
        private const string PRINTSHOPLIST = "btn-print-shopList";
        private const string CONFIRM_PRINTPROD = "//*[@id=\"report-form\"]/div/div[3]/button[2]";
        private const string VALIDATE_BTN = "btn-popup-validate";
        private const string ORDER_NAME = "//*[@id=\"dispatchTable\"]/tbody/tr[2]/td[2]";
        private const string ORDER_CATEGORY = "//*[@id=\"dispatchTable\"]/tbody/tr[2]/td[3]";
        private const string ORDER_QUANTITY = "//*[@id=\"dispatchTable\"]/tbody/tr[2]/td[4]";
        private const string VAT = "//*[@id=\"dispatchTable\"]/tbody/tr[2]/td[9]"; 
        private const string VERIFY_BUTTON = "//*[@id='div-body']/div/div[1]/div/div[2]/div/a[@class='btn btn-verify']";
        private const string CC_ADDRESSES = "CcAddresses";


        // Onglets
        private const string GENERAL_INFORMATION_TAB = "hrefTabContentInformations";

        // Tableau
        private const string NEW_ITEM = "addOrderDetailBtn";
        private const string FIRST_ITEM = "//*[@id=\"dispatchTable\"]/tbody/tr[2]";

        private const string ITEM_NAME = "[placeholder = 'Item']";
        private const string FREE_PRICE = "[Free price]";
        private const string ITEM_SELECTED = "//*[@id=\"dispatchTable\"]/tbody/tr[2]/td[2]/span/span/div/div/div[1]";
        private const string ITEM_QUANTITY = "[placeholder='Quantity']";

        
        private const string ITEM_CATEGORY = "/html/body/div[3]/div/div[3]/div[2]/div/div[1]/div[2]/div/div/table/tbody/tr[2]/td[3]/select";
        private const string ITEM_PRICE = "[placeholder='Unit price']";
        private const string ICON_SAVED = "//*[@id=\"dispatchTable\"]/tbody/tr[2]/td[15]/span[@class='fas fa-save']";

        private const string PROD_COMMENT = "//*[@id=\"dispatchTable\"]/tbody/tr[2]/td[12]/a";
        private const string BILLING_COMMENT = "//*[@id=\"dispatchTable\"]/tbody/tr[2]/td[13]/a";
        private const string COMMENT_AREA = "Comment";
        private const string VALIDATE_COMMENT = "//*[@id=\"orderDetailCommentForm\"]/div[2]/button[2]";

        private const string DELETE_ITEM = "//*[@id=\"dispatchTable\"]/tbody/tr[2]/td[14]/div/a[2]/span";
        private const string DELETE_ALL_ITEM = "deleteAllZeroQtyBtn";
        private const string MESSAGES_TAB = "//*[@id=\"hrefTabContentMessages\"]";

        private const string SELECTED_TYPE_RAPPORT = "//*[@id=\"SelectedPurchaseReportType\"]";

        private const string SELECTED_FIRST_TYPE_RAPPORT = "//*[@id=\"SelectedPurchaseReportType\"]/option[2]";
        private const string SEND_EMAIL = "btn-send-by-email-order";
        private const string SEND_EMAIL_BUTTON = "btn-init-async-send-mail";
        private const string SEND_TO_EMAIL_INPUT = "ToAddresses";
        private const string INPUT_CHOOSE_FILE = "add-attachment-input";
        private const string ATTACHMENTS_FILE = "//*[@id=\"modal-1\"]/div/div/div/form/div[2]/div/div[4]/div/div/div/div/div[*]";
        private const string PRICE_CURRENCY = "//*[@id=\"dispatchTable\"]/tbody/tr[2]/td[8]/div/div[1]/span";
       
        private const string SUGGESTION_ELEMENTS = "//*[@id=\"dispatchTable\"]/tbody/tr[2]/td[2]/span/span/div/div/div";
        private const string MENU_SUGGESTION = "//*[@id=\"dispatchTable\"]/tbody/tr[2]/td[2]/span/span/div";
        private const string LABEL1 = "//*[@id=\"Label1\"]";
        private const string LABEL2 = "//*[@id=\"Label2\"]";
        private const string LABEL3 = "//*[@id=\"Label3\"]";
        private const string LABEL4 = "//*[@id=\"Label4\"]";
        private const string LABEL5 = "//*[@id=\"Label5\"]";
        private const string FP_CHECK = "//*[@id=\"dispatchTable\"]/tbody/tr[2]/td[11]";
        private const string S_CHECK = "//*[@id=\"dispatchTable\"]/tbody/tr[3]/td[11]";
        private const string VALIDATE_Button = "btn-popup-validate";
        private const string CLOSE_BTN = "//*[@id=\"modal-1\"]/div[1]/button/span";
        
        private const string SELLING_UNIT = "/html/body/div[3]/div/div[3]/div[2]/div/div[1]/div[2]/div/div/table/tbody/tr[2]/td[5]/select";
        private const string NON_MODIFIABLE_SELLING_UNIT = "/html/body/div[3]/div/div[3]/div[2]/div/div[1]/div[2]/div/div/table/tbody/tr[2]/td[5]";
        private const string TOTAL_PRICE = "/html/body/div[3]/div/div[3]/div[2]/div/div[1]/div[2]/div/div/table/tbody/tr[2]/td[8]/div/div[2]/input";
        private const string PRICE_EXCL = "//b[1]/span";


        //__________________________________ Variables ______________________________________

        // General
        [FindsBy(How = How.XPath, Using = LABEL1)]
        private IWebElement _label1;
        [FindsBy(How = How.XPath, Using = FP_CHECK)]
        private IWebElement _fpcheck;
        [FindsBy(How = How.XPath, Using = S_CHECK)]
        private IWebElement _scheck;
        [FindsBy(How = How.XPath, Using = LABEL2)]
        private IWebElement _label2;
        [FindsBy(How = How.XPath, Using = LABEL3)]
        private IWebElement _label3;
        [FindsBy(How = How.XPath, Using = LABEL4)]
        private IWebElement _label4;
        [FindsBy(How = How.XPath, Using = LABEL5)]
        private IWebElement _label5;

        [FindsBy(How = How.XPath, Using = SELECTED_TYPE_RAPPORT)]
        private IWebElement _selectedtyperaport;

        [FindsBy(How = How.XPath, Using = SELECTED_FIRST_TYPE_RAPPORT)]
        private IWebElement _selectedfirsttyperaport;

        [FindsBy(How = How.XPath, Using = BACK_TO_LIST)]
        private IWebElement _backToList;

        [FindsBy(How = How.XPath, Using = VALIDATION_BUTTON)]
        private IWebElement _validationButton;

        [FindsBy(How = How.Id, Using = VALIDATE)]
        private IWebElement _validate;

        [FindsBy(How = How.Id, Using = VALIDATE_BTN)]
        private IWebElement _validateBtn;

        [FindsBy(How = How.XPath, Using = EXTENTED_MENU)]
        private IWebElement _extendedMenu;

        [FindsBy(How = How.XPath, Using = PRINT)]
        private IWebElement _print;

        [FindsBy(How = How.Id, Using = PRINTPROD)]
        private IWebElement _printProd;

        [FindsBy(How = How.Id, Using = PRINTSHOPLIST)]
        private IWebElement _printShopList;
        
        [FindsBy(How = How.XPath, Using = CONFIRM_PRINTPROD)]
        private IWebElement _confirmPrintProd;

        [FindsBy(How = How.Id, Using = CC_ADDRESSES)]
        private IWebElement _ccAddresses;

        // Onglets
        [FindsBy(How = How.Id, Using = GENERAL_INFORMATION_TAB)]
        private IWebElement _generalInformationTab;

        // Tableau
        [FindsBy(How = How.Id, Using = NEW_ITEM)]
        private IWebElement _createNewItem;

        [FindsBy(How = How.CssSelector, Using = ITEM_NAME)]
        private IWebElement _itemName;

        [FindsBy(How = How.XPath, Using = ITEM_SELECTED)]
        private IWebElement _itemSelected;

        [FindsBy(How = How.CssSelector, Using = ITEM_QUANTITY)]
        private IWebElement _itemQuantity; 

        [FindsBy(How = How.CssSelector, Using = ITEM_CATEGORY)]
        private IWebElement _itemCategory;

        [FindsBy(How = How.XPath, Using = ITEM_PRICE)]
        private IWebElement _itemPrice;

        [FindsBy(How = How.XPath, Using = PROD_COMMENT)]
        private IWebElement _prodComment;

        [FindsBy(How = How.XPath, Using = BILLING_COMMENT)]
        private IWebElement _billingComment;

        [FindsBy(How = How.Id, Using = COMMENT_AREA)]
        private IWebElement _commentArea;

        [FindsBy(How = How.XPath, Using = VALIDATE_COMMENT)]
        private IWebElement _validateComment;

        [FindsBy(How = How.XPath, Using = DELETE_ITEM)]
        private IWebElement _deleteItem;

        [FindsBy(How = How.Id, Using = SEND_EMAIL)]
        private IWebElement _sendEmail;

        [FindsBy(How = How.Id, Using = SEND_TO_EMAIL_INPUT)]
        private IWebElement _sendToEmailInput;

        [FindsBy(How = How.Id, Using = SEND_EMAIL_BUTTON)]
        private IWebElement _sendEmailButton;

        [FindsBy(How = How.XPath, Using = CLOSE_BTN)]
        private IWebElement _closeBtn;
        
        [FindsBy(How = How.XPath, Using = SELLING_UNIT)]
        private IWebElement _sellingUnit;

        [FindsBy(How = How.XPath, Using = TOTAL_PRICE)]
        private IWebElement _totalPrice;

        // ______________________________________________ Méthodes ______________________________________________

        // Général
        public CustomerOrderPage BackToList()
        { 
            _backToList = WaitForElementIsVisible(By.XPath(BACK_TO_LIST));
            _backToList.Click();
            WaitForLoad();

            return new CustomerOrderPage(_webDriver, _testContext);
        }
        public void SetLabels(string label1, string label2, string label3, string label4, string label5)
        {

            _label1 = WaitForElementIsVisible(By.XPath(LABEL1));
            _label1.SetValue(ControlType.TextBox, label1);
            _label2 = WaitForElementIsVisible(By.XPath(LABEL2));
            _label2.SetValue(ControlType.TextBox, label2);
            _label3 = WaitForElementIsVisible(By.XPath(LABEL3));
            _label3.SetValue(ControlType.TextBox, label3);
            _label4 = WaitForElementIsVisible(By.XPath(LABEL4));
            _label4.SetValue(ControlType.TextBox, label4);
            _label5 = WaitForElementIsVisible(By.XPath(LABEL5));
            _label5.SetValue(ControlType.TextBox, label5);
            WaitPageLoading();
            WaitForLoad();
        }
        public CustomerOrderIMessagesPage GoToMessagesTab()
        {
            var _messageTab = WaitForElementIsVisible(By.XPath(MESSAGES_TAB));
            _messageTab.Click();
            WaitForLoad();

            return new CustomerOrderIMessagesPage(_webDriver, _testContext);
        }

        public override void ShowValidationMenu()
        {
            WaitForElementExists(By.XPath(VALIDATION_BUTTON));
            var actions = new Actions(_webDriver);
            actions.MoveToElement(_validationButton).Perform();
            WaitForLoad();
        }

        public override void ShowExtendedMenu()
        {
            WaitForElementExists(By.XPath(EXTENTED_MENU));
            var actions = new Actions(_webDriver);
            actions.MoveToElement(_extendedMenu).Perform();
            WaitForLoad();
        }

        public void ValidateCustomerOrder()
        {
            ShowValidationMenu();

            _validate = WaitForElementIsVisible(By.Id(VALIDATE));
            _validate.Click();
            WaitForLoad();
            _validateBtn = WaitForElementIsVisible(By.Id(VALIDATE_BTN));
            _validateBtn.Click();
            WaitForLoad();
        }

        public PrintReportPage PrintCustomerOrder(bool printValue)
        {

            ShowExtendedMenu();

            _print = WaitForElementIsVisible(By.XPath(PRINT));
            _print.Click();
            WaitForLoad();

            var confirm = WaitForElementIsVisible(By.XPath("//*[@id=\"btn-print\"]"));
            confirm.Click();

            if (printValue)
            {
                IsFileLoaded(By.CssSelector("[target='_blank'][class='far fa-file-pdf']"));
                ClickPrintButton();
            }

            //Results are opened in a new tab, switch the driver to the newly created one
            WebDriverWait wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(60));
            wait.Until((driver) => driver.WindowHandles.Count > 1);
            _webDriver.SwitchTo().Window(_webDriver.WindowHandles[1]);

            return new PrintReportPage(_webDriver, _testContext);
        }

        public PrintReportPage PrintProdCustomerOrder(bool printValue)
        {

            ShowExtendedMenu();

            var printbtn = WaitForElementIsVisible(By.XPath("//*[@id=\"Print-Group-btn\"]"));
            printbtn.Click();

            _printProd = WaitForElementIsVisible(By.Id(PRINTPROD));
            _printProd.Click();
            WaitForLoad();

            _confirmPrintProd = WaitForElementIsVisible(By.XPath(CONFIRM_PRINTPROD));
            _confirmPrintProd.Click();
            WaitForLoad();

            if (printValue)
            {
                IsFileLoaded(By.CssSelector("[target='_blank'][class='far fa-file-pdf']"));
                ClickPrintButton();
            }

            //Results are opened in a new tab, switch the driver to the newly created one
            WebDriverWait wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(60));
            wait.Until((driver) => driver.WindowHandles.Count > 1);
            _webDriver.SwitchTo().Window(_webDriver.WindowHandles[1]);

            return new PrintReportPage(_webDriver, _testContext);
        }

        public PrintReportPage PrintShopListCustomerOrder(bool printValue)
        {

            ShowExtendedMenu();

            var printbtn = WaitForElementIsVisible(By.XPath("//*[@id=\"Print-Group-btn\"]"));
            printbtn.Click();

            _printShopList = WaitForElementIsVisible(By.Id(PRINTSHOPLIST));
            _printShopList.Click();
             WaitForLoad();

            if (printValue)
            {
                IsFileLoaded(By.CssSelector("[target='_blank'][class='far fa-file-pdf']"));
                ClickPrintButton();
            }

            //Results are opened in a new tab, switch the driver to the newly created one
            WebDriverWait wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(60));
            wait.Until((driver) => driver.WindowHandles.Count > 1);
            _webDriver.SwitchTo().Window(_webDriver.WindowHandles[1]);

            return new PrintReportPage(_webDriver, _testContext);
        }
        // Onglets
        public CustomerOrderGeneralInformationPage ClickOnGeneralInformationTab()
        {
            _generalInformationTab = WaitForElementExists(By.Id(GENERAL_INFORMATION_TAB));
            _generalInformationTab.Click();
            WaitForLoad();

            return new CustomerOrderGeneralInformationPage(_webDriver, _testContext);
        }

        // Tableau
        public bool AddNewItem(string itemName, string quantity)
        {
            try
            {
                // Click sur le bouton +
                _createNewItem = WaitForElementIsVisible(By.Id(NEW_ITEM));
                _createNewItem.Click();
                WaitForLoad();

                _itemName = WaitForElementIsVisible(By.CssSelector(ITEM_NAME));
                _itemName.SetValue(ControlType.TextBox, itemName);
                WaitForLoad();

                // Selection du premier élément de la liste
                _itemSelected = WaitForElementIsVisible(By.XPath(ITEM_SELECTED));
                _itemSelected.Click();
                WaitForLoad();

                if (isElementVisible(By.XPath("//*[@id=\"modal-1\"]")))
                {
                    var freePriceName = WaitForElementIsVisible(By.Id("Name"));
                    freePriceName.SetValue(ControlType.TextBox, itemName);
                    var Quantity = WaitForElementIsVisible(By.Id("Quantity"));
                    Quantity.SetValue(ControlType.TextBox, quantity);
                    var save = WaitForElementIsVisible(By.XPath("//*[@id=\"freePriceForm\"]/div[2]/button[2]"));
                    save.Click();
                    WaitForLoad();
                }

                // Temps d'attente obligatoire pour les informations du service créé se chargent
                _itemQuantity = WaitForElementIsVisible(By.CssSelector(ITEM_QUANTITY));
                _itemQuantity.SetValue(ControlType.TextBox, quantity);

                // Temps d'attente obligatoire pour la prise en compte de l'item
                WaitForElementIsVisible(By.XPath("//*[@id='dispatchTable']/tbody/tr[2]/td[15]/span[@class='fas fa-save']"));
            }
            catch
            {
                return false;
            }

            return true;
        }

        public bool AddNewItemWithCategory(string itemName, string quantity, string category)
        {
            try
            {
                // Click sur le bouton +
                _createNewItem = WaitForElementIsVisible(By.Id(NEW_ITEM));
                _createNewItem.Click();
                WaitForLoad();

                _itemName = WaitForElementIsVisible(By.CssSelector(ITEM_NAME));
                _itemName.SetValue(ControlType.TextBox, itemName);
                WaitForLoad();

                // Selection du premier élément de la liste
                _itemSelected = WaitForElementIsVisible(By.XPath(ITEM_SELECTED));
                _itemSelected.Click();
                WaitForLoad();

                // Temps d'attente obligatoire pour les informations du service créé se chargent
                _itemQuantity = WaitForElementIsVisible(By.CssSelector(ITEM_QUANTITY));
                _itemQuantity.SetValue(ControlType.TextBox, quantity);

                _itemCategory = WaitForElementIsVisible(By.XPath(ITEM_CATEGORY));
                _itemCategory.SetValue(ControlType.DropDownList, category);

                // Temps d'attente obligatoire pour la prise en compte de l'item
                WaitForElementIsVisible(By.XPath(ICON_SAVED));
            }
            catch
            {
                return false;
            }

            return true;
        }

        public CreateFreePriceModalPage AddFreePrice(string freePrice)
        {
            // Click sur le bouton +
            _createNewItem = WaitForElementIsVisible(By.Id(NEW_ITEM));
            _createNewItem.Click();
            WaitForLoad();

            _itemName = WaitForElementIsVisible(By.CssSelector(ITEM_NAME));
            _itemName.SetValue(ControlType.TextBox, freePrice);
            WaitForLoad();

            // Selection du premier élément de la liste
            var serviceSelected = WaitForElementIsVisible(By.XPath("//div[text()='" + FREE_PRICE + "']"));
            serviceSelected.Click();
            WaitForLoad();

            return new CreateFreePriceModalPage(_webDriver, _testContext);
        }

        public bool IsVisible()
        {
            WaitPageLoading();
            return isElementVisible(By.XPath(FIRST_ITEM));
        }

        public void SetItemPrice(string price)
        {
            _itemPrice = WaitForElementIsVisible(By.CssSelector(ITEM_PRICE));
            _itemPrice.SetValue(ControlType.TextBox, price);

            //WaitForElementIsVisible(By.XPath(ICON_SAVED));
            Thread.Sleep(1000);
        }

        public string GetItemPrice()
        {
            _itemPrice = WaitForElementIsVisible(By.CssSelector(ITEM_PRICE));
            return _itemPrice.GetAttribute("value");
        }
        public string CheckFreePriceAndS()
        {
            _fpcheck = WaitForElementIsVisible(By.XPath(FP_CHECK));
            return _fpcheck.Text;
        }
        public string SCheck()
        {
            _scheck = WaitForElementIsVisible(By.XPath(S_CHECK));
            return _scheck.Text;
        }
        public string GetLabel1()
        {
            _label1 = WaitForElementIsVisible(By.XPath(LABEL1));
            return _label1.GetAttribute("value");
        }
        public string GetLabel2()
        {

            _label2 = WaitForElementIsVisible(By.XPath(LABEL2));
            return _label2.GetAttribute("value");
        }
        public string GetLabel3()
        {
            _label3 = WaitForElementIsVisible(By.XPath(LABEL3));
            return _label3.GetAttribute("value");
        }
        public string GetLabel4()
        {
            _label4 = WaitForElementIsVisible(By.XPath(LABEL4));
            return _label4.GetAttribute("value");
        }
        public string GetLabel5()
        {
            _label5 = WaitForElementIsVisible(By.XPath(LABEL5));
            return _label5.GetAttribute("value");
        }

        public string GetItemPriceCurrency()
        {
            var currency = WaitForElementIsVisible(By.XPath(PRICE_CURRENCY));
            return currency.Text;
        }
        public void SetItemQuantity(string quantity)
        {
            _itemQuantity = WaitForElementIsVisible(By.CssSelector(ITEM_QUANTITY));
            _itemQuantity.SetValue(ControlType.TextBox, quantity);

            //WaitForElementIsVisible(By.XPath(ICON_SAVED));
            Thread.Sleep(1000);
        }

        public string GetItemQuantity()
        {
            _itemQuantity = WaitForElementIsVisible(By.CssSelector(ITEM_QUANTITY));
            return _itemQuantity.GetAttribute("value");
        }

        public void ClickProdComment()
        {
            _prodComment = WaitForElementIsVisible(By.XPath(PROD_COMMENT));
            _prodComment.Click();
            WaitForLoad();
        }

        public bool AddProdComment(string comment)
        {
            ClickProdComment();

            _commentArea = WaitForElementIsVisible(By.Id(COMMENT_AREA));
            _commentArea.SetValue(ControlType.TextBox, comment);

            _validateComment = WaitForElementIsVisible(By.XPath(VALIDATE_COMMENT));
            _validateComment.Click();
            WaitForLoad();

            try
            {
                _prodComment = WaitForElementIsVisible(By.XPath(PROD_COMMENT));
                if (!_prodComment.GetAttribute("class").Contains("has-comments"))
                    return false;  
            }
            catch
            {
                return false;
            }

            return true;
        }

        public void ClickBillingComment()
        {
            _billingComment = WaitForElementIsVisible(By.XPath(BILLING_COMMENT));
            _billingComment.Click();
            WaitForLoad();
        }

        public bool AddBillingComment(string comment)
        {
            ClickBillingComment();

            _commentArea = WaitForElementIsVisible(By.Id(COMMENT_AREA));
            _commentArea.SetValue(ControlType.TextBox, comment);

            _validateComment = WaitForElementIsVisible(By.XPath(VALIDATE_COMMENT));
            _validateComment.Click();
            WaitForLoad();

            try
            {
                _billingComment = WaitForElementIsVisible(By.XPath(BILLING_COMMENT));
                if (!_billingComment.GetAttribute("class").Contains("has-comments"))
                    return false;
            }
            catch
            {
                return false;
            }

            return true;
        }

        public string GetComment()
        {
            _commentArea = WaitForElementIsVisible(By.Id(COMMENT_AREA));
            return _commentArea.Text;
        }

        public void DeleteItem()
        {
            _deleteItem = WaitForElementIsVisible(By.XPath(DELETE_ITEM));
            _deleteItem.Click();
            WaitForLoad();
        }

        public void DeleteAllItem()
        {
            WaitPageLoading();  
            _deleteItem = WaitForElementIsVisible(By.Id(DELETE_ALL_ITEM));
            _deleteItem.Click();
            WaitForLoad();
        }

        public void DoNotVerify()
        {
            ShowExtendedMenu();

            var VerifyNot = WaitForElementIsVisible(By.XPath("//*[@id=\"div-body\"]/div/div[1]/div/div[2]/div/a[@class='btn btn-doNotVerify']"));
            VerifyNot.Click();
            WaitForLoad();
        }
        public void DoVerify()
        {
            ShowExtendedMenu();

            var Verify = WaitForElementIsVisible(By.XPath("//*[@id=\"div-body\"]/div/div[1]/div/div[2]/div/a[@class='btn btn-verify']"));
            Verify.Click();
            WaitForLoad();
        }

        public PrintReportPage PrintProdCustomerOrderWithTypeOfReport(bool printValue)
        {

            ShowExtendedMenu();

            var printbtn = WaitForElementIsVisible(By.XPath("//*[@id=\"Print-Group-btn\"]"));
            printbtn.Click();

            _printProd = WaitForElementIsVisible(By.Id(PRINTPROD));
            _printProd.Click();
            WaitForLoad();

            _selectedtyperaport = WaitForElementIsVisible(By.XPath(SELECTED_TYPE_RAPPORT));
            _selectedtyperaport.Click();
            WaitForLoad();


            _selectedfirsttyperaport = WaitForElementIsVisible(By.XPath(SELECTED_FIRST_TYPE_RAPPORT));
            _selectedfirsttyperaport.Click();
            WaitForLoad();

            _confirmPrintProd = WaitForElementIsVisible(By.XPath(CONFIRM_PRINTPROD));
            _confirmPrintProd.Click();
            WaitForLoad();

            if (printValue)
            {
                IsFileLoaded(By.CssSelector("[target='_blank'][class='far fa-file-pdf']"));
                ClickPrintButton();
            }

            //Results are opened in a new tab, switch the driver to the newly created one
            WebDriverWait wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(60));
            wait.Until((driver) => driver.WindowHandles.Count > 1);
            _webDriver.SwitchTo().Window(_webDriver.WindowHandles[1]);

            return new PrintReportPage(_webDriver, _testContext);
        }
        public bool AddNewMiltipleItemWithCategory(List<string> itemNames, string quantity, string category)
        {
            int i = 0;
            foreach (var itemName in itemNames)
            {
                i++;
                // Click sur le bouton +
                _createNewItem = WaitForElementIsVisible(By.Id(NEW_ITEM));
                _createNewItem.Click();
                WaitForLoad();

                if (isElementExists(By.XPath(string.Format("(//input[@placeholder='Item' and not(@value)])[{0}]", i))))
                {
                    _itemName = WaitForElementIsVisible(By.XPath(string.Format("(//input[@placeholder='Item' and not(@value)])[{0}]", i)));
                _itemName.SetValue(ControlType.TextBox, itemName);
                    WaitForLoad();
                }
                else
                {
                    _itemName = WaitForElementIsVisible(By.XPath($"(//input[@placeholder='Item' and @value!='{itemName}'])[{i}]"));
                    _itemName.SetValue(ControlType.TextBox, itemName);
                    WaitForLoad();
                }
                    
                _itemSelected = WaitForElementIsVisible(By.XPath(String.Format("(//*[@id=\"dispatchTable\"]/tbody/tr[*]/td[2]/span/span/div/div/div[1][contains(text(), '{0}')])[last()]", itemName)));
                _itemSelected.Click();
                WaitForLoad();

                _itemQuantity = WaitForElementIsVisible(By.XPath(string.Format("(//input[@placeholder='Quantity' and @value='0'])[{0}]",i)));
                _itemQuantity.SetValue(ControlType.TextBox, quantity);
                WaitForLoad();

                _itemCategory = WaitForElementIsVisible(By.XPath(string.Format("(//*/select[@class='serviceCategoryId'])[{0}]", i)));
                _itemCategory.SetValue(ControlType.DropDownList, category);
                WaitForLoad();

                // disquette
                Thread.Sleep(2000);
            }

            return true;
        }

        public void SendEmailPopUp()
        {
            ShowExtendedMenu();
            _sendEmail = WaitForElementIsVisible(By.Id(SEND_EMAIL));
            _sendEmail.Click();
            WaitForLoad();

        }

        public void SetEmailTo(string email)
        {
            _sendToEmailInput = WaitForElementIsVisible(By.Id(SEND_TO_EMAIL_INPUT));
            _sendToEmailInput.SetValue(ControlType.TextBox, email);
            WaitForLoad();

        }

        public void ChooseFile(string filePath)
        {
            var inputChooseFile = WaitForElementIsVisible(By.Id(INPUT_CHOOSE_FILE));
            inputChooseFile.SendKeys(filePath);
            WaitForLoad();

        }

        public void SendEmailButton()
        {
            _sendEmailButton = WaitForElementIsVisible(By.Id(SEND_EMAIL_BUTTON));
            _sendEmailButton.Click();
            WaitPageLoading();

        }

        public int GetNumberAttachementFile()
        {

            return _webDriver.FindElements(By.XPath(ATTACHMENTS_FILE)).Count;
        }
        public void SetNameItem(string name)
        {
            _createNewItem = WaitForElementIsVisible(By.Id(NEW_ITEM));
            _createNewItem.Click();
            WaitForLoad();
            _itemName = WaitForElementIsVisible(By.CssSelector(ITEM_NAME));
            _itemName.SetValue(ControlType.TextBox, name);
            WaitForLoad();
        }
        public List<string> GetSuggestionsMenuNameList()
        {
            List<string> suggestionsList = new List<string>();
            if (isElementVisible(By.XPath(MENU_SUGGESTION)))
            {
                var suggestionElements = _webDriver.FindElements(By.XPath(SUGGESTION_ELEMENTS));
                foreach (var element in suggestionElements)
                {
                    string itemName = element.Text.Trim();
                    if (itemName.StartsWith("HOT"))
                    {
                        suggestionsList.Add(itemName);
                    }
                }
            }

            return suggestionsList;
        }
        public bool IsOnCorrectPage()
        {
            return _webDriver.Url.Contains("Production/Order/Detail");
        }
        public bool AddItemDetails(string itemName, string quantity)
        {
            try
            {
                // Click sur le bouton +
                _createNewItem = WaitForElementIsVisible(By.Id(NEW_ITEM));
                _createNewItem.Click();
                WaitForLoad();

                _itemName = WaitForElementIsVisible(By.CssSelector(ITEM_NAME));
                _itemName.SetValue(ControlType.TextBox, itemName);
                WaitForLoad();

                // Selection du premier élément de la liste
                _itemSelected = WaitForElementIsVisible(By.XPath(ITEM_SELECTED));
                _itemSelected.Click();
                WaitForLoad();
    
                _itemQuantity = WaitForElementIsVisible(By.CssSelector(ITEM_QUANTITY));
                _itemQuantity.SetValue(ControlType.TextBox, quantity);
                WaitPageLoading();
            }
            catch
            {
                return false;
            }

            return true;
        }
        public void ValidateCO()
        {
            ShowValidationMenu();
            _validate = WaitForElementIsVisible(By.Id(VALIDATE));
            WaitForLoad();
            _validate.Click();
            _validate = WaitForElementIsVisible(By.Id(VALIDATE_Button));
            _validate.Click();
            WaitPageLoading();
        }
        public string GetOrderName()
        {
            IWebElement orderNameElement = _webDriver.FindElement(By.XPath(ORDER_NAME));

            return orderNameElement.Text.Trim();
        }
        public string GetOrderCategory()
        {
            IWebElement orderCategoryElement = _webDriver.FindElement(By.XPath(ORDER_CATEGORY));

            return orderCategoryElement.Text.Trim();
        }
        public int GetOrderQuantity()
        {
            IWebElement orderQuantityElement = _webDriver.FindElement(By.XPath(ORDER_QUANTITY));

            // Extract the text of the order quantity
            return int.Parse(orderQuantityElement.Text.Trim());
        }
        public string GetNumberCustomer()
        {
            var CustomerInput = WaitForElementIsVisible(By.XPath("/html/body/div[3]/div/div[1]/h1"));
            string pattern = @"\d+";

            Match match = Regex.Match(CustomerInput.Text, pattern);

            if (match.Success)
            {
                string orderNumber = match.Value;
                return orderNumber;
            }
            else
            {
                return string.Empty;
            }
        }
        public bool IsVerifyButtonVisible()
        {
           
                try
                {
                    ShowExtendedMenu();

                    var verifyButton = WaitForElementIsVisible(By.XPath(VERIFY_BUTTON));

                    if (verifyButton.Displayed)
                    {
                        WaitForLoad();
                        return true;  
                    }
                    else
                    {
                        Console.WriteLine("Le bouton 'Verify' n'est pas visible.");
                        return false; 
                    }
                }
                catch (NoSuchElementException)
                {
                    Console.WriteLine("Le bouton 'Verify' n'a pas été trouvé sur la page.");
                    return false;  
                }
                catch (ElementNotVisibleException)
                {
                    Console.WriteLine("Le bouton 'Verify' est présent mais pas visible.");
                    return false;                 
            }
         }

        public bool IsVATexcl()
        {
            var _vat = WaitForElementExists(By.XPath(VAT)) ; 
            if ( _vat.Text == "Excl.")
             return true; 
            else
              return false; 
        }

        public bool CheckCommercialManagerEmail(string mail)
        {
            _ccAddresses = WaitForElementIsVisible(By.Id(CC_ADDRESSES));
            string value = _ccAddresses.GetAttribute("value");

            return value.Contains(mail);
        }

        public void CloseCommentModal()
        {
            _closeBtn = WaitForElementIsVisible(By.XPath(CLOSE_BTN));
            _closeBtn.Click();
            WaitForLoad();
        }

        public void SelectSellingUnit(string sellingUnit)
        {
            var sellingUnitField = WaitForElementExists(By.XPath("//*[@name=\"UnitId\"]"));
            var selectSellingUnit = new SelectElement(sellingUnitField);
            selectSellingUnit.SelectByText(sellingUnit);
            WaitForLoad();
        }

        public string GetSellingUnit()
        {
            var options = _webDriver.FindElements(By.XPath("//*[@name=\"UnitId\"]/option"));
            var selectedOption = options.FirstOrDefault(op => op.Selected);
            return selectedOption?.Text;
        }
        
        public bool InteractableSellingUnit()
        {
            _sellingUnit = WaitForElementIsVisible(By.XPath(NON_MODIFIABLE_SELLING_UNIT));
            return _sellingUnit!=null;
        }

        public void SetItemPriceTotal(string totalPrice)
        {
            _totalPrice = WaitForElementIsVisible(By.Name("TotalPrice"));
            _totalPrice.SetValue(ControlType.TextBox, totalPrice);
            LoadingPage();

        }
        public string GetPriceExcl()
        {
            // Attendre que l'élément soit visible
            WebDriverWait wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(10));
            IWebElement priceExclElement = wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath(PRICE_EXCL)));
            string rawText = priceExclElement.Text.Trim();
            int euroIndex = rawText.IndexOf('€');
            return rawText.Substring(euroIndex + 1).Trim();
        }


    }
}
