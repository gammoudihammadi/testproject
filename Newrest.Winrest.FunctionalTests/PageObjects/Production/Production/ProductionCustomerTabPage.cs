using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;
using System;
using System.Collections.Generic;
using System.Threading;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Production.Production
{
    public class MappedCustomerAndRecipePax
    {
        public string customerPAX { get; set; }
        public string recipePAX { get; set; }
    }

    public class ProductionCustomerTabPage : PageBase
    {
        public ProductionCustomerTabPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        // ______________________________________ Constantes _____________________________________________
        private const string MENU_TAB = "//*[@id=\"itemTabTab\"]/li[1]";

        private const string FOLD_UNFOLD_ALL = "//*[@id=\"tabContentItemContainer\"]/div[1]/div/div/div/a[1]";
        private const string CUSTOMERNAME_COLUMN_TITLE = "//*[@id=\"list-item-with-action\"]/div[1]/div/div/div[2]/table/thead/tr/th[1]";
        private const string CUSTOMERNAME_COLUMN = "//*[@id=\"list-item-with-action\"]/div[*]/div[1]/div/div[2]/table/tbody/tr/td[1]";
        private const string FIRST_CUSTOMER_NAME = "//*[@id=\"list-item-with-action\"]/div[2]/div[1]/div/div[2]/table/tbody/tr/td[1]";
        private const string CURRENT_CUSTOMER_NAME = "//*[@id=\"list-item-with-action\"]/div[{0}]/div[1]/div/div[2]/table/tbody/tr/td[1]";
        private const string CURRENT_CUSTOMER_TOGGLE_BTN = "//*[@id=\"list-item-with-action\"]/div[{0}]/div[1]/div/div[1]";
        private const string CURRENT_CUSTOMER_MENU = "/html/body/div[3]/div/div[2]/div[2]/div/div/div/div[{0}]/div[2]/div/div/table/tbody/tr/td[3]";
        private const string FIRST_RECIPE_NAME = "//*[@id=\"list-item-with-action\"]/div[2]/div[2]/div/div/table/tbody/tr/td[1]";
        private const string CURRENT_CUSTOMER_PAX = "/html/body/div[3]/div/div[2]/div[2]/div/div/div/div[{0}]/div[2]/div/div/table/tbody/tr/td[3]";
        private const string CURRENT_CUSTOMER_QTYTOPRODUCE = "/html/body/div[3]/div/div[2]/div[2]/div/div/div/div[{0}]/div[2]/div/div/table/tbody/tr/td[5]";
        private const string CURRENT_CUSTOMER_RECIPE_QTYTOPRODUCE = "//td[1][contains(text(),'{0}')]//ancestor::div[@class='panel panel-default']//td[2][contains(text(),'{1}')]/..//span[contains(text(),'Qty to produce')]/parent::td";
        private const string CURRENT_CUSTOMER_RECIPE_PAX = "/html/body/div[3]/div/div[2]/div[2]/div/div/div/div[{0}]/div[2]/div/div/table/tbody/tr/td[4]";
        private const string CURRENT_CUSTOMER_RECIPE_ITEM_NETWEIGHT = "//*[@id=\"list-item-with-action\"]/div[2]/div[1]/div/div[2]/table/tbody/tr/td[contains(text(),'{0}')]//ancestor::div[@class='panel panel-default']//child::span[@class='itemWeight']";

        // ______________________________________ Variables _____________________________________________

        [FindsBy(How = How.XPath, Using = MENU_TAB)]
        private IWebElement _menuTab;

        [FindsBy(How = How.XPath, Using = FOLD_UNFOLD_ALL)]
        private IWebElement _foldUnfoldAll;

        [FindsBy(How = How.XPath, Using = CUSTOMERNAME_COLUMN_TITLE)]
        private IWebElement _columnTitle;

        [FindsBy(How = How.XPath, Using = FIRST_CUSTOMER_NAME)]
        private IWebElement _firstCustomerName;
        
        [FindsBy(How = How.XPath, Using = CURRENT_CUSTOMER_NAME)]
        private IWebElement _currentCustomerName;
        
        [FindsBy(How = How.XPath, Using = CURRENT_CUSTOMER_TOGGLE_BTN)]
        private IWebElement _currentCustomerToggleBtn;
        
        [FindsBy(How = How.XPath, Using = CURRENT_CUSTOMER_MENU)]
        private IWebElement _currentCustomerMenu;

        [FindsBy(How = How.XPath, Using = FIRST_RECIPE_NAME)]
        private IWebElement _firstRecipeName;

        [FindsBy(How = How.XPath, Using = CURRENT_CUSTOMER_QTYTOPRODUCE)]
        private IWebElement _currentCustomerQty;

        [FindsBy(How = How.XPath, Using = CURRENT_CUSTOMER_RECIPE_QTYTOPRODUCE)]
        private IWebElement _currentCustomerRecipeQty;

        [FindsBy(How = How.XPath, Using = CURRENT_CUSTOMER_PAX)]
        private IWebElement _currentCustomerPAX;

        [FindsBy(How = How.XPath, Using = CURRENT_CUSTOMER_RECIPE_PAX)]
        private IWebElement _currentCustomerRecipePAX;

        [FindsBy(How = How.XPath, Using = CURRENT_CUSTOMER_RECIPE_ITEM_NETWEIGHT)]
        private IWebElement _currentCustomerRecipeItemNetWeight;

        //ONGLETS
        public ProductionSearchResultsMenuTabPage GoToProductionMenuTab()
        {
            _menuTab = WaitForElementIsVisible(By.XPath(MENU_TAB));
            _menuTab.Click();
            WaitForLoad();
            return new ProductionSearchResultsMenuTabPage(_webDriver, _testContext);
        }

        public bool IsResultDisplayedByCustomer()
        {
            _columnTitle = WaitForElementIsVisible(By.XPath(CUSTOMERNAME_COLUMN_TITLE));
            _firstCustomerName = WaitForElementIsVisible(By.XPath(FIRST_CUSTOMER_NAME));

            if (_columnTitle.Text.Equals("Customer name") && _firstCustomerName.Text.Contains("Customer"))
            {
                return true;

            } else
            {
                return false;
            }
        }

        public string GetFirstCustomerName()
        {
            _firstCustomerName = WaitForElementIsVisible(By.XPath(FIRST_CUSTOMER_NAME));
            return _firstCustomerName.Text;
        }

        public Dictionary<string, string> GetCustomersNamesAndMenusQties()
        {
            int i = 0;

            Dictionary<string, string> customersNames = new Dictionary<string, string>();

            var columnCustomersNames = _webDriver.FindElements(By.XPath(CUSTOMERNAME_COLUMN));

            foreach (var customer in columnCustomersNames)
            {
                // On limite le nombre de menus remontés à 5 pour ne pas surcharger le test
                if (i >= 4)
                    break;

                _currentCustomerMenu = _webDriver.FindElement(By.XPath(String.Format(CURRENT_CUSTOMER_MENU, i + 2)));


                customersNames.Add(customer.Text, _currentCustomerMenu.Text);
                i++;
            }

            return customersNames;
        }

        // OUTPUT : <customerName, qtyToProduce>>
        public Dictionary<string, string> GetCustomersNamesAndQties()
        {
            int i = 0;

            Dictionary<string, string> customerNamesAndMappedMenusQties = new Dictionary<string, string>();

            var columnCustomersNames = _webDriver.FindElements(By.XPath(CUSTOMERNAME_COLUMN));

            foreach (var customer in columnCustomersNames)
            {
                // On limite le nombre de menus remontés à 3 pour ne pas surcharger le test
                if (i >= 8)
                    break;

                _currentCustomerToggleBtn = _webDriver.FindElement(By.XPath(String.Format(CURRENT_CUSTOMER_TOGGLE_BTN, i + 2)));
                _currentCustomerToggleBtn.Click();
                Thread.Sleep(1000);
                WaitForLoad();

                _currentCustomerQty = _webDriver.FindElement(By.XPath(String.Format(CURRENT_CUSTOMER_QTYTOPRODUCE, i + 2)));

                var qtyString = _currentCustomerQty.GetAttribute("outerText").Substring("Qty to produce : ".Length);


                customerNamesAndMappedMenusQties.Add(customer.Text, qtyString);

                i++;
            }

            return customerNamesAndMappedMenusQties;
        }

        public Dictionary<string, MappedCustomerAndRecipePax> GetCustomersPAX()
        {
            int i = 0;

            Dictionary<string, MappedCustomerAndRecipePax> cutomersPAX = new Dictionary<string, MappedCustomerAndRecipePax>();

            var columnCustomersNames = _webDriver.FindElements(By.XPath(CUSTOMERNAME_COLUMN));

            var MappedCustomerAndRecipePax = new MappedCustomerAndRecipePax();

            foreach (var customer in columnCustomersNames)
            {
                // On limite le nombre de menus remontés à 3 pour ne pas surcharger le test
                if (i >= 3)
                    break;

                _currentCustomerToggleBtn = _webDriver.FindElement(By.XPath(String.Format(CURRENT_CUSTOMER_TOGGLE_BTN, i + 2)));
                _currentCustomerToggleBtn.Click();
                Thread.Sleep(1000);
                WaitForLoad();

                _currentCustomerPAX = _webDriver.FindElement(By.XPath(String.Format(CURRENT_CUSTOMER_RECIPE_PAX, i + 2)));
                MappedCustomerAndRecipePax.customerPAX = _currentCustomerPAX.Text;

                _currentCustomerRecipePAX = _webDriver.FindElement(By.XPath(String.Format(CURRENT_CUSTOMER_RECIPE_PAX, i + 2)));
                MappedCustomerAndRecipePax.recipePAX = _currentCustomerRecipePAX.Text;

                cutomersPAX.Add(customer.Text, MappedCustomerAndRecipePax);
                i++;
            }

            return cutomersPAX;
        }

        // INPUT : Customer Name - OUTPUT : Item Net Weight
        public string GetCustomerRecipeItemNetWeight(string customerName)
        {
            FoldUnfoldAll();
            WaitForLoad();

            _currentCustomerRecipeItemNetWeight = _webDriver.FindElement(By.XPath(String.Format(CURRENT_CUSTOMER_RECIPE_ITEM_NETWEIGHT, customerName)));
            return _currentCustomerRecipeItemNetWeight.Text;
        }

        // INPUT : Customer Name - OUTPUT : Item Weight
        public string GetCustomerRecipeItemWeight(string customerName)
        {
            var foldOrUnfold = _webDriver.FindElement(By.XPath("//*[@id=\"tabContentItemContainer\"]/div[1]/div/div/div/a[1]"));
            if (foldOrUnfold.Text != "Fold All")
            {
                FoldUnfoldAll();
                WaitForLoad();
            }

            var currentCustomerRecipeItemWeights = _webDriver.FindElements(By.XPath(String.Format(CURRENT_CUSTOMER_RECIPE_ITEM_NETWEIGHT, customerName)));
            return currentCustomerRecipeItemWeights[1].Text; //[0] --> net weight [1] weight
        }

        // INPUT : Customer Name & Recipe Name - OUTPUT : Quantity to produce
        public string GetCustomerRecipeQuantityToProduce(string customerName, string recipe)
        {
            ShowExtendedMenu();
            var foldOrUnfold = _webDriver.FindElement(By.XPath("//*[@id=\"tabContentItemContainer\"]/div[1]/div/div/div/a[1]"));
            if (foldOrUnfold.Text != "Fold All")
            {
                FoldUnfoldAll();
                WaitForLoad();
            }

            _currentCustomerRecipeQty = _webDriver.FindElement(By.XPath(String.Format(CURRENT_CUSTOMER_RECIPE_QTYTOPRODUCE, customerName, recipe)));
            return _currentCustomerRecipeQty.Text;

        }

        public bool FoldUnfoldAll()
        {
            ShowExtendedMenu();
            _foldUnfoldAll = WaitForElementIsVisible(By.XPath(FOLD_UNFOLD_ALL));
            _foldUnfoldAll.Click();
            //Temps d'affichage du résultat
            WaitPageLoading();
            if (isElementVisible(By.XPath(FIRST_RECIPE_NAME)))
            {
                _firstRecipeName = WaitForElementIsVisible(By.XPath(FIRST_RECIPE_NAME));
                return true;
            }
            else
            {
                return false;
            }
        }
    }
}
