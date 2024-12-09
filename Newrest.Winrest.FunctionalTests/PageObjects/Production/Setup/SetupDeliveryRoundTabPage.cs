using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Menus.Recipes;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Threading;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Production.Setup
{
    public class MappedRecipeAndQuantity
    {
        public string DeliveryRoundName { get; set; }
        public string CustomerName { get; set; }
        public string DeliveryName { get; set; }
        public string RecipeName { get; set; }
        public string MenuVariant { get; set; }
        public int QtyToProduce { get; set; }
        public int WeightPerPack { get; set; }
        public int Total { get; set; }
    }
    public class SetupDeliveryRoundTabPage : PageBase
    {
        public SetupDeliveryRoundTabPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        // ______________________________________ Constantes _____________________________________________

        // Tableau Delivery Round
        private const string COLUMN_DELIVERYROUND_NAMES = "//*[@id=\"list-item-with-action\"]/div[*]/div[1]/div/div[2]/table/tbody/tr/td[1]";
        private const string COLUMN_DELIVERYROUND_TOGGLEBUTTON = "//*[@id=\"list-item-with-action\"]/div[{0}]/div[1]/div/div[1]";
        private const string COLUMN_DELIVERYROUND_QTYPRODUCE = "/html/body/div[2]/div/div[2]/div[2]/div/div/div[2]/div[2]/div[2]/div/div/table/tbody/tr[1]/td[4]/div[2]";
        private const string COLUMN_DELIVERYROUND_WEIGHTPERPACK = "/html/body/div[2]/div/div[2]/div[2]/div/div/div[2]/div[2]/div[2]/div/div/table/tbody/tr[1]/td[5]/div[2]";
        private const string COLUMN_DELIVERYROUND_TOTAL = "/html/body/div[2]/div/div[2]/div[2]/div/div/div[2]/div[2]/div[2]/div/div/table/tbody/tr[1]/td[6]/div[2]";
        private const string COLUMN_DELIVERYROUND_RECIPE = "/html/body/div[2]/div/div[2]/div[2]/div/div/div[2]/div[2]/div[2]/div/div/table/tbody/tr[1]/td[2]";
        private const string RECIPE_LINK = "/html/body/div[2]/div/div[2]/div[2]/div/div/div[2]/div[2]/div[2]/div/div/table/tbody/tr/td[7]/a/span[@class='pencil']";
        private const string RECIPE = "//*[@id=\"list-item-with-action\"]/div/div[2]/div";
        private const string DELIVERYROUND_SERVICE_RECIPE_NAME = "//td[1][contains(text(),'{0}')]/ancestor::div[@class='panel panel-default']//td[1][contains(text(),'{1}')]/../td[2]";
        private const string DELIVERYROUND_SERVICE_PAX = "//td[1][contains(text(),'{0}')]/ancestor::div[@class='panel panel-default']//td[1][contains(text(),'{1}')]/..//span[contains(text(),'PAX')]/parent::td";
        private const string DELIVERYROUND_SERVICE_QUANTITY_TO_PRODUCE = "//td[1][contains(text(),'{0}')]/ancestor::div[@class='panel panel-default']//td[1][contains(text(),'{1}')]/..//div[contains(text(),'Qty to produce')]/parent::td";
        private const string DELIVERYROUND_SERVICE_WEIGHT_PER_PACK = "//td[1][contains(text(),'{0}')]/ancestor::div[@class='panel panel-default']//td[1][contains(text(),'{1}')]/..//div[contains(text(),'Weight per pack')]/parent::td";
        private const string DELIVERYROUND_SERVICE_TOTAL_PER_PACK = "//td[1][contains(text(),'{0}')]/ancestor::div[@class='panel panel-default']//td[1][contains(text(),'{1}')]/..//div[contains(text(),'Total')]/parent::td";
        private const string FOLD_UNFOLD_ALL = "/html/body/div[2]/div/div[2]/div[1]/div/div/div/a[1]";
        private const string FOLD_UNFOLD_BUTTON_LABEL = "//*[@id=\"tabContentItemContainer\"]/div[1]/div/div/div/a[1]";

        // ______________________________________ Variables _____________________________________________

        // Tableau Delivery Round

        [FindsBy(How = How.XPath, Using = COLUMN_DELIVERYROUND_TOGGLEBUTTON)]
        private IWebElement _deliveryRoundColumnCurrentToggleButton;

        [FindsBy(How = How.XPath, Using = COLUMN_DELIVERYROUND_QTYPRODUCE)]
        private IWebElement _deliveryRoundColumnCurrentQtyProduce;

        [FindsBy(How = How.XPath, Using = COLUMN_DELIVERYROUND_WEIGHTPERPACK)]
        private IWebElement _deliveryRoundColumnWeightPerPack;

        [FindsBy(How = How.XPath, Using = COLUMN_DELIVERYROUND_TOTAL)]
        private IWebElement _deliveryRoundColumnTotal;

        [FindsBy(How = How.XPath, Using = COLUMN_DELIVERYROUND_RECIPE)]
        private IWebElement _deliveryRoundColumnRecipeName;

        [FindsBy(How = How.XPath, Using = RECIPE)]
        private IWebElement _recipe;

        [FindsBy(How = How.XPath, Using = RECIPE_LINK)]
        private IWebElement _recipeLink;
        
        //FOLD/UNFOLD
        [FindsBy(How = How.XPath, Using = FOLD_UNFOLD_ALL)]
        private IWebElement _foldUnfoldAll;

        [FindsBy(How = How.XPath, Using = FOLD_UNFOLD_BUTTON_LABEL)]
        private IWebElement _foldUnfoldButtonLabel;

        public MappedRecipeAndQuantity GetMenuNameAndQtyForOneDeliveryRound()
        {
            int i = 0;

            var columnDeliveryRoundName = _webDriver.FindElement(By.XPath(COLUMN_DELIVERYROUND_NAMES));

            _deliveryRoundColumnCurrentToggleButton = _webDriver.FindElement(By.XPath(string.Format(COLUMN_DELIVERYROUND_TOGGLEBUTTON, i + 2)));
            _deliveryRoundColumnCurrentToggleButton.Click();
            WaitForLoad();

            _recipe = WaitForElementIsVisible(By.XPath(RECIPE));
            _recipe.Click();

            var mappedDeliveryAndQty = new MappedRecipeAndQuantity();

            var deliveryRoundLineText = columnDeliveryRoundName.Text;
            var deliveryRoundNameEndIndex = deliveryRoundLineText.IndexOf('-');
            mappedDeliveryAndQty.DeliveryRoundName = deliveryRoundLineText.Substring(0, deliveryRoundNameEndIndex - 1);

            var customerAndDeliveryLineText = deliveryRoundLineText.Substring(deliveryRoundNameEndIndex + 2);
            var customerNameEndIndex = customerAndDeliveryLineText.IndexOf('-') + 1;
            mappedDeliveryAndQty.CustomerName = customerAndDeliveryLineText.Substring(0, customerNameEndIndex - 2);

            mappedDeliveryAndQty.DeliveryName = customerAndDeliveryLineText.Substring(customerNameEndIndex + 1);
            WaitForLoad();

            _deliveryRoundColumnRecipeName = WaitForElementIsVisible(By.XPath(string.Format(COLUMN_DELIVERYROUND_RECIPE, i + 2)));
            if (_deliveryRoundColumnRecipeName.Text.Contains("?"))
            {

                mappedDeliveryAndQty.RecipeName = _deliveryRoundColumnRecipeName.Text.Substring(0, _deliveryRoundColumnRecipeName.Text.IndexOf(" ?\r")).ToString();
            }
            else
            {
                mappedDeliveryAndQty.RecipeName = _deliveryRoundColumnRecipeName.Text.Substring(0, _deliveryRoundColumnRecipeName.Text.IndexOf("\r")).ToString();
            }
            mappedDeliveryAndQty.MenuVariant = _deliveryRoundColumnRecipeName.Text.Substring(_deliveryRoundColumnRecipeName.Text.IndexOf("\r\n") + 2);

            _deliveryRoundColumnCurrentQtyProduce = _webDriver.FindElement(By.XPath(string.Format(COLUMN_DELIVERYROUND_QTYPRODUCE, i + 2)));
            if (_deliveryRoundColumnCurrentQtyProduce.Text.Length > 0)
            {
                var qtyString = _deliveryRoundColumnCurrentQtyProduce.Text.Substring(0, _deliveryRoundColumnCurrentQtyProduce.Text.IndexOf('x')-1).Trim('(', ')');
                Int32.TryParse(qtyString, out int qtyInt);
                mappedDeliveryAndQty.QtyToProduce = qtyInt;
            }

            _deliveryRoundColumnWeightPerPack = _webDriver.FindElement(By.XPath(string.Format(COLUMN_DELIVERYROUND_WEIGHTPERPACK, i + 2)));
            if (_deliveryRoundColumnWeightPerPack.Text.Length > 0)
            {
                var weightPerPackString = _deliveryRoundColumnWeightPerPack.Text.Substring(0, _deliveryRoundColumnWeightPerPack.Text.IndexOf('g') - 1);
                Int32.TryParse(weightPerPackString, out int weightPerPackInt);
                mappedDeliveryAndQty.WeightPerPack = weightPerPackInt;
            }

            _deliveryRoundColumnTotal = _webDriver.FindElement(By.XPath(string.Format(COLUMN_DELIVERYROUND_TOTAL, i + 2)));
            if (_deliveryRoundColumnTotal.Text.Length > 0)
            {
                var totalString = _deliveryRoundColumnTotal.Text.Substring(0, _deliveryRoundColumnTotal.Text.IndexOf('g') - 1);
                Int32.TryParse(totalString, out int totalInt);
                mappedDeliveryAndQty.Total = totalInt;
            }

            return mappedDeliveryAndQty;
        }

        public Dictionary<string, MappedRecipeAndQuantity> GetMenusNamesAndQtyForAllDeliveryRound()
        {
            int i = 0;

            Dictionary<string, MappedRecipeAndQuantity> menusNamesAndQty = new Dictionary<string, MappedRecipeAndQuantity>();

            var columnDeliveryRoundName = _webDriver.FindElements(By.XPath(COLUMN_DELIVERYROUND_NAMES));

            foreach (var menu in columnDeliveryRoundName)
            {
                // On limite le nombre de menus remontés à 5 pour ne pas surcharger le test
                if (i >= 4)
                    break;

                _deliveryRoundColumnCurrentToggleButton = _webDriver.FindElement(By.XPath(String.Format(COLUMN_DELIVERYROUND_TOGGLEBUTTON, i + 2)));
                _deliveryRoundColumnCurrentToggleButton.Click();
                WaitForLoad();

                _recipe = WaitForElementIsVisible(By.XPath(RECIPE));
                _recipe.Click();

                var mappedDeliveryAndQty = new MappedRecipeAndQuantity();

                var deliveryRoundLineText = menu.Text;
                var deliveryRoundNameEndIndex = deliveryRoundLineText.IndexOf('-');
                mappedDeliveryAndQty.DeliveryRoundName = deliveryRoundLineText.Substring(0, deliveryRoundNameEndIndex - 1);

                var customerAndDeliveryLineText = deliveryRoundLineText.Substring(deliveryRoundNameEndIndex + 2);
                var customerNameEndIndex = customerAndDeliveryLineText.IndexOf('-') + 1;
                mappedDeliveryAndQty.CustomerName = customerAndDeliveryLineText.Substring(0, customerNameEndIndex - 2);

                mappedDeliveryAndQty.DeliveryName = customerAndDeliveryLineText.Substring(customerNameEndIndex + 1);

                _deliveryRoundColumnRecipeName = WaitForElementIsVisible(By.XPath(String.Format(COLUMN_DELIVERYROUND_RECIPE, i + 2)));
                if (_deliveryRoundColumnRecipeName.Text.Contains("?"))
                {

                    mappedDeliveryAndQty.RecipeName = _deliveryRoundColumnRecipeName.Text.Substring(0, _deliveryRoundColumnRecipeName.Text.IndexOf(" ?\r")).ToString();
                }
                else
                {
                    mappedDeliveryAndQty.RecipeName = _deliveryRoundColumnRecipeName.Text.Substring(0, _deliveryRoundColumnRecipeName.Text.IndexOf("\r")).ToString();
                }
                mappedDeliveryAndQty.MenuVariant = _deliveryRoundColumnRecipeName.Text.Substring(_deliveryRoundColumnRecipeName.Text.IndexOf("\r\n") + 2);

                WaitForLoad();
                menusNamesAndQty.Add(menu.Text, mappedDeliveryAndQty);
                i++;
            }

            return menusNamesAndQty;
        }

        public RecipesVariantPage EditRecipe()
        {
            _recipe = WaitForElementIsVisible(By.XPath(RECIPE));
            _recipe.Click();

            _recipeLink = WaitForElementIsVisible(By.XPath("//*/span[contains(@class,'pencil')]"));
            _recipeLink.Click();
            WaitForLoad();

            var wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(60));
            wait.Until((driver) => driver.WindowHandles.Count > 1);
            _webDriver.SwitchTo().Window(_webDriver.WindowHandles[1]);

            return new RecipesVariantPage(_webDriver, _testContext);
        }


        // Retourne vrai si les éléments du tableau sont dépliés
        // Retourne faux si les éléments du tableau sont repliés
        public bool FoldUnfoldAll()
        {
            ShowExtendedMenu();
            WaitPageLoading();

            _foldUnfoldAll = WaitForElementIsVisible(By.XPath(FOLD_UNFOLD_ALL));
            _foldUnfoldAll.Click();
            //Temps d'affichage du résultat
            Thread.Sleep(1000);

            try
            {
                var columnDeliveryRoundName = _webDriver.FindElements(By.XPath(COLUMN_DELIVERYROUND_NAMES));

                foreach (var menu in columnDeliveryRoundName)
                {
                    _recipe = WaitForElementIsVisible(By.XPath(RECIPE));
                }

                return true;
            }
            catch
            {
                return false;
            }
        }

        // INPUT : Delivery round name & service name - OUTPUT : recipe name
        public List<string> GetDeliveryRoundServiceRecipeName(string deliveryRoundName, string serviceName)
        {
            ShowExtendedMenu();
            _foldUnfoldButtonLabel = _webDriver.FindElement(By.XPath(FOLD_UNFOLD_BUTTON_LABEL));
            if (_foldUnfoldButtonLabel.Text != "Fold All")
            {
                FoldUnfoldAll();
                WaitForLoad();
            }

            List<string> deliveryRoundServiceRecipeNamesList = new List<string>();
            var deliveryRoundServiceRecipeNameElements = _webDriver.FindElements(By.XPath(String.Format(DELIVERYROUND_SERVICE_RECIPE_NAME, deliveryRoundName, serviceName)));
            if(deliveryRoundServiceRecipeNameElements.Count != 0)
            {
                foreach (var element in deliveryRoundServiceRecipeNameElements)
                {
                    string recipeName;
                    if (element.Text.Contains("?"))
                    {
                        recipeName = element.Text.Substring(0, element.Text.IndexOf(" ?"));
                    }
                    else if (element.Text.Contains("\r\n"))
                    {
                        recipeName = element.Text.Substring(0, element.Text.IndexOf("\r\n"));
                    }
                    else
                    {
                        recipeName = element.Text;
                    }
                    deliveryRoundServiceRecipeNamesList.Add(recipeName);
                }
            }
           

            return deliveryRoundServiceRecipeNamesList;
        }

        // INPUT : Delivery round name & service name - OUTPUT : PAX
        public List<string> GetDeliveryRoundServicePAX(string deliveryRoundName, string serviceName)
        {
            ShowExtendedMenu();
            _foldUnfoldButtonLabel = _webDriver.FindElement(By.XPath(FOLD_UNFOLD_BUTTON_LABEL));
            if (_foldUnfoldButtonLabel.Text != "Fold All")
            {
                FoldUnfoldAll();
                WaitForLoad();
            }

            List<string> deliveryRoundServicePaxList = new List<string>();
            var deliveryRoundServicePaxElements = _webDriver.FindElements(By.XPath(String.Format(DELIVERYROUND_SERVICE_PAX, deliveryRoundName, serviceName)));
            if(deliveryRoundServicePaxElements.Count != 0)
            {
                foreach (var element in deliveryRoundServicePaxElements)
                {
                    var pax = element.Text;
                    deliveryRoundServicePaxList.Add(pax);
                }

            }
           
            return deliveryRoundServicePaxList;
        }

        // INPUT : Delivery round name & service name - OUTPUT : quantity to produce list
        public List<string> GetDeliveryRoundServiceQuantityToProduce(string deliveryRoundName, string serviceName)
        {
            ShowExtendedMenu();
            _foldUnfoldButtonLabel = _webDriver.FindElement(By.XPath(FOLD_UNFOLD_BUTTON_LABEL));
            if (_foldUnfoldButtonLabel.Text != "Fold All")
            {
                FoldUnfoldAll();
                WaitForLoad();
            }

            List<string> deliveryRoundServiceQuantityList = new List<string>();
            var deliveryRoundServiceQuantityElements = _webDriver.FindElements(By.XPath(String.Format(DELIVERYROUND_SERVICE_QUANTITY_TO_PRODUCE, deliveryRoundName, serviceName)));
            if(deliveryRoundServiceQuantityElements.Count != 0) 
            {
                foreach (var element in deliveryRoundServiceQuantityElements)
            {
                string quantity = element.Text.Replace("\r\n", " ");
                deliveryRoundServiceQuantityList.Add(quantity);
            }
            }
            

            return deliveryRoundServiceQuantityList;
        }

        // INPUT : Delivery round name & service name - OUTPUT : weight list
        public List<string> GetDeliveryRoundServiceWeights(string deliveryRoundName, string serviceName)
        {
            ShowExtendedMenu();
            _foldUnfoldButtonLabel = _webDriver.FindElement(By.XPath(FOLD_UNFOLD_BUTTON_LABEL));
            if (_foldUnfoldButtonLabel.Text != "Fold All")
            {
                FoldUnfoldAll();
                WaitForLoad();
            }

            List<string> deliveryRoundServiceWeightList = new List<string>();
            var deliveryRoundServiceWeightElements = _webDriver.FindElements(By.XPath(String.Format(DELIVERYROUND_SERVICE_WEIGHT_PER_PACK, deliveryRoundName, serviceName)));
            if(deliveryRoundServiceWeightElements.Count != 0) 
            {
                foreach (var element in deliveryRoundServiceWeightElements)
                {
                    string weight = element.Text.Replace("\r\n", " ");
                    deliveryRoundServiceWeightList.Add(weight);
                }
            }
            

            return deliveryRoundServiceWeightList;
        }

        // INPUT : Delivery round name & service name - OUTPUT : total list
        public List<string> GetDeliveryRoundServiceTotals(string deliveryRoundName, string serviceName)
        {
            ShowExtendedMenu();
            _foldUnfoldButtonLabel = _webDriver.FindElement(By.XPath(FOLD_UNFOLD_BUTTON_LABEL));
            if (_foldUnfoldButtonLabel.Text != "Fold All")
            {
                FoldUnfoldAll();
                WaitForLoad();
            }

            List<string> deliveryRoundServiceTotalList = new List<string>();
            var deliveryRoundServiceTotalElements = _webDriver.FindElements(By.XPath(String.Format(DELIVERYROUND_SERVICE_TOTAL_PER_PACK, deliveryRoundName, serviceName)));
            if (deliveryRoundServiceTotalElements.Count != 0)
            {
                foreach (var element in deliveryRoundServiceTotalElements)
                {
                    string weight = element.Text.Replace("\r\n", " ");
                    deliveryRoundServiceTotalList.Add(weight);
                }
            }

            return deliveryRoundServiceTotalList;
        }
        public string GetRecipeNameRow()
        {
            var nameRow = WaitForElementIsVisible(By.XPath("/html/body/div[2]/div/div[2]/div[2]/div/div/div[2]/div[2]/div[2]/div/div/table/tbody/tr/td[2]"));
            return nameRow.Text.Trim();
        }
    }
}
