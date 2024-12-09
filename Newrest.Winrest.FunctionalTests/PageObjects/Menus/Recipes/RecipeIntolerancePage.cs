using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.PageObjects;
using System;
using System.Threading;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Menus.Recipes
{
    public class RecipeIntolerancePage : PageBase
    {

        public RecipeIntolerancePage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        // _____________________________________ Constantes _________________________________________________

        private const string RECIPE_INTOLERANCE_PAGE_ALLERGEN_LINE = "//*[@id=\"tabContentDetails\"]/div[2]/table/tbody/tr[*]/td[2]/label[text()='{0}']";
        private const string LIST_ALLERGENS_SELECTED = "//*[@id=\"allergenCheckBox\" and @checked]";

        // _____________________________________ Variables __________________________________________________

        [FindsBy(How = How.XPath, Using = RECIPE_INTOLERANCE_PAGE_ALLERGEN_LINE)]
        private IWebElement _allergenLine;

        // _____________________________________ Méthodes ___________________________________________________

        public bool IsRecipeAllergenChecked(string allergen)
        {

            bool result = false;

                var allergens = _webDriver.FindElements(By.XPath("//*[@id='tabContentDetails']/div[2]/table/tbody/tr[*]"));

                foreach (var element in allergens)
                {
                try
                {
                    var labelElement = element.FindElement(By.XPath(".//td[2]/label"));


                    if (labelElement != null && labelElement.Text == allergen)
                    {

                        var checkbox = element.FindElements(By.XPath(".//td[1]/div/input"));


                        if (checkbox.First().Selected)
                        {
                            return true;
                        }
                    }

                }
                catch
                {
                    result = false;
                   
                }
                  
                }

                return false; 
            }
        public int CalculateAllergens()
        {
            WaitForLoad();
            var allergensSelected = _webDriver.FindElements(By.XPath(LIST_ALLERGENS_SELECTED));
            return allergensSelected.Count;
        }

        public List<string> GetListAllergensSelected()
        {
            var allergensSelected = _webDriver.FindElements(By.XPath(LIST_ALLERGENS_SELECTED));
            var list = new List<string>();

            foreach (var element in allergensSelected)
            {
                var nameAttribute = element.GetAttribute("name");
                var match = System.Text.RegularExpressions.Regex.Match(nameAttribute, @"^([^\[]+)\[(\d+)\]");
                if (!match.Success) continue;
                string aWord = match.Groups[1].Value;  
                string number = match.Groups[2].Value;
                string xpath = $"//*[@id='AllergenName'][contains(@for, '{aWord}') and contains(@for, '{number}')]";
                var allergenElement = _webDriver.FindElement(By.XPath(xpath));
                if (allergenElement != null)
                {
                    list.Add(allergenElement.Text);
                }
            }

            return list;
        }

    }
}
