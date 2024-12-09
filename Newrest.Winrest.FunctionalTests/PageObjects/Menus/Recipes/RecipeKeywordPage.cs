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

namespace Newrest.Winrest.FunctionalTests.PageObjects.Menus.Recipes
{
    public class RecipeKeywordPage : PageBase
    {
        // ______________________________________ Constantes _____________________________________________

        private const string FIRST_KEYWORD = "//*[@id=\"tabContentDetails\"]/div/table/tbody/tr[2]/td";
        private const string LIST_KEYWORD = "//*[@id=\"tabContentDetails\"]/div/table/tbody/tr[*]/td";
        private const string LIST_KEYWORD_HEAD = "//*[@id=\"tabContentDetails\"]/div/table/tbody/tr[1]/th";


        // ______________________________________ Variables _____________________________________________

        // Général

        [FindsBy(How = How.XPath, Using = FIRST_KEYWORD)]
        private IWebElement _firstKeyword;

        public RecipeKeywordPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        public string GetFirstKeyword()
        {
            WaitForLoad();
            _firstKeyword = WaitForElementExists(By.XPath(FIRST_KEYWORD));
            return _firstKeyword.Text;
        }

        public bool IsKeyWordsAlignedInTheTable()
        {
            var headerCell = _webDriver.FindElement(By.XPath(LIST_KEYWORD_HEAD));
            int headerX = headerCell.Location.X;
            var dataCells = _webDriver.FindElements(By.XPath(LIST_KEYWORD));
            foreach (var cell in dataCells)
            {
                if (cell.Location.X != headerX)
                {
                    return false; 
                }
            }

            return true; 
        }
        public bool CheckKeyword(string keyword)
        {
            var listKeyword = _webDriver.FindElements(By.XPath(LIST_KEYWORD));

            WaitPageLoading();
            WaitPageLoading();

            foreach (var element in listKeyword)
            {
                if (element.Text.Equals(keyword, StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }

            return false;
        }

        public void AddKeyword(string keyword)
        {
            var selectBtn = WaitForElementIsVisible(By.Id("KeywordValuesIds_ms"));
            selectBtn.Click();

            var uncheckAll = WaitForElementIsVisible(By.XPath("/html/body/div[11]/div/ul/li[2]/a[@title= 'Uncheck all']"));
            uncheckAll.Click();
            if (!isElementVisible(By.XPath("/html/body/div[11]/div/div/label/input")))
            {
                selectBtn = WaitForElementIsVisible(By.Id("KeywordValuesIds_ms"));
                selectBtn.Click();
            }
            var searchKeyword = WaitForElementIsVisible(By.XPath("/html/body/div[11]/div/div/label/input"));
            searchKeyword.SetValue(ControlType.TextBox, keyword);


            Thread.Sleep(5000);
            if (!isElementVisible(By.XPath("/html/body/div[11]/div/div/label/input")))
            {
                selectBtn = WaitForElementIsVisible(By.Id("KeywordValuesIds_ms"));
                selectBtn.Click();
            }
            var keywordToCheck = WaitForElementExists(By.XPath("//span[text()='" + keyword + "']"));
            keywordToCheck.SetValue(ControlType.CheckBox, true);

            //ComboBoxSelectById("KeywordValuesIds_ms", keyword);
            WaitForLoad();
        }

    }
}
