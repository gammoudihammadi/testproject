using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.PageObjects;
using System;
using System.Collections.Generic;
using System.Threading;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Purchasing.Item
{
    public class ItemKeywordPage : PageBase
    {
        public ItemKeywordPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        private const string KEYWORD = "//*[@id=\"div-keywords\"]/div[1]/div/div/input";
        private const string ADD_KEYWORD = "//*[@id=\"div-keywords\"]/div[1]/div/div/a";
        private const string CANCEL_ADD_KEYWORD = "dataAlertCancel";
        private const string GENERAL_INFO_TAB = "hrefTabContentItem";
        private const string KEYWORDS = "KeywordValuesIds_ms";
        private const string SELECT_KEYWORDS = "KeywordValuesIdsTrueValue";
        private const string ENTER_KEYWORDS = "/html/body/div[11]/div/div/label/input";
        private const string SELECT_FIRST_KEYWORD = "ui-multiselect-0-KeywordValuesIds-option-0";
        private const string SEARCH_KEYWORD = "/html/body/div[11]/div/div/label/input";
        private const string SELECT_ALL_KEYWORD = "/html/body/div[11]/div/ul/li[1]/a/span[2]";
        private const string KEYWORD_UNSELECT_ALL = "/html/body/div[11]/div/ul/li[2]/a/span[2]";
        private const string UNCHECK = "//*[@id=\"keyword-list-content\"]/div/a";
        private const string FIRST_SITE_NAME = "//*[@id=\"div-keywords-details\"]/table/tbody/tr[1]/td[1]/span";
        private const string SELECT_ADD = "/html/body/div[3]/div/div[2]/div[2]/div/div/div/form/div[4]/table/tbody/tr/td[5]/div[1]/div/div/div/div/div/div[1]/button";
        private const string UNCHECK_ALL = "/html/body/div[12]/div/ul/li[2]";
        private const string SEARCH_KEYWORD_TEST = "/html/body/div[12]/ul/li[298]/label/input";
        private const string SEARCH_KEYWORDP = "/html/body/div[12]/div/div/label/input";
        private const string SEARCH = "SearchPattern";
        private const string PACKAGING_FIRST_ROW = "detail_ItemDetailId";
        private const string BACK_TO_LIST = "/html/body/div[2]/a/span[2]";


        [FindsBy(How = How.XPath, Using = KEYWORD)]
        private IWebElement _keyword;
        [FindsBy(How = How.Id, Using = SEARCH)]
        private IWebElement _search;

        [FindsBy(How = How.Id, Using = CANCEL_ADD_KEYWORD)]
        private IWebElement _cancelAddKeyword;

        [FindsBy(How = How.Id, Using = KEYWORDS)]
        private IWebElement _keywords;

        [FindsBy(How = How.Id, Using = SELECT_FIRST_KEYWORD)]
        private IWebElement _selectFirstKeywords;

        [FindsBy(How = How.XPath, Using = ENTER_KEYWORDS)]
        private IWebElement _enterKeywords;
        [FindsBy(How = How.XPath, Using = SEARCH_KEYWORD)]
        private IWebElement _searchKeyword;
        [FindsBy(How = How.XPath, Using = SELECT_ALL_KEYWORD)]
        private IWebElement _selectAllKeyword;
        [FindsBy(How = How.XPath, Using = KEYWORD_UNSELECT_ALL)]
        private IWebElement _KeywordUnselectAll;
        [FindsBy(How = How.XPath, Using = UNCHECK)]
        private IWebElement _uncheck;
        [FindsBy(How = How.XPath, Using = FIRST_SITE_NAME)]
        private IWebElement _firstSiteName;
        [FindsBy(How = How.XPath, Using = BACK_TO_LIST)]
        private IWebElement _backToList;


        // Onglets
        [FindsBy(How = How.Id, Using = GENERAL_INFO_TAB)]
        private IWebElement _generalInfoTab;

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

        public void CloseWindow(string originalWindow)
        {

            IList<string> allWindows = _webDriver.WindowHandles;

            // Switch to the new window
            foreach (string window in allWindows)
            {
                if (window != originalWindow)
                {
                    _webDriver.SwitchTo().Window(window);
                    break;
                }
            }
            _webDriver.Close();

            // Switch back to the original window
            _webDriver.SwitchTo().Window(originalWindow);
        }
        public void AddKeywords(List<string> keywords)
        {
            var input = WaitForElementIsVisible(By.Id("KeywordValuesIds_ms"));
            input.Click();
            WaitForLoad();

            foreach (string keyword in keywords)
            {
                if (keyword != null)
                {
                    var searchVisible = SolveVisible("//*/input[@type='search']");
                    Assert.IsNotNull(searchVisible);
                    searchVisible.SetValue(ControlType.TextBox, keyword);
                    // on ne clique pas de checkbox donc pas de rechargement de page ici
                    WaitPageLoading();

                    var select = SolveVisible("//*/label[contains(@for, 'ui-multiselect')]/span[contains(text(),'" + keyword + "')]");
                    Assert.IsNotNull(select, "Pas de sélection de " + keyword);
                    select.Click();
                    WaitForLoad();
                }
            }
        }

        public void RemoveKeyword(string keyword)
        {
            WaitForLoad();
            var remove = WaitForElementIsVisible(By.XPath(String.Format("//*/span[text()='{0}']/../a", keyword)));
            remove.Click();
            WaitForLoad();
        }

        public bool IsKeywordPresent(string keyword)
        {
            return  _webDriver.FindElements(By.XPath($"//div[@id='keyword-list-content']//span[text()='{keyword}']")).Count>0;
        }

        public ItemGeneralInformationPage ClickOnGeneralInfo()
        {
            _generalInfoTab = WaitForElementIsVisible(By.Id(GENERAL_INFO_TAB));
            _generalInfoTab.Click();
            WaitPageLoading();

            return new ItemGeneralInformationPage(_webDriver, _testContext);
        }

        public bool IsKeywordAdded(string keyword)
        {
            int KeywordAdded = _webDriver.FindElements(By.XPath("//*[@id=\"keyword-list-content\"]/div/span[text()='" + keyword + "']")).Count;

            if (KeywordAdded == 0)
                return false;

            return true;
        }
        public enum FilterType
        {

            Keywords,
            Search


        }
        public void Filter(FilterType filterType, object value)
        {
            Actions action = new Actions(_webDriver);

            switch (filterType)
            {
                case FilterType.Keywords:
                    _keyword = WaitForElementIsVisible(By.Id(KEYWORDS));
                    _keyword.Click();


                    _searchKeyword = WaitForElementIsVisible(By.XPath(SEARCH_KEYWORD));
                    _searchKeyword.SetValue(ControlType.TextBox, value);

                    var keywordToCheck = WaitForElementIsVisible(By.XPath("//span[text()='" + value + "']"));
                    keywordToCheck.SetValue(ControlType.CheckBox, true);

                    break;
                case FilterType.Search:
                    _search = WaitForElementIsVisible(By.Id(SEARCH));
                    _search.SetValue(ControlType.TextBox, value);
                    WaitForLoad();
                    break;
                default:
                    throw new ArgumentOutOfRangeException(nameof(filterType), filterType, "Invalid filter type provided.");
            }
        }
        public void ClickOnuncheckKeyword()
        {
            _uncheck = WaitForElementIsVisible(By.XPath(UNCHECK));


            _uncheck.Click();
            WaitForLoad();
        }

        public string GetFirstSiteName()
        {
            _firstSiteName = WaitForElementIsVisible(By.XPath(FIRST_SITE_NAME));
            string site = _firstSiteName.Text;
            return site;




        }

        public void AddKeywordDetails(string keyword)
        {
            var selectBtn = WaitForElementIsVisible(By.XPath(SELECT_ADD));
            selectBtn.Click();
            WaitForLoad();

            //var uncheckAll = WaitForElementIsVisible(By.XPath(UNCHECK_ALL));
            //uncheckAll.Click();
            //WaitForLoad();

            if (IsDev())
            {
                if (!isElementVisible(By.XPath(SEARCH_KEYWORD_TEST)))
                {
                    selectBtn = WaitForElementIsVisible(By.XPath(SELECT_ADD));
                    selectBtn.Click();
                    WaitForLoad();
                }
                var searchKeyword = WaitForElementIsVisible(By.XPath(SEARCH_KEYWORD_TEST));
                searchKeyword.SetValue(ControlType.TextBox, keyword);
                WaitForLoad();
            }
            else
            {

                if (!isElementVisible(By.XPath(SEARCH_KEYWORDP)))
                {
                    selectBtn = WaitForElementIsVisible(By.XPath(SELECT_ADD));
                    selectBtn.Click();
                    WaitForLoad();
                }
                var searchKeyword = WaitForElementIsVisible(By.XPath(SEARCH_KEYWORDP));
                searchKeyword.SetValue(ControlType.TextBox, keyword);
                WaitForLoad();
                var select = SolveVisible("//*/label[contains(@for, 'ui-multiselect')]/span[contains(text(),'" + keyword + "')]");
                Assert.IsNotNull(select, "Pas de sélection de " + keyword);
                select.Click();

            }
            WaitForLoad();
        }
        public bool IsPackagingDuplicated()
        {
            return isElementExists(By.Id(PACKAGING_FIRST_ROW));
        }
        public ItemPage BackToList()
        {
            _backToList = WaitForElementIsVisible(By.XPath(BACK_TO_LIST));
            _backToList.Click();
            WaitForLoad();

            return new ItemPage(_webDriver, _testContext);
        }

    }
}
