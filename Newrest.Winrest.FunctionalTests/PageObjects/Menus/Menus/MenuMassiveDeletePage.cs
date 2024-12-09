using DocumentFormat.OpenXml.Bibliography;
using Limilabs.Client.IMAP;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Menus.Recipes;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Menus.Menus
{
    public class MenuMassiveDeletePage : PageBase
    {
        public MenuMassiveDeletePage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {

        }

        private const string SEARCH = "SearchMenusBtn";
        private const string FROM = "//*[@id=\"tableMenus\"]/thead/tr/th[5]/span/a";
        private const string LIST_FROM = "//*[@id=\"tableMenus\"]/tbody/tr[*]/td[5]";
        private const string TO = "//*[@id=\"tableMenus\"]/thead/tr/th[6]/span/a";
        private const string LIST_TO = "/html/body/div[3]/div/div/div[2]/div/form/div/div[8]/div/div/table/tbody/tr[*]/td[6]";
        private const string FIRST_LINE = "//*[@id=\"tableMenus\"]/tbody/tr[1]";
        private const string DELETE = "//*[@id=\"deleteMenusBtn\"]";
        private const string FIRST_LINE_NAME = "//*[@id=\"tableMenus\"]/tbody/tr[1]/td[3]";
        private const string CONFIRME_DELETE = "//*[@id=\"dataConfirmOK\"]";
        private const string CONFIRME_AFTER_DELETE = "//*[@id=\"modal-1\"]/div[3]/button";
        private const string ITEM_NAME = "//*[@id=\"tableMenus\"]/tbody/tr[*]/td[3]";
        private const string MENU_NAME_BTN = "//*[@id=\"tableMenus\"]/thead/tr/th[3]/span/a";
        private const string LIST_VARIANT = "/html/body/div[3]/div/div/div[2]/div/form/div/div[8]/div/div/table/tbody/tr[*]/td[4]";
        private const string MAIN_VARIANT_BTN = "/html/body/div[3]/div/div/div[2]/div/form/div/div[8]/div/div/table/thead/tr/th[4]/span/a";
        private const string MASSIVE_DELETE = "//*[@id=\"tabContentItemContainer\"]/div/div/div[1]/div/a[1]";
        private const string SEARCH_MENUS_BTN = "SearchMenusBtn";
        private const string STATUS = "SelectedStatus_ms";
        private const string SEARCH_FILTER = "//*[@id=\"modal-1\"]//input[@id=\"SearchPattern\"]";
        private const string SELECT_ALL = "//*[@id=\"selectAll\"]";
        private const string CONFIRM_DELETE_ALL = "//*[@id=\"modal-1\"]/div[3]/button";
        private const string COUNT_AFTER_DELETE = "//*[@id=\"tableMenus\"]/tbody/tr[*]";
        private const string COUNT_SELECTED_MENUS = "menuCount";
        private const string UNSELECT_ALL = "//*[@id=\"unselectAll\"]";
        private const string PAGINATION2 = "//*[@id=\"list-menus-deletion\"]/nav/ul/li[4]/a";
        private const string PAGINATION3 = "//*[@id=\"list-menus-deletion\"]/nav/ul/li[5]/a";
        private const string PAGINATION4 = "//*[@id=\"list-menus-deletion\"]/nav/ul/li[6]/a";
        private const string PAGINATION5 = "//*[@id=\"list-menus-deletion\"]/nav/ul/li[7]/a";
        private const string LIGNE_PAGINATION = "//*[@id=\"modal-1\"]//select[@id=\"page-size-selector\"]";
        private const string PAGINATION_NUMBER = "//*[@id=\"modal-1\"]//select[@id=\"page-size-selector\"]/option[2]";
        private const string PAGE_SIZE = "/html/body/div[3]/div/div/div[2]/div/form/div/div[8]/div/div/nav/select";
        private const string SITE_SORT = "//*[@id=\"tableMenus\"]/thead/tr/th[2]/span/a";
        private const string SITE_COLUMN = "//*[@id=\"tableMenus\"]/tbody/tr[*]/td[2]";
        private const string SITE_FILTER_SHOW_INACTIVE_ID = "ShowInactiveSites";
        private const string SITE_FILTER_ID = "siteFilter";

        private const string TABLE_INACTIVE = "//*[@id=\"tableMenus\"]/tbody/tr[*]";
        private const string COMBOBOX_INACTIVE = "/html/body/div[17]/ul/li[*]/label/span[contains(text() ,'Inactive -')]";
        [FindsBy(How = How.XPath, Using = ITEM_NAME)]
        private IWebElement _itemName;

        [FindsBy(How = How.XPath, Using = MENU_NAME_BTN)]
        private IWebElement _menuNameBtn;

        [FindsBy(How = How.XPath, Using = MASSIVE_DELETE)]
        private IWebElement _massiveDelete;
        [FindsBy(How = How.Id, Using = STATUS)]
        private IWebElement _status;
        [FindsBy(How = How.Id, Using = COUNT_SELECTED_MENUS)]
        private IWebElement _count_selectedmenus;
        [FindsBy(How = How.Id, Using = UNSELECT_ALL)]
        private IWebElement _unselectall;
        [FindsBy(How = How.XPath, Using = SEARCH_FILTER)]
        private IWebElement _searchFilter;
        [FindsBy(How = How.XPath, Using = PAGE_SIZE)]
        private IWebElement _pageSize;
        [FindsBy(How = How.XPath, Using = PAGINATION2)]
        private IWebElement _pagination2;
        [FindsBy(How = How.XPath, Using = PAGINATION3)]
        private IWebElement _pagination3;
        [FindsBy(How = How.XPath, Using = PAGINATION4)]
        private IWebElement _pagination4;
        [FindsBy(How = How.XPath, Using = PAGINATION5)]
        private IWebElement _pagination5;
        [FindsBy(How = How.XPath, Using = LIGNE_PAGINATION)]
        private IWebElement _paginationline;
        [FindsBy(How = How.XPath, Using = PAGINATION_NUMBER)]
        private IWebElement _paginationnumber;

        [FindsBy(How = How.XPath, Using = SITE_SORT)]
        private IWebElement _siteSortButton;


        [FindsBy(How = How.XPath, Using = TABLE_INACTIVE)]
        private IWebElement _tableInactive;

        public void ClickOnSearch()
        {
            var _searchclick = WaitForElementExists(By.Id(SEARCH));
            _searchclick.Click();
            WaitPageLoading();
            WaitForLoad();
        }
        public void ClickOnFrom()
        {
            var _fromclick = WaitForElementExists(By.XPath(FROM));
            _fromclick.Click();
            WaitPageLoading();
            WaitForLoad();
        }
        public void ClickOnTo()
        {
            var _fromclick = WaitForElementExists(By.XPath(TO));
            _fromclick.Click();
            WaitPageLoading();
            WaitForLoad();
        }
        public bool CheckOrderFrom()
        {
            var list_from = _webDriver.FindElements(By.XPath(LIST_FROM));
            for (int i = 0; i < list_from.Count - 1; i++)
            {
                DateTime currentDate = DateTime.ParseExact(list_from[i].Text, "dd/MM/yyyy", null);
                DateTime nextDate = DateTime.ParseExact(list_from[i + 1].Text ,"dd/MM/yyyy", null);

                if (currentDate > nextDate)
                {
                    return false;
                }
            }
            return true;
        }
        public bool CheckOrderToDsc()
        {
            var list_To = _webDriver.FindElements(By.XPath(LIST_TO));
            for (int i = 0; i < list_To.Count - 1; i++)
            {
                DateTime currentDate = DateTime.ParseExact(list_To[i].Text, "dd/MM/yyyy", CultureInfo.InvariantCulture);
                DateTime nextDate = DateTime.ParseExact(list_To[i + 1].Text, "dd/MM/yyyy", CultureInfo.InvariantCulture);

                if (currentDate < nextDate) 
                {
                    return false;
                }
            }
            return true;
        }

        public bool CheckOrderToASC()
        {
            var list_To = _webDriver.FindElements(By.XPath(LIST_TO));
            for (int i = 0; i < list_To.Count - 1; i++)
            {
                DateTime currentDate = DateTime.ParseExact(list_To[i].Text, "dd/MM/yyyy", CultureInfo.InvariantCulture);
                DateTime nextDate = DateTime.ParseExact(list_To[i + 1].Text, "dd/MM/yyyy", CultureInfo.InvariantCulture);

                if (currentDate > nextDate) 
                {
                    return false;
                }
            }
            return true;
        }

        public bool IsSortedByName()
        {
            var ancientName = "";
            int compteur = 1;

            var elements = _webDriver.FindElements(By.XPath(ITEM_NAME));

            if (elements.Count == 0)
                return false;

            foreach (var elm in elements)
            {
                if (compteur == 1)
                    ancientName = elm.GetAttribute("InnerText");

                if (string.Compare(ancientName, elm.GetAttribute("InnerText")) > 0)
                    return false;

                ancientName = elm.GetAttribute("InnerText");
                compteur++;
            }

            return true;
        }

        public bool IsSortedByVariant()
        {
            var ancientName = "";
            int compteur = 1;

            var elements = _webDriver.FindElements(By.XPath(LIST_VARIANT));

            if (elements.Count == 0)
                return false;

            foreach (var elm in elements)
            {
                if (compteur == 1)
                    ancientName = elm.GetAttribute("InnerText");

                if (string.Compare(ancientName, elm.GetAttribute("InnerText")) > 0)
                    return false;

                ancientName = elm.GetAttribute("InnerText");
                compteur++;
            }

            return true;
        }
        public void ClickOnFirstLine()
        {
            var _firstline = WaitForElementExists(By.XPath(FIRST_LINE));
            _firstline.Click();
        }
        public void ClickonSelectAll()
        {
            var _selectall = WaitForElementExists(By.XPath(SELECT_ALL));
            _selectall.Click();
        }

        public void ClickOnDelete()
        {
            var _delete = WaitForElementExists(By.XPath(DELETE));
            WaitPageLoading();
            _delete.Click();
            WaitPageLoading();
        }
        public void ConfirmDeleteAll()
        {
            var _deleteall = WaitForElementExists(By.XPath(CONFIRM_DELETE_ALL));
            _deleteall.Click();
        }

        public void ClickOnConfirmDelete()
        {
            var _Confirmdelete = WaitForElementExists(By.XPath(CONFIRME_DELETE));
            _Confirmdelete.Click();
            WaitPageLoading();



        }
        public void ClickOnConfirmAfterDelete()
        {
            var _Confirmdelete = WaitForElementExists(By.XPath(CONFIRME_AFTER_DELETE));
            _Confirmdelete.Click();
            WaitPageLoading();
        }
        public string FirstLineName()
        {
            var _name = WaitForElementExists(By.XPath(FIRST_LINE_NAME)).Text;
            return _name;
        }

        public void MenuNameTriParClick()
        {
            _menuNameBtn = WaitForElementIsVisible(By.XPath(MENU_NAME_BTN));
            _menuNameBtn.Click();
            WaitForLoad();
        }  public void MenuVariantTriParClick()
        {
            var  menuVariantBtn = WaitForElementIsVisible(By.XPath(MAIN_VARIANT_BTN));
                 menuVariantBtn.Click();
            WaitForLoad();
        }

        public void MassiveDeleteSiteSearch(string site)
        {
            ComboBoxSelectById(new ComboBoxOptions("SelectedSiteIds_ms", site, false));

            var searchMenusbtn = WaitForElementIsVisible(By.Id(SEARCH_MENUS_BTN));
            searchMenusbtn.Click();
            WaitForLoad();
        }

        public bool SiteComboBoxChecker(string site)
        {
            var searchSite = WaitForElementIsVisible(By.Id("SelectedSiteIds_ms"));
            searchSite.Click();
            searchSite.SetValue(ControlType.TextBox, site);
            WaitPageLoading();
            searchSite.Click();

            if (isElementVisible(By.XPath(COMBOBOX_INACTIVE)))
            {
                var listInactive = _webDriver.FindElements(By.XPath(COMBOBOX_INACTIVE));
                foreach (var item in listInactive)
                {
                    if (!item.Text.Contains(site))
                    {
                        return false;
                    }
                }
                return true;
            }
            else
            {
                return false;
            }
        }

        public enum FilterType
        {
            SearchMenu,
            Status

        }

        public void Filter(FilterType filterType, object value)
        {
            Actions action = new Actions(_webDriver);

            switch (filterType)
            {
                case FilterType.SearchMenu:
                    _searchFilter = WaitForElementIsVisible(By.XPath(SEARCH_FILTER));
                    _searchFilter.SetValue(ControlType.TextBox, value);
                    break;
                case FilterType.Status:
                    ComboBoxSelectById(new ComboBoxOptions(STATUS, (string)value,false));
                    _status.Click();
                    break;
                default:
                    throw new ArgumentOutOfRangeException(nameof(FilterType), filterType, null);
            }

            WaitPageLoading();
            WaitForLoad();
        }
        public int CountRows()
        {
            var rows = _webDriver.FindElements(By.XPath(COUNT_AFTER_DELETE));
            return rows.Count;
        }
        public void PageSizeMenuMassiveDelete(string size)
        {
            if (size == "1000")
            {   // Test
                IJavaScriptExecutor js = (IJavaScriptExecutor)_webDriver;
                js.ExecuteScript("$('#" + PAGE_SIZE + "').append($('<option>', {value: 1000,text: '1000'}),'');");
            }

            var wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(30));
            try
            {
                wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath(PAGE_SIZE)));
            }
            catch
            {
                // tableau vide : pas de PageSize
                return;
            }
            Actions action = new Actions(_webDriver);
            action.MoveToElement(_pageSize).Perform();
            _pageSize.SetValue(ControlType.DropDownList, size);

            WaitPageLoading();
            WaitForLoad();
        }
        public bool Pagination()
        {

             var _paginate2 = WaitForElementExists(By.XPath(PAGINATION2));
            if (_paginate2 == null) return false;   
            _paginate2.Click();
            WaitForLoad();

            var _paginate3 = WaitForElementExists(By.XPath(PAGINATION3));
            if (_paginate3 == null) return false;  
            _paginate3.Click();
            WaitForLoad();

            var _paginate4 = WaitForElementExists(By.XPath(PAGINATION4));
            if (_paginate4 == null) return false;   
            _paginate4.Click();
            WaitForLoad();

        
            var _paginate5 = WaitForElementExists(By.XPath(PAGINATION5));
            if (_paginate5 == null) return false;   
            _paginate5.Click();
            WaitForLoad();

             return true;
        }

        public int CheckSelectedMenus()
        {
            WaitForLoad();
            _count_selectedmenus = WaitForElementExists(By.Id(COUNT_SELECTED_MENUS));
            int nombre = Int32.Parse(_count_selectedmenus.Text);
            return nombre;
        }
        public void ClickonUnSelectAll()
        {
             _unselectall = WaitForElementExists(By.XPath(UNSELECT_ALL));
            _unselectall.Click();
        }
        public void ChangeNumberPagination()
        {
            var _firstline = WaitForElementExists(By.XPath(LIGNE_PAGINATION));
            _firstline.Click();
            var _firstlines = WaitForElementExists(By.XPath(PAGINATION_NUMBER));
            _firstlines.Click();

        }
        public void SortBySite()
        {
            _siteSortButton = WaitForElementIsVisible(By.XPath(SITE_SORT));
            _siteSortButton.Click();
            WaitForLoad();
        }
        public bool IsSorted()
        {
            var elements = _webDriver.FindElements(By.XPath(SITE_COLUMN));
            for (int i = 0; i < elements.Count - 1; i++) 
                if (string.Compare(elements[i].Text, elements[i + 1].Text) > 0)
                    return false;
            return true;
        }
        public bool CheckAllRowsAreInactive()
        {
            var tableInactive = _webDriver.FindElements(By.XPath(TABLE_INACTIVE));

            foreach (var row in tableInactive)
            {
                if (!row.GetAttribute("title").Contains("Inactive"))
                {
                    return false;
                }
            }
            return true;
        }
        public void ClickOnInactiveSiteCheck()
        {
            IWebElement checkBoxInactiveSite = WaitForElementExists(By.Id(SITE_FILTER_SHOW_INACTIVE_ID));
            checkBoxInactiveSite.Click();
            WaitForLoad();
        }
        public void SelectAllInactiveSites()
        {
            ComboBoxOptions cbOpt = new ComboBoxOptions(SITE_FILTER_ID, "Inactive", false)
            { ClickCheckAllAtStart = false, ClickCheckAllAfterSelection = true };
            ComboBoxSelectById(cbOpt);
        }
        public void CheckAllSites()
        {
            IWebElement siteDropdown = WaitForElementExists(By.Id(SITE_FILTER_ID)); 
            siteDropdown.Click();

            IWebElement checkAllOption = WaitForElementExists(By.XPath("/html/body/div[17]/div/ul/li[1]/a"));
            checkAllOption.Click();
        }
       

    }
}

