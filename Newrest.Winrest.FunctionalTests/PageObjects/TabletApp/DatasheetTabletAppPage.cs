using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Newrest.Winrest.FunctionalTests.PageObjects.TabletApp
{
    public class DatasheetTabletAppPage :PageBase
    {
        public DatasheetTabletAppPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {


        }

        //--------------------------------------------------CONSTANTES---------------------------------------------------------------

        // Général
        private const string GUEST_TYPE_SEARCH = "/html/body/app-root/mat-sidenav-container/mat-sidenav-content/main/search/div/div[2]/div/div[2]/selectable-list/div/div[2]/input";
        private const string GUEST_TYPE_LIST = " /html/body/app-root/mat-sidenav-container/mat-sidenav-content/main/search/div/div[2]/div/div[2]/selectable-list/div/div[4]/div[*]/span";



        //____________________________________________________Variables______________________________________________________________

        // Général
        [FindsBy(How = How.XPath, Using = GUEST_TYPE_SEARCH)]
        private IWebElement _guestTypeSearch;

        [FindsBy(How = How.XPath, Using = GUEST_TYPE_LIST)]
        private IWebElement _guestTypeList;


        public bool CheckGuestTypeExist(string guestType)
        {
            _guestTypeSearch = WaitForElementIsVisible(By.XPath(GUEST_TYPE_SEARCH));
            _guestTypeSearch.SetValue(ControlType.TextBox,guestType);
            WaitLoading();
            var _listGuestType = _webDriver.FindElements(By.XPath(GUEST_TYPE_LIST));

            if (_listGuestType.Count() == 0)
                return false;

            foreach (var elm in _listGuestType)
            {
                if (!elm.Text.Contains(guestType))
                {
                    return false;
                }
            }
            return true;
        }

    }
}
