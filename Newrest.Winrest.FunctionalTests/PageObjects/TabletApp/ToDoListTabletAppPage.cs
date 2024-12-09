using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Newrest.Winrest.FunctionalTests.PageObjects.TabletApp
{
    public class ToDoListTabletAppPage : PageBase
    { 
        //---------------------------------------CONSTANTES---------------------------------------------------------------
        public const string FIRST_HOUR = "/html/body/app-root/mat-sidenav-container/mat-sidenav-content/main/lounge/div/div[2]/div/div/table/tbody/tr[1]/td[1]";
        public const string SECOND_HOUR = "/html/body/app-root/mat-sidenav-container/mat-sidenav-content/main/lounge/div/div[2]/div/div/table/tbody/tr[2]/td[1]";
        public const string FINAL_HOUR = "/html/body/app-root/mat-sidenav-container/mat-sidenav-content/main/lounge/div/div[2]/div/div/table/tbody/tr[3]/td[1]";

        public ToDoListTabletAppPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
  
          
        }

        public string GetFirstHour()
        {
              var FirstHour = WaitForElementIsVisible(By.XPath(FIRST_HOUR));
              return FirstHour.Text;
            

        }
        public string GetSecondHour()
        {
            var secondHour = WaitForElementIsVisible(By.XPath(SECOND_HOUR));
            return secondHour.Text;


        }
        public string GetFinalHour()
        {
            var finalHour = WaitForElementIsVisible(By.XPath(FINAL_HOUR));
            return finalHour.Text;


        }

        public List<string> GetAllStartTimeByTaskName(string taskName)
        {
            WaitLoading();
            var startTime = _webDriver.FindElements(By.XPath(string.Format("/html/body/app-root/mat-sidenav-container/mat-sidenav-content/main/lounge/div/div[2]/div/div/table/tbody/tr[td[2][contains(text(), '{0}')]]/td[1]",taskName)));
            WaitPageLoading();
            return startTime.Select(e => e.Text).ToList();
        }

        public bool VerifyAllHoursExist(string taskName)
        {
            var FirstHour = "06:00";
            var SecondHour = "12:00";
            var FinalHour = "18:00";
            WaitLoading();
            var allHours = GetAllStartTimeByTaskName(taskName);
            WaitPageLoading();
            if (!allHours.Contains(FirstHour) ||
            !allHours.Contains(SecondHour) ||
            !allHours.Contains(FinalHour))
            {
                return false;
            }
            return true;
        }

    }
}
