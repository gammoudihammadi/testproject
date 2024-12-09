using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Keys = OpenQA.Selenium.Keys;

namespace Newrest.Winrest.FunctionalTests.PageObjects.ToDoList.Scheduler
{
    public class ScheduleDetailsPage : PageBase
    {
        public ScheduleDetailsPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        //__________________________________Constantes_______________________________________

        // General

        private const string NAME = "Name";
        public const string START_TIME = "TimeFrom";
        public const string END_TIME = "TimeTo";
        //public const string END_TIME = "/html/body/div[3]/div/div/div/div/form/div/div/div[1]/div/div[10]/div/input";
        public const string REALISATION_TIME = "RealisationDateTime";
        private const string REALISATION_DATE = "realisation-date-picker";
        private const string TASK = "taskDefinitionSelect";

        public const string BACK_TO_LIST = "/html/body/div[2]/a/span[1]";
        public const string STATUS = "Status";
        private const string SITE = "SelectedSites_ms";

        private const string PERIOD_FROM = "datefrom-date-picker";
        private const string PERIOD_TO = "dateto-date-picker";




        //__________________________________Variables_______________________________________

        // Général

        [FindsBy(How = How.Id, Using = NAME)]
        private IWebElement _name;
        [FindsBy(How = How.Id, Using = START_TIME)]
        private IWebElement _starttime;
        [FindsBy(How = How.Id, Using = END_TIME)]
        private IWebElement _endtime;

        [FindsBy(How = How.Id, Using = REALISATION_TIME)]
        private IWebElement _realisationTime;

        [FindsBy(How = How.Id, Using = REALISATION_DATE)]
        private IWebElement _realisationDate;


        [FindsBy(How = How.Id, Using = STATUS)]
        private IWebElement _status;

        [FindsBy(How = How.Id, Using = TASK)]
        private IWebElement _taskField;


        [FindsBy(How = How.Id, Using = SITE)]
        private IWebElement _site;

        [FindsBy(How = How.Id, Using = PERIOD_FROM)]
        private IWebElement _from;

        [FindsBy(How = How.Id, Using = PERIOD_TO)]
        private IWebElement _to;
        //__________________________________Methodes_______________________________________

        public void EditName(string name)
        {
            _name = WaitForElementIsVisible(By.Id(NAME));
            _name.SetValue(PageBase.ControlType.TextBox, name);
            WaitLoading();

        }
        public void EditStartTime(string updateTime)
        {
            _starttime = WaitForElementIsVisible(By.Id(START_TIME));
            _starttime.Clear();
            _starttime.SendKeys(updateTime);
            WaitLoading();
        }

        public void EditEndTime(string updateTime)
        {
            WaitLoading();
            _endtime = WaitForElementIsVisible(By.Id(END_TIME));
            _endtime.Clear();
            _endtime.SendKeys(updateTime);
            WaitLoading();
        }

        public string GetStartTime()
        {
            _starttime = WaitForElementIsVisible(By.Id(START_TIME));
            return _starttime.GetAttribute("value");
        }
        public string GetEndTime()
        {
            WaitPageLoading();
            _endtime = WaitForElementIsVisible(By.Id(END_TIME));
            return _endtime.GetAttribute("value");
        }


        public void EditRealisationTime()
        {
            _realisationTime.SendKeys("14:30");
            WaitLoading();

        }

        public string GetRealisationTime()
        {
            _realisationTime = WaitForElementIsVisible(By.Id(REALISATION_TIME));
            return _realisationTime.GetAttribute("value");
        }
        public void EditRealisationDate(DateTime updateRealisationDate)
        {
            _realisationDate = WaitForElementIsVisible(By.Id(REALISATION_DATE));
            _realisationDate.SetValue(PageBase.ControlType.DateTime, updateRealisationDate);
            WaitLoading();

        }

        public string GetRealisationDate()
        {
            _realisationDate = WaitForElementIsVisible(By.Id(REALISATION_DATE));
            return _realisationDate.GetAttribute("value");
        }
        public void EditStatus(string status)
        {
            _status = WaitForElementIsVisible(By.Id(STATUS));
            _status.SetValue(PageBase.ControlType.DropDownList, status);
            WaitLoading();
        }
        public SchedulerPage BackToList()
        {
            var backToListBtn = WaitForElementIsVisible(By.XPath(BACK_TO_LIST));
            backToListBtn.Click();
            WaitLoading();
            return new SchedulerPage(_webDriver, _testContext);
        }

        public bool IsTaskFieldDisabled()
        {
            _taskField = WaitForElementIsVisible(By.Id(TASK));

            return _taskField.GetAttribute("disabled") == "true";
        }
        //public void EditSite(string site) //List<string> sitesInput)
        //{
        //    //foreach (var site in sitesInput)
        //        ComboBoxSelectMultipleById(new ComboBoxOptions(SITES, site, false));
        //}
        public void ComboBoxSelectMultipleById(ComboBoxOptions cbOpt)
        {
            IWebElement input = WaitForElementIsVisible(By.Id(cbOpt.XpathId));
            input.Click();
            WaitForLoad();

            //if (cbOpt.ClickCheckAllAtStart)
            //{
            //    var checkAllVisible = SolveVisible("//*/span[text()='Check all']");
            //    Assert.IsNotNull(checkAllVisible);
            //    checkAllVisible.Click();
            //}
            //else if (cbOpt.ClickUncheckAllAtStart)
            //{
            //    var uncheckAllVisible = SolveVisible("//*/span[text()='Uncheck all']");
            //    Assert.IsNotNull(uncheckAllVisible);
            //    uncheckAllVisible.Click();
            //}

            if (cbOpt.IsUsedInFilter)
            {
                WaitPageLoading();
                WaitForLoad();
            }
            else if (cbOpt.ClickUncheckAllAtStart)
            {
                WaitForLoad();
            }

            bool selectionWasModified = false;

            if (cbOpt.SelectionValue != null)
            {
                var searchVisible = SolveVisible("//*/input[@type='search']");
                Assert.IsNotNull(searchVisible);
                searchVisible.SetValue(ControlType.TextBox, cbOpt.SelectionValue);
                // on ne clique pas de checkbox donc pas de rechargement de page ici
                Thread.Sleep(3000);
                WaitForLoad();

                var select = SolveVisible("//*/label[contains(@for, 'ui-multiselect')]/span[contains(text(),'" + cbOpt.SelectionValue + "')]");
                Assert.IsNotNull(select, "Pas de sélection de " + cbOpt.SelectionValue);

                if (cbOpt.ClickCheckAllAfterSelection)
                {
                    var checkAllVisible = SolveVisible("//*/span[text()='Check all']");
                    Assert.IsNotNull(checkAllVisible);
                    checkAllVisible.Click();
                }
                else if (cbOpt.ClickUncheckAllAfterSelection)
                {
                    var uncheckAllVisible = SolveVisible("//*/span[text()='Uncheck all']");
                    Assert.IsNotNull(uncheckAllVisible);
                    uncheckAllVisible.Click();
                }
                else
                {
                    select.Click();
                }
                selectionWasModified = true;
            }

            if (selectionWasModified)
            {
                if (cbOpt.IsUsedInFilter)
                {
                    WaitPageLoading();
                    WaitForLoad();
                }
                else
                {
                    WaitForLoad();
                }
            }

            input = WaitForElementIsVisible(By.Id(cbOpt.XpathId));

            try
            {
                input.SendKeys(Keys.Enter);
            }
            catch
            {
                //Silent catch: sometimes there's no associated action with "enter" key
                input.Click();
            }

            WaitForLoad();
        }

        public void EditSite(string site)
        {
            _site = WaitForElementExists(By.Id(SITE));
            _site.Click();
            WaitForLoad();
            var unselectAll = WaitForElementIsVisible(By.XPath("/html/body/div[11]/div/ul/li[2]/a/span[2]"));
            unselectAll.Click();
            // secousse message rouge validator
            WaitForLoad();
            var search = WaitForElementIsVisible(By.XPath("/html/body/div[11]/div/div/label/input"));
            search.SetValue(ControlType.TextBox, site);
            WaitForLoad();
            var siteToCheck = WaitForElementIsVisible(By.XPath("//span[contains(text(), '" + site + "')]"));
            siteToCheck.SetValue(ControlType.CheckBox, true);
            WaitPageLoading();
            WaitForLoad();
        }
        public void EditValidPeriod(DateTime periodFrom , DateTime periodTo)
        {
            _from = WaitForElementIsVisible(By.Id(PERIOD_FROM));
            _from.SetValue(ControlType.DateTime, periodFrom);
            _from.SendKeys(Keys.Enter);
            _to = WaitForElementIsVisible(By.Id(PERIOD_TO));
            _to.SetValue(ControlType.DateTime, periodTo);
            _to.SendKeys(Keys.Enter);
            WaitPageLoading();
            WaitForLoad();
          
        }
        public string GePeriodFromDate()
        {
            _from = WaitForElementIsVisible(By.Id(PERIOD_FROM));
            return _from.GetAttribute("value") ;
        }
        public string GePeriodToDate()
        {
            _to = WaitForElementIsVisible(By.Id(PERIOD_TO));
            return _to.GetAttribute("value");
        }


    }
}