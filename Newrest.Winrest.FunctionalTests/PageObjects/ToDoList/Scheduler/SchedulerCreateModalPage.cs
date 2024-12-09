using DocumentFormat.OpenXml.Drawing;
using Limilabs.Client.IMAP;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Purchasing;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Threading;
using System.Windows.Forms;
using Keys = OpenQA.Selenium.Keys;

namespace Newrest.Winrest.FunctionalTests.PageObjects.ToDoList.Scheduler
{
    public class SchedulerCreateModalPage : PageBase
    {
        public SchedulerCreateModalPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        //---------------------------------------CONSTANTES---------------------------------------------------------------
        public const string NAME = "Name";
        public const string TASK = "taskDefinitionSelect";
        public const string SITES = "SelectedSites_ms";
        public const string REALISATION_DATE = "realisation-date-picker";
        public const string REALISATION_TIME = "RealisationDateTime";
        public const string STATUS = "//select[@name=\"Status\"]";
        public const string VALID_FROM = "datefrom-date-picker";
        public const string VALID_TO = "dateto-date-picker";
        public const string START_TIME = "TimeFrom";
        public const string END_TIME = "TimeTo";
        public const string REPEAT_TIME = "TimeValueRepeat";
        public const string ACTIVE_ON_MONDAY = "IsActiveOnMonday";
        public const string ACTIVE_ON_TUESDAY = "IsActiveOnTuesday";
        public const string ACTIVE_ON_WEDNESDAY = "IsActiveOnWednesday";
        public const string ACTIVE_ON_THURSDAY = "IsActiveOnThursday";
        public const string ACTIVE_ON_FRIDAY = "IsActiveOnFriday";
        public const string ACTIVE_ON_SATURDAY = "IsActiveOnSaturday";
        public const string ACTIVE_ON_SUNDAY = "IsActiveOnSunday";
        public const string CREATE_BTN = "/html/body/div[3]/div/div/div[2]/div/div/form/div/div/div[2]/div/button[2]";
        //tableau
        private const string ITEM = "/[@id=\"list-item-with-action\"]/div/div[2]/div[*]/div/div/form/div[2]/div[1]";
        private const string ITEM_NAME = "/[@id=\"list-item-with-action\"]/div/div[2]/div[*]/div/div/form/div[2]/div[3]/span";

        // ____________________________________ Variables _______________________________________________
        [FindsBy(How = How.Id, Using = NAME)]
        private IWebElement _name;

        [FindsBy(How = How.Id, Using = TASK)]
        private IWebElement _task;

        [FindsBy(How = How.XPath, Using = STATUS)]
        private IWebElement _status;

        [FindsBy(How = How.Id, Using = SITES)]
        private IWebElement _sites;

        [FindsBy(How = How.Id, Using = START_TIME)]
        private IWebElement _startTime;

        [FindsBy(How = How.Id, Using = END_TIME)]
        private IWebElement _endTime;

        [FindsBy(How = How.Id, Using = REALISATION_DATE)]
        private IWebElement _realisationDate;

        [FindsBy(How = How.Id, Using = REALISATION_TIME)]
        private IWebElement _realisationTime;

        [FindsBy(How = How.Id, Using = VALID_FROM)]
        private IWebElement _validFrom;

        [FindsBy(How = How.Id, Using = VALID_TO)]
        private IWebElement _validTo;

        [FindsBy(How = How.Id, Using = REPEAT_TIME)]
        private IWebElement _repeatTime;

        [FindsBy(How = How.Id, Using = ACTIVE_ON_MONDAY)]
        private IWebElement _activeOnMonday;

        [FindsBy(How = How.Id, Using = ACTIVE_ON_TUESDAY)]
        private IWebElement _activeOnTuesday;

        [FindsBy(How = How.Id, Using = ACTIVE_ON_WEDNESDAY)]
        private IWebElement _activeOnWednesday;

        [FindsBy(How = How.Id, Using = ACTIVE_ON_THURSDAY)]
        private IWebElement _activeOnThursday;

        [FindsBy(How = How.Id, Using = ACTIVE_ON_FRIDAY)]
        private IWebElement _activeOnFriday;

        [FindsBy(How = How.Id, Using = ACTIVE_ON_SATURDAY)]
        private IWebElement _activeOnSaturday;

        [FindsBy(How = How.Id, Using = ACTIVE_ON_SUNDAY)]
        private IWebElement _activeOnSunday;

        //tableau

        [FindsBy(How = How.XPath, Using = ITEM_NAME)]
        private IWebElement _itemName;

        [FindsBy(How = How.XPath, Using = ITEM)]
        private IWebElement _firstItem;

        public ScheduleDetailsPage CreateNewSchedule(string nameInput, string taskInput, string sitesInput, DateTime realisationDateInput,
            string statusinput, DateTime validFromInput, DateTime validToInput, string repeatTimeInput, bool selectall = false)
        {
            _name = WaitForElementIsVisible(By.Id(NAME));
            _name.SetValue(ControlType.TextBox, nameInput);

            _task = WaitForElementIsVisible(By.Id(TASK));
            _task.SetValue(ControlType.DropDownList, taskInput);
            _task.SendKeys(Keys.Tab);
            WaitForLoad();
            if (selectall)
            {
                var site = WaitForElementExists(By.Id(SITES));
                site.Click();

                var selectAll = WaitForElementIsVisible(By.XPath("/html/body/div[12]/div/ul/li[1]/a/span[2]"));
                selectAll.Click();

                site = WaitForElementExists(By.Id(SITES));
                site.Click();
                WaitForLoad();
            }
            else
            {
                ComboBoxSelectById(new ComboBoxOptions(SITES, sitesInput, false));
                WaitForLoad();

            }
            _realisationDate = WaitForElementIsVisible(By.Id(REALISATION_DATE));
            _realisationDate.SetValue(ControlType.DateTime, realisationDateInput);
            _realisationDate.SendKeys(Keys.Tab);

            _realisationTime.SendKeys(Keys.Tab);

            _status = WaitForElementIsVisible(By.XPath(STATUS));
            _status.SetValue(ControlType.DropDownList, statusinput);

            _validFrom = WaitForElementIsVisible(By.Id(VALID_FROM));
            _validFrom.SetValue(ControlType.DateTime, validFromInput);
            _validFrom.SendKeys(Keys.Tab);

            _validTo = WaitForElementIsVisible(By.Id(VALID_TO));
            _validTo.SetValue(ControlType.DateTime, validToInput);
            _validTo.SendKeys(Keys.Tab);

            _startTime.SendKeys(Keys.Tab);

            _endTime.SendKeys(Keys.Tab);

            _activeOnMonday.SendKeys(Keys.Tab);

            _activeOnTuesday.SendKeys(Keys.Tab);

            _activeOnWednesday.SendKeys(Keys.Tab);

            _activeOnThursday.SendKeys(Keys.Tab);

            _activeOnFriday.SendKeys(Keys.Tab);

            _activeOnSaturday.SendKeys(Keys.Tab);

            _activeOnSunday.SendKeys(Keys.Tab);

            _repeatTime = WaitForElementIsVisible(By.Id(REPEAT_TIME));
            _repeatTime.SetValue(ControlType.TextBox, repeatTimeInput);

            var createBtn = WaitForElementIsVisible(By.XPath(CREATE_BTN));
            createBtn.Click();
            WaitPageLoading();
            WaitLoading();
            return new ScheduleDetailsPage(_webDriver, _testContext);
        }
        public ScheduleDetailsPage CreateNewScheduleForTabletApp(
          string nameInput, string taskInput, string sitesInput, DateTime realisationDateInput,
          string statusInput, DateTime validFromInput, DateTime validToInput, string repeatTimeInput,
          string realisationTimeInput, string endTimeInput, bool selectAll = false)
        {
            _name = WaitForElementIsVisible(By.Id(NAME));
            _name.SetValue(ControlType.TextBox, nameInput);

            _task = WaitForElementIsVisible(By.Id(TASK));
            _task.SetValue(ControlType.DropDownList, taskInput);
            _task.SendKeys(Keys.Tab);
            WaitForLoad();

            if (selectAll)
            {
                var site = WaitForElementExists(By.Id(SITES));
                site.Click();

                var selectAllOption = WaitForElementIsVisible(By.XPath("/html/body/div[12]/div/ul/li[1]/a/span[2]"));
                selectAllOption.Click();

                site = WaitForElementExists(By.Id(SITES));
                site.Click();
            }
            else
            {
                ComboBoxSelectById(new ComboBoxOptions(SITES, sitesInput, false));
                WaitForLoad();
            }
            WaitForLoad();

            _realisationDate = WaitForElementIsVisible(By.Id(REALISATION_DATE));
            _realisationDate.SetValue(ControlType.DateTime, realisationDateInput);
            _realisationDate.SendKeys(Keys.Tab);

            var _startTime = WaitForElementIsVisible(By.Id(START_TIME));
            _startTime.Clear();
            _startTime.SendKeys(realisationTimeInput);

            var _endTime = WaitForElementIsVisible(By.Id(END_TIME));
            _endTime.Clear();
            _endTime.SendKeys(endTimeInput);

            _status = WaitForElementIsVisible(By.XPath(STATUS));
            _status.SetValue(ControlType.DropDownList, statusInput);

            _validFrom = WaitForElementIsVisible(By.Id(VALID_FROM));
            _validFrom.SetValue(ControlType.DateTime, validFromInput);
            _validFrom.SendKeys(Keys.Tab);

            _validTo = WaitForElementIsVisible(By.Id(VALID_TO));
            _validTo.SetValue(ControlType.DateTime, validToInput);
            _validTo.SendKeys(Keys.Tab);

            _activeOnMonday = WaitForElementExists(By.Id("IsActiveOnMonday"));
            _activeOnMonday.SetValue(ControlType.CheckBox, true);

            _activeOnTuesday = WaitForElementExists(By.Id("IsActiveOnTuesday"));
            _activeOnTuesday.SetValue(ControlType.CheckBox, true);

            _activeOnWednesday = WaitForElementExists(By.Id("IsActiveOnWednesday"));
            _activeOnWednesday.SetValue(ControlType.CheckBox, true);

            _activeOnThursday = WaitForElementExists(By.Id("IsActiveOnThursday"));
            _activeOnThursday.SetValue(ControlType.CheckBox, true);


            _activeOnFriday.SendKeys(Keys.Tab);
            _activeOnSaturday.SendKeys(Keys.Tab);
            _activeOnSunday.SendKeys(Keys.Tab);

            _repeatTime = WaitForElementIsVisible(By.Id(REPEAT_TIME));
            _repeatTime.SetValue(ControlType.TextBox, repeatTimeInput);

            var repeatTimeHour = WaitForElementIsVisible(By.Id("TypeOfTimeRepeat"));
            repeatTimeHour.Click();
            var element = WaitForElementIsVisible(By.XPath("//select[@id='TypeOfTimeRepeat']/option[@value='1']"));
            repeatTimeHour.SetValue(ControlType.DropDownList, element.Text);
            repeatTimeHour.Click();

            var createBtn = WaitForElementIsVisible(By.XPath(CREATE_BTN));
            createBtn.Click();
            WaitPageLoading();
            WaitLoading();

            return new ScheduleDetailsPage(_webDriver, _testContext);
        }

        public void ComboBoxSelectMultipleById(ComboBoxOptions cbOpt)
        {
            IWebElement input = WaitForElementIsVisible(By.Id(cbOpt.XpathId));
            input.Click();
            WaitForLoad();

            //if (cbOpt.ClickCheckAllAtStart)
            //{
            //    var checkAllVisible = SolveVisible("//span[text()='Check all']");
            //    Assert.IsNotNull(checkAllVisible);
            //    checkAllVisible.Click();
            //}
            //else if (cbOpt.ClickUncheckAllAtStart)
            //{
            //    var uncheckAllVisible = SolveVisible("//span[text()='Uncheck all']");
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
        public string GetFirstItemName()
        {
            if (isElementExists(By.XPath(ITEM_NAME)))
            {
                _itemName = WaitForElementExists(By.XPath(ITEM_NAME));
                return _itemName.GetAttribute("site");
            }
            else
            {
                return "";
            }
        }
        public string SelectFirstItem()
        {
            _firstItem = WaitForElementIsVisible(By.XPath(ITEM));
            var value = _firstItem.Text;
            _firstItem.Click();
            WaitForLoad();
            return value;
        }
       
    }
}
