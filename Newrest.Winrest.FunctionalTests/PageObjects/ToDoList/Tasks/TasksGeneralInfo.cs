using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Threading;

namespace Newrest.Winrest.FunctionalTests.PageObjects.ToDoList.Tasks
{
    public class TasksGeneralInfo : PageBase
    {
        public const string BACK_TO_LIST = "/html/body/div[2]/a/span[2]";
        public const string CLICKON_GENERALINFO = "generalInformation";
        public const string CLICKON_SCHEDULERS = "//*/a[text()='Schedulers']/..";
        public const string NAME = "Name";
        public const string SITE = "SelectedSites_ms";
        public const string SITE_SELECTED = "//*[@id=\"SelectedSitesTrueValue\"]/option[@selected]";
        public const string SITE_VALUE = "SelectedSitesTrueValue";
        public const string CUSTOMER = "CustomerId";
        public const string VEHICULE = "VehiculeId";
        public const string ENFORCE_ORDER = "EnforceOrder";
        public const string AIRCRAFT = "SelectedAircrafts_ms";
        public const string AIRCRAFT_VALUE = "SelectedAircraftsTrueValue";
        public const string DESCRIPTION = "Description";
        public const string COMMENT = "Comment";
        public const string ACTIVE = "IsActive";
        public const string DURATION = "taskDuration";
        //---------------------------------------ADD NEW STEP CONSTANTES---------------------------------------------------------------
        public const string ADD_NEW_STEP_PLUS = "btn-add-task-step";
        public const string NAME_NEW_STEP = "/html/body/div[4]/div/div/div/div/form/div[2]/div/div/div[1]/div/div[1]/div/input";
        public const string PHOTO_REQUIRED_NEW_STEP = "PhotoResultRequired";
        public const string INSTRUCTIONS = "Instructions";
        public const string COMMENT_EDIT = "//div[@class='item-create-container']/div[1]/div[1]/div[4]/div/textarea[@id=\"Comment\"]";
        public const string PHOTO_NEW_STEP = "FileSent";
        public const string CREATE_NEW_STEP_BTN = "/html/body/div[4]/div/div/div/div/form/div[2]/div/div/div[2]/div/a[2]";
        public const string FIRST_LINE_STEP_SELECTED = "/html/body/div[3]/div/div/div/div/div[3]/div[1]/table/tbody/tr";
        public const string DELETE_FIRST_STEP = "/html/body/div[3]/div/div/div/div/div[3]/div[1]/table/tbody/tr[1]/td[4]/a[1]";
        public const string DELETE_STEP_CONFIRM = "/html/body/div[13]/div/div/div[3]/a[1]";
        public const string CONFIRM_ACTION_DELETE_STEP = "/html/body/div[13]/div/div/div[3]/a[1]";
        public const string LIST_LINE_STEP_DELETE = "/html/body/div[3]/div/div/div/div/div[3]/div[1]/table/tbody/tr[*]/td[5]/a[1]/span";
        public const string ITEM_DETAILS_BODY = "/html/body/div[3]/div/div/div/div/div[3]/div[1]/table/tbody";
        public const string STYLO_EDIT = "//span[@class='fas fa-pencil-alt']";
        public const string VALIDATE_BUTTON = "/html/body/div[4]/div/div/div[2]/div/form/div[2]/button[2]";
        public const string CANCEL_BUTTON = "//*[@id=\"stepCreateContainer\"]/div/div/div[2]/div/button[1]";
        public const string TASK_DURATION = "taskDuration";
        public const string STEP_NAME = "/html/body/div[3]/div/div/div/div/div[3]/div[1]/table/tbody/tr[1]/td[2]";
        public const string STEP_NAME_INPUT = "/html/body/div[3]/div/div/div/div/div[3]/div[1]/table/tbody/tr/td[2]/input[4]";

        public const string STEP_PICTURE_REQUIRED = "//*[@id=\"step_PhotoResultRequired\"]";

        public const string NAME_STEP = "/html/body/div[4]/div/div/div/div/form/div[2]/div/div/div[1]/div[1]/div[1]/div/input";
        public const string INSTRUCTION_STEP = "//*[@id=\"Instructions\"]";
        public const string UPDATE_BUTTON = "//*/button[@value='Update']";
        public const string UPDATE_BUTTON_PATCH = "//*[@id=\"stepCreateContainer\"]/div/div/div[2]/div/a";
        // il y a plusieurs ID Comment dans la meme interface Add New Step c'est pour ça on a utilisé full XPath
        public const string COMMENT_NEW_STEP = "/html/body/div[4]/div/div/div/div/form/div[2]/div/div/div[1]/div[1]/div[4]/div/textarea";
        public const string REALISATION_TIME = "RealisationDateTime";
        public const string REALISATION_DATE = "realisation-date-picker";
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
        public const string VALID_FROM = "datefrom-date-picker";
        public const string VALID_TO = "dateto-date-picker";
        public const string TASK = "taskDefinitionSelect";
        public const string SITES = "SelectedSites_ms";
        public const string STATUS = "//select[@name=\"Status\"]";
        public const string PLUS_BTN = "/html/body/div[2]/div/div[2]/div/div[1]/div/div/button";
        public const string FIRSTSCHUDLERNAME = "//*[@id=\"list-item-with-action\"]/table/tbody/tr/td[2]";
        public const string FIRSTTASKNAMEONSCHUDLER = "//*[@id=\"list-item-with-action\"]/table/tbody/tr/td[3]";
        public const string FIRSTSITENAMEONSCHUDELR = "//*[@id=\"list-item-with-action\"]/table/tbody/tr/td[4]";
        public const string FIRSTSTATUSONSCHEDULERS = "//*[@id=\"list-item-with-action\"]/table/tbody/tr/td[5]";
        public const string MONDAY = "//*[@id=\"item-filter-from\"]/div/div/div[1]/div/div[11]/div/div[1]/div[1]/div";
        public const string TUESDAY = "//*[@id=\"item-filter-from\"]/div/div/div[1]/div/div[11]/div/div[2]/div[1]/div";
        public const string WEDNESDAY = "//*[@id=\"item-filter-from\"]/div/div/div[1]/div/div[11]/div/div[3]/div[1]/div";
        public const string THURSDAY = "//*[@id=\"item-filter-from\"]/div/div/div[1]/div/div[11]/div/div[1]/div[2]/div";
        public const string FRIDAY = "//*[@id=\"item-filter-from\"]/div/div/div[1]/div/div[11]/div/div[2]/div[2]/div";
        public const string SATURDAY = "//*[@id=\"item-filter-from\"]/div/div/div[1]/div/div[11]/div/div[3]/div[2]/div";
        public const string SUNDAY = "//*[@id=\"item-filter-from\"]/div/div/div[1]/div/div[11]/div/div[1]/div[3]/div";


        // ____________________________________ Variables _______________________________________________

        [FindsBy(How = How.XPath, Using = FIRSTSCHUDLERNAME)]
        private IWebElement _firstschduerlname;
        [FindsBy(How = How.XPath, Using = FIRSTTASKNAMEONSCHUDLER)]
        private IWebElement _firsttasknameschudeler;
        [FindsBy(How = How.XPath, Using = FIRSTSITENAMEONSCHUDELR)]
        private IWebElement _firstsitenameonschduler;
        [FindsBy(How = How.XPath, Using = FIRSTSTATUSONSCHEDULERS)]
        private IWebElement _firststatutonschduler;
        [FindsBy(How = How.Id, Using = ADD_NEW_STEP_PLUS)]
        private IWebElement _addNewStepPlus;

        [FindsBy(How = How.Id, Using = NAME_NEW_STEP)]
        private IWebElement _nameNewStep;

        [FindsBy(How = How.Id, Using = PHOTO_REQUIRED_NEW_STEP)]
        private IWebElement _photoRequiredNewStep;

        [FindsBy(How = How.Id, Using = INSTRUCTIONS)]
        private IWebElement _instruction;

        [FindsBy(How = How.Id, Using = COMMENT)]
        private IWebElement _comment;

        [FindsBy(How = How.Id, Using = PHOTO_NEW_STEP)]
        private IWebElement _photoNewStep;

        [FindsBy(How = How.XPath, Using = COMMENT_NEW_STEP)]
        private IWebElement _commentNewStep;

        [FindsBy(How = How.Id, Using = REALISATION_TIME)]
        private IWebElement _realisationTime;

        [FindsBy(How = How.Id, Using = START_TIME)]
        private IWebElement _startTime;

        [FindsBy(How = How.Id, Using = END_TIME)]
        private IWebElement _endTime;

        [FindsBy(How = How.Id, Using = REALISATION_DATE)]
        private IWebElement _realisationDate;

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

        [FindsBy(How = How.XPath, Using = STATUS)]
        private IWebElement _status;

        [FindsBy(How = How.XPath, Using = PLUS_BTN)]
        private IWebElement _plusBtn;


        public TasksGeneralInfo(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        public TasksPage BackToList()
        {
            var backToListBtn = WaitForElementIsVisible(By.XPath(BACK_TO_LIST));
            backToListBtn.Click();
            WaitPageLoading();
            return new TasksPage(_webDriver, _testContext);
        }

        public void ClickOnGeneralInfoTab()
        {
            var generalInfo = WaitForElementIsVisible(By.Id(CLICKON_GENERALINFO));
            generalInfo.Click();
            WaitPageLoading();
        }

        public void ClickOnSchedulersTab()
        {
            _webDriver.FindElement(By.TagName("html")).SendKeys(Keys.PageUp);
            _webDriver.FindElement(By.TagName("html")).SendKeys(Keys.PageUp);
            // le tableau bouge
            Thread.Sleep(1000);
            var schedulers = WaitForElementIsVisible(By.XPath(CLICKON_SCHEDULERS));
            schedulers.Click();
            WaitPageLoading();
        }

        public string GetName()
        {
            var name = WaitForElementIsVisible(By.Id(NAME));
            WaitForLoad();
            return name.GetAttribute("value");
        }
        public string GetNameOnSchedulers()
        {
            _firstschduerlname = WaitForElementIsVisible(By.XPath(FIRSTSCHUDLERNAME));
            return _firstschduerlname.Text;
        }
        public string GetTaskNameOnSchedulers()
        {
            _firsttasknameschudeler = WaitForElementIsVisible(By.XPath(FIRSTTASKNAMEONSCHUDLER));
            return _firsttasknameschudeler.Text;
        }
        public string GetSiteOnSchedulers()
        {
            _firstsitenameonschduler = WaitForElementIsVisible(By.XPath(FIRSTSITENAMEONSCHUDELR));
            return _firstsitenameonschduler.Text;
        }
        public string GetStatutOnSchedulers()
        {
            _firststatutonschduler = WaitForElementIsVisible(By.XPath(FIRSTSTATUSONSCHEDULERS));
            return _firststatutonschduler.Text;
        }
        public void SetName(string value)
        {
            var name = WaitForElementIsVisible(By.Id(NAME));
            name.SetValue(ControlType.TextBox, value);
            WaitPageLoading();
        }

        public bool GetSite(string site)
        {
            var sites = WaitForElementExists(By.Id(SITE_VALUE));

            var liste = new SelectElement(sites).AllSelectedOptions;
            foreach (IWebElement element in liste)
            {
                //if (element.GetAttribute("value")==site)
                if (element.GetAttribute("innerHTML") == site)
                {
                    return true;
                }
            }
            return false;
        }

        public void SetSite(string value)
        {
            try
            {
                // premier coup ça vide
                ComboBoxSelectById(new ComboBoxOptions(SITE, value));
                WaitPageLoading();
            }
            catch
            {
                WaitPageLoading();
                // second coup ça rempli
                ComboBoxOptions cbOpt = new ComboBoxOptions(SITE, value) { ClickUncheckAllAtStart = false };
                ComboBoxSelectById(cbOpt);
                WaitPageLoading();
            }
        }
                
        public string GetCustomer()
        {
            var customer = WaitForElementIsVisible(By.Id(CUSTOMER));
            WaitForLoad();
            return new SelectElement(customer).SelectedOption.Text;
        }
        public void SetCustomer(string value)
        {
            var customer = WaitForElementIsVisible(By.Id(CUSTOMER));
            new SelectElement(customer).SelectByText(value);
            WaitPageLoading();
        }
        public string GetVehiculeRegistration()
        {
            var vehicule = WaitForElementIsVisible(By.Id(VEHICULE));
            return new SelectElement(vehicule).SelectedOption.Text;
        }

        public void SetVehiculeRegistration(string value)
        {
            var vehicule = WaitForElementIsVisible(By.Id(VEHICULE));
            new SelectElement(vehicule).SelectByText(value);
            WaitPageLoading();
        }
        public bool IsEnforceOrder()
        {
            var enforce = WaitForElementExists(By.Id(ENFORCE_ORDER));
            return enforce.Selected;
        }
        public void SetEnforceOrder(bool value)
        {
            var enforce = WaitForElementExists(By.Id(ENFORCE_ORDER));
            enforce.SetValue(ControlType.CheckBox, value);
            WaitPageLoading();
        }
        public bool GetAirCraft(string aircraft)
        {
            var aircrafts = WaitForElementExists(By.Id(AIRCRAFT_VALUE));
            var liste = new SelectElement(aircrafts).AllSelectedOptions;
            foreach (var element in liste)
            {
                if (element.GetAttribute("innerHTML") == aircraft)
                {
                    return true;
                }
            }
            return false;

        }
        public void SetAirCraft(string value)
        {
            try
            {
                // premier coup ça vide
                ComboBoxSelectById(new ComboBoxOptions(AIRCRAFT, value));
                WaitPageLoading();
            }
            catch
            {
                WaitPageLoading();
                // second coup ça rempli
                ComboBoxOptions cbOpt = new ComboBoxOptions(AIRCRAFT, value) { ClickUncheckAllAtStart = false };
                ComboBoxSelectById(cbOpt);
                WaitPageLoading();
            }
        }
        public string GetDescription()
        {
            var description = WaitForElementIsVisible(By.Id(DESCRIPTION));
            return description.Text;
        }
        public void SetDescription(string value)
        {
            var description = WaitForElementIsVisible(By.Id(DESCRIPTION));
            description.SetValue(ControlType.TextBox, value);
            WaitPageLoading();
            Thread.Sleep(2000);
        }
        public string GetComment()
        {
            var comment = WaitForElementIsVisible(By.Id(COMMENT));
            return comment.Text;
        }
        public void SetComment(string value)
        {
            var comment = WaitForElementIsVisible(By.Id(COMMENT));
            comment.SetValue(ControlType.TextBox, value);
            WaitPageLoading();
        }
        public bool IsActive()
        {
            var active = WaitForElementExists(By.Id(ACTIVE));
            return active.Selected;

        }
        public void SetActive(bool value)
        {
            var active = WaitForElementExists(By.Id(ACTIVE));
            active.SetValue(ControlType.CheckBox, value);
            Thread.Sleep(1000);
        }
        public string GetDuration()
        {
            var duration = WaitForElementIsVisible(By.Id(DURATION));
            return duration.GetAttribute("value");
        }
        public void SetDuration(string value)
        {
            var duration = WaitForElementIsVisible(By.Id(DURATION));
            duration.SetValue(ControlType.TextBox, value);
            Thread.Sleep(1000);
        }
        public void AddNewStep(string name, bool photorequired, string instruction, string comment, FileInfo fileInfo)
        {
            _addNewStepPlus = WaitForElementIsVisible(By.Id(ADD_NEW_STEP_PLUS));
            _addNewStepPlus.Click();

            _nameNewStep = WaitForElementIsVisible(By.XPath(NAME_NEW_STEP));
            // _nameNewStep.Click();
            _nameNewStep.SendKeys(name);
            _nameNewStep.SendKeys(Keys.Tab);

            _instruction = WaitForElementIsVisible(By.Id(INSTRUCTIONS));
            _instruction.SendKeys(instruction);
            _instruction.SendKeys(Keys.Tab);

            // il ya plusieurs ID "Comment" dans la meme interface 
            _commentNewStep = WaitForElementExists(By.XPath(COMMENT_NEW_STEP));
            _commentNewStep.SendKeys(comment);
            _commentNewStep.SendKeys(Keys.Tab);

          
            var iconePhoto = WaitForElementIsVisible(By.XPath("//*[@id=\"stepCreateContainer\"]/div/div/div[1]/div[2]/div/div/div[1]/div/div/span[2]"));
            iconePhoto.Click();
          
            _photoNewStep = WaitForElementIsVisible(By.Id(PHOTO_NEW_STEP));
            _photoNewStep.SendKeys(fileInfo.FullName);

            if (IsDev())
            {
                var createNewStepBtn = WaitForElementIsVisible(By.XPath("//*/a[text()='Create']"));
                createNewStepBtn.Click();
            }
            else
            {
                var createNewStepBtn = WaitForElementIsVisible(By.XPath(CREATE_NEW_STEP_BTN));
                createNewStepBtn.Click();
            }
            WaitPageLoading();
            WaitForLoad();
        }
        public bool IsExistTaskDetails()
        {
            var listLineStepsSelected = _webDriver.FindElements(By.XPath("//*[@id=\"stepsTable\"]/tbody/tr[*]/td[2]"));
            if (listLineStepsSelected.Count != 0)
            {
                return true;
            }
            return false;
        }
        public void DeleteStep_TaskDetails()
        {
            var deleteStepList = _webDriver.FindElements(By.XPath(LIST_LINE_STEP_DELETE));
            if (deleteStepList.Count != 0)
            {
                foreach (var deleteStep in deleteStepList)
                {
                    // la page se recharge pour chaque entrée
                    //deleteStep.Click();
                    var deleteFirst = WaitForElementIsVisible(By.XPath(LIST_LINE_STEP_DELETE));
                    deleteFirst.Click();
                    WaitForLoad();
                    var deleteStepConfirm = WaitForElementIsVisible(By.Id("dataConfirmOK"));
                    deleteStepConfirm.Click();
                    WaitForLoad();
                    var confirmDeleteStep = WaitForElementIsVisible(By.Id("dataConfirmOK"));
                    confirmDeleteStep.Click();
                    WaitPageLoading();
                }
                WaitForLoad();
                WaitPageLoading();
            }
        }
        public bool VerifyDeleteTasksDetails()
        {
            var listLineStepsSelected = _webDriver.FindElements(By.XPath(ITEM_DETAILS_BODY));
            if (listLineStepsSelected.Count == 0)
            {
                return true;
            }
            return false;
        }
        public void EditTask(string newStepName, string newInstruction, string newComment)
        {
            WaitForElementIsVisible(By.XPath(STEP_NAME)).Click();
            var stepNameInput = WaitForElementIsVisible(By.XPath(STEP_NAME_INPUT));
            stepNameInput.SetValue(ControlType.TextBox, newStepName);
            WaitPageLoading();

            WaitForElementIsVisible(By.Id(TASK_DURATION)).SetValue(ControlType.TextBox,"3");
            WaitPageLoading();

            WaitForLoad();
            var editStylo = WaitForElementIsVisible(By.XPath(STYLO_EDIT));
            editStylo.Click();
            WaitPageLoading();
            _instruction = WaitForElementIsVisible(By.Id(INSTRUCTIONS));
            _instruction.SetValue(ControlType.TextBox, newInstruction);
            _instruction.SendKeys(Keys.Tab);
             
            var comment = WaitForElementIsVisible(By.XPath(COMMENT_EDIT));
            comment.SetValue(ControlType.TextBox, newComment);
            comment.SendKeys(Keys.Tab);

            if (IsDev())
            {
                var updateBtn = WaitForElementIsVisible(By.XPath("//*/a[text()='Update']"));
                updateBtn.Click();
            }
            else
            {
                var updateBtn = WaitForElementIsVisible(By.XPath(UPDATE_BUTTON_PATCH));
                updateBtn.Click();
            }
            WaitForLoad();
          
        }
        public bool VerifyEditTask(string newStepName, string newInstruction, string newComment)
        {
            //WaitForElementIsVisible(By.XPath(STEP_NAME)).Click();
            var stepNameInput = WaitForElementIsVisible(By.XPath(STEP_NAME));
            if (stepNameInput.Text != newStepName)
            {
                return false;
            }

            var duration = WaitForElementIsVisible(By.Id(TASK_DURATION)).GetAttribute("value");
            if (duration == null || duration == "0")
            {
                return false;
            }

            var editStylo = WaitForElementIsVisible(By.XPath(STYLO_EDIT));
            editStylo.Click();
            WaitForLoad();
            _instruction = WaitForElementIsVisible(By.Id(INSTRUCTIONS));
            if (_instruction.Text != newInstruction)
            {
                return false;
            }

            var comment = WaitForElementIsVisible(By.XPath(COMMENT_EDIT)).Text;
            if (comment != newComment)
            {
                return false;
            }

            var cancelBtn = WaitForElementIsVisible(By.XPath(CANCEL_BUTTON));
            cancelBtn.Click();
            WaitForLoad();
            return true;
        }


        public bool VerifyPictureRequiredFunctionality()
        {
            var CheckBoxList  = _webDriver.FindElements(By.XPath(STEP_PICTURE_REQUIRED));
            foreach (var checkBoxElement in CheckBoxList) {
                if (checkBoxElement.Enabled == true)
                {
                    return false;
                }
                if (checkBoxElement.GetProperty("title") != "To Modify Click on the pen icon")
                {
                    return false;
                }
            }

            return true;
            
        }

        public bool CheckStepInfoAreCorrect(string name, string instruction)
        {
            WaitForLoad();
            var editStylo = WaitForElementIsVisible(By.XPath(STYLO_EDIT));
            editStylo.Click();
            var testName = WaitForElementIsVisible(By.XPath(NAME_STEP)).GetProperty("value");
            var test1Instruction = WaitForElementIsVisible(By.XPath(INSTRUCTION_STEP)).Text;
            if(name== testName && instruction == test1Instruction)
            {
                return true;
            }
            else
            {
                return false;
            }

        }

        public void AddPicture(FileInfo fileInfo) {
            WaitForLoad();
            var editStylo = WaitForElementIsVisible(By.XPath(STYLO_EDIT));
            var iconePhoto = WaitForElementIsVisible(By.XPath("//*[@id=\"stepCreateContainer\"]/div/div/div[1]/div[2]/div/div/div[1]/div/div/span[2]"));
            iconePhoto.Click();
            _photoNewStep = WaitForElementIsVisible(By.Id(PHOTO_NEW_STEP));
            _photoNewStep.SendKeys(fileInfo.FullName);
            var updateButton = WaitForElementIsVisible(By.XPath("//*[@id=\"stepCreateContainer\"]/div/div/div[2]/div/a"));
            updateButton.Click();
            WaitPageLoading();
            WaitForLoad();
        }
        public string GetSiteValue()
        {
            var siteselected = WaitForElementExists(By.XPath(SITE_SELECTED));
            WaitForLoad();
            return siteselected.GetAttribute("innerHTML");
        }
        public bool VerifyTaskNotCreate()
        {
            if(isElementExists(By.Id("generalInformation")))
            {
                return true ;
            }
            return false;
        }

        public List<string> CreateNewSchedule(string nameInput, string taskInput, string sitesInput, DateTime realisationDateInput,
    string statusInput, DateTime validFromInput, DateTime validToInput, string repeatTimeInput,
    bool activeMonday = false, bool activeTuesday = false, bool activeWednesday = false, bool activeThursday = false,
    bool activeFriday = false, bool activeSaturday = false, bool activeSunday = false, bool selectall = false)
        {
            List<string> plannedDays = new List<string>();

            var _name = WaitForElementIsVisible(By.Id(NAME));
            _name.SetValue(ControlType.TextBox, nameInput);

            var _task = WaitForElementIsVisible(By.Id(TASK));
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

            var _realisationDate = WaitForElementIsVisible(By.Id(REALISATION_DATE));
            _realisationDate.SetValue(ControlType.DateTime, realisationDateInput);
            _realisationDate.SendKeys(Keys.Tab);

            _realisationTime.SendKeys(Keys.Tab);

            _status = WaitForElementIsVisible(By.XPath(STATUS));
            _status.SetValue(ControlType.DropDownList, statusInput);

            _validFrom = WaitForElementIsVisible(By.Id(VALID_FROM));
            _validFrom.SetValue(ControlType.DateTime, validFromInput);
            _validFrom.SendKeys(Keys.Tab);

            _validTo = WaitForElementIsVisible(By.Id(VALID_TO));
            _validTo.SetValue(ControlType.DateTime, validToInput);
            _validTo.SendKeys(Keys.Tab);

            _startTime.SendKeys(Keys.Tab);
            _endTime.SendKeys(Keys.Tab);

            // Handling the "Active on" checkboxes and collecting planned days
            if (activeMonday)
            {
                var mondayCheckbox = WaitForElementIsVisible(By.XPath(MONDAY));
                if (!mondayCheckbox.Selected)
                {
                    mondayCheckbox.Click();
                }
                plannedDays.Add("Monday");
            }

            if (activeTuesday)
            {
                var tuesdayCheckbox = WaitForElementIsVisible(By.XPath(TUESDAY));
                if (!tuesdayCheckbox.Selected)
                {
                    tuesdayCheckbox.Click();
                }
                plannedDays.Add("Tuesday");
            }

            if (activeWednesday)
            {
                var wednesdayCheckbox = WaitForElementIsVisible(By.XPath(WEDNESDAY));
                if (!wednesdayCheckbox.Selected)
                {
                    wednesdayCheckbox.Click();
                }
                plannedDays.Add("Wednesday");
            }

            if (activeThursday)
            {
                var thursdayCheckbox = WaitForElementIsVisible(By.XPath(THURSDAY));
                if (!thursdayCheckbox.Selected)
                {
                    thursdayCheckbox.Click();
                }
                plannedDays.Add("Thursday");
            }

            if (activeFriday)
            {
                var fridayCheckbox = WaitForElementIsVisible(By.XPath(FRIDAY));
                if (!fridayCheckbox.Selected)
                {
                    fridayCheckbox.Click();
                }
                plannedDays.Add("Friday");
            }

            if (activeSaturday)
            {
                var saturdayCheckbox = WaitForElementIsVisible(By.XPath(SATURDAY));
                if (!saturdayCheckbox.Selected)
                {
                    saturdayCheckbox.Click();
                }
                plannedDays.Add("Saturday");
            }

            if (activeSunday)
            {
                var sundayCheckbox = WaitForElementIsVisible(By.XPath(SUNDAY));
                if (!sundayCheckbox.Selected)
                {
                    sundayCheckbox.Click();
                }
                plannedDays.Add("Sunday");
            }

            var _repeatTime = WaitForElementIsVisible(By.Id(REPEAT_TIME));
            _repeatTime.SetValue(ControlType.TextBox, repeatTimeInput);

            var createBtn = WaitForElementIsVisible(By.XPath("//*[@id=\"item-filter-from\"]/div/div/div[2]/div/button[2]"));
            createBtn.Click();
            WaitPageLoading();
            WaitLoading();

            // Retourner les jours planifiés
            return plannedDays;
        }


        public bool IsElementSelected(By by)
        {
            try
            {
                var element = WaitForElementIsVisible(by);
                return element.Selected;
            }
            catch (NoSuchElementException)
            {
                return false;
            }
        }

        public void ShowPlusScheduler()
        {
            _plusBtn = WaitForElementIsVisible(By.XPath("//*[@id=\"list-item-with-action\"]/div[1]/div/div/button"));
            _plusBtn.Click();
            WaitForLoad();
        }

        public void CreateSchedulerModalPage()
        {
            ShowPlusScheduler();
           var _createBtn = WaitForElementIsVisible(By.XPath("//*[@id=\"list-item-with-action\"]/div[1]/div/div/div/a"));
            _createBtn.Click();
            WaitForLoad();
        }
    }
}
