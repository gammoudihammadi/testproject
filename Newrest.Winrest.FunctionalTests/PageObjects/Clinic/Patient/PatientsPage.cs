using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.PageObjects;
using System;
using System.Globalization;
using System.Threading;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Clinic.Patient
{
    public class PatientsPage : PageBase
    {
        public PatientsPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        //______________</div>_____________________ Constantes _____________________________________________

        //Filtres
        private const string CL_PATIENT_RESET_FILTER_BUTTON = "//a[contains(text(), 'Reset Filter')]";
        private const string SEARCH_FILTER = "Filters_SearchText";
        private const string PATIENT_DATE_FROM = "datepicker-dateFrom";
        private const string PATIENT_DATE_TO = "datepicker-dateTo";
        private const string PATIENT_DIET_MONITORING_ALL = "//input[contains(@value, 'All')]";
        private const string PATIENT_DIET_MONITORING_NOT_CONCERNED = "//input[contains(@value, 'NotConcerned')]";
        private const string PATIENT_DIET_MONITORING_DONE = "//input[contains(@value, 'Checked')]";
        private const string PATIENT_DIET_MONITORING_DONE_DEV = "/html/body/div[2]/div/div[1]/div[1]/div/form/div[6]/div/div[2]/label/input";
        private const string PATIENT_DIET_MONITORING_NOT_DONE = "//input[contains(@value, 'NotChecked')]";
        private const string PATIENT_DIET_MONITORING_NOT_DONE_DEV = "/html/body/div[2]/div/div[1]/div[1]/div/form/div[6]/div/div[3]/label/input";
        private const string DIET_MONITORING_FILTER = "//*[@id=\"patient-filter-form\"]/div[6]/a";
        // General
        private const string PLUS_BUTTON = "//div[contains(@class,'dropdown dropdown-add-button')]/button[contains(@class,'dropbtn')]";
        private const string NEW_PATIENT = "New patient";
        private const string FIRST_PATIENT_FIRST_NAME = "//*[@id=\"patient-table\"]/tbody/tr[2]/td[2]";
        private const string FIRST_PATIENT_LAST_NAME = "//*[@id=\"patient-table\"]/tbody/tr[2]/td[3]";
        private const string FIRST_PATIENT_IPP = "//*[@id=\"patient-table\"]/tbody/tr[2]/td[4]";
        private const string FIRST_PATIENT_VISIT_NB = "//*[@id=\"patient-table\"]/tbody/tr[2]/td[5]";
        private const string FIRST_PATIENT_FLOOR = "//*[@id=\"patient-table\"]/tbody/tr[2]/td[6]";
        private const string FIRST_PATIENT_ROOM_BED = "//*[@id=\"patient-table\"]/tbody/tr[2]/td[7]";
        private const string FIRST_PATIENT_START_DATE = "//*[@id=\"patient-table\"]/tbody/tr[2]/td[8]";
        private const string FIRST_PATIENT_END_DATE = "//*[@id=\"patient-table\"]/tbody/tr[2]/td[9]";
        private const string FIRST_PATIENT_DIET = "//*[@id=\"patient-table\"]/tbody/tr[2]/td[10]/span";
        private const string PATIENT_DIET_NOT_CONCERNED = "//*[@id=\"patient-table\"]/tbody/tr[{0}]/td[10]/span[contains(@class, '{1}')]";
        private const string PATIENT_DELETE_BUTTON = "//tr[contains(td[2], '{0}')]//span[contains(@class, 'trash')]";
        private const string PATIENT_DELETE_CONFIRM_BUTTON = "first";
        //_____________________________________ Variables _____________________________________________

        //Filtres
        [FindsBy(How = How.XPath, Using = CL_PATIENT_RESET_FILTER_BUTTON)]
        private IWebElement _resetFilterButton;

        [FindsBy(How = How.Id, Using = SEARCH_FILTER)]
        private IWebElement _searchFilter;

        [FindsBy(How = How.Id, Using = PATIENT_DATE_FROM)]
        private IWebElement _dateFrom;

        [FindsBy(How = How.Id, Using = PATIENT_DATE_TO)]
        private IWebElement _dateTo;

        [FindsBy(How = How.XPath, Using = PATIENT_DIET_MONITORING_ALL)]
        private IWebElement _dietAll;

        [FindsBy(How = How.XPath, Using = PATIENT_DIET_MONITORING_NOT_CONCERNED)]
        private IWebElement _dietNotConcerned;

        [FindsBy(How = How.XPath, Using = PATIENT_DIET_MONITORING_DONE)]
        private IWebElement _dietDone;

        [FindsBy(How = How.XPath, Using = PATIENT_DIET_MONITORING_NOT_DONE)]
        private IWebElement _dietNotDone;

        // General
        [FindsBy(How = How.Id, Using = PLUS_BUTTON)]
        private IWebElement _plusButton;

        [FindsBy(How = How.XPath, Using = NEW_PATIENT)]
        private IWebElement _createNewPatient;

        [FindsBy(How = How.XPath, Using = FIRST_PATIENT_FIRST_NAME)]
        private IWebElement _firstPatientFirstName;

        [FindsBy(How = How.XPath, Using = FIRST_PATIENT_LAST_NAME)]
        private IWebElement _firstPatientLastName;

        [FindsBy(How = How.XPath, Using = FIRST_PATIENT_IPP)]
        private IWebElement _firstPatientIpp;

        [FindsBy(How = How.XPath, Using = FIRST_PATIENT_VISIT_NB)]
        private IWebElement _firstPatientVisitNb;

        [FindsBy(How = How.XPath, Using = FIRST_PATIENT_FLOOR)]
        private IWebElement _firstPatientFloor;

        [FindsBy(How = How.XPath, Using = FIRST_PATIENT_ROOM_BED)]
        private IWebElement _firstPatientRoomBed;

        [FindsBy(How = How.XPath, Using = FIRST_PATIENT_START_DATE)]
        private IWebElement _firstPatientStartDate;

        [FindsBy(How = How.XPath, Using = FIRST_PATIENT_END_DATE)]
        private IWebElement _firstPatientEndDate;

        [FindsBy(How = How.XPath, Using = FIRST_PATIENT_DIET)]
        private IWebElement _firstPatientDiet;

        [FindsBy(How = How.XPath, Using = PATIENT_DELETE_BUTTON)]
        private IWebElement _patientDeleteButton;

        [FindsBy(How = How.XPath, Using = PATIENT_DELETE_CONFIRM_BUTTON)]
        private IWebElement _patientDeleteConfirmButton;

        //_______________________________________ Methodes ____________________________________________

        // General

        public enum FilterType
        {
            Search,
            DateFrom,
            DateTo,
            ShowAll,
            DietMonitoringNotConcerned,
            DietMonitoringDone,
            DietMonitoringNotDone
        }

        public void ResetFilters()
        {
            _resetFilterButton = WaitForElementIsVisible(By.XPath(CL_PATIENT_RESET_FILTER_BUTTON), nameof(CL_PATIENT_RESET_FILTER_BUTTON));
            var javaScriptExecutor = _webDriver as IJavaScriptExecutor;
            javaScriptExecutor.ExecuteScript("arguments[0].click();", _resetFilterButton);
            WaitForLoad();

            // batch passage à minuit-2h UTC Renaud
            if (DateUtils.IsBeforeMidnight())
            {
                Filter(FilterType.DateTo, DateUtils.Now);
            }
        }

        public void Filter(FilterType filterType, object value)
        {
            switch (filterType)
            {
                case FilterType.Search:
                    _searchFilter = WaitForElementIsVisible(By.Id(SEARCH_FILTER));
                    _searchFilter.SetValue(ControlType.TextBox, value);
                    WaitLoading();
                    break;
                case FilterType.DateFrom:
                    _dateFrom = WaitForElementIsVisible(By.Id(PATIENT_DATE_FROM));
                    _dateFrom.SetValue(ControlType.DateTime, value);
                    _dateFrom.SendKeys(Keys.Tab);
                    WaitLoading();
                    break;
                case FilterType.DateTo:
                    _dateTo = WaitForElementIsVisible(By.Id(PATIENT_DATE_TO));
                    _dateTo.SetValue(ControlType.DateTime, value);
                    _dateTo.SendKeys(Keys.Tab);
                    WaitLoading();
                    break;
                case FilterType.ShowAll:
                    _dietAll = WaitForElementExists(By.XPath(PATIENT_DIET_MONITORING_ALL));
                    _dietAll.SetValue(ControlType.RadioButton, value);
                    WaitLoading();
                    break;
                case FilterType.DietMonitoringNotConcerned:
                    _dietNotConcerned = WaitForElementExists(By.XPath(PATIENT_DIET_MONITORING_NOT_CONCERNED));
                    _dietNotConcerned.SetValue(ControlType.RadioButton, value);
                    WaitLoading();
                    break;
                case FilterType.DietMonitoringDone:
                    _dietDone = WaitForElementExists(By.XPath(PATIENT_DIET_MONITORING_DONE_DEV));
                    _dietDone.SetValue(ControlType.RadioButton, value);
                    WaitLoading();
                    break;
                case FilterType.DietMonitoringNotDone:
                    _dietNotDone = WaitForElementExists(By.XPath(PATIENT_DIET_MONITORING_NOT_DONE_DEV));
                    _dietNotDone.SetValue(ControlType.RadioButton, value);
                    WaitLoading();
                    break;
                default:
                    break;
            }

            WaitPageLoading();
        }
        public void ClickDietMonitoring()
        {
            var openDietMonitoring = WaitForElementIsVisible(By.XPath(DIET_MONITORING_FILTER));
            openDietMonitoring.Click();
            Thread.Sleep(1000);
        }
        public PatientCreateModalPage PatientCreatePage()
        {
            ShowPlusMenu();

            _createNewPatient = WaitForElementIsVisible(By.LinkText(NEW_PATIENT));
            _createNewPatient.Click();
            WaitForLoad();

            return new PatientCreateModalPage(_webDriver, _testContext);
        }

        public override void ShowPlusMenu()
        {
            var actions = new Actions(_webDriver);

            _plusButton = WaitForElementIsVisible(By.XPath(PLUS_BUTTON));
            actions.MoveToElement(_plusButton).Perform();
            WaitForLoad();

        }

        public string GetFirstPatientFirstName()
        {
            _firstPatientFirstName = WaitForElementExists(By.XPath(FIRST_PATIENT_FIRST_NAME));
            return _firstPatientFirstName.Text;
        }

        public string GetFirstPatientLastName()
        {
            _firstPatientLastName = WaitForElementIsVisible(By.XPath(FIRST_PATIENT_LAST_NAME), nameof(FIRST_PATIENT_LAST_NAME));
            return _firstPatientLastName.Text;
        }
        public string GetFirstPatientIpp()
        {
            _firstPatientIpp = WaitForElementIsVisible(By.XPath(FIRST_PATIENT_IPP), nameof(FIRST_PATIENT_IPP));
            return _firstPatientIpp.Text;
        }
        public string GetFirstPatientVisitNumber()
        {
            _firstPatientVisitNb = WaitForElementIsVisible(By.XPath(FIRST_PATIENT_VISIT_NB), nameof(FIRST_PATIENT_VISIT_NB));
            return _firstPatientVisitNb.Text;
        }

        public string GetFirstPatientFloor()
        {
            _firstPatientFloor = WaitForElementIsVisible(By.XPath(FIRST_PATIENT_FLOOR), nameof(FIRST_PATIENT_FLOOR));
            return _firstPatientFloor.Text;
        }
        public string GetFirstPatientRoomBed()
        {
            _firstPatientRoomBed = WaitForElementIsVisible(By.XPath(FIRST_PATIENT_ROOM_BED), nameof(FIRST_PATIENT_ROOM_BED));
            return _firstPatientRoomBed.Text;
        }
        public string GetFirstPatientStartDate()
        {
            _firstPatientStartDate = WaitForElementIsVisible(By.XPath(FIRST_PATIENT_START_DATE), nameof(FIRST_PATIENT_START_DATE));
            return _firstPatientStartDate.Text;
        }
        public string GetFirstPatientEndDate()
        {
            _firstPatientEndDate = WaitForElementIsVisible(By.XPath(FIRST_PATIENT_END_DATE), nameof(FIRST_PATIENT_END_DATE));
            return _firstPatientEndDate.Text;
        }

        public string GetFirstPatientDietMonitoring()
        {
            _firstPatientDiet = WaitForElementIsVisible(By.XPath(FIRST_PATIENT_DIET), nameof(FIRST_PATIENT_DIET));
            return _firstPatientDiet.GetAttribute("class");
        }
        public void DeletePatient(string patientName)
        {
            _patientDeleteButton = WaitForElementIsVisible(By.XPath(string.Format(PATIENT_DELETE_BUTTON, patientName)), nameof(PATIENT_DELETE_BUTTON));
            _patientDeleteButton.Click();
            WaitForLoad();

            _patientDeleteConfirmButton = WaitForElementIsVisible(By.Id(PATIENT_DELETE_CONFIRM_BUTTON), nameof(PATIENT_DELETE_CONFIRM_BUTTON));
            _patientDeleteConfirmButton.Click();
            WaitForLoad();
        }

        public bool IsPatientDisplayed()
        {
            if (isElementVisible(By.XPath(FIRST_PATIENT_FIRST_NAME)))
            {
                _firstPatientFirstName = _webDriver.FindElement((By.XPath(FIRST_PATIENT_FIRST_NAME)));
                return _firstPatientFirstName.Displayed;
            }
            else
            {
                return false;
            }
        }

        public Boolean IsDateRespected(DateTime fromDate, DateTime toDate)
        {
            // Take the date format from the datepicker element and use it to format the date column to avoid date errors
            _dateFrom = WaitForElementIsVisible(By.Id(PATIENT_DATE_FROM));
            var dateFormat = _dateFrom.GetAttribute("data-date-format");
            CultureInfo ci = dateFormat.Equals("dd/mm/yyyy") ? new CultureInfo("fr-FR") : new CultureInfo("en-US");

            int tot;

            if (CheckTotalNumber() > 100)
            {
                tot = 100;
            }
            else
            {
                tot = CheckTotalNumber();
            }
            if (tot == 0)
                return false;

            var patientElements = _webDriver.FindElements(By.XPath("//*[@id=\"patient-table\"]/tbody/tr[*]"));
            foreach (var elm in patientElements)
            {
                if (!elm.Text.Contains("First name"))
                {
                    try
                    {
                        string startDateText = elm.Text.Substring(elm.Text.Length - 10);
                        string endDateText = elm.Text.Substring(elm.Text.Length - 21, 10);

                        DateTime startDate = DateTime.Parse(startDateText, ci).Date;
                        DateTime endDate = DateTime.Parse(endDateText, ci).Date;

                        if (startDate.Date.CompareTo(fromDate.Date) < 0 && endDate.Date.CompareTo(fromDate.Date) < 0 ||
                            endDate.Date.CompareTo(toDate.Date) > 0 && startDate.Date.CompareTo(toDate.Date) > 0)
                        {
                            return false;
                        }
                    }
                    catch
                    {
                        return false;
                    }
                }
            }
            return true;
        }

        public bool CheckDietMonitoring(string type)
        {
            bool isValidated = true;
            int tot = CheckTotalNumber() > 100 ? 100 : CheckTotalNumber();

            if (tot == 0)
                return false;

            for (int i = 1; i < tot; i++)
            {
                if (isElementVisible(By.XPath(string.Format(PATIENT_DIET_NOT_CONCERNED, i + 1, type))))
                {
                    _webDriver.FindElement(By.XPath(string.Format(PATIENT_DIET_NOT_CONCERNED, i + 1, type)));
                }
                else
                {
                    isValidated = false;
                }
            }

            return isValidated;
        }
    }
}
