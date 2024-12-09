using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.PageObjects;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Flights.Crew
{
    public class CrewPage : PageBase
    {
        public CrewPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        // __________________________________________ Constantes _________________________________________

        public const string EXTEND_ADD_MENU = "/html/body/div[2]/div/div[2]/div[1]/div/div[2]/button";
        public const string CREATE_NEW_CREW_MENU_SELECT = "/html/body/div[2]/div/div[2]/div[1]/div/div[2]/div/a";
        public const string SEARCH_NUMBER = "SearchNumber";
        public const string SHOW_ALL = "show-only-all";
        public const string SHOW_ONLY_ACTIVE = "show-only-active";
        public const string SHOW_ONLY_INACTIVE = "show-only-inactive";
        public const string FIRST_CREW_DISPLAY_NAME = "/html/body/div[2]/div/div[2]/div[2]/div/table/tbody/tr[1]/td[2]";
        public const string FIRST_CREW_EMP_NUMBER = "/html/body/div[2]/div/div[2]/div[2]/div/table/tbody/tr[1]/td[1]";
        public const string FIRST_CREW_AIRLINE = "/html/body/div[2]/div/div[2]/div[2]/div/table/tbody/tr[2]/td[3]";
        public const string NUMBER_CREWS_HEADER = "/html/body/div[2]/div/div[2]/div[1]/h1/span";
        public const string DELETE_BUTTON = "/html/body/div[2]/div/div[2]/div[2]/div/table/tbody/tr/td[4]/a";
        public const string CONFIRM_DELETE_BUTTON = "/html/body/div[10]/div/div/div[3]/a[1]";
        public const string RESET_FILTER = "ResetFilter";
        public const string LIST_CREW_GRID = "/html/body/div[2]/div/div[2]/div[2]/div/table/tbody/tr[*]/td[1]";
        public const string EXTEND_MENU_EXPORT = "/html/body/div[2]/div/div[2]/div[1]/div/div[1]/button";
        public const string EXPORT_BUTTON = "/html/body/div[2]/div/div[2]/div[1]/div/div[1]/div/a[2]";
        public const string IMPORT_BUTTON = "/html/body/div[2]/div/div[2]/div[1]/div/div[1]/div/a[1]";
        public const string INPUT_CHOOSE_FILE = "/html/body/div[3]/div/div/div[2]/div/form/div[1]/div[1]/div/input";
        public const string CHECK_FILE_BUTTON = "/html/body/div[3]/div/div/div[2]/div/form/div[2]/button[2]";
        public const string CLOSE_IMPORT_MODAL = "/html/body/div[3]/div/div/div[2]/div[3]/button";

        //Create new crew member modal/Edit crew General information

        public const string AIRLINE = "crew_CustomerId";
        public const string DISPLAY_NAME = "crew_DisplayName";
        public const string FIRST_NAME = "crew_FirstName";
        public const string EMP_NUMBER = "crew_EmployeeNumber";
        public const string LOGON_NAME = "crew_LogonName";
        public const string USER_ID = "crew_UserId";
        public const string PASSWORD = "crew_Password";
        public const string SAVE_BUTTON = "//button[text()='Save']";

        //Crew general information

        public const string AIRLINE_GENERAL_INFORMATION_SELECTED = "//*[@id=\"crew_CustomerId\"]/option[@selected='selected']";
        public const string ENABLED_YES = "//label[text()='Enabled']/..//span[contains(@class,'bootstrap-switch-handle-on')]";
        public const string ENABLED_NO = "//label[text()='Enabled']/..//span[contains(@class,'bootstrap-switch-handle-off')]";
        public const string BACK_TO_LIST = "/html/body/div[2]/a/span[1]";

        //_____________________________________________ Variables _____________________________________

        [FindsBy(How = How.Id, Using = AIRLINE)]
        private IWebElement _airline;

        [FindsBy(How = How.XPath, Using = DISPLAY_NAME)]
        private IWebElement _displayname;

        [FindsBy(How = How.XPath, Using = FIRST_NAME)]
        private IWebElement _firstName;

        [FindsBy(How = How.XPath, Using = EMP_NUMBER)]
        private IWebElement _empNumber;

        [FindsBy(How = How.XPath, Using = LOGON_NAME)]
        private IWebElement _logonName;

        [FindsBy(How = How.XPath, Using = USER_ID)]
        private IWebElement _userid;

        [FindsBy(How = How.XPath, Using = PASSWORD)]
        private IWebElement _password;

        [FindsBy(How = How.Id, Using = SEARCH_NUMBER)]
        private IWebElement _search;

        [FindsBy(How = How.Id, Using = SHOW_ALL)]
        private IWebElement _showAll;

        [FindsBy(How = How.Id, Using = SHOW_ONLY_ACTIVE)]
        private IWebElement _showOnlyActive;

        [FindsBy(How = How.Id, Using = SHOW_ONLY_INACTIVE)]
        private IWebElement _showOnlyInactive;

        [FindsBy(How = How.Id, Using = RESET_FILTER)]
        private IWebElement _resetFilter;

        // ____________________________________________ Méthodes ___________________________________________

        public enum FilterType
        {
            Search,
            ShowAll,
            ShowOnlyActive,
            ShowOnlyInactive
        }

        public void Filter(FilterType filterType, object value)

        {
            switch (filterType)
            {
                case FilterType.Search:

                    _search = WaitForElementIsVisible(By.Id(SEARCH_NUMBER));
                    _search.SetValue(ControlType.TextBox, value);
                    break;

                case FilterType.ShowAll:
                    _showAll = WaitForElementIsVisible(By.Id(SHOW_ALL));
                    _showAll.SetValue(ControlType.RadioButton, value);
                    break;

                case FilterType.ShowOnlyActive:
                    _showOnlyActive = WaitForElementIsVisible(By.Id(SHOW_ONLY_ACTIVE));
                    _showOnlyActive.SetValue(ControlType.RadioButton, value);
                    break;

                case FilterType.ShowOnlyInactive:
                    _showOnlyInactive = WaitForElementIsVisible(By.Id(SHOW_ONLY_INACTIVE));
                    _showOnlyInactive.SetValue(ControlType.RadioButton, value);
                    break;

                default:
                    break;

            }
            WaitPageLoading();
            WaitForLoad();
        }

        public object GetFilterValue(FilterType filterType)
        {
            switch (filterType)
            {
                case FilterType.Search:
                    _search = WaitForElementIsVisible(By.Id(SEARCH_NUMBER));
                    return _search.GetAttribute("value");

                case FilterType.ShowAll:
                     _showAll = WaitForElementIsVisible(By.Id(SHOW_ALL));
                    return _showAll.Selected;

                case FilterType.ShowOnlyActive:
                     _showOnlyActive = WaitForElementIsVisible(By.Id(SHOW_ONLY_ACTIVE));
                    return _showOnlyActive.Selected;

                case FilterType.ShowOnlyInactive:
                     _showOnlyInactive = WaitForElementIsVisible(By.Id(SHOW_ONLY_INACTIVE));
                    return _showOnlyInactive.Selected;
            }
            return null;
        }

        public void ResetFilters()
        {
            _resetFilter = WaitForElementIsVisible(By.Id(RESET_FILTER));
            _resetFilter.Click();

            WaitForLoad();
            // batch passage à minuit-2h UTC Renaud
            if (DateUtils.IsBeforeMidnight())
            {
                //pas de date
            }
        }

        public string GetNumberCrewsInHeader()
        {
            var numberCrews = WaitForElementIsVisible(By.XPath(NUMBER_CREWS_HEADER));
            return numberCrews.Text;
        }

        public void CreateNewCrewMember(string airlineInput, string displayNameInput, string firstNameInput, string empNumberInput, string logonNameIput,
            string userIdInput, string passwordInput)
        {
            //open modal Create new crew member
            var extendMenu = WaitForElementIsVisible(By.XPath(EXTEND_ADD_MENU));
            Actions actions = new Actions(_webDriver);
            actions.MoveToElement(extendMenu).Perform();
            var createNewCrewMenuSelect = WaitForElementIsVisible(By.XPath(CREATE_NEW_CREW_MENU_SELECT));
            createNewCrewMenuSelect.Click();
            WaitForLoad();
            //create crew (remplir les champs obligatoires)
            _airline = WaitForElementIsVisible(By.Id(AIRLINE));
            _airline.SetValue(ControlType.DropDownList, airlineInput);

            _displayname = WaitForElementIsVisible(By.Id(DISPLAY_NAME));
            _displayname.SetValue(ControlType.TextBox, displayNameInput);

            
            _firstName = WaitForElementIsVisible(By.Id(FIRST_NAME));
            _firstName.SetValue(ControlType.TextBox, firstNameInput);

            _empNumber = WaitForElementIsVisible(By.Id(EMP_NUMBER));
            _empNumber.SetValue(ControlType.TextBox, empNumberInput);

            _logonName = WaitForElementIsVisible(By.Id(LOGON_NAME));
            _logonName.SetValue(ControlType.TextBox, logonNameIput);

            _userid = WaitForElementIsVisible(By.Id(USER_ID));
            _userid.SetValue(ControlType.TextBox, userIdInput);

            _password = WaitForElementIsVisible(By.Id(PASSWORD));
            _password.SetValue(ControlType.TextBox, passwordInput);

            var saveBtn = WaitForElementIsVisible(By.XPath(SAVE_BUTTON));
            saveBtn.Click();
            WaitPageLoading();
        }
        public void DeleteCrew()
        {
            WaitForLoad();
            var deleteBtn = WaitForElementIsVisible(By.XPath(DELETE_BUTTON));
            deleteBtn.Click();
            WaitForLoad();
            var confirmDelete = WaitForElementIsVisible(By.XPath(CONFIRM_DELETE_BUTTON));
            confirmDelete.Click();
            WaitForLoad();
        }
        public bool VerifyShowOnlyActiveFilter()
        {
            var listCrewGrid = _webDriver.FindElements(By.XPath(LIST_CREW_GRID));
            for(var i = 0; i< listCrewGrid.Count(); i++)
            {
                var crews = _webDriver.FindElements(By.XPath(String.Format(LIST_CREW_GRID, i)));
                crews[i].Click();
                WaitForLoad();
                var enabledBtnGreen = WaitForElementIsVisible(By.XPath(ENABLED_YES));
                if (enabledBtnGreen.Text != "Yes")
                {
                    return false;
                }
                var backToList = WaitForElementIsVisible(By.XPath(BACK_TO_LIST));
                backToList.Click();
                WaitForLoad();
            }
            return true;
        }
        public bool VerifyShowOnlyInActiveFilter()
        {
            var listCrewGrid = _webDriver.FindElements(By.XPath(LIST_CREW_GRID));
            for (var i = 0; i < listCrewGrid.Count(); i++)
            {
                var crews = _webDriver.FindElements(By.XPath(String.Format(LIST_CREW_GRID, i)));
                crews[i].Click();
                WaitForLoad();
                var enabledBtnGray = WaitForElementIsVisible(By.XPath(ENABLED_NO));
                if (enabledBtnGray.Text != "No")
                {
                    return false;
                }
                var backToList = WaitForElementIsVisible(By.XPath(BACK_TO_LIST));
                backToList.Click();
                WaitForLoad();
            }
            return true;
        }
        public void BackToList()
        {
            var backToList = WaitForElementIsVisible(By.XPath(BACK_TO_LIST));
            backToList.Click();
            WaitForLoad();
        }
        public void SetAirlineCrew(string airlineInp)
        {
            _airline = WaitForElementIsVisible(By.Id(AIRLINE));
            _airline.SetValue(ControlType.DropDownList, airlineInp);
            WaitForLoad();
        }
        public void SelectFirstCrew()
        {
            var firstCrewNumber = WaitForElementIsVisible(By.XPath(FIRST_CREW_EMP_NUMBER));
            firstCrewNumber.Click();
            WaitForLoad();
        }
        public void EditDetailCrew(string airlineInp, string logonNameInp)
        {
            SetAirlineCrew(airlineInp);

            _logonName = WaitForElementIsVisible(By.Id(LOGON_NAME));
            _logonName.SetValue(ControlType.TextBox, logonNameInp);
            
            WaitForLoad();
            WaitPageLoading();
        }
        public bool VerifyModifyDetail(string airlineInp, string logonNameInp)
        {
            var airline = WaitForElementIsVisible(By.XPath(AIRLINE_GENERAL_INFORMATION_SELECTED));
            var airlineResult = airline.GetAttribute("text");

            _logonName = WaitForElementIsVisible(By.Id(LOGON_NAME));
            var logonNameResult = _logonName.GetAttribute("value");

            if(airlineResult != airlineInp || logonNameResult != logonNameInp)
            {
                return false;
            }
            return true;
        }
        public IEnumerable<string> GetListNumbersCrew()
        {
            var listNumbers = _webDriver.FindElements(By.XPath(LIST_CREW_GRID));
            return listNumbers.Select(li => li.Text);
        }
        public void Export()
        {
            var extendMenu = WaitForElementIsVisible(By.XPath(EXTEND_MENU_EXPORT));
            Actions actions = new Actions(_webDriver);
            actions.MoveToElement(extendMenu).Perform();
            var exportBtn = WaitForElementIsVisible(By.XPath(EXPORT_BUTTON));
            exportBtn.Click();

            IsFileLoaded(By.CssSelector("[target='_blank'][class='far fa-file-excel']"));
            ClickPrintButton();
            WaitForDownload();
            Close();
            WaitForLoad();
        }

        public FileInfo GetExportExcelFile(FileInfo[] taskFiles)
        {
            List<FileInfo> correctDownloadFiles = new List<FileInfo>();

            StringBuilder sb = new StringBuilder();

            foreach (var file in taskFiles)
            {
                sb.Append(file.Name + " ");
                //  Test REGEX
                if (IsExcelFileCorrect(file.Name))
                {
                    correctDownloadFiles.Add(file);
                }
            }

            if (correctDownloadFiles.Count == 0)
            {
                return null;
            }
            var time = correctDownloadFiles[0].LastWriteTimeUtc;
            var correctFile = correctDownloadFiles[0];

            correctDownloadFiles.ForEach(file =>
            {
                if (time < file.LastWriteTimeUtc)
                {
                    time = file.LastWriteTimeUtc;
                    correctFile = file;
                }
            });

            return correctFile;
        }

        public bool IsExcelFileCorrect(string filePath)
        {
            Regex r = new Regex("[export?flights\\s\\d.-]", RegexOptions.IgnoreCase | RegexOptions.Singleline);
            Match m = r.Match(filePath);
            return m.Success;
        }
        public string[] GetNumberInString(string input)
        {
            string[] numbers = Regex.Matches(input, @"\d+").Cast<Match>().Select(m => m.Value).ToArray();
            return numbers;
        }

        public bool VerifyExcel(IEnumerable<string> numbersExcel, List<string> numberExcel)
        {
            for (int i = 0; i < numbersExcel.Count(); i++)
            {
                var number = GetNumberInString(numbersExcel.ElementAt(i)).First();
                if (number != null)
                {
                    if (number != numberExcel[i])
                        return false;
                }
            }
            return true;
        }
        public string GenerateString()
        {
            string letters = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ";
            string numbers = "0123456789";
            var random = new Random();
            char[] charArray = new char[8];

            for (int i = 0; i < 6; i++)
            {
                int index = random.Next(letters.Length);
                charArray[i] = letters[index];
            }

            for (int i = 6; i < 8; i++)
            {
                int index = random.Next(numbers.Length);
                charArray[i] = numbers[index];
            }

            return new string(charArray);
        }
        public void Import(string filePath)
        {
            var extendMenu = WaitForElementIsVisible(By.XPath(EXTEND_MENU_EXPORT));
            Actions actions = new Actions(_webDriver);
            actions.MoveToElement(extendMenu).Perform();
            var importBtn = WaitForElementIsVisible(By.XPath(IMPORT_BUTTON));
            importBtn.Click();
            WaitForLoad(); 
            var inputChooseFile = WaitForElementIsVisible(By.XPath(INPUT_CHOOSE_FILE));
            inputChooseFile.SendKeys(filePath);
            WaitForLoad();
            var checkFileButton = WaitForElementIsVisible(By.XPath(CHECK_FILE_BUTTON));
            checkFileButton.Click();
            WaitForLoad();
            var closeImportBtn = WaitForElementIsVisible(By.XPath(CLOSE_IMPORT_MODAL));
            closeImportBtn.Click();
            WaitForLoad();
        }
        public bool IsAddedFileVerif(int numberRawsListCrews)
        {
            var numberCrewsInHeader = GetNumberCrewsInHeader();
            if (int.Parse(numberCrewsInHeader) != numberRawsListCrews + 1)
            {
                return false;
            }
            return true;
        }

    }
}
