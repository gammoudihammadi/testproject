using DocumentFormat.OpenXml.Bibliography;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Customers.Service;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;

namespace Newrest.Winrest.FunctionalTests.PageObjects.CustomerPortal
{
    public class CustomerPortalQtyByWeekPage : PageBase
    {

        public CustomerPortalQtyByWeekPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        // _________________________________________ Constantes _______________________________________________



        private const string QTY_MANAGEMENT_SEARCH = "//*[@id=\"SearchPattern\"]";
        private const string QTY_MANAGEMENT_SEARCH_RESULT_EMPTY = "//*[@id=\"tabContentItemContainer\"]/div[2]/div[2]/div/p";
        private const string QTY_MANAGEMENT_SEARCH_FIRST_RESULT = "//*[@id=\"tabContentItemContainer\"]/div[2]/div[2]/div/table/tbody/tr[2]/td[1]";

        private const string QTY_MANAGEMENT_FILTER_DELIVERY_COUNT = "//*[@id=\"collapseFlightDeliveriesFilter\"]/div";
        private const string QTY_MANAGEMENT_FILTER_DELIVERY_LABEL = "//*[@id=\"collapseFlightDeliveriesFilter\"]/div[{0}]/label";

        private const string QTY_MANAGEMENT_CHECKBOX = "//*[@id=\"FlightDeliveries_{0}__IsSelected\"]";

        private const string PREVIOUS_WEEK = "//*[@id=\"selectDate\"]/a[1]";
        private const string WEEK = "//*[@id=\"selectDate\"]/span[1]";
        private const string NEXT_WEEK = "//*[@id=\"selectDate\"]/a[2]";
        private const string QTY_MANAGEMENT_CALENDAR = "//*[@id=\"selectDate\"]/span[2]";

        private const string DAYS = "/html/body/div[3]/div/div[2]/div/div[2]/div[1]/table/tbody/tr[2]/td[*]/div/p/input";
        private const string VALIDATE_DAYS = "//*[starts-with(@id,\"btn-validate-\")]";
        private const string DAY = "//*/input[contains(@id,'{0}')]";
        private const string VALIDATE = "btn-validate-CP-{0}";

        private const string DUPLICATE_TITLE = "//*/span[text()='{0}']";
        private const string DUPLICATE_IMG = "//*[@id='th{0}']/span[2]/img";
        private const string DUPLICATE_WITH_COMMENT = "radioDuplicateYes";
        private const string DUPLICATE_NEXTDAY = "//*/div[contains(@class,'original')]/following-sibling::div[@class != 'clickableDay previousOrNextMonth']";
        private const string DUPLICATE = "btn-duplicate";

        private const string CHECK_QUANTITY = "//td[contains(@class, '{0}')]/div/p";
        private const string CHECK_QUANTITY_EDIT = "//td[contains(@class, 'day')]/div/p/input";
        private const string CHECK_COMMENT = "//td[contains(@class, '{0}')]/div/a";
        private const string COMMENT_ICON = "//*/input[contains(@id,'{0}')]/parent::p/parent::div/a";
        private const string COMMENT_TEXT = "textarea-comment";
        private const string COMMENT_SAVE = "//*/button[text()='Save']";
        
        private const string SEND_MAIL_ICON = "//*[@id=\"btnSendMail\"]/span";
        private const string RECEIVER = "msg_subject";
        private const string MAIL_SUBJECT = "Subject";
        private const string MAIL_ATTACHMENT_TITLE = "btnViewFile";
        private const string SEND_BTN = "btn-init-async-send-mail";
        
        private const string ALL_DATE_FROM_TO = "/html/body/div[3]/div/div[2]/div[1]/div[2]/div[3]/span[2]";
        private const string PREVIOUS_WEEK_BTN = "/html/body/div[3]/div/div[2]/div[1]/div[2]/div[3]/a[1]/span";
        private const string NEXT_WEEK_BTN = "/html/body/div[3]/div/div[2]/div[1]/div[2]/div[3]/a[2]/span";
        private const string CHECK_INPUT_BY_DAY_NAME = "//*/input[contains(@id,'{0}')]";
        private const string QTY_MANAGEMENT_FILTER_DELIVERY_LABEL_DEV = "/html/body/div[4]/ul/li[{0}]/label/span";
        private const string CAPTION_WEEK_CONTENT = "/html/body/div[3]/div[1]/div[2]/div/div[1]/div[1]/div/span[2]";
        private const string ROWS_XPAth = "//*/tr[string(@data-service-id)]";
        private const string QUANTITY_TO_PRODUCE = "/html/body/div[3]/div/div[2]/div[2]/div[2]/div/table/tbody/tr[4]/td[2]/div/p/input";
        private const string QUANTITY_BY_DAY = "//*[@id=\"tabContentItemContainer\"]/div[2]/div[2]/div/table/tbody/tr[2]/td[1][contains(text(),'{0}')]/..//input";

        // _________________________________________ Variables _________________________________________________

        [FindsBy(How = How.XPath, Using = PREVIOUS_WEEK)]
        private IWebElement _previousWeek;

        [FindsBy(How = How.XPath, Using = WEEK)]
        private IWebElement _week;

        [FindsBy(How = How.XPath, Using = NEXT_WEEK)]
        private IWebElement _nextWeek;

        [FindsBy(How = How.XPath, Using = CHECK_QUANTITY)]
        private IWebElement _duplicatedQuantity;

        [FindsBy(How = How.XPath, Using = CHECK_COMMENT)]
        private IWebElement _duplicatedComment;

        [FindsBy(How = How.Id, Using = SEND_MAIL_ICON)]
        private IWebElement _sendMailIcon;

        [FindsBy(How = How.Id, Using = RECEIVER)]
        private IWebElement _receiver;

        [FindsBy(How = How.Name, Using = MAIL_SUBJECT)]
        private IWebElement _mailSubject;

        [FindsBy(How = How.Id, Using = MAIL_ATTACHMENT_TITLE)]
        private IWebElement _mailAttachment;

        [FindsBy(How = How.Id, Using = SEND_BTN)]
        private IWebElement _sendBtn;

        // _________________________________________ Méthodes __________________________________________________



        // Barre des dates
        public void ClickOnNextWeek()
        {
            _nextWeek = WaitForElementIsVisible(By.XPath(NEXT_WEEK));
            _nextWeek.Click();
            WaitForLoad();
        }

        public void ClickOnPreviousWeek()
        {
            _previousWeek = WaitForElementIsVisible(By.XPath(PREVIOUS_WEEK));
            _previousWeek.Click();
            WaitForLoad();
        }

        public int GetNumWeek()
        {
            _week = WaitForElementIsVisible(By.XPath(WEEK));
            string text = _week.Text.Substring(0, _week.Text.IndexOf(":")).Replace("Week ", "").Trim();
            return int.Parse(text);
        }

        public void GoOnCurrentWeek(int weekNumber)
        {            
            bool sameWeek = false;

            while (!sameWeek)
            {
                var week = GetNumWeek();

                if (weekNumber == week)
                    sameWeek = true;
                else if (weekNumber < week)
                {
                    ClickOnPreviousWeek();

                }
                else
                {
                    ClickOnNextWeek();
                }
            }

        }

        // Tableau quantités
        public void SetQuantities(string newQty)
        {
            Actions action = new Actions(_webDriver);
            IEnumerable<IWebElement> days = new List<IWebElement>();
            days = _webDriver.FindElements(By.XPath(DAYS));
            
            foreach(var day in days)
            {
                action.MoveToElement(day).Perform();
                day.SetValue(ControlType.TextBox, newQty);
                WaitForLoad();
            }

            // Temps d'enregistrement des résultats
            Thread.Sleep(2000);
            WaitForLoad();
        }
        public void SetQuantitytoproduce(string value , string service)
        {
            var element = WaitForElementIsVisible(By.XPath(String.Format(QUANTITY_BY_DAY,service)));
            element.Clear();
            element.SetValue(ControlType.TextBox,value);
        }
        public void SetQuantityByweek(string serviceName, DateTime date , string value)
        {
            string dayOfWeek = date.ToString("dddd", CultureInfo.InvariantCulture);
            string xpath = $@"//*[@id='tabContentItemContainer']/div/div[2]/div[2]/div/table/tbody/tr[*][@data-step='Production']/td[1][contains(text(),'{serviceName}')]/../td/div//*[contains(@id,'Qty{dayOfWeek}')]";
            var element = WaitForElementIsVisible(By.XPath(xpath));
            //element.Clear();
            element.SetValue(ControlType.TextBox, value);
        }

        public void ValidateQuantities()
        {

            Actions action = new Actions(_webDriver);
            var validateBtns = _webDriver.FindElements(By.XPath(VALIDATE_DAYS));

            foreach (var validate in validateBtns)
            {
                action.MoveToElement(validate).Perform();
                validate.Click();
                // Temps d'enregistrement des résultats
                WaitPageLoading();
            }
        }

        public string SendQtyByMail(string userMail, string mailSubject = null)
        {
            Actions a = new Actions(_webDriver);
            ShowExtendedMenu();
            _sendMailIcon = WaitForElementIsVisible(By.XPath("//*[@id=\"btnSendMail\"]"));
            a.MoveToElement(_sendMailIcon).Click().Perform();           
            WaitForLoad();
            
            _receiver = WaitForElementIsVisible(By.Id(RECEIVER));
            _receiver.SetValue(ControlType.TextBox, userMail);

            if (mailSubject != null)
            {
                _mailSubject = WaitForElementIsVisible(By.Name(MAIL_SUBJECT));
                _mailSubject.Clear();
                _mailSubject.SendKeys(mailSubject);
                WaitForLoad();
            }

            _mailAttachment = WaitForElementIsVisible(By.Id(MAIL_ATTACHMENT_TITLE));
            WaitForLoad();

            return _mailAttachment.Text;
        }

        public string GetMailSubject()
        {
            _mailSubject = WaitForElementIsVisible(By.Name(MAIL_SUBJECT));
            return _mailSubject.GetAttribute("value");
        }

        public void SendMail()
        {
            _sendBtn = WaitForElementIsVisible(By.Id(SEND_BTN));
            _sendBtn.Click();
            WaitForLoad();

            Thread.Sleep(5000);
        }

        public override void ShowExtendedMenu()
        {
            var _extendedButton = WaitForElementExists(By.XPath("//button[contains(text(), '...')]"));
            var actions = new Actions(_webDriver);
            actions.MoveToElement(_extendedButton).Click().Perform();
            WaitForLoad();
        }

        public void Export(string site)
        {
            ShowExtendedMenu();
            //1.Cliquer sur "export"(icone imprimante)
            var imprimante = WaitForElementIsVisible(By.XPath("//*[@id=\"btnExport\"]"));
            imprimante.Click();
           
            //2.Remplir les champs
            var sitesSelect = WaitForElementIsVisible(By.Id("ddlSites"));
            sitesSelect.SetValue(ControlType.DropDownList, site);
            // date from/to déjà positionné à cette semaine
            var dateFrom = WaitForElementIsVisible(By.Id("date-picker-exportstart"));
            dateFrom.SetValue(ControlType.DateTime, DateUtils.Now.AddDays(0));
            var dateTo = WaitForElementIsVisible(By.Id("date-picker-exportend"));
            dateTo.SetValue(ControlType.DateTime, DateUtils.Now.AddDays(1));
            //3.Cliquer sur "export"
            var exportButton = WaitForElementIsVisible(By.Id("print-export"));
            exportButton.Click();
            Thread.Sleep(2000);
            var exportCloseButton = WaitForElementIsVisible(By.XPath("//button[@class = 'close']"));
            exportCloseButton.Click();
            WaitPageLoading();
        }
        public FileInfo GetQMExportExcelFile(FileInfo[] taskFiles)
        {
            List<FileInfo> correctDownloadFiles = new List<FileInfo>();

            foreach (var file in taskFiles)
            {
                //  Test REGEX
                if (IsQMExcelFileCorrect(file.Name))
                {
                    correctDownloadFiles.Add(file);
                }
            }

            if (correctDownloadFiles.Count <= 0)
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

        public bool IsQMExcelFileCorrect(string filePath)
        {
            //exemple Dispatch_Export_2022-06-23 13-31-26
            //("Dispatch_Export_"+DateUtils.Now.ToString("yyyy-MM-dd")

            string mois = "(?:0[1-9]|1[0-2])";         // mois
            string space = "(\\s)";                    // Espace
            string annee = "\\d{4}";                   // annee YYYY
            string jour = "[0-3]\\d";                  // jour
            string heure = "(?:0[0-9]|1[0-9]|2[0-3])"; // heure
            string minutes = "[0-5]\\d";               // minutes
            string secondes = "[0-5]\\d";              // secondes

            Regex r = new Regex("^Dispatch_Export_" + annee + "-" + mois + "-" + jour + space + heure + "-" + minutes + "-" + secondes + ".xlsx$", RegexOptions.IgnoreCase | RegexOptions.Singleline);
            Match m = r.Match(filePath);

            return m.Success;
        }
        public void Search(string text)
        {
            var searchText = _webDriver.FindElement(By.XPath("//*[@id=\"tbSearchPattern\"]"));
            searchText.Clear();
            searchText.SendKeys(text);
            
            WaitPageLoading();
        }

        public bool IsSearchResultEmpty()
        {
            IWebElement message;
            if (isElementVisible(By.XPath(QTY_MANAGEMENT_SEARCH_RESULT_EMPTY)))
            {
                message = WaitForElementIsVisible(By.XPath(QTY_MANAGEMENT_SEARCH_RESULT_EMPTY));
            }
            else
            {
                message = WaitForElementIsVisible(By.XPath("//*[@id=\"tabContentItemContainer\"]/div/div[2]/div[2]/div/p"));
            }
            return message.Text == "No previsional quantity";
        }

        public string SearchFirstResult()
        {
            IWebElement tableauText; 
            if (isElementVisible(By.XPath(QTY_MANAGEMENT_SEARCH_FIRST_RESULT)))
            {
                tableauText = WaitForElementIsVisible(By.XPath(QTY_MANAGEMENT_SEARCH_FIRST_RESULT));
            }
            else
            {
                tableauText = WaitForElementIsVisible(By.XPath("//*[@id=\"tabContentItemContainer\"]/div/div[2]/div[2]/div/table/tbody/tr[2]/td[1]"));
            }
            return tableauText.Text;
        }

        public void SearchFilterDelivery(string mot)
        {
            if (!isElementVisible(By.XPath("//*/div[contains(@class,'collapse')][contains(@class, 'show')]")))
            {
                var deliveryFilter = WaitForElementIsVisible(By.XPath("//*[@id=\"item-filter-form\"]/div[2]/div/a"));
                deliveryFilter.Click();
            }
            Thread.Sleep(2000);
            WaitForLoad();
            
            ComboBoxSelectById(new ComboBoxOptions("SelectedFlightDeliveries_ms", mot, false));
            WaitPageLoading();
        }

        internal DateTime GetCalendarBegin()
        {
            var dateRange = By.XPath(QTY_MANAGEMENT_CALENDAR);
            string dateRangeText = _webDriver.FindElement(dateRange).Text;
            return DateTime.ParseExact(dateRangeText.Substring(0, 10), "dd/MM/yyyy", CultureInfo.InvariantCulture);
        }

        internal DateTime GetCalendarEnd()
        {
            var dateRange = By.XPath(QTY_MANAGEMENT_CALENDAR);
            string dateRangeText = _webDriver.FindElement(dateRange).Text;
            return DateTime.ParseExact(dateRangeText.Substring(14, 10), "dd/MM/yyyy", CultureInfo.InvariantCulture);
        }

        public int FilterDeliveryCount()
        {
            //var selectDelivery = WaitForElementExists(By.XPath("//*[@id=\"SelectByDeliveryNameOption_ms\"]"));
            //selectDelivery.Click();
            var xPathList = _webDriver.FindElements(By.XPath("//*[@id=\"SelectedFlightDeliveriesTrueValue\"]/option"));
            //IReadOnlyCollection<IWebElement> options = xPathList.FindElements(By.TagName("li"));
            int optionCount = xPathList.Count;
            return optionCount;        
        }
        public void clickDeliveryFilter()
        {
            if (!isElementVisible(By.XPath("//*[@id=\"collapseSelectByDeliveryNameOptionFilter\"][contains(@class, 'show')]")))
            {
                var deliveryFilter = WaitForElementIsVisible(By.XPath("//*[@id=\"item-filter-form\"]/div[2]/div/a"));
                deliveryFilter.Click();
            }    
            // animation
            Thread.Sleep(2000);
            WaitForLoad();
        }
        internal string FilterDeliveryLabel(int offset)
        {
            var xPathLabel = By.XPath(QTY_MANAGEMENT_FILTER_DELIVERY_LABEL_DEV.Replace("{0}", offset.ToString()));
            return _webDriver.FindElement(xPathLabel).GetAttribute("innerText");
        }

        public void SetQuantities(string day, string qty)
        {
            var _day = WaitForElementIsVisible(By.XPath(string.Format(DAY,day)));
            _day.Clear();
            _day.SendKeys(qty);
            WaitForLoad();
        }

        internal void SetComment(string day, string comment)
        {
            var _day = WaitForElementIsVisible(By.XPath(string.Format(DAY, day)));
            new Actions(_webDriver).MoveToElement(_day).Perform();
            var commentIcon = WaitForElementIsVisible(By.XPath(string.Format(COMMENT_ICON, day)));
            commentIcon.Click();
            var commentText = WaitForElementIsVisible(By.Id(COMMENT_TEXT));
            commentText.Clear();
            commentText.SendKeys(comment);
            var commentSave = WaitForElementIsVisible(By.XPath(COMMENT_SAVE));
            commentSave.Click();
        }

        public void Validate(int offsetDay)
        {
            // offsetDay 3 : Wednesday
            var validate = WaitForElementIsVisible(By.XPath(string.Format("/html/body/div[3]/div/div[2]/div/div[2]/div[2]/div/table/tfoot/tr/td[*]/button[@id=\"btn-validate-CP-{0}\"]", offsetDay)));
            validate.Click();
            WaitPageLoading();
        }

        public void DuplicateOneDay(string shortDay, int offsetDay)
        {
            var title = WaitForElementIsVisible(By.XPath(string.Format(DUPLICATE_TITLE,shortDay)));
            new Actions(_webDriver).MoveToElement(title).Perform();
            var duplicateImg = WaitForElementIsVisible(By.XPath(String.Format(DUPLICATE_IMG,offsetDay)));
            duplicateImg.Click();
            var withComment = WaitForElementIsVisible(By.Id(DUPLICATE_WITH_COMMENT));
            withComment.Click();
            var nextDay = WaitForElementIsVisible(By.XPath(DUPLICATE_NEXTDAY));
            nextDay.Click();
            var duplicateButton = WaitForElementIsVisible(By.Id(DUPLICATE));
            duplicateButton.Click();
            WaitPageLoading();
        }

        public string GetQuantity(string day)
        {
            _duplicatedQuantity = WaitForElementIsVisible(By.XPath(String.Format(CHECK_QUANTITY,day)));
            return _duplicatedQuantity.Text;
        }

        public string GetQuantityEditble(string serviceName, DateTime date)
        {
            string dayOfWeek = date.ToString("dddd", CultureInfo.InvariantCulture);
            string xpath = $@"//*[@id='tabContentItemContainer']/div/div[2]/div[2]/div/table/tbody/tr[*][@data-step='Production']/td[1][contains(text(),'{serviceName}')]/../td/div//*[contains(@id,'Qty{dayOfWeek}')]";
            _duplicatedQuantity = WaitForElementIsVisible(By.XPath(xpath));
            return _duplicatedQuantity.GetAttribute("value");
        }  
        public string GetQuantitytoproduce(string service)
        {
            var quantity = WaitForElementIsVisible(By.XPath(String.Format(QUANTITY_BY_DAY,service)));
            return quantity.GetAttribute("value");
        }

        public string GetComment(string day)
        {
            _duplicatedComment = WaitForElementIsVisible(By.XPath(String.Format(CHECK_COMMENT, day)));
            return _duplicatedComment.GetAttribute("title");
        }
        public void GoToDate(DateTime date)
        {
            (DateTime fromDate, DateTime toDate) = ExtractDates();
            if(date < fromDate)
            {
                GoToPreviousWeek();
            }
            if(date>toDate)
            {
                GoToNextWeek();
            }
        }
        public bool CheckIfModifiable(DateTime date)
        {
            string dayOfWeek = date.ToString("dddd", CultureInfo.InvariantCulture);
            //set first letter to uppercase
            if (!string.IsNullOrEmpty(dayOfWeek))
            {
                dayOfWeek = char.ToUpper(dayOfWeek[0]) + dayOfWeek.Substring(1);
            }

            if (isElementExists(By.XPath(string.Format(CHECK_INPUT_BY_DAY_NAME, dayOfWeek))))
            {
                var inputQuantity = WaitForElementIsVisible(By.XPath(string.Format(CHECK_INPUT_BY_DAY_NAME, dayOfWeek)));
                var disabledAttribute = inputQuantity.GetAttribute("disabled");
                if (disabledAttribute == "true")
                {
                    return false;
                }
                else
                {
                    return true;
                }
            }
            else
            {
                return isElementExists(By.XPath(string.Format(CHECK_INPUT_BY_DAY_NAME, dayOfWeek)));
            }
        }

        public bool CheckIfModifiableByweek(DateTime date, string serviceName)
        {
            string dayOfWeek = date.ToString("dddd", CultureInfo.InvariantCulture);
            //set first letter to uppercase
            if (!string.IsNullOrEmpty(dayOfWeek))
            {
                dayOfWeek = char.ToUpper(dayOfWeek[0]) + dayOfWeek.Substring(1);
            }
            string xpath = $@"//*[@id='tabContentItemContainer']/div/div[2]/div[2]/div/table/tbody/tr[*][@data-step='Production']
                       /td[1][contains(text(),'{serviceName}')]/../td/div//*[contains(@id,'Qty{dayOfWeek}')]";

            if (isElementExists(By.XPath(xpath)))
            {
                var inputQuantity = WaitForElementIsVisible(By.XPath(xpath));
                var disabledAttribute = inputQuantity.GetAttribute("disabled");
                if (disabledAttribute == "true")
                {
                    return false;
                }
                else
                {
                    return true;
                }
            }
            else
            {
                return isElementExists(By.XPath(xpath));
            }
        }

        //Utils
        public (DateTime fromDate, DateTime toDate) ExtractDates()
        {
            var inputString = WaitForElementIsVisible(By.XPath("//*[@id=\"selectDate\"]/span[2]")).Text;
            // Split the input string by the "To" keyword
            string[] dateStrings = inputString.Split(new string[] { " To " }, StringSplitOptions.None);

            // Parse the "from" and "to" date strings as DateTime objects
            DateTime fromDate = DateTime.ParseExact(dateStrings[0], "dd/MM/yyyy", CultureInfo.InvariantCulture);
            DateTime toDate = DateTime.ParseExact(dateStrings[1], "dd/MM/yyyy", CultureInfo.InvariantCulture);

            // Return a tuple containing the "from" and "to" dates
            return (fromDate, toDate);
        }

        public void GoToPreviousWeek()
        {
            var previousWeekBtn = WaitForElementIsVisible(By.XPath(PREVIOUS_WEEK_BTN));
            previousWeekBtn.Click();
        }

        public void GoToNextWeek()
        {
            var nextWeekBtn = WaitForElementIsVisible(By.XPath(NEXT_WEEK_BTN));
            nextWeekBtn.Click();
        }
        public string GetDateToUpdate()
        {
            var IdsWithInputs = _webDriver.FindElements(By.XPath("//*[contains(@id, 'editZone')]"));
            int i = 1;
            string dayPart = string.Empty;
            foreach (var row in IdsWithInputs)
            {

                string completeId = row.GetAttribute("id");

                dayPart = completeId.Split('-').Last();

                 
                i++;
                if (i == 2)
                {
                    break;
                }
               

            }

            return dayPart;
        }
        public enum Days
        {
            Mon,
            Tue,
            Wed,
            Thu,
            Fri,
            Sat,
            Sun
        }
        public DateTime GetDateOfInput(Days days , DateTime date)
        {

            switch (days)
            {
                case Days.Mon:
                    return date;
                    
                case Days.Tue:
                    return date.AddDays(1);

                case Days.Wed:
                    return date.AddDays(2);

                case Days.Thu:
                    return date.AddDays(3);

                case Days.Fri:
                    return date.AddDays(4);

                case Days.Sat:
                    return date.AddDays(5);

                case Days.Sun:
                    return date.AddDays(6);

                default:
                    return date;
            }

        }
        public DateTime GetCaptionWeeekContent()
        {
            var captionWeek = WaitForElementIsVisible(By.XPath(CAPTION_WEEK_CONTENT));
            var allDate = captionWeek.Text;

            string[] dates = allDate.Split(new string[] { " To " }, StringSplitOptions.RemoveEmptyEntries); ;

            string startDateString = dates[0].Trim();

            string format = "dd/MM/yyyy";

            CultureInfo culture = new CultureInfo("fr-FR");

            DateTime startDate = DateTime.ParseExact(startDateString, format, culture);

            return startDate;
        }
        public int GetCounter()
        {
            var lignes = _webDriver.FindElements(By.XPath(ROWS_XPAth));
            return lignes.Count();

        }

    }
}
