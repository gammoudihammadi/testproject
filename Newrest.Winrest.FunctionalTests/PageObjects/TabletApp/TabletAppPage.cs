using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.Menus;
using Newrest.Winrest.FunctionalTests.PageObjects.Purchasing;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;

namespace Newrest.Winrest.FunctionalTests.PageObjects.TabletApp
{
    public class TabletAppPage : PageBase
    {
        private const string FILL_FILENAME3 = "//*[@id=\"mat-dialog-3\"]/haccp--modal-favorite/input";
        private const string FILL_FILENAME0 = "//*[@id=\"mat-dialog-0\"]/haccp--modal-favorite/input";
        private const string CLICK_BASEBUTTON = "//*[contains(text(), '{0}')]/parent::base-button";
        private const string CLICK_BUTTON = "//*[contains(text(), '{0}')]/parent::button";
        private const string BUTTON_ADDLINE = "//*[contains(text(), '+Add line')]";
        private const string NOTES = "//*[contains(text(),'note')]";
        private const string MESSAGEBOX_SAVE = "//*[@id=\"mat-dialog-{0}\"]/confirm-dialog/div[2]/button[2]";
        private const string NEWSELECT_OPTION = "//*[contains(text(), '{0}')]";
        private const string SELECT_OPTION = "//*[@id=\"mat-select-value-3\"]/span";
        private const string MESSAGEBOX_SAVE_PRODINFLIGHT = "//*[@id=\"mat-dialog-0\"]/confirm-dialog/mat-dialog-actions/button[2]";
        private const string FILL_FILENAME_PRODINFLIGHT0 = "//*/production--modal-local-save/input";
        private const string FILL_FILENAME_PRODINFLIGHT1 = "//*/production--modal-local-save/input";
        private const string FILTER_TOGGLE = "//*[@id=\"cdk-accordion-child-{0}\"]/parent::mat-expansion-panel";
        private const string FILTER_SELECT_ALL = "//*[@id=\"cdk-accordion-child-{0}\"]/div/div[2]";
        private const string FILTER_SEARCH = "//*[@id=\"cdk-accordion-child-{0}\"]/div/div[1]/input";
        private const string FILTER_TABLEAU_0 = "//*[text()[contains(.,'{0}')]]/parent::div";
        private const string FILTER_TABLEAU = "//*[normalize-space(text())='{0}']/parent::div";

        private const string FILTER_SITE = "/html/body/app-root/mat-sidenav-container/mat-sidenav-content/main/home/div/div[1]/mat-form-field/div/div[1]/div/mat-select/div/div[1]";
        private const string SITES = "/html/body/div[2]/div[2]/div/div/div/mat-option[*]/span";
        private const string TABLET_APP_FLIGHT = "//*[contains(text(),'Flight')]/parent::div";
        private const string NAME_INGREDIENT = "//*[@id=\"cdk-overlay-0\"]/div/div[2]/datasheet-recipe/div/div/div/div/div/div/div/div[1]/p";
        private const string ALL_NAME_INGREDIENT = "/html/body/app-root/mat-sidenav-container/mat-sidenav-content/main/haccp-thawing/div/virtual-scroller/div[2]/div[*]/div[2]/div[1]";
        private const string TABLET_APP_TODOLIST = "//*[contains(text(),'To-do List')]/parent::div";
        private const string FILE_TEMPORARY = "//*[@id=\"mat-select-value-5\"]";
        private const string PRODUCT_FIELD = "mat-input-0";
        private const string COMMENTS_BTN = "/html/body/app-root/mat-sidenav-container/mat-sidenav-content/main/app-haccp-document/div/table/tbody/tr/td[7]/app-column-editor-blank-doc/div/mat-icon";
        private const string TEXT_EDIT = "mat-input-4";
        private const string CORRECTIVE_BTN = "/html/body/app-root/mat-sidenav-container/mat-sidenav-content/main/app-haccp-document/div/table/tbody/tr/td[8]/app-column-editor-blank-doc/div/mat-icon";
        private const string TEXT_CORRECTIVE = "mat-input-5";
        private const string PREPAREDBY_BTN = "/html/body/app-root/mat-sidenav-container/mat-sidenav-content/main/app-haccp-document/div/table/tbody/tr/td[9]/app-column-editor-blank-doc/div/mat-icon";
        private const string TEXT_PREPAREDBY = "mat-input-6";
        private const string QTY_FIELD = "//*[@id=\"mat-input-1\"]";
        private const string DISINFECTION_1_FIELD = "mat-input-2";
        private const string DISINFECTION_2_FIELD = "mat-input-3";

      
        public enum FilterType
        {
            Services,
            Customers,
            Favorite,
            From,
            To,
            StartTime,
            EndTime,
            GuestType,
            ItemGroups,
            RecipeType,
            ServiceCategories,
            ShowHiddenArticles,
            ShowNormalAndVacuumProduction,
            ShowNormalProductionOnly,
            ShowServicesWithoutDatasheetOnly,
            ShowVacuumProductionOnly,
            ValidatedFlightsOnly,
            Workshop
        }


        public TabletAppPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        public void WaitForLoad()
        {
            IWait<IWebDriver> wait = new OpenQA.Selenium.Support.UI.WebDriverWait(_webDriver, TimeSpan.FromSeconds(30.00));
            wait.Until(driver1 => ((IJavaScriptExecutor)_webDriver).ExecuteScript("return document.readyState").Equals("complete"));
        }

        public void GotoTabletApp_DocumentTab()
        {
            //Results are opened in a new tab, switch the driver to the newly created one
            var wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(60));
            wait.Until((driver) => driver.WindowHandles.Count > 1);
            _webDriver.SwitchTo().Window(_webDriver.WindowHandles.Last());
        }

        public void DatePickerFrom(DateTime date)
        {
            string xPathDatePicker1 = "//*/mat-datepicker-toggle[@data-mat-calendar='mat-datepicker-0']/button";
            var datePicker1 = WaitForElementIsVisible(By.XPath(xPathDatePicker1));
            datePicker1.Click();
            DateTime reference = DateUtils.Now.Date;
            while (reference.Year > date.Year)
            {
                reference = reference.AddMonths(-1);
                string xPathDatePicker1Month = "//*[@id=\"mat-datepicker-0\"]/mat-calendar-header/div/div/button[2]";
                var datePicker1Month = WaitForElementIsVisible(By.XPath(xPathDatePicker1Month));
                datePicker1Month.Click();
                WaitForLoad();
            }
            while (reference.Month > date.Month)
            {
                reference = reference.AddMonths(-1);
                string xPathDatePicker1Month = "//*[@id=\"mat-datepicker-0\"]/mat-calendar-header/div/div/button[2]";
                var datePicker1Month = WaitForElementIsVisible(By.XPath(xPathDatePicker1Month));
                datePicker1Month.Click();
                WaitForLoad();
            }
            while (reference.Year < date.Year)
            {
                reference = reference.AddMonths(1);
                string xPathDatePicker2Month = "//*[@id=\"mat-datepicker-0\"]/mat-calendar-header/div/div/button[3]";
                var datePicker1Month = WaitForElementIsVisible(By.XPath(xPathDatePicker2Month));
                datePicker1Month.Click();
                WaitForLoad();
            }
            while (reference.Month < date.Month)
            {
                reference = reference.AddMonths(1);
                string xPathDatePicker1Month = "//*[@id=\"mat-datepicker-0\"]/mat-calendar-header/div/div/button[3]";
                var datePicker1Month = WaitForElementIsVisible(By.XPath(xPathDatePicker1Month));
                datePicker1Month.Click();
                WaitForLoad();
            }
            int jour = date.Day;
            // on peut faire mieux, tirer le jour du premier coup
            string xPathDatePicker1Day;
            xPathDatePicker1Day = "//*[@id='mat-datepicker-0']//span[text()=' " + jour + " ']/parent::button";
            var datePicker1Day = WaitForElementIsVisible(By.XPath(xPathDatePicker1Day));
            datePicker1Day.Click();
            Thread.Sleep(2000);
            WaitForLoad();
            if (isElementVisible(By.Id("mat-datepicker-0")))
            {
                //parfois ne se repli pas
                datePicker1Day = WaitForElementIsVisible(By.XPath(xPathDatePicker1Day));
                datePicker1Day.Click();
            }
            WaitForLoad();
        }

        public void DatePickerTo(DateTime date)
        {
            string xPathDatePicker2 = "//*/mat-datepicker-toggle[@data-mat-calendar='mat-datepicker-1']/button";
            var datePicker2 = WaitForElementIsVisible(By.XPath(xPathDatePicker2));
            new Actions(_webDriver).MoveToElement(datePicker2).Perform();
            datePicker2.Click();
            WaitForLoad();
            DateTime reference = DateUtils.Now.Date;
            while (reference.Year > date.Year)
            {
                reference = reference.AddMonths(-1);
                string xPathDatePicker1Month = "//*[@id=\"mat-datepicker-1\"]/mat-calendar-header/div/div/button[2]";
                var datePicker1Month = WaitForElementIsVisible(By.XPath(xPathDatePicker1Month));
                WaitForLoad();
                datePicker1Month.Click();
            }
            while (reference.Month > date.Month)
            {
                reference = reference.AddMonths(-1);
                string xPathDatePicker1Month = "//*[@id=\"mat-datepicker-1\"]/mat-calendar-header/div/div/button[2]";
                var datePicker1Month = WaitForElementIsVisible(By.XPath(xPathDatePicker1Month));
                WaitForLoad();
                datePicker1Month.Click();
            }
            while (reference.Year < date.Year)
            {
                reference = reference.AddMonths(1);
                string xPathDatePicker2Month = "//*[@id=\"mat-datepicker-1\"]/mat-calendar-header/div/div/button[3]";
                var datePicker1Month = WaitForElementIsVisible(By.XPath(xPathDatePicker2Month));
                datePicker1Month.Click();
                WaitForLoad();
            }
            while (reference.Month < date.Month)
            {
                reference = reference.AddMonths(1);
                string xPathDatePicker1Month = "//*[@id=\"mat-datepicker-1\"]/mat-calendar-header/div/div/button[3]";
                var datePicker1Month = WaitForElementIsVisible(By.XPath(xPathDatePicker1Month));
                datePicker1Month.Click();
                WaitForLoad();
            }
            int jour = date.Day;
            // on peut faire mieux, tirer le jour du premier coup
            string xPathDatePicker1Day;
            xPathDatePicker1Day = "//*[@id='mat-datepicker-1']//span[text()=' " + jour + " ']/parent::button";
            var datePicker1Day = WaitForElementIsVisible(By.XPath(xPathDatePicker1Day));
            datePicker1Day.Click();
            Thread.Sleep(2000);
            WaitForLoad();
            if (isElementVisible(By.Id("mat-datepicker-1")))
            {
                //parfois ne se repli pas
                datePicker1Day = WaitForElementIsVisible(By.XPath(xPathDatePicker1Day));
                datePicker1Day.Click();
            }
            WaitForLoad();
        }

        public bool Select(string xIdDocComboBox, string DocTitle)
        {
            var docComboBox = WaitForElementIsVisibleNew(By.Id(xIdDocComboBox));

            int countdown = 30;
            while (docComboBox.Text != DocTitle)
            {
                countdown--;
                if (countdown == 0)
                {
                    break;
                }
                docComboBox.SendKeys(Keys.ArrowDown);
            }
            WaitForLoad();
            WaitPageLoading();
            return docComboBox.Text == DocTitle;
        }

        public void SelectV2(string xIdDocComboBox, string DocTitle)
        {
            var docComboBox = WaitForElementIsVisible(By.Id(xIdDocComboBox));
            SelectElement se = new SelectElement(docComboBox);
            se.SelectByValue(DocTitle);
            WaitForLoad();
        }

        public void SelectV3(string xIdDocComboBox, string DocTitle)
        {
            _webDriver.FindElement(By.Id(xIdDocComboBox)).Click();
            //_webDriver.FindElement(By.XPath("mat-option-0")).Click();
            _webDriver.FindElement(By.XPath("//*/span[contains(text(),'" + DocTitle + "')]/parent::mat-option")).Click();
            Thread.Sleep(2000);
            WaitForLoad();
        }

        public void SelectAction(string xIdDocComboBox, string docTitle, bool newSelect = false)
        {
            Actions action = new Actions(_webDriver);
            var cible = WaitForElementIsVisibleNew(By.Id(xIdDocComboBox));
            action.MoveToElement(cible).Click().Perform();
            Thread.Sleep(2000);
            WaitForLoad();

            if (newSelect)
            {
                // de bleu à ouvert
                Actions action3 = new Actions(_webDriver);
                cible = WaitForElementIsVisible(By.Id(xIdDocComboBox));
                action3.MoveToElement(cible).Click().Perform();
                // chargement Ajax
                Thread.Sleep(4000);
                WaitForLoad();
                WaitPageLoading();
                Actions action2 = new Actions(_webDriver);
                var option = WaitForElementIsVisible(By.XPath(string.Format(NEWSELECT_OPTION, docTitle)));
                action2.MoveToElement(option).Click().Perform();
            }
            else
            {
                Actions action2 = new Actions(_webDriver);
                var option = WaitForElementIsVisible(By.XPath(string.Format(SELECT_OPTION, docTitle)));
                action2.MoveToElement(option).Click().Perform();
            }
            Thread.Sleep(2000);
            WaitForLoad();
        }

        /**
         * A lancer à la fin sinon le combobox est bleu et "NEXT STEP" bug.
         */
        public void Purge(int nb)
        {
            string xIdDocComboBox = "mat-select-" + nb;
            // ouvre le combobox
            Actions action = new Actions(_webDriver);
            var cible = WaitForElementIsVisible(By.Id(xIdDocComboBox));
            action.MoveToElement(cible).Click().Perform();
            WaitForLoad();

            // de bleu à ouvert
            Actions action3 = new Actions(_webDriver);
            cible = WaitForElementIsVisible(By.Id(xIdDocComboBox));
            action3.MoveToElement(cible).Click().Perform();
            Thread.Sleep(1000);
            WaitForLoad();

            // scan delete
            string xPath;
            xPath = "//*[@id='mat-select-" + nb + "-panel']/mat-option[*]/mat-icon[text()=' delete ']";
            ReadOnlyCollection<IWebElement> liste = null;
            try
            {
                liste = _webDriver.FindElements(By.XPath(xPath));
                for (int i = liste.Count - 1; i >= 0; i--)
                {
                    liste[i].Click();
                    Thread.Sleep(1000);
                    WaitForLoad();
                }

            }
            catch
            {
                return;
            }

            //referme le combobox
            Actions action4 = new Actions(_webDriver);
            var cible2 = WaitForElementIsVisible(By.Id(xIdDocComboBox));
            action.MoveToElement(cible2).Click().Perform();
          //  action.MoveToElement(cible2).Click().Perform();
            WaitForLoad();
        }


        public void addLineSanitization(string Product, string Qty, string Disinfection1, string Disinfection2, string Comments, string Corrective, string PreparedBy, int offset = 0)
        {
            string xPathAddLine = BUTTON_ADDLINE;
            var addLine = WaitForElementIsVisible(By.XPath(xPathAddLine), "BUTTON_ADDLINE");
            addLine.Click();
            WaitForLoad();

            var colonneProduct = WaitForElementIsVisible(By.Id("mat-input-" + offset));
            colonneProduct.SendKeys(Product);
            offset++;
            WaitForLoad();

            var colonneQty = WaitForElementIsVisible(By.Id("mat-input-" + offset));
            colonneQty.SendKeys(Qty);
            offset++;
            WaitForLoad();

            var colonneDisinfection1 = WaitForElementIsVisible(By.Id("mat-input-" + offset));
            colonneDisinfection1.SendKeys(Disinfection1);
            offset++;
            WaitForLoad();

            var colonneDisinfection2 = WaitForElementIsVisible(By.Id("mat-input-" + offset));
            colonneDisinfection2.SendKeys(Disinfection2);
            offset++;
            WaitForLoad();

            if (offset == 4)
            {
                IWebElement mat_checkbox;

                for (int i = 1; i <= 2; i++)
                {
                    mat_checkbox = WaitForElementExists(By.XPath("//*[@id='mat-mdc-checkbox-" + i + "']"));
                    mat_checkbox.Click();
                    WaitForLoad();
                }
            }


            // Trois icones
            var notes = _webDriver.FindElements(By.XPath(NOTES));

            notes[notes.Count - 3].Click();
            var colonneComments = WaitForElementIsVisible(By.Id("mat-input-" + offset));
            colonneComments.SendKeys(Comments);
            string xPathMessageBoxSave;
            xPathMessageBoxSave = string.Format("//*[@id=\"mat-mdc-dialog-{0}\"]/div/div/confirm-dialog/div[2]/button[2]", 0);
            var colonneCommentsSave = WaitForElementIsVisible(By.XPath(xPathMessageBoxSave));
            colonneCommentsSave.Click();
            offset++;
            WaitForLoad();

            notes[notes.Count - 2].Click();
            var colonneCorrectiveAction = WaitForElementIsVisible(By.Id("mat-input-" + offset));
            colonneCorrectiveAction.SendKeys(Corrective);
            xPathMessageBoxSave = string.Format("//*[@id=\"mat-mdc-dialog-{0}\"]/div/div/confirm-dialog/div[2]/button[2]", 1);

            colonneCommentsSave = WaitForElementIsVisible(By.XPath(xPathMessageBoxSave));
            colonneCommentsSave.Click();
            offset++;
            WaitForLoad();

            notes[notes.Count - 1].Click();
            var colonnePreparedBy = WaitForElementIsVisible(By.Id("mat-input-" + offset));
            colonnePreparedBy.Clear();
            colonnePreparedBy.SendKeys(PreparedBy);
            xPathMessageBoxSave = string.Format("//*[@id=\"mat-mdc-dialog-{0}\"]/div/div/confirm-dialog/div[2]/button[2]", 2);

            colonneCommentsSave = WaitForElementIsVisible(By.XPath(xPathMessageBoxSave));
            colonneCommentsSave.Click();
            WaitForLoad();
        }


        public void addLineSanitizationProdInFlight(string Item, string Qty, string DesinfectionQty, string DisinfectionTime, bool Rising, bool Final, string Comments, string PreparedBy)
        {
            string xPathMandarinas = "//*/div[contains(text(),'" + Item + "')]";
            var mandarinas = WaitForElementIsVisible(By.XPath(xPathMandarinas));
            mandarinas.Click();

            var colonneQty = WaitForElementIsVisible(By.Id("mat-input-1"));
            colonneQty.SendKeys(Qty);

            var colonneDesinfectionQty = WaitForElementIsVisible(By.Id("mat-input-3"));
            colonneDesinfectionQty.SendKeys(DesinfectionQty);

            var colonneDisinfectionTime = WaitForElementIsVisible(By.Id("mat-input-5"));
            colonneDisinfectionTime.SendKeys(DisinfectionTime);

            IWebElement colonneRising;
            colonneRising = WaitForElementIsVisible(By.XPath("//*/input[@id='mat-mdc-checkbox-1-input']/parent::div"));
            if ((colonneRising.Enabled && !Rising) || (!colonneRising.Enabled && Rising))
            {
                colonneRising.Click();
            }

            IWebElement colonneFinal;
            colonneFinal = WaitForElementIsVisible(By.XPath("//*/input[@id='mat-mdc-checkbox-2-input']/parent::div"));
            if ((colonneFinal.Enabled && !Final) || (!colonneFinal.Enabled && Final))
            {
                colonneFinal.Click();
            }

            var notes = _webDriver.FindElements(By.XPath(NOTES));
            notes[0].Click();

            var colonneComments = WaitForElementIsVisible(By.XPath("//*/textarea[contains(@class,'comment-text')]"));
            colonneComments.SendKeys(Comments);
            var colonneCommentsSave = WaitForElementIsVisible(By.XPath("//*/mat-dialog-actions/button[contains(@class,'validate-button')]"));
            colonneCommentsSave.Click();

            var colonnePreparedBy = WaitForElementIsVisible(By.Id("mat-input-10"));
            colonnePreparedBy.SendKeys(PreparedBy);
        }

        public void addLineThrowning(string Product, string ExpDate, string BatchNbr, string Qty, string ThrawingStartDate, string MaxEndUsing, string Comments, string Corrective, string PreparedBy, int offset = 0)
        {
            string xPathAddLine = BUTTON_ADDLINE;
            var addLine = WaitForElementIsVisible(By.XPath(xPathAddLine), "BUTTON_ADDLINE");
            addLine.Click();

            var colonneProduct = WaitForElementIsVisible(By.Id("mat-input-" + offset));
            colonneProduct.SendKeys(Product);
            offset++;

            var colonneExpDate = WaitForElementIsVisible(By.Id("mat-input-" + offset));
            colonneExpDate.SendKeys(ExpDate);
            offset++;

            var colonneBatchNbr = WaitForElementIsVisible(By.Id("mat-input-" + offset));
            colonneBatchNbr.SendKeys(BatchNbr);
            offset++;

            var colonneQty = WaitForElementIsVisible(By.Id("mat-input-" + offset));
            colonneQty.SendKeys(Qty);
            offset++;

            var colonneThrawingStartDate = WaitForElementIsVisible(By.Id("mat-input-" + offset));
            colonneThrawingStartDate.SendKeys(ThrawingStartDate);
            offset++;

            var colonneMaxEndUsing = WaitForElementIsVisible(By.Id("mat-input-" + offset));
            colonneMaxEndUsing.SendKeys(MaxEndUsing);
            offset++;

            // Trois icones
            var notes = _webDriver.FindElements(By.XPath(NOTES));

            notes[0].Click();
            var colonneComments = WaitForElementIsVisible(By.Id("mat-input-" + offset));
            colonneComments.SendKeys(Comments);
            string xPathMessageBoxSave;
            xPathMessageBoxSave = string.Format("//*[@id=\"mat-mdc-dialog-{0}\"]/div/div/confirm-dialog/div[2]/button[2]", 0);
            var colonneCommentsSave = WaitForElementIsVisible(By.XPath(xPathMessageBoxSave));
            colonneCommentsSave.Click();
            offset++;

            notes[1].Click();
            var colonneCorrectiveAction = WaitForElementIsVisible(By.Id("mat-input-" + offset));
            colonneCorrectiveAction.SendKeys(Corrective);
            xPathMessageBoxSave = string.Format("//*[@id=\"mat-mdc-dialog-{0}\"]/div/div/confirm-dialog/div[2]/button[2]", 1);
            colonneCommentsSave = WaitForElementIsVisible(By.XPath(xPathMessageBoxSave));
            colonneCommentsSave.Click();
            offset++;

            notes[2].Click();
            var colonnePreparedBy = WaitForElementIsVisible(By.Id("mat-input-" + offset));
            colonnePreparedBy.Clear();
            colonnePreparedBy.SendKeys(PreparedBy);
            xPathMessageBoxSave = string.Format("//*[@id=\"mat-mdc-dialog-{0}\"]/div/div/confirm-dialog/div[2]/button[2]", 2);
            colonneCommentsSave = WaitForElementIsVisible(By.XPath(xPathMessageBoxSave));
            colonneCommentsSave.Click();
        }

        public void addLineThawingProdInFlight(string Item, string Qty, string ExpDate, string BatchNumber, string StartDate, string EndDate, string Comments, string PreparedBy, int offset = 0)
        {
            int increment = 2;
            offset++; // zapper le filtre en hauts

            string xPathCanapes = "//*/div[contains(text(),'" + Item + "')]";
            var mandarinas = WaitForElementIsVisible(By.XPath(xPathCanapes));
            mandarinas.Click();

            var colonneQty = WaitForElementIsVisible(By.Id("mat-input-" + offset));
            colonneQty.SendKeys(Qty);
            offset += increment;

            var colonneExpDate = WaitForElementIsVisible(By.Id("mat-input-" + offset));
            colonneExpDate.SendKeys(ExpDate);
            offset += increment;

            var colonneBatchNumber = WaitForElementIsVisible(By.Id("mat-input-" + offset));
            colonneBatchNumber.SendKeys(BatchNumber);
            offset += increment;

            var colonneStartDate = WaitForElementIsVisible(By.Id("mat-input-" + offset));
            colonneStartDate.SendKeys(StartDate);
            offset += increment;

            var colonneEndDate = WaitForElementIsVisible(By.Id("mat-input-" + offset));
            colonneEndDate.SendKeys(EndDate);
            offset += increment;

            var colonneComments = WaitForElementIsVisible(By.Id("mat-input-" + offset));
            colonneComments.SendKeys(Comments);
            offset += increment;

            var colonnePreparedBy = WaitForElementIsVisible(By.Id("mat-input-" + offset));
            colonnePreparedBy.SendKeys(PreparedBy);
        }

        public void addLineTest(bool CNA, bool CNC, int offset = 1, int line = 0)
        {
            offset = offset + line;
            WaitForLoad();
            string xPathAddLine = BUTTON_ADDLINE;
            var addLine = WaitForElementIsVisible(By.XPath(xPathAddLine), "BUTTON_ADDLINE");
            addLine.Click();
            WaitForLoad();

            IWebElement colonneCNA;

            colonneCNA = WaitForElementIsVisible(By.XPath("//*/input[@id=\"mat-mdc-checkbox-" + offset + "-input\"]/parent::div"));
            if (CNA)
            {
                colonneCNA.Click();
                WaitForLoad();
            }
            offset++;

            IWebElement colonneCNC;
            colonneCNC = WaitForElementIsVisible(By.XPath("//*/input[@id=\"mat-mdc-checkbox-" + offset + "-input\"]/parent::div"));
            if (CNC)
            {
                colonneCNC.Click();
                WaitForLoad();
            }

            var colonneMultiDate = _webDriver.FindElements(By.XPath("//*[contains(text(),'Open calendar')]"));
            var col = colonneMultiDate[colonneMultiDate.Count - 1];
            col.Click();
            WaitForLoad();

            // je lance un jet de clicks
            for (
                int i = 10; i < 17; i++)
            {
                IWebElement arrayList;
                arrayList = WaitForElementIsVisible(By.XPath("//*[contains(@class,'mat-calendar-body-cell-content')][contains(text(),'" + i + "')]"));
                if (!arrayList.GetAttribute("class").Contains("selected"))
                {
                    WaitPageLoading();
                    arrayList.Click();
                    WaitForLoad();
                }
            }
            // sortir du MultiDate
            var xPathSortieMultiDate = "//*[contains(text(),'CheckboxCNA')]";
            Actions action = new Actions(_webDriver);
            var cible = WaitForElementIsVisible(By.XPath(xPathSortieMultiDate));
            action.MoveToElement(cible).Click().Perform();
        }

        public void EditLineTraySetup()
        {
            // ouvrir le mode edition
            var xPathEdit = "//*/p[contains(text(), 'Service')]";
            var cible1 = WaitForElementIsVisible(By.XPath(xPathEdit));
            cible1.Click();
        }

        public void addLineTraySetup(bool CNC, bool CNA)
        {
            // ouvrir le mode edition
            EditLineTraySetup();


            IWebElement colonneCNC;
            colonneCNC = WaitForElementIsVisible(By.XPath("//*/input[@id='mat-mdc-checkbox-1-input']/parent::div"));
            if (CNC)
            {
                WaitForLoad();
                colonneCNC.Click();
            }

            IWebElement colonneCNA;
            colonneCNA = WaitForElementIsVisible(By.XPath("//*/input[@id='mat-mdc-checkbox-2-input']/parent::div"));
            if (CNA)
            {
                WaitForLoad();
                colonneCNA.Click();
            }

            var colonneMultiDate = WaitForElementIsVisible(By.XPath("//*[contains(text(),'Open calendar')]"));
            colonneMultiDate.Click();

            // je lance un jet de clicks
            ReadOnlyCollection<IWebElement> arrayMultiDate;
            for (int i = 10; i < 17; i++)
            {
                arrayMultiDate = _webDriver.FindElements(By.XPath("//*[contains(@class,'mat-calendar-body-cell-content')][contains(text(),'" + i + "')]/parent::button"));
                IWebElement arrayList = arrayMultiDate.ToList<IWebElement>()[arrayMultiDate.Count - 1];
                if (!arrayList.GetAttribute("class").Contains("selected"))
                {
                    WaitPageLoading();
                    arrayList.Click();
                }
            }
            // sortir du MultiDate
            var xPathSortieMultiDate = "//*[contains(text(),'Service - Cycle')]";
            Actions action = new Actions(_webDriver);
            var cible = WaitForElementIsVisible(By.XPath(xPathSortieMultiDate));
            action.MoveToElement(cible).Click().Perform();
        }

        public void addLineModifiedTexture(string Product, string StartMixingBegin, string StartMixingEnd, string EndMixingBegin, string EndMixingEnd, string Comments, string Corrective, string PreparedBy, int offset = 0)
        {
            string xPathAddLine = BUTTON_ADDLINE;
            var addLine = WaitForElementIsVisible(By.XPath(xPathAddLine), "BUTTON_ADDLINE");
            addLine.Click();
            WaitForLoad();

            var colonneProduct = WaitForElementIsVisible(By.Id("mat-input-" + offset));
            colonneProduct.SendKeys(Product);
            offset++;
            WaitForLoad();

            var colonneStartMixingBegin = WaitForElementIsVisible(By.Id("mat-input-" + offset));
            colonneStartMixingBegin.SendKeys(StartMixingBegin);
            offset++;
            WaitForLoad();

            var colonneStartMixingEnd = WaitForElementIsVisible(By.Id("mat-input-" + offset));
            colonneStartMixingEnd.SendKeys(StartMixingEnd);
            offset++;
            WaitForLoad();

            var colonneEndMixingBegin = WaitForElementIsVisible(By.Id("mat-input-" + offset));
            colonneEndMixingBegin.SendKeys(EndMixingBegin);
            offset++;
            WaitForLoad();

            var colonneEndMixingEnd = WaitForElementIsVisible(By.Id("mat-input-" + offset));
            colonneEndMixingEnd.SendKeys(EndMixingEnd);
            offset++;
            WaitForLoad();

            // Trois icones
            var notes = _webDriver.FindElements(By.XPath(NOTES));

            notes[notes.Count - 3].Click();
            var colonneComments = WaitForElementIsVisible(By.Id("mat-input-" + offset));
            colonneComments.SendKeys(Comments);
            string xPathMessageBoxSave;
            xPathMessageBoxSave = string.Format("//*[@id=\"mat-mdc-dialog-{0}\"]/div/div/confirm-dialog/div[2]/button[2]", 0);
            var colonneCommentsSave = WaitForElementIsVisible(By.XPath(xPathMessageBoxSave));
            colonneCommentsSave.Click();
            offset++;
            WaitForLoad();

            notes[notes.Count - 2].Click();
            var colonneCorrectiveAction = WaitForElementIsVisible(By.Id("mat-input-" + offset));
            colonneCorrectiveAction.SendKeys(Corrective);
            xPathMessageBoxSave = string.Format("//*[@id=\"mat-mdc-dialog-{0}\"]/div/div/confirm-dialog/div[2]/button[2]", 1);
            colonneCommentsSave = WaitForElementIsVisible(By.XPath(xPathMessageBoxSave));
            colonneCommentsSave.Click();
            offset++;
            WaitForLoad();

            notes[notes.Count - 1].Click();
            var colonnePreparedBy = WaitForElementIsVisible(By.Id("mat-input-" + offset));
            colonnePreparedBy.Clear();
            colonnePreparedBy.SendKeys(PreparedBy);
            xPathMessageBoxSave = string.Format("//*[@id=\"mat-mdc-dialog-{0}\"]/div/div/confirm-dialog/div[2]/button[2]", 2);
            colonneCommentsSave = WaitForElementIsVisible(By.XPath(xPathMessageBoxSave));
            colonneCommentsSave.Click();
            WaitForLoad();
        }

        /**
         * withIsVisible pas certain que le bouton est là
         */
        public void ClickButton(string name, bool withIsVisible = false)
        {
            WaitLoading();
            string xPathDocSave = string.Format(CLICK_BUTTON, name);
            if (withIsVisible)
            {
                if (!isElementVisible(By.XPath(xPathDocSave)))
                {
                    return;
                }
            }
            var docSave = WaitForElementIsVisibleNew(By.XPath(xPathDocSave), "CLICK_BUTTON " + name);
            docSave.Click();
            WaitPageLoading();
            WaitPageLoading();
        }

        public void ClickBaseButton(string name)
        {
            string xPathNext = string.Format(CLICK_BASEBUTTON, name);
            var next = WaitForElementIsVisibleNew(By.XPath(xPathNext), CLICK_BASEBUTTON);
            next.Click();
        }

        public void FillFileName3(string docFileName)
        {
            string xPathFileName = FILL_FILENAME3;
            IWebElement fileName;
            fileName = WaitForElementIsVisible(By.XPath("//*[@id=\"mat-mdc-dialog-3\"]//haccp--modal-favorite/input"), "FILL_FILENAME");
            fileName.Clear();
            fileName.SendKeys(docFileName);
        }
        public void FillFileName0(string docFileName)
        {
            string xPathFileName = FILL_FILENAME0;
            IWebElement fileName;
            fileName = WaitForElementIsVisible(By.XPath("//*[@id=\"mat-mdc-dialog-0\"]/div/div/haccp--modal-favorite/input"), "FILL_FILENAME");
            fileName.Clear();
            fileName.SendKeys(docFileName);
        }

        public void FileFileNameProdInFlight0(string docFileName)
        {
            string xPathFileName = FILL_FILENAME_PRODINFLIGHT0;
            var fileName = WaitForElementIsVisible(By.XPath(xPathFileName), "FILL_FILENAME");
            fileName.Clear();
            fileName.SendKeys(docFileName);
        }

        public void FileFileNameProdInFlight1(string docFileName)
        {
            string xPathFileName = FILL_FILENAME_PRODINFLIGHT1;
            var fileName = WaitForElementIsVisible(By.XPath(xPathFileName), "FILL_FILENAME");
            fileName.Clear();
            fileName.SendKeys(docFileName);
        }

        private int MiddleFilter(int noChild, string value)
        {
            //Toggle
            var services = WaitForElementIsVisible(By.XPath(string.Format(FILTER_TOGGLE, noChild)));
            services.Click();
            // Tout déselectionner
            var selectAll = WaitForElementIsVisible(By.XPath(string.Format(FILTER_SELECT_ALL, noChild)));
            selectAll.Click();
            // search
            var search = WaitForElementIsVisible(By.XPath(string.Format(FILTER_SEARCH, noChild)));
            search.Clear();
            search.SendKeys(value);
            // cocher les résultats
            string xPathTableau;
            if (noChild == 0)
            {
                xPathTableau = string.Format(FILTER_TABLEAU_0, value);
            }
            else
            {
                xPathTableau = string.Format(FILTER_TABLEAU, value);
            }
            var tableau = _webDriver.FindElements(By.XPath(xPathTableau));
            // Select All
            foreach (var ligne in tableau)
            {
                if (!ligne.GetAttribute("class").Contains("selected"))
                {
                    ligne.Click();
                }
            }
            return tableau.Count;

        }
        public int Filter(TabletAppPage.FilterType filterType, object value)
        {
            switch (filterType)
            {
                case FilterType.Services:
                    return MiddleFilter(1, (string)value);
                case FilterType.Customers:
                    return MiddleFilter(0, (string)value);
                case FilterType.ServiceCategories:
                    return MiddleFilter(5, (string)value);
                case FilterType.GuestType:
                    return MiddleFilter(2, (string)value);
                case FilterType.RecipeType:
                    return MiddleFilter(4, (string)value);
                case FilterType.ItemGroups:
                    return MiddleFilter(6, (string)value);
                case FilterType.Workshop:
                    return MiddleFilter(3, (string)value);
                case FilterType.ShowNormalProductionOnly:
                    var radioProduction = WaitForElementIsVisible(By.XPath("//*[@id=\"mat-radio-4\"]/div/div/div[1]"));
                    radioProduction.Click();
                    break;
                case FilterType.ShowVacuumProductionOnly:
                    var radioVacuum = WaitForElementIsVisible(By.XPath("//*[@id=\"mat-radio-3\"]/div/div/div[1]"));
                    radioVacuum.Click();
                    break;
                case FilterType.ShowNormalAndVacuumProduction:
                    var radioProductionVacuum = WaitForElementIsVisible(By.Id("mat-radio-2"));
                    radioProductionVacuum.Click();
                    break;
                case FilterType.ValidatedFlightsOnly:
                    IWebElement checkboxVal;
                    checkboxVal = WaitForElementExists(By.XPath("//*/input[@id='mat-mdc-checkbox-2-input']/parent::div"));
                    checkboxVal.SetValue(ControlType.CheckBox, value);
                    break;
                case FilterType.ShowServicesWithoutDatasheetOnly:
                    var checkboxVal3 = WaitForElementIsVisible(By.XPath("//*[@id='mat-checkbox-3']/label"));
                    checkboxVal3.Click();
                    break;
                case FilterType.From:
                    DatePickerFrom((DateTime)value);
                    break;
                case FilterType.To:
                    DatePickerTo((DateTime)value);
                    break;
                case FilterType.StartTime:
                    var startTime = WaitForElementIsVisible(By.XPath("//*/label[text()='Start time']/following-sibling::input"));
                    startTime.SendKeys((string)value);
                    break;
                case FilterType.EndTime:
                    var endTime = WaitForElementIsVisible(By.XPath("//*/label[text()='End time']/following-sibling::input"));
                    endTime.SendKeys((string)value);
                    break;

                case FilterType.ShowHiddenArticles:
                    string ShowHiddenArticlesForClick;
                    ShowHiddenArticlesForClick = "//*[@id='mat-mdc-checkbox-1-input']/parent::div";
                    var ShowHiddenCheckBoxClick = WaitForElementIsVisible(By.XPath(ShowHiddenArticlesForClick));
                    ShowHiddenCheckBoxClick.SetValue(ControlType.CheckBox, value);
                    break;
            }
            return -1;
        }

        public string CheckValue(int offset)
        {
            var columnProject = WaitForElementIsVisible(By.Id("mat-input-" + offset));
            return columnProject.GetProperty("value");
        }

        public bool CheckValueBox(int offset)
        {
            try
            {
                var columnProject = _webDriver.FindElement(By.Id("mat-mdc-checkbox-" + offset + "-input"));
                return columnProject.GetAttribute("class").Contains("mdc-checkbox--selected");
            }
            catch
            {
                return false;
            }
        }

        public string CheckNote(int offsetClick, int offset)
        {
            var notes = _webDriver.FindElements(By.XPath(NOTES));

            notes[offsetClick].Click();
            var colonneComments = WaitForElementIsVisible(By.Id("mat-input-" + offset));
            var result = colonneComments.GetProperty("value");
            ClickButton("Cancel");
            return result;
        }

        public bool CheckMultiDate()
        {
            var colonneMultiDate = WaitForElementIsVisible(By.XPath("//*[contains(text(),'Open calendar')]"));
            colonneMultiDate.Click();
            bool valide = true;
            for (int i = 10; i < 17; i++)
            {
                IWebElement arrayMultiDate;
                arrayMultiDate = WaitForElementIsVisible(By.XPath("//*[contains(@class,'mat-calendar-body-cell-content')][contains(text(),'" + i + "')]/parent::button"));

                if (!arrayMultiDate.GetAttribute("class").Contains("selected"))
                {
                    valide = false;
                }
            }
            for (int i = 18; i < 20; i++)
            {
                IWebElement arrayMultiDate;
                arrayMultiDate = WaitForElementIsVisible(By.XPath("//*[contains(@class,'mat-calendar-body-cell-content')][contains(text(),'" + i + "')]/parent::button"));
                if (arrayMultiDate.GetAttribute("class").Contains("selected"))
                {
                    valide = false;
                }
            }
            // sortir du MultiDate
            Actions action = new Actions(_webDriver);
            var cible = WaitForElementIsVisible(By.XPath("//*[contains(text(),'Open calendar')]"));
            action.MoveToElement(cible).Click().Perform();
            return valide;
        }



        public int FilterCounter(string value)
        {
            string xPathCheck = "//*[contains(text(),'" + value + "')]";
            string countServiceCategoris = WaitForElementIsVisible(By.XPath(xPathCheck)).Text;
            Regex r = new Regex(value.ToUpper() + " \\(([0-9]+)\\)");
            Match m = r.Match(countServiceCategoris);
            return int.Parse(m.Groups[1].Value);
        }

        public void GotoTabletApp_BlankDoc()
        {
            string xPathLink = "//*[contains(text(),'HACCP Blank docs')]/parent::div";
            var ProdInFlight = WaitForElementIsVisible(By.XPath(xPathLink));
            ProdInFlight.Click();
            Thread.Sleep(2000);
            WaitForLoad();
        }

        public void GotoTabletApp_ProdInFlight()
        {
            string xPathLink = "//*[contains(text(),'Prod. inflight')]/parent::div";
            var ProdInFlight = WaitForElementIsVisibleNew(By.XPath(xPathLink));
            ProdInFlight.Click();
            WaitForLoad();
            WaitPageLoading();
        }
        public TimeBlockTabletAppPage GotoTabletApp_TimeBlock()
        {
            string xPathLink = "//*[contains(text(),'Time Block')]/parent::div";
            var TimeBlock = WaitForElementIsVisible(By.XPath(xPathLink));
            TimeBlock.Click();
            WaitForLoad();

            return new TimeBlockTabletAppPage(_webDriver, _testContext);
        }

        public FlightTabletAppPage GotoTabletApp_Flight()
        {
            var TimeBlock = WaitForElementIsVisibleNew(By.XPath(TABLET_APP_FLIGHT));
            TimeBlock.Click();
            WaitForLoad();

            return new FlightTabletAppPage(_webDriver, _testContext);
        }

        public ToDoListTabletAppPage GotoTabletApp_ToDoList()
        {
            var TimeBlock = WaitForElementIsVisible(By.XPath(TABLET_APP_TODOLIST));
            TimeBlock.Click();
            WaitForLoad();

            return new ToDoListTabletAppPage(_webDriver, _testContext);
        }

        public DatasheetTabletAppPage GotoTabletApp_Datasheet()
        {
            string xPathLink = "//*[contains(text(),'Datasheet')]/parent::div";
            var ProdInFlight = WaitForElementIsVisibleNew(By.XPath(xPathLink));
            ProdInFlight.Click();
            WaitForLoad();
            WaitPageLoading();
            return new DatasheetTabletAppPage(_webDriver, _testContext);

        }

        // Purge pour le document Print
        public void Purge(string downloadsPath, string BeginFileName)
        {
            string[] files = Directory.GetFiles(downloadsPath);
            foreach (string f in files)
            {
                FileInfo fi = new FileInfo(f);
                if (fi.Name.StartsWith(BeginFileName))
                {
                    fi.Delete();
                }
            }
        }
        public FileInfo GetReportPdf(FileInfo[] taskFiles, string fileNameStart)
        {
            List<FileInfo> correctDownloadFiles = new List<FileInfo>();

            foreach (var file in taskFiles)
            {
                //  Test REGEX name_date_time
                if (IsPDFFileCorrectDateTime(file.Name, fileNameStart))
                {
                    correctDownloadFiles.Add(file);
                }
            }
            if (correctDownloadFiles.Count <= 0)
            {
                foreach (var file in taskFiles)
                {
                    //  Test REGEX name_date_date
                    if (IsPDFFileCorrectDateDate(file.Name, fileNameStart))
                    {
                        correctDownloadFiles.Add(file);
                    }
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
        public bool IsPDFFileCorrectDateTime(string filePath, string fileNameStart)
        {
            // HACCP3 Sanitization6_28032022_073730

            string mois = "(?:0[1-9]|1[0-2])";         // mois
            string annee = "\\d{4}";                   // annee YYYY
            string jour = "[0-3]\\d";                  // jour
            string heure = "(?:0[0-9]|1[0-9]|2[0-3])"; // heure
            string minutes = "[0-5]\\d";               // minutes
            string secondes = "[0-5]\\d";              // secondes

            Regex r = new Regex(fileNameStart + jour + mois + annee + "_" + heure + minutes + secondes, RegexOptions.IgnoreCase | RegexOptions.Singleline);
            Match m = r.Match(filePath);

            return m.Success;
        }
        public bool IsPDFFileCorrectDateDate(string filePath, string fileNameStart)
        {
            // HACCP3 Sanitization6_28032022_073730

            string mois = "(?:0[1-9]|1[0-2])";         // mois
            string annee = "\\d{4}";                   // annee YYYY
            string jour = "[0-3]\\d";                  // jour

            Regex r = new Regex(fileNameStart + "_" + jour + mois + annee + "_" + jour + mois + annee, RegexOptions.IgnoreCase | RegexOptions.Singleline);
            Match m = r.Match(filePath);

            return m.Success;
        }

        public void SavefavoriteDocument(string DocFileName)
        {
            ClickButton("Favorite");
            //cliquer sur ""validate""
            FillFileName0(DocFileName);
            ClickButton("Validate");
            WaitForLoad();
            Thread.Sleep(2000);
            ClickButton("Save and validate");
            ClickButton("Save favorite");
            WaitForLoad();
            WaitHACCPHorizontalProgressBar();
            ClickButton("Ok");
        }

        public void WaitHACCPHorizontalProgressBar()
        {
            // attente de la progress bar
            int compteur = 1;
            bool vueSablier = false;
            while (compteur <= 1000)
            {
                try
                {
                    _webDriver.FindElement(By.ClassName("progress"));
                    vueSablier = true;
                    break;
                }
                catch
                {
                    compteur++;
                }
            }

            // attente de la fin de la progress bar
            compteur = 1;

            while (compteur <= 600 && vueSablier)
            {
                try
                {
                    var wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(1));
                    wait.Until(ExpectedConditions.ElementIsVisible(By.ClassName("progress")));
                    compteur++;
                    // Ophélie : ajout d'un sleep pour augmenter le temps d'attente (équivalent à 1 minute max au total)
                    Thread.Sleep(100);
                }
                catch
                {
                    vueSablier = false;
                }
            }

            if (vueSablier)
            {
                throw new Exception("Délai d'attente dépassé pour le chargement de la page.");
            }
            WaitForLoad();
        }
        public void SetSite(string siteToSet)
        {
            IWebElement siteBtn;
            siteBtn = WaitForElementExists(By.XPath("//*[@id=\"mat-select-0\"]/div/div[1]"));
            siteBtn.Click();
            ReadOnlyCollection<IWebElement> sites = _webDriver.FindElements(By.XPath(SITES));
            sites = _webDriver.FindElements(By.XPath("//*[@id=\"mat-select-0-panel\"]/mat-option[*]/span"));
            foreach (var site in sites)
            {
                if (site.Text.Contains(siteToSet))
                {
                    site.Click();
                    break;
                }
            }
        }

        public bool CanNextStep()
        {
            var printButton = WaitForElementIsVisible(By.XPath("//*[contains(text(),'NEXT STEP')]"));
            return !printButton.GetAttribute("class").Contains("hidden");
        }

        public void WaitForDownloadHACCP()
        {
            WaitForDownload();
            Close();
        }
        public void MenuTestModal()
        {
            if (IsDev())
            {
                string xPathLink = "/html/body/app-root/mat-sidenav-container/mat-sidenav-content/main/haccp-tray-setup/div/virtual-scroller/div[2]/div[6]/div[1]/span";
                var MenuTest = WaitForElementIsVisible(By.XPath(xPathLink));
                MenuTest.Click();
                Thread.Sleep(2000);
                WaitForLoad();
            }
            else
            {
                string xPathLink = "/html/body/app-root/mat-sidenav-container/mat-sidenav-content/main/haccp-hot-kitchen/div/virtual-scroller/div[2]/div[2]/div[1]/span";
                var MenuTest = WaitForElementIsVisible(By.XPath(xPathLink));
                MenuTest.Click();
                Thread.Sleep(2000);
                WaitForLoad();
            }
        }
        public void ToggleTablet()
        {
            if (IsDev())
            {
                var _toggleTablet = WaitForElementIsVisible(By.XPath("//*[@id=\"cdk-overlay-0\"]/div/div[2]/datasheet-popup/div/div/div/mat-list[1]/details/summary/div/div[1]/span"));
                _toggleTablet.Click();
            }
        }
        public string GetNameIngredient()
        {
            var nameField = WaitForElementIsVisible(By.XPath(NAME_INGREDIENT));
            return nameField.Text;
        }
        public List<string> GetAllNameIngredient()
        {
            List<string> result = new List<string>();
            var elements = _webDriver.FindElements(By.XPath(ALL_NAME_INGREDIENT));
            if (elements.Count == 0) return new List<string> { "NO DATA" };

            else
            {
                var regex = new Regex(@"more_vert\r\n(.*)");

                foreach (var element in elements)
                {
                    var match = regex.Match(element.Text);
                    if (match.Success)
                    {
                        // Add the captured group (the part after "more_vert\r\n")
                        result.Add(match.Groups[1].Value.Trim());
                    }
                }
                return result;
            }
        }

        public void SelectDayInToDoList(DateTime date)
        {
            WaitPageLoading();
            //Cliquer sur l'icône du calendrier pour ouvrir le sélecteur de date
            var calendarButton = _webDriver.FindElement(By.XPath("/html/body/app-root/mat-sidenav-container/mat-sidenav-content/main/lounge/div/div[1]/div/div[2]/div"));
            calendarButton.Click();

            // Convert the DateTime parameter to the desired format
            string formattedDate = date.ToString("MMMM d, yyyy", CultureInfo.CreateSpecificCulture("en-US"));
            // XPath to find the button based on the aria-label attribute
            string xpath = $"//button[@aria-label='{formattedDate}']";

            // Find the button using the XPath
            var button = _webDriver.FindElement(By.XPath(xpath));
            WaitPageLoading();
            // Click the button
            button.Click();
        }

        public List<string> GetTaskNames()
        {
            WaitPageLoading();

            // XPath of the tbody that contains the rows
            string tbodyXPath = "/html/body/app-root/mat-sidenav-container/mat-sidenav-content/main/lounge/div/div[2]/div/div/table/tbody";

            // Create a list to store the task names
            List<string> taskNames = new List<string>();

            try
            {
                // Try to locate the tbody element
                IWebElement tbody = _webDriver.FindElement(By.XPath(tbodyXPath));

                // Find all the rows in the tbody
                IList<IWebElement> rows = tbody.FindElements(By.TagName("tr"));

                // Iterate through each row and get the task name from the second <td> element
                foreach (IWebElement row in rows)
                {
                    IWebElement taskNameElement = row.FindElement(By.XPath("td[2]"));

                    string taskName = taskNameElement.Text;
                    taskNames.Add(taskName);
                }
            }
            catch (NoSuchElementException)
            {
                Console.WriteLine("Tbody element not found.");
            }

            return taskNames;
        }

        public void clickNextStep(string docTitle)
        {
            var counter = 30;
            // check Number of services 
            var serviceEelementCount = WaitForElementIsVisible(By.XPath("/html/body/app-root/mat-sidenav-container/mat-sidenav-content/main/search/div/div[2]/div/div[2]/selectable-list[2]/div/mat-expansion-panel/mat-expansion-panel-header/span[1]/mat-panel-title/div"));
            var serviceCountInput = serviceEelementCount.Text;

            while (serviceCountInput.Contains("0") && counter>0)
            {
                // get FromDate
                var dateFromInput = WaitForElementIsVisible(By.XPath("/html/body/app-root/mat-sidenav-container/mat-sidenav-content/main/search/div/div[2]/div/div[1]/div/div[1]/div[2]/div[2]/mat-form-field/div[1]/div[2]/div[1]/input"));
                var dateFromValue = dateFromInput.GetAttribute("value");
                DateTime date = DateTime.ParseExact(dateFromValue, "dd/MM/yyyy", CultureInfo.InvariantCulture);
                var newDateFrom = new DateTime(date.Year, date.Month, date.Day - 1);
                Filter(FilterType.From, newDateFrom);
                WaitLoading();
                serviceEelementCount = WaitForElementIsVisible(By.XPath("/html/body/app-root/mat-sidenav-container/mat-sidenav-content/main/search/div/div[2]/div/div[2]/selectable-list[2]/div/mat-expansion-panel/mat-expansion-panel-header/span[1]/mat-panel-title/div"));
                serviceCountInput = serviceEelementCount.Text;
                counter--;
            }
            if (counter == 0 && serviceCountInput.Contains("0")) 
            {
                Assert.IsFalse(serviceCountInput.Contains("0"), "Aucun service trouvé");
            }
            else
            {
                if (counter < 30)
                {

                    Select("mat-select-2", docTitle);
                    SelectAction("mat-select-2", docTitle);
                }
                ClickBaseButton("NEXT STEP");

            }
        }

       public bool isTabletAppPageDisplayed ()
        {
            WaitPageLoading();
            return isElementVisible(By.XPath("/html/body/app-root/mat-sidenav-container"));
        }

        public string GetFileTemporary()
        {
            var nameField = WaitForElementIsVisible(By.XPath(FILE_TEMPORARY));
            return nameField.Text;
        }
        public string GetProduct()
        {
            WaitPageLoading();
            var productField = WaitForElementIsVisible(By.Id(PRODUCT_FIELD));
            return productField.Text;
        }
        public string GetComments()
        {
            var productBtn = WaitForElementIsVisible(By.XPath(COMMENTS_BTN));
            productBtn.Click();
            var textField = WaitForElementIsVisible(By.XPath(TEXT_EDIT));
            return textField.Text;
        }
        public string GetCorrective()
        {
            var correctiveBtn = WaitForElementIsVisible(By.XPath(CORRECTIVE_BTN));
            correctiveBtn.Click();
            var textCorrective = WaitForElementIsVisible(By.XPath(TEXT_CORRECTIVE));
            return textCorrective.Text;
        }
        public string GetFieldPrepared()
        {
            var correctiveBtn = WaitForElementIsVisible(By.XPath(PREPAREDBY_BTN));
            correctiveBtn.Click();
            var textCorrective = WaitForElementIsVisible(By.XPath(TEXT_PREPAREDBY));
            return textCorrective.Text;
        }

        /*public Dictionary<string, string> GetSanitizationLineValues(int offset = 0)
        {
            var fieldValues = new Dictionary<string, string>();

            // Extraire la valeur du champ Product
            var productField = WaitForElementIsVisible(By.Id("mat-input-" + offset));
            fieldValues["Product"] = productField.GetAttribute("value") ?? string.Empty;
            offset++;
            WaitForLoad();

           // Extraire la valeur du champ Quantity
            var qtyField = WaitForElementIsVisible(By.Id("mat-input-" + offset));
            fieldValues["Quantity"] = qtyField.GetAttribute("value") ?? string.Empty;
            offset++;
            WaitForLoad();

            // Extraire la valeur du champ Disinfection1
            var disinfection1Field = WaitForElementIsVisible(By.Id("mat-input-" + offset));
            fieldValues["Disinfection1"] = disinfection1Field.GetAttribute("value") ?? string.Empty;
            offset++;
            WaitForLoad();

            // Extraire la valeur du champ Disinfection2
            var disinfection2Field = WaitForElementIsVisible(By.Id("mat-input-" + offset));
            fieldValues["Disinfection2"] = disinfection2Field.GetAttribute("value") ?? string.Empty;
            offset++;
            WaitForLoad();

            // Extraire la valeur du champ Comments
            var notes = _webDriver.FindElements(By.XPath(NOTES));
            notes[notes.Count - 3].Click();
            var commentsField = WaitForElementIsVisible(By.Id("mat-input-" + offset));
            fieldValues["Comments"] = commentsField.GetAttribute("value") ?? string.Empty;
            notes[notes.Count - 3].Click(); // Fermer le dialogue
            offset++;
            WaitForLoad();

            // Extraire la valeur du champ Corrective
            notes[notes.Count - 2].Click();
            var correctiveField = WaitForElementIsVisible(By.Id("mat-input-" + offset));
            fieldValues["Corrective"] = correctiveField.GetAttribute("value") ?? string.Empty;
            notes[notes.Count - 2].Click(); // Fermer le dialogue
            offset++;
            WaitForLoad();

            // Extraire la valeur du champ Prepared By
            notes[notes.Count - 1].Click();
            var preparedByField = WaitForElementIsVisible(By.Id("mat-input-" + offset));
            fieldValues["PreparedBy"] = preparedByField.GetAttribute("value") ?? string.Empty;
            notes[notes.Count - 1].Click(); // Fermer le dialogue
            WaitForLoad();

            return fieldValues;
        }*/
        public string GetGuestType()
        {
            var guestType = WaitForElementIsVisibleNew(By.XPath("/html/body/app-root/mat-sidenav-container/mat-sidenav-content/main/haccp-tray-setup/div/virtual-scroller/div[2]/div[5]/div[1]/p"));
            return guestType.Text.Trim();
        }


    }
}
