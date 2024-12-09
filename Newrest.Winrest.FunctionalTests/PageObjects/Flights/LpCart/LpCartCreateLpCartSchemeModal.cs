using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;
using System.Linq;
using System.Threading;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Flights.LpCart
{
    public class LpCartCreateLpCartSchemeModal : PageBase
    {
        // ____________________________________ Constantes __________________________________________

        //LpCart Scheme
        private const string TITLE = "Title";
        private const string PREP_ZONE = "PrepZoneConcat";
        private const string COMMENT = "Comment";
        private const string SEALS_NUMBER = "sealsNumber";
        private const string LABEL_PAGE = "labelPageNbr";
        private const string ROWS = "rowsNumber";
        private const string COLUMN = "colsNumber";
        private const string GENERATE = "btnGenerate";
        private const string CONFIRM = "create_btn";
        private const string VALIDATE = "dataConfirmOK";
        private const string FRONT_BTN = "BTN-F0";
        private const string BACK_BTN = "BTN-B0";
        private const string INPUT_SCHEMA_TROLLEY = "//input[contains(@class, 'inputElem')]";
        private const string EQP_NAME_LABEL = "//*[@id=\"EquipmentNameLabel\"]";
        private const string CART_COLOR = "//*[@id=\"trolley-label-color\"]";
        private const string MAIN_POSITION = "//*[@id=\"GlobalPosition\"]";
        private const string SHORT_COMMENT = "/html/body/div[4]/div/div/form/div[2]/div[4]/div[1]/div/div/input";

        // ____________________________________ Variables ___________________________________________


        //LpCart Scheme

        [FindsBy(How = How.Id, Using = TITLE)]
        private IWebElement _title;

        [FindsBy(How = How.Id, Using = PREP_ZONE)]
        private IWebElement _prepZone;

        [FindsBy(How = How.Id, Using = COMMENT)]
        private IWebElement _comment;

        [FindsBy(How = How.Id, Using = SEALS_NUMBER)]
        private IWebElement _sealNumber;

        [FindsBy(How = How.Id, Using = LABEL_PAGE)]
        private IWebElement _labelPage;

        [FindsBy(How = How.Id, Using = ROWS)]
        private IWebElement _rows;

        [FindsBy(How = How.Id, Using = COLUMN)]
        private IWebElement _column;

        [FindsBy(How = How.Id, Using = GENERATE)]
        private IWebElement _generate;

        [FindsBy(How = How.Id, Using = VALIDATE)]
        private IWebElement _validate;

        [FindsBy(How = How.Id, Using = CONFIRM)]
        private IWebElement _confirm;

        [FindsBy(How = How.Id, Using = FRONT_BTN)]
        private IWebElement _front_Btn;

        [FindsBy(How = How.XPath, Using = EQP_NAME_LABEL)]
        private IWebElement _eqpNamelabel;

        [FindsBy(How = How.XPath, Using = CART_COLOR)]
        private IWebElement _cartColor;

        [FindsBy(How = How.XPath, Using = MAIN_POSITION)]
        private IWebElement _mainPosition;

        [FindsBy(How = How.XPath, Using = SHORT_COMMENT)]
        private IWebElement _shortComment;

        // ___________________________________  Méthodes ___________________________________________
        public LpCartCreateLpCartSchemeModal(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        public void ClickGenerate()
        {
            _generate = WaitForElementIsVisible(By.Id(GENERATE));
            _generate.Click();
            WaitPageLoading();
            WaitForLoad();
        }

        public void ClickValidate()
        {
            _validate = WaitForElementIsVisible(By.Id(VALIDATE));
            _validate.Click();
            WaitPageLoading();
            WaitForLoad();

        }

        public void ClickConfirm()
        {
            _confirm = WaitForElementIsVisible(By.Id(CONFIRM));
            _confirm.Click();
            WaitPageLoading();
            WaitForLoad();
        }

        public void CreateLpCartscheme(string title, string rows, string columns, string prepZone = null, string shortComment = null, string comment = null, string sealNbre = null, string labelpage = null, string eqpNamelabel = null, string cartColor = null, string mainPos = null)
        {

            _title = WaitForElementIsVisible(By.Id(TITLE));
            _title.SetValue(ControlType.TextBox, title);
            WaitForLoad();

            if (prepZone != null)
            {
                _prepZone = WaitForElementIsVisible(By.Id(PREP_ZONE));
                _prepZone.SetValue(ControlType.TextBox, prepZone);
                WaitForLoad();
            }
            WaitForLoad();
            if (shortComment != null)
            {
                _shortComment = WaitForElementIsVisible(By.XPath(SHORT_COMMENT));
                _shortComment.SetValue(ControlType.TextBox, shortComment);
                WaitForLoad();
            }

            if (comment != null)
            {
                _comment = WaitForElementIsVisible(By.Id(COMMENT));
                _comment.SetValue(ControlType.TextBox, comment);
                WaitForLoad();
            }

            if (sealNbre != null) {
                _sealNumber = WaitForElementIsVisible(By.Id(SEALS_NUMBER));
                _sealNumber.SetValue(ControlType.TextBox, sealNbre);
                WaitForLoad();
            }

            if (labelpage != null) {
                _labelPage = WaitForElementIsVisible(By.Id(LABEL_PAGE));
                _labelPage.SetValue(ControlType.TextBox, labelpage);
                WaitForLoad();
            }

            if (eqpNamelabel != null)
            {
                _eqpNamelabel = WaitForElementIsVisible(By.XPath(EQP_NAME_LABEL));
                _eqpNamelabel.SetValue(ControlType.TextBox, eqpNamelabel);
                WaitForLoad();
            }

            if (cartColor != null)
            {
                _cartColor = WaitForElementIsVisible(By.XPath(CART_COLOR));
                _cartColor.SetValue(ControlType.TextBox, cartColor);
                WaitForLoad();
            }

            if (mainPos != null) {
                _mainPosition = WaitForElementIsVisible(By.XPath(MAIN_POSITION));
                _mainPosition.SetValue(ControlType.TextBox, mainPos);
                WaitForLoad();
            }

            _rows = WaitForElementIsVisible(By.Id(ROWS));
            _rows.SetValue(ControlType.TextBox, rows);
            WaitForLoad();

            _column = WaitForElementIsVisible(By.Id(COLUMN));
            _column.SetValue(ControlType.TextBox, columns);
            WaitForLoad();

            ClickGenerate();
            ClickValidate();

            //Add Values in trolley géneré 
            if (string.IsNullOrEmpty(title))
            {
                AddValueInTrolley("TROLLEY");
            }
            else
            {
                AddValueInTrolley(title);
            }

            ClickFrontBtn();
            ClickConfirm();

        }

        public void CreateLpCartschemeWithoutDrawes(string title)
        {

            _title = WaitForElementIsVisible(By.Id(TITLE));
            _title.SetValue(ControlType.TextBox, title);
            WaitForLoad();

            //Add Values in trolley géneré 
            if (string.IsNullOrEmpty(title))
            {
                AddValueInTrolley("TROLLEY");
            }
            else
            {
                AddValueInTrolley(title);
            }
            ClickConfirm();

        }



        public void AddValueInTrolley(string value)
        {
            //Set Front and Back Btn
            //_Front_Btn = WaitForElementIsVisible(By.Id(FRONT_BTN));
            //_Front_Btn.Click();
            //WaitForLoad();

            //var _Back_Btn = WaitForElementIsVisible(By.Id(BACK_BTN));
            //_Back_Btn.Click();
            //WaitForLoad();

            var inputs = _webDriver.FindElements(By.XPath(INPUT_SCHEMA_TROLLEY)).ToList<IWebElement>();

            foreach (var input in inputs)
            {
                input.Clear();
                input.SendKeys(value);
                WaitForLoad();
            }
        }
        public void ClickFrontBtn()
        {
            _front_Btn = WaitForElementIsVisible(By.Id(FRONT_BTN));
            _front_Btn.Click();
            WaitPageLoading();
            WaitForLoad();

        }
    }
}
