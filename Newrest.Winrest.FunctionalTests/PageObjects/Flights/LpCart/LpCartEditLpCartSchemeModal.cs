using DocumentFormat.OpenXml.Bibliography;
using iText.StyledXmlParser.Jsoup.Nodes;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.PageObjects;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Flights.LpCart
{
    public class LpCartEditLpCartSchemeModal : PageBase
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
        private const string INPUT_SCHEMA_TROLLEY = "/html/body/div[4]/div/div/form/div[2]/div[10]/table[2]/tbody/tr[*]/td[*]/input";
        private const string FUSION_ON_BTN = "fusionOn";
        private const string FUSION_BTN = "fusionOn";
        private const string FUSE_BTN = "fuseBtn";
        private const string UNFUSE_BTN = "unfuseBtn";
        private const string INPUTS = "//*[@id=\"schemeTable\"]/tbody/tr[2]/td[*]";
        private const string DRAWER = "//*[@id=\"schemeTable\"]/tbody/tr[*]/td[@class = 'scheme-details-td']";
        private const string CHOOSE_LABEL_COLORS = "/html/body/div[4]/div/div/form/div[3]/span/button[2]";
        private const string SELECTED_COLORS = "/html/body/div[4]/div/div/form/div[2]/div[12]/table[2]/tbody/tr[*]/td[*]/span[1]/select/option[@selected=\"selected\"]";
        private const string CONFIRM_BTN = "create_btn";
        private const string FIRST_COLOR = "//*[@id=\"COLOR-C0R0\"]";
        private const string SECOND_COLOR = "//*[@id=\"COLOR-C1R0\"]";
        private const string THIRD_COLOR = "//*[@id=\"COLOR-C0R1\"]";
        private const string FOURTH_COLOR = "//*[@id=\"COLOR-C1R1\"]";
        private const string FIRST_POSITION = "POS-C0R0";
        private const string FIRST_PP = "PP-C0R0";
        private const string OPEN_DETAILS_BTN = "//*[@id=\"TD-C0R0\"]/div";
        private const string DETAILS_TEXT_AREA = "//*[@id=\"editPositionDetailForm\"]/div[2]/div[1]/div/div[2]/div/div[3]/div[2]/p";
        private const string CONFIRM_DETAILS_BTN = "last";
        private const string POSITIONS_DETAILS_TEXT = "//*[@id=\"TD-C0R0\"]/div/p";
        private const string CONFIRM_LABEL_POPUP = "dataConfirmLabel";
        private const string POSITION = "printPositions";
        private const string EQRNAMELABEL = "//*[@id=\"EquipmentNameLabel\"]";
        private const string CARTCOLOR = "//*[@id=\"trolley-label-color\"]";
        private const string LINECOUNT = "//*[@id=\"LPCartDetailsTable\"]/tbody/tr[2]/td[9]";
        private const string COLUMNCOUNT = "//*[@id=\"LPCartDetailsTable\"]/tbody/tr[2]/td[10]";
        private const string WEIGHT = "Weight";
        private const string OLDPOSITION = "OldPosition";
        private const string TEXTCLICK = "//*[@id=\"TD-C0R0\"]/div";
        private const string NOTE_EDITABLE = "//*[@id=\"editPositionDetailForm\"]/div[2]/div[1]/div[4]/div[2]/div/div[3]/div[2]/p";
        private const string MODAL_FORM = "//*[@id=\"editPositionDetailForm\"]";
        private const string CLOSE_FORM = "//*[@id=\"report-form\"]";
        private const string DESACTIVATE_FUSION = "fusionOff";
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
        private IWebElement _Front_Btn;

        [FindsBy(How = How.Id, Using = FUSION_ON_BTN)]
        private IWebElement _fusionOn_Btn;

        [FindsBy(How = How.Id, Using = FUSION_BTN)]
        private IWebElement _fusion_Btn;

        [FindsBy(How = How.Id, Using = FUSE_BTN)]
        private IWebElement _fuse_Btn;

        [FindsBy(How = How.Id, Using = UNFUSE_BTN)]
        private IWebElement _unfuse_Btn;

        [FindsBy(How = How.Id, Using = POSITION)]
        private IWebElement _position;

        [FindsBy(How = How.XPath, Using = EQRNAMELABEL)]
        private IWebElement _eqrnamelabel;

        [FindsBy(How = How.XPath, Using = CARTCOLOR)]
        private IWebElement _cart_color;

        [FindsBy(How = How.XPath, Using = LINECOUNT)]
        private IWebElement _line_count;

        [FindsBy(How = How.XPath, Using = COLUMNCOUNT)]
        private IWebElement _column_count;

        [FindsBy(How = How.XPath, Using = TEXTCLICK)]
        private IWebElement _text_click;

        [FindsBy(How = How.XPath, Using = NOTE_EDITABLE)]
        private IWebElement _note_editable;
        [FindsBy(How = How.XPath, Using = MODAL_FORM)]
        private IWebElement _modalForm;
        [FindsBy(How = How.XPath, Using = CLOSE_FORM)]
        private IWebElement _closeForm;

        // ___________________________________  Méthodes ___________________________________________
        public LpCartEditLpCartSchemeModal(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        public void ClickGenerate()
        {
            _generate = WaitForElementIsVisible(By.Id(GENERATE));
            _generate.Click();
            WaitForLoad();

        }

        public void ClickValidate()
        {
            _validate = WaitForElementIsVisible(By.Id(VALIDATE));
            _validate.Click();
            WaitForLoad();

        }

        public void ClickConfirm()
        {
            var javaScriptExecutor = _webDriver as IJavaScriptExecutor;
            _confirm = WaitForElementIsVisible(By.Id(CONFIRM));
            javaScriptExecutor.ExecuteScript("arguments[0].scrollIntoView(true);", _confirm);
            _confirm.Click();
            WaitForLoad();

        }

        public void EditLpCartscheme(string value, string sealsNum, string LabelPageNb, string rows, string columns, string colorCart)
        {

            _title = WaitForElementIsVisible(By.Id(TITLE));
            _title.SetValue(ControlType.TextBox, value);

            _prepZone = WaitForElementIsVisible(By.Id(PREP_ZONE));
            _prepZone.SetValue(ControlType.TextBox, value);

            _comment = WaitForElementIsVisible(By.Id(COMMENT));
            _comment.SetValue(ControlType.TextBox, value);

            _sealNumber = WaitForElementIsVisible(By.Id(SEALS_NUMBER));
            _sealNumber.SendKeys(sealsNum);

            _labelPage = WaitForElementIsVisible(By.Id(LABEL_PAGE));
            _labelPage.SetValue(ControlType.TextBox, LabelPageNb);

            _rows = WaitForElementIsVisible(By.Id(ROWS));
            _rows.SetValue(ControlType.TextBox, rows);

            _column = WaitForElementIsVisible(By.Id(COLUMN));
            _column.SetValue(ControlType.TextBox, columns);

            _position = WaitForElementIsVisible(By.Id(POSITION)); 
            _position.SetValue(ControlType.CheckBox, true);

            _eqrnamelabel = WaitForElementExists(By.XPath(EQRNAMELABEL));
            _eqrnamelabel.SetValue(ControlType.TextBox,value);

            _cart_color = WaitForElementExists(By.XPath(CARTCOLOR));
            _cart_color.SetValue(ControlType.TextBox, colorCart);

            WaitForLoad();

            ClickGenerate();
            ClickValidate();

            for (int x = 0; x < int.Parse(rows); x++)
            {
                //POS-C0R0 //PP-C0R0 //POS-C1R0 //PP-C1R0
                //POS-C0R1 //PP-C0R1 //POS-C1R1 //PP-C1R1

                for (int y = 0; y < int.Parse(columns); y++)
                {
                    var inputPOS = WaitForElementIsVisible(By.Id("POS-C" + y + "R" + x));
                    inputPOS.SetValue(ControlType.TextBox, "A" + y + x);

                    var inputPP = WaitForElementIsVisible(By.Id("PP-C" + y + "R" + x));
                    inputPP.SetValue(ControlType.TextBox, "B" + y + x);
                }
            }

			ClickConfirm();

        }

        public void EditLpCartschemeWithoutConfirm(string value, string sealsNum, string LabelPageNb, string rows, string columns, string colorCart)
        {

            _title = WaitForElementIsVisible(By.Id(TITLE));
            _title.SetValue(ControlType.TextBox, value);

            _prepZone = WaitForElementIsVisible(By.Id(PREP_ZONE));
            _prepZone.SetValue(ControlType.TextBox, value);

            _comment = WaitForElementIsVisible(By.Id(COMMENT));
            _comment.SetValue(ControlType.TextBox, value);

            _sealNumber = WaitForElementIsVisible(By.Id(SEALS_NUMBER));
            _sealNumber.SendKeys(sealsNum);

            _labelPage = WaitForElementIsVisible(By.Id(LABEL_PAGE));
            _labelPage.SetValue(ControlType.TextBox, LabelPageNb);

            _rows = WaitForElementIsVisible(By.Id(ROWS));
            _rows.SetValue(ControlType.TextBox, rows);

            _column = WaitForElementIsVisible(By.Id(COLUMN));
            _column.SetValue(ControlType.TextBox, columns);

            _position = WaitForElementIsVisible(By.Id(POSITION));
            _position.SetValue(ControlType.CheckBox, true);

            _eqrnamelabel = WaitForElementExists(By.XPath(EQRNAMELABEL));
            _eqrnamelabel.SetValue(ControlType.TextBox, value);

            _cart_color = WaitForElementExists(By.XPath(CARTCOLOR));
            _cart_color.SetValue(ControlType.TextBox, colorCart);

            WaitForLoad();

            ClickGenerate();
            ClickValidate();

            for (int x = 0; x < int.Parse(rows); x++)
            {
                //POS-C0R0 //PP-C0R0 //POS-C1R0 //PP-C1R0
                //POS-C0R1 //PP-C0R1 //POS-C1R1 //PP-C1R1

                for (int y = 0; y < int.Parse(columns); y++)
                {
                    var inputPOS = WaitForElementIsVisible(By.Id("POS-C" + y + "R" + x));
                    inputPOS.SetValue(ControlType.TextBox, "A" + y + x);

                    var inputPP = WaitForElementIsVisible(By.Id("PP-C" + y + "R" + x));
                    inputPP.SetValue(ControlType.TextBox, "B" + y + x);
                }
            }

        }



        public List<string> GetLpCartschemeValues()
        {
            List<string> SchemeValues = new List<string>();

            _title = WaitForElementIsVisible(By.Id(TITLE));
            SchemeValues.Add(_title.GetAttribute("value"));

            _prepZone = WaitForElementIsVisible(By.Id(PREP_ZONE));
            SchemeValues.Add(_prepZone.GetAttribute("value"));

            _comment = WaitForElementIsVisible(By.Id(COMMENT));
            SchemeValues.Add(_comment.GetAttribute("value"));

            _sealNumber = WaitForElementIsVisible(By.Id(SEALS_NUMBER));
            SchemeValues.Add(_sealNumber.GetAttribute("value"));

            _labelPage = WaitForElementIsVisible(By.Id(LABEL_PAGE));
            SchemeValues.Add(_labelPage.GetAttribute("value"));

            _rows = WaitForElementIsVisible(By.Id(ROWS));
            SchemeValues.Add(_rows.GetAttribute("value"));

            _column = WaitForElementIsVisible(By.Id(COLUMN));
            SchemeValues.Add(_column.GetAttribute("value"));


            WaitForLoad();

            ClickGenerate();
            ClickValidate();

            ClickConfirm();

            var result = SchemeValues.Distinct().ToList();
            return result;
        }


        public void AddValueInTrolley(string value)
        {
            //Set Front and Back Btn
            _Front_Btn = WaitForElementIsVisible(By.Id(FRONT_BTN));
            _Front_Btn.Click();

            var inputs = _webDriver.FindElements(By.XPath(INPUT_SCHEMA_TROLLEY));

            foreach (var input in inputs)
            {
                input.SetValue(ControlType.TextBox, value);
            }
        }

        public void ActivateFusion()
        {
            //Activate fusion
            _fusionOn_Btn = WaitForElementIsVisible(By.Id(FUSION_ON_BTN));
            _fusionOn_Btn.Click();

            WaitForLoad();

        }

        public void SelectFusionElement()
        {
            // Liste des id des inputs que vous souhaitez cliquer
            List<string> listId = new List<string> { "POS-C0R0", "PP-C1R0", "PP-C0R1", "PP-C1R1" };
            // cliquer sur input de chaque drawer pour rendre selectionner pour la fusion
            ReadOnlyCollection<IWebElement> elements = _webDriver.FindElements(By.XPath("//*/td[@class='scheme-details-td']"));
            foreach (var element in elements)
            {
                ReadOnlyCollection<IWebElement> inputElements = element.FindElements(By.XPath(".//input"));

                // Parcourir chaque input trouvé
                foreach (var input in inputElements)
                {
                    // Vérifier si l'id de l'input est dans la liste
                    string inputId = input.GetAttribute("id");
                    if (listId.Contains(inputId))
                    {
                        // Cliquer sur l'input correspondant
                        input.Click();
                        break;  // Sortir du loop après avoir cliqué sur le premier input correspondant
                    }
                }
            }
        }

        public int GetDrawerCount()
        {
            ReadOnlyCollection<IWebElement> drawer;
            drawer = _webDriver.FindElements(By.XPath("//*/div[@title='Details...' or @title='Details ...']"));

            return drawer.Count;
        }

        public void ClickFusion()
        {

            _fuse_Btn = WaitForElementIsVisible(By.Id(FUSE_BTN));
            _fuse_Btn.Click();

            WaitForLoad();

            ClickConfirm();
        }

        public void ClickDesactivateFusion()
        {

            IWebElement deasctivateFusionButton = WaitForElementIsVisible(By.Id(DESACTIVATE_FUSION));
            deasctivateFusionButton.Click();

            WaitForLoad();
        }

        public void ClickFusionWithoutConfirm()
        {

            _fuse_Btn = WaitForElementIsVisible(By.Id(FUSE_BTN));
            _fuse_Btn.Click();

            WaitForLoad();
        }

        public void ClickUndoFusion()
        {

            _unfuse_Btn = WaitForElementIsVisible(By.Id(UNFUSE_BTN));
            _unfuse_Btn.Click();

            WaitForLoad();

            ClickConfirm();
        }

        public void ClickUndoFusionWithoutConfirm()
        {

            _unfuse_Btn = WaitForElementIsVisible(By.Id(UNFUSE_BTN));
            _unfuse_Btn.Click();

            WaitForLoad();
        }

        public string GetValueInTrolley()
        {
            //Set Front and Back Btn
            List<string> TrolleyValues = new List<string>();
            _Front_Btn = WaitForElementIsVisible(By.Id(FRONT_BTN));
            _Front_Btn.Click();

            var inputs = _webDriver.FindElements(By.XPath(INPUT_SCHEMA_TROLLEY));

            foreach (var input in inputs)
            {
                TrolleyValues.Add(input.Text);
            }
          
            var result = TrolleyValues.Distinct().FirstOrDefault();
            return result;
        }

        public void ClickChooseLabelColors()
        {
            var btn = WaitForElementIsVisible(By.XPath(CHOOSE_LABEL_COLORS));
            btn.Click();
        }
        public IEnumerable<object> GetColors()
        {
            var listObj = new List<object>();
            var list =   _webDriver.FindElements(By.XPath(SELECTED_COLORS)).Select(e=>e.Text);
            foreach (var item in list)
            {
                var itemOgj = (object)item;
                listObj.Add(itemOgj);

            }
            var btnConfirm = WaitForElementIsVisible(By.Id(CONFIRM_BTN));
            btnConfirm.Click();
            WaitForLoad();
            return listObj;
        }
        public void GenerateRowsColumns(int rows, int columns)
        {
            var rowsInput = WaitForElementIsVisible(By.Id("rowsNumber"));
            rowsInput.Clear();
            rowsInput.SendKeys(rows.ToString());
            var columnsInput = WaitForElementIsVisible(By.Id("colsNumber"));
            columnsInput.Clear();
            columnsInput.SendKeys(columns.ToString());
            ClickGenerate();
            ClickValidate();
        }
        public void ChangeColors(string color1 , string color2 , string color3 , string color4)
        {
            var firstColorSelect = WaitForElementIsVisible(By.XPath(FIRST_COLOR));
            firstColorSelect.SetValue(ControlType.DropDownList, color1);

            var secondColorSelect = WaitForElementIsVisible(By.XPath(SECOND_COLOR));
            secondColorSelect.SetValue(ControlType.DropDownList, color2);

            var thirdColorSelect = WaitForElementIsVisible(By.XPath(THIRD_COLOR));
            thirdColorSelect.SetValue(ControlType.DropDownList, color3);

            var fourthColorSelect = WaitForElementIsVisible(By.XPath(FOURTH_COLOR));
            fourthColorSelect.SetValue(ControlType.DropDownList, color4);
        }

        public void EditFirstPosition(string position , string pandp , string details)
        {
            var positionInput = WaitForElementIsVisible(By.Id(FIRST_POSITION));
            positionInput.Clear();
            positionInput.SendKeys(position);
            var pandpInput = WaitForElementIsVisible(By.Id(FIRST_PP));
            pandpInput.Clear();
            pandpInput.SendKeys(pandp);
            var detail = WaitForElementIsVisible(By.XPath(OPEN_DETAILS_BTN));
            detail.Click();
            var textDetails = WaitForElementIsVisible(By.XPath(DETAILS_TEXT_AREA));
            //select all 
            textDetails.Click();
            Actions actions = new Actions(_webDriver);
            actions.KeyDown(Keys.Control).SendKeys("a").KeyUp(Keys.Control).Build().Perform();
            
            textDetails.SendKeys(details);
            var confirm = WaitForElementIsVisible(By.Id(CONFIRM_DETAILS_BTN));
            confirm.Click();
            WaitForLoad();
            var confirmAllBtn = WaitForElementIsVisible(By.Id(CONFIRM_BTN));
            confirmAllBtn.Click();
            WaitForLoad();

            //if popup confirm your action
            if (isElementVisible(By.Id(CONFIRM_LABEL_POPUP)))
            {
                _validate = WaitForElementIsVisible(By.Id(VALIDATE));
                _validate.Click();
                WaitForLoad();

            }
        }
        public bool VerifyPosition(string position ,string pnp , string positionDetail)
        {
            var positionValue = WaitForElementIsVisible(By.Id(FIRST_POSITION)).GetAttribute("value");
            var pnpValue = WaitForElementIsVisible(By.Id(FIRST_PP)).GetAttribute("value");
            var detailsValue = WaitForElementIsVisible(By.XPath(POSITIONS_DETAILS_TEXT)).Text;
            var confirmBtn = WaitForElementIsVisible(By.Id(CONFIRM));

            if(position.Equals(positionValue) && pnp.Equals(pnpValue) && positionDetail.Equals(detailsValue))
            {
                confirmBtn.Click();
                WaitForLoad();
                return true;
            }
            confirmBtn.Click();
            WaitForLoad();
            return false;
        }

        public bool VerifyCountLineColumn(string row, string col)
        {
            _line_count = WaitForElementIsVisible(By.XPath(LINECOUNT));
            var line_count = _line_count.Text;

            _column_count = WaitForElementIsVisible(By.XPath(COLUMNCOUNT));
            var column_count = _column_count.Text;
         
            if (row == line_count && col == column_count)
            {
                return true;
            }
            return false;
        }
        public void EditWeightPosition(string position, string pandp, string details, string weightValue, string oldPositionValue)
        {
            var positionInput = WaitForElementIsVisible(By.Id(FIRST_POSITION));
            positionInput.Clear();
            positionInput.SendKeys(position);
            var pandpInput = WaitForElementIsVisible(By.Id(FIRST_PP));
            pandpInput.Clear();
            pandpInput.SendKeys(pandp);
            var detail = WaitForElementIsVisible(By.XPath(OPEN_DETAILS_BTN));
            detail.Click();
            var weightInput = WaitForElementIsVisible(By.Id(WEIGHT));
            //select all 
            weightInput.SetValue(ControlType.TextBox, weightValue);
            Actions actions = new Actions(_webDriver);
            var oldPositionInput = WaitForElementIsVisible(By.Id(OLDPOSITION));
            //select all 
            oldPositionInput.SetValue(ControlType.TextBox, oldPositionValue);

        }

        public bool VerifyWeightPosition(string position, string pandp, string details, string weightValue, string oldPositionValue)
        {
            var positionInput = WaitForElementIsVisible(By.Id(FIRST_POSITION));
            positionInput.Clear();
            positionInput.SendKeys(position);
            var pandpInput = WaitForElementIsVisible(By.Id(FIRST_PP));
            pandpInput.Clear();
            pandpInput.SendKeys(pandp);
            var detail = WaitForElementIsVisible(By.XPath(OPEN_DETAILS_BTN));
            detail.Click();
            var weightInput = WaitForElementIsVisible(By.Id(WEIGHT));
            //select all 
            var weight=weightInput.GetAttribute("value");
            Actions actions = new Actions(_webDriver);
            var oldPositionInput = WaitForElementIsVisible(By.Id(OLDPOSITION));
            //select all 
            var positionX= oldPositionInput.GetAttribute("value");

            return positionX.Equals(oldPositionValue) && weight.Equals(weightValue);
        }

        public void DisplayEditPSDetails()
        {
            _text_click = WaitForElementIsVisible(By.XPath(TEXTCLICK));
            _text_click.Click();

            WaitForLoad();
        }

        public bool FieldsPSDetailsIsEmpty()
        {
            WaitPageLoading();
            DisplayEditPSDetails();
            var previousPosition = WaitForElementIsVisible(By.Id(OLDPOSITION));
            var oldPosition = previousPosition.GetAttribute("value");
            var weightInput = WaitForElementIsVisible(By.Id(WEIGHT));
            var weight=weightInput.GetAttribute("value");
            var noteEditable = WaitForElementIsVisible(By.XPath(NOTE_EDITABLE));

            return string.IsNullOrEmpty(oldPosition) && string.IsNullOrEmpty(weight) && string.IsNullOrEmpty(noteEditable.Text);
        }


        public List<string> Insert_Postion_PP_Line(string rows, string columns)
        {
            List<string> SchemeValues = new List<string>();

            WaitForLoad();

            for (int x = 0; x < int.Parse(rows); x++)
            {
                //POS-C0R0 //PP-C0R0 //POS-C1R0 //PP-C1R0
                //POS-C0R1 //PP-C0R1 //POS-C1R1 //PP-C1R1

                for (int y = 0; y < int.Parse(columns); y++)
                {
                    var inputPOS = WaitForElementIsVisible(By.Id("POS-C" + y + "R" + x));
                    inputPOS.SetValue(ControlType.TextBox, "A" + y + x);                    
                    SchemeValues.Add(inputPOS.GetAttribute("value"));

                    var inputPP = WaitForElementIsVisible(By.Id("PP-C" + y + "R" + x));
                    // mêmes valeurs pour la seconde column
                    inputPP.SetValue(ControlType.TextBox, "B" + "0" + "0");
                    SchemeValues.Add(inputPP.GetAttribute("value"));
                }
            }            

            ClickConfirm();

            return SchemeValues;
        }

        public List<string> ValueOf_Postion_PP_Line_After_Insert()
        {
            List<string> SchemeValues = new List<string>();

            _rows = WaitForElementIsVisible(By.Id(ROWS));

            var rows = _rows.GetAttribute("value");

            _column = WaitForElementIsVisible(By.Id(COLUMN));

            var columns = _column.GetAttribute("value");
           
            WaitForLoad();

            for (int x = 0; x < int.Parse(rows); x++)
            {
                for (int y = 0; y < int.Parse(columns); y++)
                {
                    //POS-C0R0 //PP-C0R0 //POS-C1R0 //PP-C1R0
                    //POS-C0R1 //PP-C0R1 //POS-C1R1 //PP-C1R1
                    var inputPOS = WaitForElementIsVisible(By.Id("POS-C" + y + "R" + x));
                    SchemeValues.Add(inputPOS.GetAttribute("A" + x + y));

                    var inputPP = WaitForElementIsVisible(By.Id("PP-C" + y + "R" + x));
                    SchemeValues.Add(inputPP.GetAttribute("A" + x + y));

                }
            }

            return SchemeValues;

        }


        public List<string> List_Of_Postion_PP_Line_After_Insert()
        {
            List<string> SchemeValues = new List<string>();

            _rows = WaitForElementIsVisible(By.Id(ROWS));

            var rows = _rows.GetAttribute("value");

            _column = WaitForElementIsVisible(By.Id(COLUMN));

            var columns = _column.GetAttribute("value");

            WaitForLoad();

            for (int x = 0; x < int.Parse(rows); x++)
            {
                //POS-C0R0 //PP-C0R0 //POS-C1R0 //PP-C1R0
                //POS-C0R1 //PP-C0R1 //POS-C1R1 //PP-C1R1

                for (int y = 0; y < int.Parse(columns); y++)
                {
                    var inputPOS = WaitForElementIsVisible(By.Id("POS-C" + y + "R" + x));
                    SchemeValues.Add(inputPOS.GetAttribute("value"));

                    var inputPP = WaitForElementIsVisible(By.Id("PP-C" + y + "R" + x));
                    SchemeValues.Add(inputPP.GetAttribute("value"));
                }
            }

            return SchemeValues;

        }
        public bool ColorsTableIsVisible()
        {
            var isVisible = isElementVisible(By.Id("TD-C0R0"));
            return isVisible;
        }
        public bool IsEditPositionDetailsPopupDisplayed()
        {
            _modalForm = _webDriver.FindElement(By.XPath(MODAL_FORM));
            return _modalForm.Displayed;
        }

        public void ClickOnAddLpCartSheme()
        {
            WaitForLoad();
            var confirmAllBtn = WaitForElementIsVisible(By.Id(CONFIRM_BTN));
            confirmAllBtn.Click();
            WaitForLoad();
        }
        public void ClickOnSavePositionDetails()
        {
            WaitForLoad();
            var confirm = WaitForElementIsVisible(By.Id(CONFIRM_DETAILS_BTN));
            confirm.Click();
        }
        public bool IsEditPositionPopupClosed()
        {
            _closeForm = _webDriver.FindElement(By.XPath(CLOSE_FORM));
            return _closeForm.Displayed;
        }

        public void EditDrawInputs(string value)
        {
            // Liste des IDs des inputs 
            List<string> listId = new List<string> { "POS-C0R0", "PP-C0R0" };
            // Gett the  all inputs
            IWebElement tdElement = _webDriver.FindElement(By.XPath("//*[@id='schemeTable']/tbody/tr[2]/td"));
            ReadOnlyCollection<IWebElement> inputElements = tdElement.FindElements(By.XPath(".//input"));
           // modifier  les input de chaque drawer 
            foreach (var input in inputElements)
            {
                // Récupérer l'id de chaque input
                string inputId = input.GetAttribute("id");

                // Vérifier si l'id est dans la liste
                if (listId.Contains(inputId))
                {
                    // Cliquer ou tabuler dans l'input
                    input.SetValue(ControlType.TextBox, value);
                    WaitForLoad();
                }
            }
        }

        public bool IsDrawersEdited(string value)
        {
            bool isEdited = true;
            // Liste des IDs des inputs 
            List<string> listId = new List<string> { "POS-C0R0", "PP-C0R0" };

            // Gett the  all inputs
            IWebElement tdElement = _webDriver.FindElement(By.XPath("//*[@id='schemeTable']/tbody/tr[2]/td"));
            ReadOnlyCollection<IWebElement> inputElements = tdElement.FindElements(By.XPath(".//input"));

            //verifier si les input de drawer sont modfier
            foreach (var input in inputElements)
            {
                // Récupérer l'id de chaque input
                string inputId = input.GetAttribute("id");

                // Vérifier si l'id est dans la liste
                if (listId.Contains(inputId))
                {
                    // check if the input draw is edited
                     string draw = input.GetAttribute("value");
                     if(draw != value)
                    {
                        isEdited =  false;
                    }
                }
            }
            return isEdited;
        }
    }
}
