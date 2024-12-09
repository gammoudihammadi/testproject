using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.Office2010.ExcelAc;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.PageObjects;
using System.Collections.Generic;
using System.Linq;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Admin
{
    public class HACCPDocumentSettingsPage : PageBase
    {
        public HACCPDocumentSettingsPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        //____________________________________________________Constantes_____________________________________________________________

        // HACCP Documents List
        private const string APP_SETTINGS_HACCP_DOCUMENT = "//*[@id=\"div-body\"]/div/div[1]/div/table/tbody/tr[*]/td[1][contains(text(),'{0}')]";
        private const string APP_SETTINGS_HACCP_ADD_DOCUMENT_BUTTON = "//*[@id=\"div-body\"]/div/div[1]/a";

        //HACCP Create new document
        private const string APP_SETTINGS_HACCP_NEW_DOCUMENT_TITLE = "Title";
        private const string APP_SETTINGS_HACCP_NEW_DOCUMENT_ORDER = "Order";
        private const string APP_SETTINGS_HACCP_NEW_DOCUMENT_FOOTER = "Footer";
        private const string APP_SETTINGS_HACCP_NEW_DOCUMENT_ISAVAILABLE = "IsAvailable";
        private const string APP_SETTINGS_HACCP_NEW_DOCUMENT_CREATE_BUTTON = "//*[@id=\"modal-1\"]/div/div/div[2]/div/form/div[2]/button[2]";

        //HACCP Document Column
        private const string APP_SETTINGS_HACCP_DOCUMENT_COLUMN = "//*[@id=\"document-details-panel\"]/table/tbody/tr[*]/td[1][contains(text(),'{0}')]/../td[3][contains(text(),'{1}')]";
        private const string APP_SETTINGS_HACCP_DOCUMENT_ADD_COLUMN_BUTTON = "//*[@id=\"document-details-panel\"]/div/a";
        private const string APP_SETTINGS_HACCP_DOCUMENT_NEW_COLUMN_TITLE = "Title";
        private const string APP_SETTINGS_HACCP_DOCUMENT_NEW_COLUMN_ORDER = "Order";
        private const string APP_SETTINGS_HACCP_DOCUMENT_NEW_COLUMN_SIZE = "Size";
        private const string APP_SETTINGS_HACCP_DOCUMENT_NEW_COLUMN_TYPE = "Type";
        private const string APP_SETTINGS_HACCP_NEW_COLUMN_CREATE_BUTTON = "//*[@id=\"create-column-form\"]/div[2]/button[2]";

        //____________________________________________________Variables______________________________________________________________

        // HACCP Documents List

        [FindsBy(How = How.XPath, Using = APP_SETTINGS_HACCP_ADD_DOCUMENT_BUTTON)]
        private IWebElement _addDocument;

        //HACCP Create new document
        [FindsBy(How = How.Id, Using = APP_SETTINGS_HACCP_NEW_DOCUMENT_TITLE)]
        private IWebElement _documentTitle;

        [FindsBy(How = How.Id, Using = APP_SETTINGS_HACCP_NEW_DOCUMENT_ORDER)]
        private IWebElement _documentOrder;

        [FindsBy(How = How.Id, Using = APP_SETTINGS_HACCP_NEW_DOCUMENT_FOOTER)]
        private IWebElement _documentFooter;

        [FindsBy(How = How.Id, Using = APP_SETTINGS_HACCP_NEW_DOCUMENT_ISAVAILABLE)]
        private IWebElement _documentIsAvailable;

        [FindsBy(How = How.Id, Using = APP_SETTINGS_HACCP_NEW_DOCUMENT_CREATE_BUTTON)]
        private IWebElement _createDocument;

        //HACCP Doucment Column

        [FindsBy(How = How.XPath, Using = APP_SETTINGS_HACCP_DOCUMENT_ADD_COLUMN_BUTTON)]
        private IWebElement _addColumn;

        [FindsBy(How = How.Id, Using = APP_SETTINGS_HACCP_DOCUMENT_NEW_COLUMN_TITLE)]
        private IWebElement _columnTitle;

        [FindsBy(How = How.Id, Using = APP_SETTINGS_HACCP_DOCUMENT_NEW_COLUMN_ORDER)]
        private IWebElement _columnOrder;

        [FindsBy(How = How.Id, Using = APP_SETTINGS_HACCP_DOCUMENT_NEW_COLUMN_SIZE)]
        private IWebElement _columnSize;

        [FindsBy(How = How.Id, Using = APP_SETTINGS_HACCP_DOCUMENT_NEW_COLUMN_TYPE)]
        private IWebElement _columnType;

        [FindsBy(How = How.XPath, Using = APP_SETTINGS_HACCP_NEW_COLUMN_CREATE_BUTTON)]
        private IWebElement _createColumn;

        //____________________________________________________Méthodes_______________________________________________________________
        public bool isDocumentExist(string documentName)
        {
            if (isElementVisible(By.XPath(string.Format(APP_SETTINGS_HACCP_DOCUMENT, documentName))))
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        public void addDocument(string documentName, string order, string footer, bool isAvailbale)
        {
            //check if document exist
            if (!isDocumentExist(documentName))
            {
                //open Create new document modal
                _addDocument = WaitForElementIsVisible(By.XPath(APP_SETTINGS_HACCP_ADD_DOCUMENT_BUTTON), nameof(APP_SETTINGS_HACCP_ADD_DOCUMENT_BUTTON));
                _addDocument.Click();
                WaitForLoad();

                //add Title
                _documentTitle = WaitForElementIsVisible(By.Id(APP_SETTINGS_HACCP_NEW_DOCUMENT_TITLE), nameof(APP_SETTINGS_HACCP_NEW_DOCUMENT_TITLE));
                _documentTitle.SetValue(ControlType.TextBox, documentName);

                //add Order
                _documentOrder = WaitForElementIsVisible(By.Id(APP_SETTINGS_HACCP_NEW_DOCUMENT_ORDER), nameof(APP_SETTINGS_HACCP_NEW_DOCUMENT_ORDER));
                _documentOrder.SetValue(ControlType.TextBox, order);

                //add Footer
                _documentFooter = WaitForElementIsVisible(By.Id(APP_SETTINGS_HACCP_NEW_DOCUMENT_FOOTER), nameof(APP_SETTINGS_HACCP_NEW_DOCUMENT_FOOTER));
                _documentFooter.SetValue(ControlType.TextBox, footer);

                //is Available
                _documentIsAvailable = WaitForElementExists(By.Id(APP_SETTINGS_HACCP_NEW_DOCUMENT_ISAVAILABLE));
                _documentIsAvailable.SetValue(ControlType.CheckBox, isAvailbale);

                _createDocument = WaitForElementIsVisible(By.XPath(APP_SETTINGS_HACCP_NEW_DOCUMENT_CREATE_BUTTON), nameof(APP_SETTINGS_HACCP_NEW_DOCUMENT_CREATE_BUTTON));
                _createDocument.Click();
                WaitForLoad();
            }
        }

        public void selectDocument(string documentName)
        {
            var document = WaitForElementIsVisible(By.XPath(string.Format(APP_SETTINGS_HACCP_DOCUMENT, documentName)), nameof(APP_SETTINGS_HACCP_DOCUMENT));
            document.Click();
            WaitForLoad();
        }

        public bool isColumnExist(string columnName, string columnOrder)
        {
            if (isElementVisible(By.XPath(string.Format(APP_SETTINGS_HACCP_DOCUMENT_COLUMN, columnName, columnOrder))))
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        public void addColumn(string columnName, string columnOrder, string columnSize, string columnType, string subtitle = null, string info = null)
        {
            //check if column exist
            if (!isColumnExist(columnName, columnOrder))
            {
                //open Create new column modal
                Actions action = new Actions(_webDriver);
                _addColumn = WaitForElementIsVisible(By.XPath(APP_SETTINGS_HACCP_DOCUMENT_ADD_COLUMN_BUTTON), nameof(APP_SETTINGS_HACCP_DOCUMENT_ADD_COLUMN_BUTTON));
                action.MoveToElement(_addColumn).Perform();
                _addColumn.Click();
                WaitForLoad();

                //add Title
                _columnTitle = WaitForElementIsVisible(By.Id(APP_SETTINGS_HACCP_DOCUMENT_NEW_COLUMN_TITLE), nameof(APP_SETTINGS_HACCP_DOCUMENT_NEW_COLUMN_TITLE));
                _columnTitle.SetValue(ControlType.TextBox, columnName);

                //add Order
                _columnOrder = WaitForElementIsVisible(By.Id(APP_SETTINGS_HACCP_DOCUMENT_NEW_COLUMN_ORDER), nameof(APP_SETTINGS_HACCP_DOCUMENT_NEW_COLUMN_ORDER));
                _columnOrder.SetValue(ControlType.TextBox, columnOrder);

                //add Size
                _columnSize = WaitForElementIsVisible(By.Id(APP_SETTINGS_HACCP_DOCUMENT_NEW_COLUMN_SIZE), nameof(APP_SETTINGS_HACCP_DOCUMENT_NEW_COLUMN_SIZE));
                _columnSize.SetValue(ControlType.TextBox, columnSize);

                //add Type
                _columnType = WaitForElementIsVisible(By.Id(APP_SETTINGS_HACCP_DOCUMENT_NEW_COLUMN_TYPE), nameof(APP_SETTINGS_HACCP_DOCUMENT_NEW_COLUMN_TYPE));
                _columnType.SetValue(ControlType.DropDownList, columnType);

                _createColumn = WaitForElementIsVisible(By.XPath(APP_SETTINGS_HACCP_NEW_COLUMN_CREATE_BUTTON), nameof(APP_SETTINGS_HACCP_NEW_COLUMN_CREATE_BUTTON));
                _createColumn.Click();
                WaitForLoad();
            }
        }

        public void AddAllDocumentsAndAllColumns()
        {
            //HACCP3 Sanitization
            addDocument("HACCP3 Sanitization", "31", "TD-U-10.1 <br></br> Target for active chlorine = 100 ppm, with contact time = 5 minutes <br> </ br > -If the chlorine concentration is less than the dilution, disinfectant must be Added<br> </ br > -If the chlorine concentration exceeds the dilution, water must be added until the desired dilution is obtained<br> </ br > -If degraded products are observed during final visual inspection they must be discarded / scrapped<br> </ br >", true);
            selectDocument("HACCP3 Sanitization");
            addColumn("Product", "1", "2", "Text");
            addColumn("Quantity (kg)", "2", "1", "Decimal");
            addColumn("Disinfection", "3", "1", "Integer", "Qty (ppm/%)", "-100ppm (Con exposición por un tiempo de 5 mimutos / Recomendsado )");
            addColumn("Disinfection", "4", "1", "Integer", "Contact time (min)", "100ppm (Con exposición por un tiempo de 5 mimutos / Recomendsado )");
            addColumn("Rinsing", "5", "1", "CheckboxCNC", "", "enjuague con agua limpia");
            addColumn("Final Visual Inspection", "6", "1", "CheckboxCNC", "", "ningún objeto extraño");
            addColumn("Comments (incl. correction)", "7", "1", "Memo");
            addColumn("Corrective action", "8", "1", "Memo", "", "Si la concentración <50 ppm o> 100 ppm: prepare una nueva solución, comprobar la concentración, y proceder a la desinfección ");
            addColumn("Prepared by", "9", "1", "Signature");

            //HACCP3 Thawing
            addDocument("HACCP3 Thawing", "34", "TD-U-11.1 <br></br> Thawing at 4°C <br> </ br > -Meat 3 days<br> </ br > -Fish 2 days<br> </ br >", true);
            selectDocument("HACCP3 Thawing");
            addColumn("Product", "0", "1", "Text");
            addColumn("Exp. date", "1", "1", "Date");
            addColumn("Batch Nbr (optional)", "2", "1", "Text");
            addColumn("Quantity (kg)", "3", "1", "Decimal");
            addColumn("Thawing Starting Date", "4", "1", "Date", "", "Descongelar a 4 ° C Carne 3 días Pescado 2 días Solo en vuelo: descongelación a ≤8ºC durante 2 días para carne cruda descongelar a ≤8 ° C durante 1 día para el pescado crudo Descongelar RTE Congelado a <4 ° C durante 1 día / 24 horas");
            addColumn("Maximum End using", "5", "1", "Date", "", "Descongelar a 4 ° C Carne 3 días Pescado 2 días Solo en vuelo: descongelación a ≤8ºC durante 2 días para carne cruda descongelar a ≤8 ° C durante 1 día para el pescado crudo Descongelar RTE Congelado a <4 ° C durante 1 día / 24 horas ");
            addColumn("Comments (incl. correction)", "6", "1", "Memo");
            addColumn("Corrective action", "7", "1", "Memo", "", "Deseche los productos alimenticios potencialmente peligrosos si se excede la temperatura de la superficie o el tiempo de descongelación");
            addColumn("Prepared by", "8", "1", "Signature");

            //HACCP3 Test
            addDocument("HACCP3 Test", "20", "Document to test different column type", true);
            selectDocument("HACCP3 Test");
            addColumn("CheckboxCNA", "0", "1", "CheckboxCNA");
            addColumn("CheckboxCNC", "1", "1", "CheckboxCNC");
            addColumn("MultiDate", "2", "1", "MultiDate");

            //HACCP3 Modified texture
            addDocument("HACCP3 Modified texture", "26", "Document to test different column type", true);
            selectDocument("HACCP3 Modified texture");
            addColumn("Products", "0", "1", "Text");
            addColumn("Start of mixing", "1", "1", "Time", "T°");
            addColumn("Start of mixing", "2", "1", "Time", "H");
            addColumn("End of mixing", "3", "1", "Temperature", "T°");
            addColumn("End of mixing", "4", "1", "Temperature", "H");
            addColumn("Comments (incl. correction)", "5", "1", "Memo");
            addColumn("Corrective action", "6", "1", "Memo");
            addColumn("Prepared by", "7", "1", "Signature");

        }
        public bool IsHACCP_CONFIG() 
        {
            string documentHACCP3_Sanitization = "HACCP3 Sanitization";
            string documentHACCP3_Thawing = "HACCP3 Thawing";
            string documentHACCP3_Test = "HACCP3 Test";
            string documentHACCP3_Modified_Texture = "HACCP3 Modified texture";

            var documentsTitle = _webDriver.FindElements(By.XPath("/html/body/div[2]/div/div[1]/div/table/tbody/tr[*]/td[1]"));
      

            var result = documentsTitle.Select(c => c.Text).Contains(documentHACCP3_Sanitization)
                          && documentsTitle.Select(c => c.Text).Contains(documentHACCP3_Thawing)
                          && documentsTitle.Select(c => c.Text).Contains(documentHACCP3_Test)
                          && documentsTitle.Select(c => c.Text).Contains(documentHACCP3_Modified_Texture);
            
            return result;
        }
    }
}
