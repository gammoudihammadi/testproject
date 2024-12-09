using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Admin
{
    public class CustomizableColumnsPage : PageBase
    {
        public CustomizableColumnsPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        //____________________________________________________Constantes_____________________________________________________________

        private const string APP_SETTINGS_CUSTOMIZABLE_COLUMN = "//*[@id=\"div-body\"]/table/tbody/tr[*]/td[1][contains(text(),'{0}')]/../td[2][contains(text(),'{1}')]";
        private const string APP_SETTINGS_CUSTOMIZABLE_COLUMN_ADD_BUTTON = "btn-add-column";
        private const string APP_SETTINGS_NEW_CUSTOMIZABLE_COLUMN_NAME = "first";
        private const string APP_SETTINGS_NEW_CUSTOMIZABLE_COLUMN_CATEGORY = "CategoryId";
        private const string APP_SETTINGS_NEW_CUSTOMIZABLE_COLUMN_TYPE = "CustomizableColumnType";
        private const string APP_SETTINGS_NEW_CUSTOMIZABLE_COLUMN_SIZE = "ColumnSize";
        private const string APP_SETTINGS_NEW_CUSTOMIZABLE_COLUMN_ORDER = "Order";
        private const string APP_SETTINGS_NEW_CUSTOMIZABLE_COLUMN_INFO = "/html/body/div[3]/div/div/div/div/form/div[2]/div[7]/div/input";
        private const string APP_SETTINGS_NEW_CUSTOMIZABLE_COLUMN_SAVE_BUTTON = "last";


        //____________________________________________________Variables______________________________________________________________

        [FindsBy(How = How.Id, Using = APP_SETTINGS_CUSTOMIZABLE_COLUMN_ADD_BUTTON)]
        private IWebElement _addCustomizableColumn;

        [FindsBy(How = How.Id, Using = APP_SETTINGS_NEW_CUSTOMIZABLE_COLUMN_NAME)]
        private IWebElement _newCustomizableColumnName;

        [FindsBy(How = How.Id, Using = APP_SETTINGS_NEW_CUSTOMIZABLE_COLUMN_CATEGORY)]
        private IWebElement _newCustomizableColumnCategory;

        [FindsBy(How = How.Id, Using = APP_SETTINGS_NEW_CUSTOMIZABLE_COLUMN_TYPE)]
        private IWebElement _newCustomizableColumnType;

        [FindsBy(How = How.Id, Using = APP_SETTINGS_NEW_CUSTOMIZABLE_COLUMN_SIZE)]
        private IWebElement _newCustomizableColumnSize;

        [FindsBy(How = How.Id, Using = APP_SETTINGS_NEW_CUSTOMIZABLE_COLUMN_ORDER)]
        private IWebElement _newCustomizableColumnOrder;

        [FindsBy(How = How.XPath, Using = APP_SETTINGS_NEW_CUSTOMIZABLE_COLUMN_INFO)]
        private IWebElement _newCustomizableColumnInfo;

        [FindsBy(How = How.Id, Using = APP_SETTINGS_NEW_CUSTOMIZABLE_COLUMN_SAVE_BUTTON)]
        private IWebElement _newCustomizableColumnSave;

        //____________________________________________________Méthodes_______________________________________________________________
        public bool IsCustomizableColumnExist(string columnName, string columnCategory)
        {
            try
            {
                WaitForElementIsVisible(By.XPath(string.Format(APP_SETTINGS_CUSTOMIZABLE_COLUMN, columnName, columnCategory)), nameof(APP_SETTINGS_CUSTOMIZABLE_COLUMN));
                return true;
            }
            catch
            {
                return false;
            }
        }

        public void CreateCustomizableColums(string columnName, string columnCategory, string columnType, string columnSize, string columnOrder, string columnInfo = null)
        {
            //check if column exist
            if (!IsCustomizableColumnExist(columnName, columnCategory))
            {
                //open Create new document modal
                _addCustomizableColumn = WaitForElementIsVisible(By.Id(APP_SETTINGS_CUSTOMIZABLE_COLUMN_ADD_BUTTON), nameof(APP_SETTINGS_CUSTOMIZABLE_COLUMN_ADD_BUTTON));
                _addCustomizableColumn.Click();
                WaitForLoad();

                //add Name
                _newCustomizableColumnName = WaitForElementIsVisible(By.Id(APP_SETTINGS_NEW_CUSTOMIZABLE_COLUMN_NAME), nameof(APP_SETTINGS_NEW_CUSTOMIZABLE_COLUMN_NAME));
                _newCustomizableColumnName.SetValue(ControlType.TextBox, columnCategory);

                //add Category
                _newCustomizableColumnCategory = WaitForElementIsVisible(By.Id(APP_SETTINGS_NEW_CUSTOMIZABLE_COLUMN_CATEGORY), nameof(APP_SETTINGS_NEW_CUSTOMIZABLE_COLUMN_CATEGORY));
                _newCustomizableColumnCategory.SetValue(ControlType.DropDownList, columnName);

                //add Type
                _newCustomizableColumnType = WaitForElementIsVisible(By.Id(APP_SETTINGS_NEW_CUSTOMIZABLE_COLUMN_TYPE), nameof(APP_SETTINGS_NEW_CUSTOMIZABLE_COLUMN_TYPE));
                _newCustomizableColumnType.SetValue(ControlType.DropDownList, columnType);

                //add column size
                _newCustomizableColumnSize = WaitForElementIsVisible(By.Id(APP_SETTINGS_NEW_CUSTOMIZABLE_COLUMN_SIZE), nameof(APP_SETTINGS_NEW_CUSTOMIZABLE_COLUMN_SIZE));
                _newCustomizableColumnSize.SetValue(ControlType.TextBox, columnSize);

                //add column order
                _newCustomizableColumnOrder = WaitForElementIsVisible(By.Id(APP_SETTINGS_NEW_CUSTOMIZABLE_COLUMN_ORDER), nameof(APP_SETTINGS_NEW_CUSTOMIZABLE_COLUMN_ORDER));
                _newCustomizableColumnOrder.SetValue(ControlType.TextBox, columnOrder);

                //add info
                if (columnInfo != null)
                {
                    _newCustomizableColumnInfo = WaitForElementIsVisible(By.XPath(APP_SETTINGS_NEW_CUSTOMIZABLE_COLUMN_INFO), nameof(APP_SETTINGS_NEW_CUSTOMIZABLE_COLUMN_INFO));
                    _newCustomizableColumnInfo.SetValue(ControlType.TextBox, columnInfo);
                }

                _newCustomizableColumnSave = WaitForElementIsVisible(By.Id(APP_SETTINGS_NEW_CUSTOMIZABLE_COLUMN_SAVE_BUTTON), nameof(APP_SETTINGS_NEW_CUSTOMIZABLE_COLUMN_SAVE_BUTTON));
                _newCustomizableColumnSave.Click();
                WaitForLoad();
            }
        }

        public void AddAllCustomizableColumns()
        {
            //HACCP Sanitization
            CreateCustomizableColums("HACCP - Sanithization", "Quantity (Kg)", "Number", "1", "0");
            CreateCustomizableColums("HACCP - Sanithization", "Desinfection Quantity (ppm %)", "Number", "1", "1", "100ppm (Con exposición por un tiempo de 5 mimutos / Recomendsado )");
            CreateCustomizableColums("HACCP - Sanithization", "Disinfection Time (min)", "Number", "1", "2", "100ppm (Con exposición por un tiempo de 5 mimutos / Recomendsado )");
            CreateCustomizableColums("HACCP - Sanithization", "Rinsing", "Checkbox C/NC", "1", "3", "enjuague con agua limpia");
            CreateCustomizableColums("HACCP - Sanithization", "Final visual inspection", "Checkbox C/NC", "1", "4", "ningún objeto extraño");
            CreateCustomizableColums("HACCP - Sanithization", "Comments/ corrective action", "Long text", "1", "5", "Si la concentración <50 ppm o> 100 ppm: prepare una nueva solución, comprobar la concentración, y proceder a la desinfección ");
            CreateCustomizableColums("HACCP - Sanithization", "Prepared by", "Text", "1", "6");

            //HACCP Thawing
            CreateCustomizableColums("HACCP - Thawing", "Real Quantity", "Number", "1", "1");
            CreateCustomizableColums("HACCP - Thawing", "Exp. Date", "Date", "1", "2");
            CreateCustomizableColums("HACCP - Thawing", "Batch number if available", "Text", "2", "3");
            CreateCustomizableColums("HACCP - Thawing", "Thawing starting date", "Date", "2", "4", "Descongelar a 4 ° C Carne 3 días Pescado 2 días Solo en vuelo: descongelación a ≤8ºC durante 2 días para carne cruda descongelar a ≤8 ° C durante 1 día para el pescado crudo Descongelar RTE Congelado a <4 ° C durante 1 día / 24 horas");
            CreateCustomizableColums("HACCP - Thawing", "Maximum End using", "Date", "2", "5", "Descongelar a 4 ° C Carne 3 días Pescado 2 días Solo en vuelo: descongelación a ≤8ºC durante 2 días para carne cruda descongelar a ≤8 ° C durante 1 día para el pescado crudo Descongelar RTE Congelado a <4 ° C durante 1 día / 24 horas ");
            CreateCustomizableColums("HACCP - Thawing", "Comments/ corrective action", "Text", "3", "6", "Deseche los productos alimenticios potencialmente peligrosos si se excede la temperatura de la superficie o el tiempo de descongelación");
            CreateCustomizableColums("HACCP - Thawing", "Prepared By", "Text", "2", "7");

            //HACCP Test Differents Types
            CreateCustomizableColums("HACCP - TSU", "Checkbox C/NC", "Checkbox C/NC", "1", "1", "created for testauto");
            CreateCustomizableColums("HACCP - TSU", "Checkbox C/NA", "Checkbox C/NA", "1", "2", "created for testauto");
            CreateCustomizableColums("HACCP - TSU", "MultiDate", "MultiDate", "2", "3", "created for testauto");
        }
    }
}
