using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium.Support.UI;
using Newrest.Winrest.FunctionalTests.PageObjects.Warehouse.Inventory;


namespace Newrest.Winrest.FunctionalTests.PageObjects.Warehouse.Inventory
{
    public class PrintPreparationSheetModalPage : PageBase
    {
        public PrintPreparationSheetModalPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        // _____________________________________ Constantes _________________________________________________

        private const string PRINT_MODE = "SelectedPrintMode";
        private const string DISPLAY_METHOD = "SelectedDisplayMethod";
        private const string THEORICAL_QUANTITIES = "ShowTheoricalQuantities";
        private const string MAIN_SUPPLIER = "ShowMainSupplier";
        private const string AVERAGE_PRICE = "ShowAveragePrice";

        // ______________________________________ Variables _________________________________________________  

        [FindsBy(How = How.Id, Using = PRINT_MODE)]
        private IWebElement _printMode;

        [FindsBy(How = How.Id, Using = DISPLAY_METHOD)]
        private IWebElement _displayMethod;

        [FindsBy(How = How.Id, Using = THEORICAL_QUANTITIES)]
        private IWebElement _showTheoricalQuantities;

        [FindsBy(How = How.Id, Using = MAIN_SUPPLIER)]
        private IWebElement _showMainSupplier;

        [FindsBy(How = How.Id, Using = AVERAGE_PRICE)]
        private IWebElement _showAveragePrice;

        // _______________________________________ METHODES _____________________________________________________

        public enum PrintParameters
        {
            PrintMode,
            DisplayMethod,
            ShowTheoricalQuantities,
            ShowMainSupplier,
            ShowAveragePrice
        }

        public void Print()
        {
            
        }

        public void SetParameters(PrintParameters parameter, object value)
        {
            switch (parameter)
            {
                case PrintParameters.PrintMode:
                    _printMode = WaitForElementIsVisible(By.Id(PRINT_MODE));
                    _printMode.SetValue(ControlType.DropDownList, value);
                    break;
                case PrintParameters.DisplayMethod:
                    _displayMethod = WaitForElementIsVisible(By.Id(DISPLAY_METHOD));
                    _displayMethod.SetValue(ControlType.DropDownList, value);
                    break;
                case PrintParameters.ShowAveragePrice:
                    _showAveragePrice = WaitForElementExists(By.Id(AVERAGE_PRICE));
                    _showAveragePrice.SetValue(ControlType.CheckBox, value);
                    break;
                case PrintParameters.ShowMainSupplier:
                    _showMainSupplier = WaitForElementExists(By.Id(MAIN_SUPPLIER));
                    _showMainSupplier.SetValue(ControlType.CheckBox, value);
                    break;
                case PrintParameters.ShowTheoricalQuantities:
                    _showTheoricalQuantities = WaitForElementExists(By.Id(THEORICAL_QUANTITIES));
                    _showTheoricalQuantities.SetValue(ControlType.CheckBox, value);
                    break;
                default:
                    throw new ArgumentOutOfRangeException(nameof(PrintParameters), parameter, null);
            }

            WaitForLoad();
        }

    }
}
