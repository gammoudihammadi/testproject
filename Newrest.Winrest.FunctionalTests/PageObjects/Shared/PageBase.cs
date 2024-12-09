using DocumentFormat.OpenXml.ExtendedProperties;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObject.Parameters.Customer;
using Newrest.Winrest.FunctionalTests.PageObjects.Accounting.Edi;
using Newrest.Winrest.FunctionalTests.PageObjects.Accounting.FreePrice;
using Newrest.Winrest.FunctionalTests.PageObjects.Accounting.Invoice;
using Newrest.Winrest.FunctionalTests.PageObjects.Accounting.Reporting;
using Newrest.Winrest.FunctionalTests.PageObjects.Accounting.SupplierInvoices;
using Newrest.Winrest.FunctionalTests.PageObjects.Admin;
using Newrest.Winrest.FunctionalTests.PageObjects.Clinic.Patient;
using Newrest.Winrest.FunctionalTests.PageObjects.CustomerPortal;
using Newrest.Winrest.FunctionalTests.PageObjects.Customers.Catalogs;
using Newrest.Winrest.FunctionalTests.PageObjects.Customers.Customer;
using Newrest.Winrest.FunctionalTests.PageObjects.Customers.Delivery;
using Newrest.Winrest.FunctionalTests.PageObjects.Customers.DeliveryRound;
using Newrest.Winrest.FunctionalTests.PageObjects.Customers.FoodPackets;
using Newrest.Winrest.FunctionalTests.PageObjects.Customers.PriceList;
using Newrest.Winrest.FunctionalTests.PageObjects.Customers.Reconciliation;
using Newrest.Winrest.FunctionalTests.PageObjects.Customers.Reinvoice;
using Newrest.Winrest.FunctionalTests.PageObjects.Customers.Service;
using Newrest.Winrest.FunctionalTests.PageObjects.Flights.Crew;
using Newrest.Winrest.FunctionalTests.PageObjects.Flights.Flight;
using Newrest.Winrest.FunctionalTests.PageObjects.Flights.LoadingPlans;
using Newrest.Winrest.FunctionalTests.PageObjects.Flights.LpCart;
using Newrest.Winrest.FunctionalTests.PageObjects.Flights.Schedule;
using Newrest.Winrest.FunctionalTests.PageObjects.Flights.TimeBlock;
using Newrest.Winrest.FunctionalTests.PageObjects.Flights.Trolleys;
using Newrest.Winrest.FunctionalTests.PageObjects.Interim.InterimOrders;
using Newrest.Winrest.FunctionalTests.PageObjects.Interim.InterimReceptions;
using Newrest.Winrest.FunctionalTests.PageObjects.Jobs.Jobs;
using Newrest.Winrest.FunctionalTests.PageObjects.Jobs.ScheduledJobs;
using Newrest.Winrest.FunctionalTests.PageObjects.Jobs.Settings;
using Newrest.Winrest.FunctionalTests.PageObjects.Menus.Datasheet;
using Newrest.Winrest.FunctionalTests.PageObjects.Menus.FoodCost;
using Newrest.Winrest.FunctionalTests.PageObjects.Menus.Menus;
using Newrest.Winrest.FunctionalTests.PageObjects.Menus.Recipes;
using Newrest.Winrest.FunctionalTests.PageObjects.Parameters.Accounting;
using Newrest.Winrest.FunctionalTests.PageObjects.Parameters.Clinic;
using Newrest.Winrest.FunctionalTests.PageObjects.Parameters.Flights;
using Newrest.Winrest.FunctionalTests.PageObjects.Parameters.GlobalSettings;
using Newrest.Winrest.FunctionalTests.PageObjects.Parameters.Logs;
using Newrest.Winrest.FunctionalTests.PageObjects.Parameters.Portal;
using Newrest.Winrest.FunctionalTests.PageObjects.Parameters.Production;
using Newrest.Winrest.FunctionalTests.PageObjects.Parameters.Purchasing;
using Newrest.Winrest.FunctionalTests.PageObjects.Parameters.Sites;
using Newrest.Winrest.FunctionalTests.PageObjects.Parameters.Tablet;
using Newrest.Winrest.FunctionalTests.PageObjects.Parameters.User;
using Newrest.Winrest.FunctionalTests.PageObjects.Production.CustomerOrder;
using Newrest.Winrest.FunctionalTests.PageObjects.Production.Dispatch;
using Newrest.Winrest.FunctionalTests.PageObjects.Production.EarlyProduction;
using Newrest.Winrest.FunctionalTests.PageObjects.Production.Production;
using Newrest.Winrest.FunctionalTests.PageObjects.Production.ProductionCO;
using Newrest.Winrest.FunctionalTests.PageObjects.Production.ProductionManagement;
using Newrest.Winrest.FunctionalTests.PageObjects.Production.Setup;
using Newrest.Winrest.FunctionalTests.PageObjects.Purchasing;
using Newrest.Winrest.FunctionalTests.PageObjects.Purchasing.Needs;
using Newrest.Winrest.FunctionalTests.PageObjects.Purchasing.Product;
using Newrest.Winrest.FunctionalTests.PageObjects.TabletApp;
using Newrest.Winrest.FunctionalTests.PageObjects.ToDoList.Scheduler;
using Newrest.Winrest.FunctionalTests.PageObjects.ToDoList.Tasks;
using Newrest.Winrest.FunctionalTests.PageObjects.Warehouse.Claims;
using Newrest.Winrest.FunctionalTests.PageObjects.Warehouse.Inventory;
using Newrest.Winrest.FunctionalTests.PageObjects.Warehouse.OutputForm;
using Newrest.Winrest.FunctionalTests.PageObjects.Warehouse.ReceiptNotes;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.Extensions;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading;
using UglyToad.PdfPig.Content;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Shared
{
    public abstract class PageBase
    {
        protected readonly IWebDriver _webDriver;
        protected readonly TestContext _testContext;

        protected PageBase(IWebDriver webDriver, TestContext testContext)
        {
            _webDriver = webDriver;
            _testContext = testContext;
            PageFactory.InitElements(_webDriver, this);
        }

        // TODO : Récupérer les iD pour chaque sous menu quand ils seront disponibles (pas les liens)

        //Menus - Sous menus

        // 1 - Purchasing
        private const string PURCHASING_MENU = "TabPurchasing";
        private const string PU_ITEM = "ItemTabNav";
        private const string PU_ITEM_LINK = "ItemLinkDashBoard";
        private const string PU_SUPPLIERS = "SupplierTabNav";
        private const string PU_SUPPLIERS_LINK = "SupplierLinkDashBoard";
        private const string PU_SUPPLY_ORDER = "SupplyOrderTabNav";
        private const string PU_SUPPLY_ORDER_LINK = "SupplyOrderLinkDashBoard";
        private const string PU_PURCHASE_ORDER = "PurchaseOrderTabNav";
        private const string PU_PURCHASE_ORDER_LINK = "PurchaseOrderLinkDashBoard";
        private const string PU_NEEDS = "NeedRawMaterialsTabNav";
        private const string PU_NEEDS_LINK = "NeedRawMaterialsLinkDashBoard";
        private const string PRODUCT_TAB_NAV = "ProductTabNav";
        private const string PRODUCT_LINK_DASHBOARD = "NeedRawMaterialsLinkDashBoard";
        private const string PRODUCTION_MENU_LIST = "//*[@id=\"CustomerOrderProductionLinkDashBoard\"]";
        private const string JOBS_JOBS = "JobsTabNav";
        private const string JOBS_JOBS_LINK = "JobsLinkDashBoard";
        [FindsBy(How = How.Id, Using = PURCHASING_MENU)]
        private IWebElement _purchasing;

        [FindsBy(How = How.Id, Using = JOBS_JOBS)]
        private IWebElement _jobs_jobs;
        [FindsBy(How = How.Id, Using = JOBS_JOBS_LINK)]
        private IWebElement _jobs_jobs_link;
        [FindsBy(How = How.Id, Using = PU_ITEM)]
        private IWebElement _purchasing_Item;

        [FindsBy(How = How.Id, Using = PU_ITEM_LINK)]
        private IWebElement _purchasing_Item_link;

        [FindsBy(How = How.Id, Using = PU_SUPPLIERS)]
        private IWebElement _purchasing_Suppliers;

        [FindsBy(How = How.Id, Using = PU_SUPPLIERS_LINK)]
        private IWebElement _purchasing_Suppliers_link;

        [FindsBy(How = How.Id, Using = PU_SUPPLY_ORDER)]
        private IWebElement _purchasing_SupplyOrders;

        [FindsBy(How = How.Id, Using = PU_SUPPLY_ORDER_LINK)]
        private IWebElement _purchasing_SupplyOrders_link;

        [FindsBy(How = How.Id, Using = PU_PURCHASE_ORDER)]
        private IWebElement _purchasing_PurchaseOrders;

        [FindsBy(How = How.Id, Using = PU_PURCHASE_ORDER_LINK)]
        private IWebElement _purchasing_PurchaseOrders_link;

        [FindsBy(How = How.Id, Using = PU_NEEDS)]
        private IWebElement _purchasing_Needs;

        [FindsBy(How = How.Id, Using = PU_NEEDS_LINK)]
        private IWebElement _purchasing_Needs_link;

        [FindsBy(How = How.Id, Using = PRODUCT_LINK_DASHBOARD)]
        private IWebElement _product_link_dashboard;
        [FindsBy(How = How.Id, Using = PRODUCT_TAB_NAV)]
        private IWebElement _product_tab_nav;

        // 2 - Flights
        private const string FLIGHT_MENU = "TabFlights";
        private const string FL_FLIGHTS = "FlightTabNav";
        private const string FL_FLIGHTS_LINK = "FlightLinkDashBoard";
        private const string FL_LOADING_PLAN = "LoadingPlanTabNav";
        private const string FL_LOADING_PLAN_LINK = "LoadingPlanLinkDashBoard";
        private const string FL_FLIGHT_SELECTION = "InflightScheduleTabNav";
        private const string FL_CREW = "CrewTabNav";
        private const string FL_CREW_LINK = "CrewLinkDashBoard";
        private const string FL_FLIGHT_SELECTION_LINK = "InflightScheduleLinkDashBoard";
        private const string FL_FLIGHT_LP_CART = "LPCartTabNav";
        private const string FL_FLIGHT_LP_CART_LINK = "LPCartLinkDashBoard";
        private const string FL_FLIGHT_TROLLEY = "TrolleyTabNav";
        private const string FL_FLIGHT_TROLLEY_LINK = "TrolleyLinkDashBoard";
        private const string FL_FLIGHT_TIMEBLOCK = "FlightWorkshopsTabNav";
        private const string FL_FLIGHT_TIMEBLOCK_LINK = "FlightWorkshopsLinkDashBoard";
        private const string JOBS_SETTINGS_LINK = "JobSettingsLinkDashBoard";
        private const string JOBS = "//*[@id=\"TabJobs\"]/a";
        private const string JOBS_SETTINGS_TAB = "//*[@id=\"JobSettingsTabNav\"]/a";
        private const string JOBS_SCHEDULED_JOBS_LINK = "ScheduledJobsLinkDashBoard";
        private const string JOBS_SCHEDULED_JOBS_TAB = "//*[@id=\"ScheduledJobsTabNav\"]/a";
        private const string SEARCH_FIELD = "//*[@id=\"form-copy-from-tbSearchName\"]";
        private const string GROSS_QTY = "/html/body/div[3]/div/div[2]/div[2]/div/div/div[2]/div[2]/div/ul/li[1]/div/div/div/form/div[3]/div[1]/div[8]/div/div[2]/input";
        private const string NET_WEIGHT = "/html/body/div[3]/div/div[2]/div[2]/div/div/div[2]/div[2]/div/ul/li[1]/div/div/div/form/div[3]/div[1]/div[6]/div/div[2]/input";

        private const string NEWS = "modal-1";
        private const string VALIDATE_NEWS = "btnMessageValidation";

        [FindsBy(How = How.Id, Using = FLIGHT_MENU)]
        private IWebElement _flights;

        [FindsBy(How = How.Id, Using = FL_FLIGHTS_LINK)]
        private IWebElement _flights_Flights_link;

        [FindsBy(How = How.Id, Using = FL_LOADING_PLAN)]
        private IWebElement _flights_LoadingPlans;

        [FindsBy(How = How.Id, Using = FL_LOADING_PLAN_LINK)]
        private IWebElement _flights_LoadingPlans_link;

        [FindsBy(How = How.Id, Using = FL_FLIGHT_SELECTION)]
        private IWebElement _flights_Flight_Selection;

        [FindsBy(How = How.Id, Using = FL_FLIGHT_SELECTION_LINK)]
        private IWebElement _flights_Flight_Selection_link;

        [FindsBy(How = How.Id, Using = FL_FLIGHT_LP_CART)]
        private IWebElement _flights_Flight_LpCart;

        [FindsBy(How = How.Id, Using = FL_FLIGHT_LP_CART_LINK)]
        private IWebElement _flights_Flight_LpCart_link;

        [FindsBy(How = How.Id, Using = FL_FLIGHT_TROLLEY)]
        private IWebElement _flights_Flight_Trolley;

        [FindsBy(How = How.Id, Using = FL_FLIGHT_TROLLEY_LINK)]
        private IWebElement _flights_Flight_Trolley_link;

        [FindsBy(How = How.Id, Using = FL_FLIGHT_TIMEBLOCK)]
        private IWebElement _flights_Flight_TimeBlock;

        [FindsBy(How = How.Id, Using = FL_FLIGHT_TIMEBLOCK_LINK)]
        private IWebElement _flights_Flight_TimeBlock_link;

        [FindsBy(How = How.Id, Using = JOBS_SETTINGS_LINK)]
        private IWebElement _settings;

        [FindsBy(How = How.XPath, Using = JOBS)]
        private IWebElement _jobs;

        [FindsBy(How = How.XPath, Using = JOBS_SETTINGS_TAB)]
        private IWebElement _jobs_settings;

        [FindsBy(How = How.Id, Using = JOBS_SCHEDULED_JOBS_LINK)]
        private IWebElement _jobs_scheduledJobs;

        [FindsBy(How = How.XPath, Using = JOBS_SCHEDULED_JOBS_TAB)]
        private IWebElement _scheduledJobs;
        // 3 - Menus 
        private const string MENUS_MENU = "TabRecipes";
        private const string ME_RECIPES = "RecipeTabNav";
        private const string ME_RECIPES_LINK = "RecipeLinkDashBoard";
        private const string ME_MENU = "MenuTabNav";
        private const string ME_MENU_LINK = "MenuLinkDashBoard";
        private const string ME_DATASHEET = "DatasheetTabNav";
        private const string ME_DATASHEET_LINK = "DatasheetLinkDashBoard";
        private const string ME_FOODCOST = "FoodCostTabNav";
        private const string ME_FOODCOST_LINK = "FoodCostLinkDashBoard";

        [FindsBy(How = How.Id, Using = MENUS_MENU)]
        private IWebElement _menu;

        [FindsBy(How = How.Id, Using = ME_RECIPES)]
        private IWebElement _menu_Recipes;

        [FindsBy(How = How.Id, Using = ME_RECIPES_LINK)]
        private IWebElement _menu_Recipes_link;

        [FindsBy(How = How.Id, Using = ME_MENU)]
        private IWebElement _menu_Menus;

        [FindsBy(How = How.Id, Using = ME_MENU_LINK)]
        private IWebElement _menu_Menus_link;

        [FindsBy(How = How.Id, Using = ME_DATASHEET)]
        private IWebElement _menu_Datasheet;

        [FindsBy(How = How.Id, Using = ME_DATASHEET_LINK)]
        private IWebElement _menu_Datasheet_link;

        // 4 - Production
        private const string PRODUCTION_MENU = "TabProduction";
        private const string PR_CUSTOMER_ORDER = "OrderTabNav";
        private const string PR_CUSTOMER_ORDER_LINK = "OrderLinkDashBoard";
        private const string PR_DISPATCH = "DispatchTabNav";
        private const string PR_PRODUCTION = "ProductionTabNav";
        private const string PR_EARLY_PRODUCTION = "InflightEarlyProductionTabNav";
        private const string PR_PRODUCTION_CO = "CustomerOrderProductionTabNav";

        private const string PR_DISPATCH_LINK = "DispatchLinkDashBoard";
        private const string PR_PRODUCTION_LINK = "ProductionLinkDashBoard";
        private const string PR_EARLY_PRODUCTION_LINK = "InflightEarlyProductionLinkDashBoard";
        private const string PR_PRODUCTION_MANAGEMENT = "InflightRawMaterialsTabNav";
        private const string PR_PRODUCTION_MANAGEMENT_LINK = "InflightRawMaterialsLinkDashBoard";
        private const string PR_SETUP = "SetupTabNav";
        private const string PR_SETUP_LINK = "SetupLinkDashBoard";
        private const string PR_PRODUCTION_CO_LINK = "CustomerOrderProductionLinkDashBoard";
        private const string PR_PRODUCTION_CO_FROM_LIST = "";


        [FindsBy(How = How.Id, Using = PRODUCTION_MENU)]
        private IWebElement _production;

        [FindsBy(How = How.Id, Using = PR_CUSTOMER_ORDER)]
        private IWebElement _production_Customer_Order;

        [FindsBy(How = How.Id, Using = PR_CUSTOMER_ORDER_LINK)]
        private IWebElement _production_Customer_Order_link;

        [FindsBy(How = How.Id, Using = PR_DISPATCH)]
        private IWebElement _production_dipatch;

        [FindsBy(How = How.Id, Using = PR_PRODUCTION)]
        private IWebElement _production_production;

        [FindsBy(How = How.Id, Using = PR_DISPATCH_LINK)]
        private IWebElement _production_dipatch_link;

        [FindsBy(How = How.Id, Using = PR_PRODUCTION_LINK)]
        private IWebElement _production_production_link;

        [FindsBy(How = How.Id, Using = PR_PRODUCTION_MANAGEMENT)]
        private IWebElement _production_productionManagement;

        [FindsBy(How = How.Id, Using = PR_PRODUCTION_MANAGEMENT_LINK)]
        private IWebElement _production_productionManagement_link;

        [FindsBy(How = How.Id, Using = PR_SETUP)]
        private IWebElement _production_setup;

        [FindsBy(How = How.Id, Using = PR_SETUP_LINK)]
        private IWebElement _production_setup_link;

        [FindsBy(How = How.Id, Using = PR_PRODUCTION_CO)]
        private IWebElement _production_Production_CO;

        [FindsBy(How = How.Id, Using = PR_CUSTOMER_ORDER_LINK)]
        private IWebElement _production_Production_CO_link;

        // 5 - Warehouse
        private const string WAREHOUSE_MENU = "//*[@id=\"TabWarehouse\"]";
        private const string WA_RECIPT_NOTES_LINK = "//*[@id=\"test\"]/div[4]/div/div/ul/li[2]/a";
        private const string WA_CLAIM = "//*[@id=\"TabWarehouse\"]/div/div/ul/li[2]/a";

        [FindsBy(How = How.XPath, Using = WAREHOUSE_MENU)]
        private IWebElement _wareHouse;
        [FindsBy(How = How.XPath, Using = WA_RECIPT_NOTES_LINK)]
        private IWebElement _wareHouse_Receipt_Note_link;

        [FindsBy(How = How.XPath, Using = WA_CLAIM)]
        private IWebElement _wareHouse_Claim;

        // 6 -Customers
        private const string CU_FOOD_PACKET_LINK = "/html/body/div[2]/div/div[1]/div/div[6]/div/div/ul/li[7]/a";
        private const string CU_DELIVERY_ROUND_LINK = "//*[@id=\"test\"]/div[5]/div/div/ul/li[5]/a";

        [FindsBy(How = How.XPath, Using = CU_DELIVERY_ROUND_LINK)]
        private IWebElement _customer_Delivery_Round_link;

        // 7 - Accounting
        private const string ACCOUNTING_MENU = "TabAccounting";
        private const string AC_SUPPLIER_INVOICE = "SupplierInvoiceTabNav";
        private const string AC_SUPPLIER_INVOICE_LINK = "SupplierInvoiceLinkDashBoard";
        private const string AC_INVOICE = "InvoiceTabNav";
        private const string AC_INVOICE_LINK = "InvoiceLinkDashBoard";
        private const string AC_FREE_PRICE = "FreePriceTabNav";
        private const string AC_FREE_PRICE_LINK = "FreePriceLinkDashBoard";
        private const string AC_REPORTING = "ReportingTabNav";
        private const string AC_REPORTING_LINK = "ReportingLinkDashBoard";

        [FindsBy(How = How.Id, Using = ACCOUNTING_MENU)]
        private IWebElement _accounting;

        [FindsBy(How = How.Id, Using = AC_SUPPLIER_INVOICE)]
        private IWebElement _accounting_Supplier_Invoice;

        [FindsBy(How = How.Id, Using = AC_SUPPLIER_INVOICE_LINK)]
        private IWebElement _accounting_Supplier_Invoice_link;

        [FindsBy(How = How.Id, Using = AC_INVOICE)]
        private IWebElement _accounting_Invoice;

        [FindsBy(How = How.Id, Using = AC_INVOICE_LINK)]
        private IWebElement _accounting_Invoice_link;

        [FindsBy(How = How.Id, Using = AC_FREE_PRICE)]
        private IWebElement _accounting_FreePrice;

        [FindsBy(How = How.Id, Using = AC_FREE_PRICE_LINK)]
        private IWebElement _accounting_FreePrice_link;

        [FindsBy(How = How.Id, Using = AC_REPORTING)]
        private IWebElement _accounting_Reporting;

        [FindsBy(How = How.Id, Using = AC_REPORTING_LINK)]
        private IWebElement _accounting_Reporting_link;

        // - Clinic
        private const string CLINIC_TAB = "TabClinics";
        private const string CLINIC_PATIENT_TAB = "PatientTabNav";
        private const string CLINIC_PATIENT_LINK = "PatientLinkDashBoard";

        [FindsBy(How = How.Id, Using = CLINIC_TAB)]
        private IWebElement _clinicTab;

        [FindsBy(How = How.Id, Using = CLINIC_PATIENT_TAB)]
        private IWebElement _clinicPatientTab;

        [FindsBy(How = How.Id, Using = CLINIC_PATIENT_LINK)]
        private IWebElement _clinicPatientLink;

        // 8 - Paremeters
        private const string PARAMETERS_MENU = "//*[@id=\"TabParameters\"]";
        private const string PA_PURCHASING = "//*[@id=\"TabParameters\"]/div/div/ul/li[6]/a";
        private const string PA_SITES_LINK = "//*[@id=\"test\"]/div[10]/div/div/ul/li[8]/a";
        private const string PA_CLINIC = "ClinicTabNav";
        private const string PA_CLINIC_LINK = "ClinicLinkDashBoard";

        [FindsBy(How = How.XPath, Using = PARAMETERS_MENU)]
        private IWebElement _parameters;

        [FindsBy(How = How.XPath, Using = PA_PURCHASING)]
        private IWebElement _parameters_Purchasing;

        [FindsBy(How = How.Id, Using = PA_SITES_LINK)]
        private IWebElement _parameters_Sites_link;

        [FindsBy(How = How.Id, Using = PA_CLINIC)]
        private IWebElement _parameters_Clinic;

        // 9 - CustomerPortal
        private const string CUSTOMER_PORTAL = "//*[@id=\"div-body\"]/div/div[2]/div[3]/div/h2/a";

        [FindsBy(How = How.XPath, Using = CUSTOMER_PORTAL)]
        private IWebElement _customerPortal;

        // 10 - TabletApp
        private const string TABLET_APP = "//*[@id=\"tabletLink\"]/parent::a";

        [FindsBy(How = How.XPath, Using = TABLET_APP)]
        private IWebElement _tabletApp;

        // 11 - EDI
        private const string EDI = "EditLinkDashBoard";

        [FindsBy(How = How.Id, Using = EDI)]
        private IWebElement _edi;

        // 11 - Divers
        private const string VERSION_DB_APPLICATION = "//*[@id=\"DbName-header\"]/div/span[3]";
        private const string PAGE_SIZE = "page-size-selector";
        private const string PLUS_BTN = "//*[@id=\"div-body\"]/div/div[2]/div[1]/div/div[2]/button";
        private const string EXTENDED_BTN = "//button[contains(text(), '...')]";
        private const string VALIDATE_BTN = "//*[@id=\"div-body\"]/div/div[1]/div/div[2]/button";
        private const string ERROR_MESSAGE_PRINT = "/html/body/div[*]/div/div/div[2]/p";
        private const string CLOSE_BTN = "btn-cancel-popup";
        private const string PRINT_BUTTON = "//*[@id=\"header-print-button\"]";
        private const string PRINT_BUTTON_ID = "header-print-button";
        private const string PRINT_POPUP = "//h3[text() = 'Print list']";
        private const string PAGE_SIZE_SELECTED = "//*[@id=\"page-size-selector\"]/option[@selected]";

        [FindsBy(How = How.XPath, Using = VERSION_DB_APPLICATION)]
        private IWebElement _versionDBApplication;

        [FindsBy(How = How.ClassName, Using = "counter")]
        private IWebElement _totalNumber;

        //Boutons + et ... en haut à droite des pages
        [FindsBy(How = How.XPath, Using = EXTENDED_BTN)]
        private IWebElement _extendedButton;

        [FindsBy(How = How.XPath, Using = PLUS_BTN)]
        private IWebElement _plusButton;

        //Bouton "✓"  en haut à droite ( d'un purchase order, ou d'un supply order, ...)
        [FindsBy(How = How.XPath, Using = VALIDATE_BTN)]
        private IWebElement _validateButton;

        [FindsBy(How = How.Id, Using = PAGE_SIZE)]
        private IWebElement _pageSize;

        // 13 - TODOLIST
        private const string TODOLIST = "TabLounge";
        private const string SCHEDULER_LINK = "SchedulerLinkDashBoard";
        private const string TASK_LINK = "TaskLinkDashBoard";
        private const string SCHEDULER_TAB = "SchedulerTabNav";
        private const string TASK_TAB = "TaskTabNav";

        [FindsBy(How = How.Id, Using = TODOLIST)]
        private IWebElement _toDoList;

        [FindsBy(How = How.Id, Using = SCHEDULER_LINK)]
        private IWebElement _schedulerLink;

        [FindsBy(How = How.Id, Using = SCHEDULER_TAB)]
        private IWebElement _schedulerTab;

        [FindsBy(How = How.Id, Using = TASK_LINK)]
        private IWebElement _taskLink;

        [FindsBy(How = How.Id, Using = TASK_TAB)]
        private IWebElement _taskTab;

        // ______________________________________________________UTILITAIRE______________________________________________________________________

        public virtual void ShowPlusMenu()
        {
            WaitForLoad();
            _plusButton = WaitForElementIsVisibleNew(By.CssSelector("#div-body > div > div.sideview-centralPane.selectable-container > div.title-bar > div > div.dropdown.dropdown-add-button > button"));
            var actions = new Actions(_webDriver);
            actions.MoveToElement(_plusButton).Perform();
            WaitForLoad();
        }

        public virtual void ShowValidationMenu()
        {
            WaitForLoad();
            _validateButton = WaitForElementExists(By.XPath(VALIDATE_BTN));
            var actions = new Actions(_webDriver);
            actions.MoveToElement(_validateButton).Perform();
            WaitForLoad();
        }

        public virtual void ShowExtendedMenu()
        {
            _extendedButton = WaitForElementExists(By.XPath(EXTENDED_BTN));
            var actions = new Actions(_webDriver);
            actions.MoveToElement(_extendedButton).Perform();
            WaitForLoad();
        }

        public virtual int CheckTotalNumber()
        {
            WaitPageLoading();
            _totalNumber = WaitForElementExists(By.ClassName("counter"));
            int nombre = Int32.Parse(_totalNumber.GetAttribute("innerText"));
            return nombre;
        }
        public int CheckTotalNumber_customerPortal()
        {
            WaitForLoad();
            _totalNumber = WaitForElementExists(By.XPath("//*/span[@class='counter']"));
            int nombre = Int32.Parse(_totalNumber.GetAttribute("innerText"));
            return nombre;
        }

        public int CheckTotalNumber_freePrice()
        {
            WaitForLoad();
            _totalNumber = WaitForElementExists(By.Id("countFreePrice"));
            int nombre = Int32.Parse(_totalNumber.GetAttribute("innerText"));
            return nombre;
        }

        public int CheckTotalNumber_display()
        {
            WaitForLoad();
            _totalNumber = WaitForElementExists(By.XPath("//*[@id=\"div-body\"]/div/div[2]/div[1]/h1/span"));
            int nombre = Int32.Parse(_totalNumber.GetAttribute("innerText"));
            return nombre;
        }
        protected void WaitForLoad()
        {
            var javaScriptExecutor = _webDriver as IJavaScriptExecutor;
            var wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(61));

            Func<IWebDriver, bool> readyCondition = webDriver =>
                (bool)javaScriptExecutor.ExecuteScript("return (document.readyState == 'complete' && jQuery.active == 0)");

            wait.Until(readyCondition);
        }

        private ReadOnlyCollection<IWebElement> elementList;

        protected void WaitForDownload()
        {
            /*var javaScriptExecutor = _webDriver as IJavaScriptExecutor;

            //Switch vers le premier onglet
            _webDriver.SwitchTo().Window(_webDriver.WindowHandles[0]);

            //Création d'un nouvel onglet
            javaScriptExecutor.ExecuteScript("window.open('about:blank','_blank');");

            //Switch vers le nouvel onglet et navigation vers la page de download
            _webDriver.SwitchTo().Window(_webDriver.WindowHandles.Last());
            _webDriver.Navigate().GoToUrl("chrome://downloads/");

            //Selection du premier élément downloads-item
            var manager = _webDriver.FindElement(By.XPath("/html/body/downloads-manager"));*/

            ScanFilesNotDownloadedYet();
        }

        private void ScanFilesNotDownloadedYet()
        {
            bool trouveTmp = true;
            bool trouveUnFichier = false;
            var downloadsPath = _testContext.Properties["DownloadsPath"].ToString();
            int counter = 0;
            while ((!trouveUnFichier || trouveTmp) && counter < 20)
            {
                trouveTmp = false;
                foreach (string f in Directory.GetFiles(downloadsPath))
                {
                    var fi = new FileInfo(f);
                    if (!fi.Exists)
                    {
                        counter++;
                        continue;
                    }
                    else if (fi.CreationTime<DateTime.Now.AddDays(-2))
                    {
                        counter++;
                        continue;
                    }
                    if (f.EndsWith(".crdownload") || f.EndsWith(".tmp") )
                    {
                        trouveTmp = true;
                        break;
                    }
                    trouveUnFichier = true;
                }
                if (trouveTmp)
                {
                    Thread.Sleep(2000);
                    counter++;
                }
            }
        }

        // problème menu Download "Show in folder"+"Delete from history" 12/06/2024
        private void WaitForTextToBePresent(IWebElement manager, string text)
        {
            // Initialisation des variables
            int compteur = 1;
            bool isFound = false;

            var wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(3));
            while (compteur <= 20 && !isFound)
            {
                try
                {
                    wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.TextToBePresentInElement(manager, text));
                    isFound = true;
                }
                catch
                {
                    compteur++;
                }
            }

            if (!isFound)
            {
                Console.WriteLine($"[{DateTime.UtcNow.ToShortDateString()} {DateTime.UtcNow.ToLongTimeString()}] [WaitForTextToBePresent] Text not found : {text}");
                throw new Exception("L'élément n'a pas été téléchargé.");
            }
        }

        public string GetApplicationDbVersion()
        {
            _versionDBApplication = WaitForElementIsVisible(By.XPath(VERSION_DB_APPLICATION));
            var value = _versionDBApplication.Text.Replace(" - ", "/").Split('/');

            return value[1];
        }

        public void PageSize(string size)
        {
            if (size == "1000")
            {   // Test
                IJavaScriptExecutor js = (IJavaScriptExecutor)_webDriver;
                js.ExecuteScript("$('#" + PAGE_SIZE + "').append($('<option>', {value: 1000,text: '1000'}),'');");
            }

            var wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(30));
            try
            {
                wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.Id(PAGE_SIZE)));
            }
            catch
            {
                // tableau vide : pas de PageSize
                return;
            }
            _pageSize = WaitForElementExists(By.Id(PAGE_SIZE));
            Actions action = new Actions(_webDriver);
            action.MoveToElement(_pageSize).Perform();
            _pageSize = WaitForElementExists(By.Id(PAGE_SIZE));
            _pageSize.SetValue(ControlType.DropDownList, size);

            WaitPageLoading();
            WaitForLoad();
            // pour écran plus petit que 8 lignes affiché
            PageUp();
        }

        public void ClearDownloads()
        {
            int compteur = 1;
            bool isVisible = false;

            while (compteur <= 5 && !isVisible)
            {
                var wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(5));

                if (!IsPrintListOpen())
                {
                    ClickPrintButton();
                }
                try
                {
                    var clearDownloadButton = wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.CssSelector("[onclick=\"clearPrintList()\"]")));
                    clearDownloadButton.Click();
                    WaitForLoad();
                    isVisible = true;

                    ClosePrintButton();
                    // Ophélie : Temps de fermeture de la fenêtre
                    //_webDriver.Navigate().Refresh();
                }
                catch
                {
                    ClickPrintButton();
                    //David : Double clic trop rapide si on n'attend pas
                    Thread.Sleep(500);
                    compteur++;
                }
            }

            if (!isVisible)
            {
                Console.WriteLine($"[{DateTime.UtcNow.ToShortDateString()} {DateTime.UtcNow.ToLongTimeString()}] [ClearDownloads] Clear button not visible");
                throw new Exception("La fonction 'Clear' du Print n'est pas accessible.");
            }
        }

        public void PurgeDownloads()
        {
            string downloadsPath = _testContext.Properties["DownloadsPath"].ToString();
            foreach (string f in Directory.GetFiles(downloadsPath))
            {
                var fi = new FileInfo(f);
                fi.Delete();
            }
        }

        public bool IsPrintListOpen()
        {
            if (isElementVisible(By.XPath(PRINT_POPUP)))
            {
                return true;
            }

            return false;
        }

        public void ClickPrintButton()
        {
            var wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(3));
            int compteur = 1;
            bool isVisible = false;

            while (compteur <= 5 && !isVisible)
            {
                try
                {

                    var printListButton = wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementExists(By.XPath("//*[@id=\"header-print-button\"]")));
                    printListButton.Click();

                    //On attend de voir le bouton ClearList qui signifie que la liste des prints est chargée
                    wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.CssSelector("[onclick=\"clearPrintList()\"]")));

                    isVisible = true;
                }
                catch (Exception)
                {
                    compteur++;

                    //Le bouton printList (icone imprimante) est censé être sur la page masi webdriver parfois n'arrive pas à le trouver
                    //L'effet clignotant est peut-être la cause ?
                    //Du coup si on n'a pas réussi à le trouver, on attends 500ms et on retente

                    // si on clique plusieur fois trop rapidement, l'icone d'impression ne marche plus, on attends 2 secondes et on retente
                    Thread.Sleep(2000);
                }
            }

            if (!isVisible)
            {
                Console.WriteLine($"[{DateTime.UtcNow.ToShortDateString()} {DateTime.UtcNow.ToLongTimeString()}] [ClickPrintButton] Print button not visible");
                throw new Exception("Le bouton 'Print' n'est pas accessible.");
            }
        }

        public void ClosePrintButton()
        {
            if (isElementVisible(By.XPath(PRINT_POPUP)))
            {
                WaitForElementIsVisibleNew(By.XPath("//*/th[1]/span[@class='header' and text()='Date']"));
                // fenêtre fantome
                var _printButton = WaitForElementIsVisibleNew(By.Id(PRINT_BUTTON_ID));
                _printButton.Click();
                WaitForLoad();
                // non accessible en javascript
                var javaScriptExecutor = _webDriver as IJavaScriptExecutor;
                javaScriptExecutor.ExecuteScript("$('div[role=tooltip]').remove();");
                WaitForLoad();
            }
        }

        public bool IsFileLoadedVisible(By elementReceherche)
        {
            int compteur = 1;
            bool isVisible = false;

            while (compteur <= 30 && !isVisible)
            {
                try
                {
                    var downloadFileBtn = _webDriver.FindElement(elementReceherche);
                    isVisible = true;
                    WaitForLoad();
                    return isVisible;
                }
                catch (Exception ex)
                {
                    ClosePrintButton();
                    compteur++;

                    //David : Le fichier n'est pas prêt, on attend 5s pour ré-ouvrir la popup
                    Thread.Sleep(5000);
                    return isVisible;
                }
            }
            Thread.Sleep(5000);

            return isVisible;
        }

        public void IsFileLoaded(By elementReceherche)
        {
            int compteur = 1;
            bool isVisible = false;

            while (compteur <= 30 && !isVisible)
            {
                ClickPrintButton();

                try
                {
                    var downloadFileBtn = _webDriver.FindElement(elementReceherche);
                    downloadFileBtn.Click();
                    isVisible = true;
                    WaitForLoad();
                }
                catch (Exception ex)
                {
                    ClosePrintButton();
                    compteur++;

                    //David : Le fichier n'est pas prêt, on attend 5s pour ré-ouvrir la popup
                    Thread.Sleep(5000);
                }
            }
            Thread.Sleep(5000);
            if (!isVisible)
            {
                Console.WriteLine($"[{DateTime.UtcNow.ToShortDateString()} {DateTime.UtcNow.ToLongTimeString()}] [IsFileLoaded] Element not found : {elementReceherche}");
                throw new Exception("Délai d'attente dépassé pour le téléchargement du fichier.");
            }

        }

        public string IsFileInError(By elementReceherche, bool messageErreur = false)
        {
            int compteur = 1;
            bool isVisible = false;
            string errorMessage = "";

            while (compteur <= 10 && !isVisible)
            {
                ClickPrintButton();

                try
                {
                    var downloadFileBtn = _webDriver.FindElement(elementReceherche);

                    if (messageErreur)
                    {
                        downloadFileBtn.Click();
                        var error = WaitForElementIsVisible(By.XPath(ERROR_MESSAGE_PRINT));
                        errorMessage = error.Text;

                        var close = WaitForElementIsVisible(By.Id(CLOSE_BTN));
                        close.Click();
                        WaitForLoad();
                    }

                    isVisible = true;
                }
                catch
                {
                    ClickPrintButton();
                    compteur++;

                    //David : Le fichier n'est pas prêt, on attend 10s pour ré-ouvrir la popup
                    Thread.Sleep(3000);
                }
            }

            if (!isVisible)
            {
                Console.WriteLine($"[{DateTime.UtcNow.ToShortDateString()} {DateTime.UtcNow.ToLongTimeString()}] [IsFileInError] Element not found : {elementReceherche}");
                throw new Exception("Aucun fichier en erreur n'a été trouvé.");
            }

            return errorMessage;
        }

        public void WaitPageLoading()
        {
            // On attend que le sablier apparaisse dans la page
            bool isLoading = WaitLoading();

            // Si le sablier n'est pas détecté, on arrête le traitement
            if (isLoading)
            {
                // On attend que le sablier disparaisse pour le chargement de la page
                LoadingPage();
            }
        }

        public bool WaitLoading()
        {
            int compteur = 1;

            while (compteur <= 300)
            {
                try
                {
                    _webDriver.FindElement(By.ClassName("busy-indicator"));
                    return true;
                }
                catch
                {
                    compteur++;
                }
            }
            return false;
        }

        public void LoadingPage()
        {
            bool vueSablier = true;
            int compteur = 1;

            while (compteur <= 600 && vueSablier)
            {
                try
                {
                    var wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(1));
                    wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.ClassName("busy-indicator")));
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
                Console.WriteLine($"[{DateTime.UtcNow.ToShortDateString()} {DateTime.UtcNow.ToLongTimeString()}] [LoadingPage] Infinite Loading");
                throw new Exception("Délai d'attente dépassé pour le chargement de la page.");
            }
        }

        public enum ControlType
        {
            DropDownList,
            TextBox,
            CheckBox,
            DateTime,
            RadioButton,
            Number
        }

        public void GoToWinrestHome()
        {
            var url = _testContext.Properties["Winrest_URL"].ToString();
            _webDriver.Navigate().GoToUrl(url);
            WaitForLoad();

        }

        public ApplicationSettingsPage GoToApplicationSettings()
        {
            var applicationSettingUrl = _testContext.Properties["Winrest_Admin_Settings"].ToString();
            _webDriver.Navigate().GoToUrl(applicationSettingUrl);

            return new ApplicationSettingsPage(_webDriver, _testContext);
        }

        public void CloseNewsPopup()
        {
            try
            {
                // <button id="btnMessageValidation" class="btn btn-primary disabled" onclick="validateMessages()">
                IJavaScriptExecutor executor = (IJavaScriptExecutor)_webDriver;
                executor.ExecuteScript("validateMessages();");
            }
            catch
            {
                //no news, go on
            }
        }

        public MailPage RedirectToOutlookMailbox()
        {
            var javaScriptExecutor = _webDriver as IJavaScriptExecutor;

            // Création d'un nouvel onglet pour ouvrir la messagerie
            javaScriptExecutor.ExecuteScript("window.open('about:blank','_blank');");

            //Switch vers le nouvel onglet et navigation vers la page de connexion GMail
            _webDriver.SwitchTo().Window(_webDriver.WindowHandles.Last());
            _webDriver.Navigate().GoToUrl("https://outlook.office.com/mail/");

            return new MailPage(_webDriver, _testContext);
        }

        //______________________________________________________PURCHASING________________________________________________________________________

        public ItemPage GoToPurchasing_ItemPage()
        {
            WaitPageLoading();
            //if popup news
            CloseNewsPopup();
            try
            {
                // On tente de cliquer sur le menu à partir de la barre des menus
                _purchasing = WaitForElementIsVisibleNew(By.Id(PURCHASING_MENU));
                _purchasing.Click();


                _purchasing_Item = WaitForElementIsVisibleNew(By.Id(PU_ITEM));
                _purchasing_Item.Click();
            }
            catch
            {
                //CloseNewsPopup();

                // Retour page d'accueil
                GoToWinrestHome();

                _purchasing_Item_link = WaitForElementIsVisible(By.Id(PU_ITEM_LINK));
                _purchasing_Item_link.Click();
            }
            WaitForLoad();

            return new ItemPage(_webDriver, _testContext);
        }

        public SuppliersPage GoToPurchasing_SuppliersPage()
        {
            //if popup news
            CloseNewsPopup();

            try
            {

                // On tente de cliquer sur le menu à partir de la barre des menus
                _purchasing = WaitForElementIsVisibleNew(By.Id(PURCHASING_MENU));
                _purchasing.Click();

                _purchasing_Suppliers = WaitForElementIsVisibleNew(By.Id(PU_SUPPLIERS));
                _purchasing_Suppliers.Click();

            }
            catch
            {
                // Retour page d'accueil
                GoToWinrestHome();

                _purchasing_Suppliers_link = WaitForElementIsVisibleNew(By.Id(PU_SUPPLIERS_LINK));
                _purchasing_Suppliers_link.Click();
            }
            WaitForLoad();

            return new SuppliersPage(_webDriver, _testContext);
        }

        public SupplyOrderPage GoToPurchasing_SupplyOrderPage()
        {
            //if popup news
            CloseNewsPopup();

            try
            {
                // On tente de cliquer sur le menu à partir de la barre des menus
                _purchasing = WaitForElementIsVisible(By.Id(PURCHASING_MENU));
                _purchasing.Click();

                _purchasing_SupplyOrders = WaitForElementIsVisible(By.Id(PU_SUPPLY_ORDER));
                _purchasing_SupplyOrders.Click();
            }
            catch
            {
                // Retour page d'accueil
                GoToWinrestHome();

                _purchasing_SupplyOrders_link = WaitForElementIsVisible(By.Id(PU_SUPPLY_ORDER_LINK));
                _purchasing_SupplyOrders_link.Click();
            }
            WaitForLoad();

            return new SupplyOrderPage(_webDriver, _testContext);
        }

        public PurchaseOrdersPage GoToPurchasing_PurchaseOrdersPage()
        {
            //if popup news
            CloseNewsPopup();

            try
            {

                // On tente de cliquer sur le menu à partir de la barre des menus
                _purchasing = WaitForElementIsVisible(By.Id(PURCHASING_MENU));
                _purchasing.Click();

                _purchasing_PurchaseOrders = WaitForElementIsVisible(By.Id(PU_PURCHASE_ORDER));
                _purchasing_PurchaseOrders.Click();

            }
            catch
            {
                // Retour page d'accueil
                GoToWinrestHome();

                _purchasing_PurchaseOrders_link = WaitForElementIsVisible(By.Id(PU_PURCHASE_ORDER_LINK));
                _purchasing_PurchaseOrders_link.Click();

            }
            WaitForLoad();

            return new PurchaseOrdersPage(_webDriver, _testContext);
        }

        public FilterAndFavoritesNeedsPage GoToPurchasing_NeedsPage()
        {
            try
            {

                // On tente de cliquer sur le menu à partir de la barre des menus
                _purchasing = WaitForElementIsVisible(By.Id(PURCHASING_MENU));
                _purchasing.Click();

                _purchasing_Needs = WaitForElementIsVisible(By.Id(PU_NEEDS));
                _purchasing_Needs.Click();

            }
            catch
            {
                // Retour page d'accueil
                GoToWinrestHome();

                _purchasing_Needs_link = WaitForElementIsVisible(By.Id(PU_NEEDS_LINK));
                _purchasing_Needs_link.Click();

            }
            WaitForLoad();

            return new FilterAndFavoritesNeedsPage(_webDriver, _testContext);
        }

        public ProductPage GoToPurchasing_ProductPage()
        {
            //if popup news
            CloseNewsPopup();

            try
            {

                // On tente de cliquer sur le menu à partir de la barre des menus
                _purchasing = WaitForElementIsVisible(By.Id(PURCHASING_MENU));
                _purchasing.Click();

                _product_tab_nav = WaitForElementIsVisible(By.Id(PRODUCT_TAB_NAV));
                _product_tab_nav.Click();

            }
            catch
            {
                // Retour page d'accueil
                GoToWinrestHome();

                _purchasing_PurchaseOrders_link = WaitForElementIsVisible(By.Id(PRODUCT_LINK_DASHBOARD));
                _purchasing_PurchaseOrders_link.Click();

            }
            WaitLoading();

            return new ProductPage(_webDriver, _testContext);
        }

        //______________________________________________________FLIGHTS________________________________________________________________________

        public FlightPage GoToFlights_FlightPage()
        {
            try
            {
                //On tente de cliquer sur le menu à partir de la barre des menus
                _flights = WaitForElementIsVisibleNew(By.Id(FLIGHT_MENU));
                _flights.Click();

                _flights_LoadingPlans = WaitForElementIsVisibleNew(By.Id(FL_FLIGHTS));
                _flights_LoadingPlans.Click();
            }
            catch
            {
                //Retour page d'accueil
                GoToWinrestHome();

                _flights_Flights_link = WaitForElementIsVisibleNew(By.Id(FL_FLIGHTS_LINK));
                _flights_Flights_link.Click();

            }

            WaitForLoad();
            return new FlightPage(_webDriver, _testContext);
        }

        public LoadingPlansPage GoToFlights_LoadingPlansPage()
        {
            try
            {

                // On tente de cliquer sur le menu à partir de la barre des menus
                _flights = WaitForElementIsVisible(By.Id(FLIGHT_MENU));
                _flights.Click();
                WaitForLoad();

                _flights_LoadingPlans = WaitForElementIsVisible(By.Id(FL_LOADING_PLAN));
                _flights_LoadingPlans.Click();


            }
            catch
            {
                // Retour page d'accueil
                GoToWinrestHome();

                _flights_LoadingPlans_link = WaitForElementIsVisible(By.Id(FL_LOADING_PLAN_LINK));
                _flights_LoadingPlans_link.Click();

            }
            WaitForLoad();

            return new LoadingPlansPage(_webDriver, _testContext);
        }

        public SchedulePage GoToFlights_FlightSelectionPage()
        {
            try
            {

                // On tente de cliquer sur le menu à partir de la barre des menus
                _flights = WaitForElementIsVisible(By.Id(FLIGHT_MENU));
                _flights.Click();

                _flights_Flight_Selection = WaitForElementIsVisible(By.Id(FL_FLIGHT_SELECTION));
                _flights_Flight_Selection.Click();


            }
            catch
            {
                // Retour page d'accueil
                GoToWinrestHome();

                _flights_Flight_Selection_link = WaitForElementIsVisible(By.Id(FL_FLIGHT_SELECTION_LINK));
                _flights_Flight_Selection_link.Click();

            }
            WaitForLoad();

            return new SchedulePage(_webDriver, _testContext);
        }

        public CrewPage GoToFlights_CrewPage()
        {
            try
            {

                // On tente de cliquer sur le menu à partir de la barre des menus
                _flights = WaitForElementIsVisible(By.Id(FLIGHT_MENU));
                _flights.Click();

                var flights_Crew = WaitForElementIsVisible(By.Id(FL_CREW));
                flights_Crew.Click();


            }
            catch
            {
                // Retour page d'accueil
                GoToWinrestHome();

                var flights_Crew_link = WaitForElementIsVisible(By.Id(FL_CREW_LINK));
                flights_Crew_link.Click();

            }
            WaitForLoad();

            return new CrewPage(_webDriver, _testContext);
        }

        public LpCartPage GoToFlights_LpCartPage()
        {
            try
            {

                // On tente de cliquer sur le menu à partir de la barre des menus
                _flights = WaitForElementIsVisible(By.Id(FLIGHT_MENU));
                _flights.Click();

                _flights_Flight_LpCart = WaitForElementIsVisible(By.Id(FL_FLIGHT_LP_CART));
                _flights_Flight_LpCart.Click();

            }
            catch
            {
                // Retour page d'accueil
                GoToWinrestHome();

                _flights_Flight_LpCart_link = WaitForElementIsVisible(By.Id(FL_FLIGHT_LP_CART_LINK));
                _flights_Flight_LpCart_link.Click();

            }
            WaitForLoad();

            return new LpCartPage(_webDriver, _testContext);
        }

        public TrolleyLightLabelPage GoToFlights_TrolleyPage()
        {
            try
            {
                // On tente de cliquer sur le menu à partir de la barre des menus
                _flights = WaitForElementIsVisible(By.Id(FLIGHT_MENU));
                _flights.Click();
                WaitForLoad();

                _flights_Flight_Trolley = WaitForElementIsVisible(By.Id(FL_FLIGHT_TROLLEY));
                _flights_Flight_Trolley.Click();
            }
            catch
            {
                // Retour page d'accueil
                GoToWinrestHome();

                _flights_Flight_Trolley_link = WaitForElementIsVisible(By.Id(FL_FLIGHT_TROLLEY_LINK));
                _flights_Flight_Trolley_link.Click();

            }
            WaitForLoad();

            return new TrolleyLightLabelPage(_webDriver, _testContext);
        }

        public TimeBlockPage GoToFlight_TimeBlockPage()
        {
            try
            {
                // On tente de cliquer sur le menu à partir de la barre des menus
                _flights = WaitForElementIsVisible(By.Id(FLIGHT_MENU));
                _flights.Click();

                _flights_Flight_TimeBlock = WaitForElementIsVisible(By.Id(FL_FLIGHT_TIMEBLOCK));
                _flights_Flight_TimeBlock.Click();

            }
            catch
            {
                // Retour page d'accueil
                GoToWinrestHome();

                _flights_Flight_TimeBlock_link = WaitForElementIsVisible(By.Id(FL_FLIGHT_TIMEBLOCK_LINK));
                _flights_Flight_TimeBlock_link.Click();

            }
            WaitForLoad();
            return new TimeBlockPage(_webDriver, _testContext);
        }


        //______________________________________________________MENUS________________________________________________________________________

        public RecipesPage GoToMenus_Recipes()
        {
            CloseNewsPopup();

            try
            {

                // On tente de cliquer sur le menu à partir de la barre des menus
                _menu = WaitForElementIsVisibleNew(By.Id(MENUS_MENU));
                _menu.Click();

                _menu_Recipes = WaitForElementIsVisibleNew(By.Id(ME_RECIPES));
                _menu_Recipes.Click();

            }
            catch
            {
                // Retour page d'accueil
                GoToWinrestHome();

                _menu_Recipes_link = WaitForElementIsVisibleNew(By.Id(ME_RECIPES_LINK));
                _menu_Recipes_link.Click();

            }
            WaitForLoad();

            return new RecipesPage(_webDriver, _testContext);
        }

        public MenusPage GoToMenus_Menus()
        {
            try
            {

                // On tente de cliquer sur le menu à partir de la barre des menus
                _menu = WaitForElementIsVisibleNew(By.Id(MENUS_MENU));
                _menu.Click();

                _menu_Menus = WaitForElementIsVisibleNew(By.Id(ME_MENU));
                _menu_Menus.Click();


            }
            catch
            {
                // Retour page d'accueil
                GoToWinrestHome();

                _menu_Menus_link = WaitForElementIsVisibleNew(By.Id(ME_MENU_LINK));
                _menu_Menus_link.Click();


            }
            WaitForLoad();

            return new MenusPage(_webDriver, _testContext);
        }

        public DatasheetPage GoToMenus_Datasheet()
        {
            try
            {
                var modal = _webDriver.FindElement(By.Id(NEWS));
                if (modal.Displayed)
                {
                    var validateButton = WaitForElementIsVisible(By.Id(VALIDATE_NEWS));
                    //validateButton.Click();
                    WebDriverWait wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(10));
                    //SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath(PRICESERVICE))
                    wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.Id(VALIDATE_NEWS)));
                    validateButton.Click();

                }
                // On tente de cliquer sur le menu à partir de la barre des menus
                _menu = WaitForElementIsVisible(By.Id(MENUS_MENU));
                _menu.Click();

                _menu_Datasheet = WaitForElementIsVisible(By.Id(ME_DATASHEET));
                _menu_Datasheet.Click();

            }
            catch
            {
                // Retour page d'accueil
                GoToWinrestHome();

                _menu_Datasheet_link = WaitForElementIsVisible(By.Id(ME_DATASHEET_LINK));
                _menu_Datasheet_link.Click();

            }
            WaitForLoad();

            return new DatasheetPage(_webDriver, _testContext);
        }

        public FoodCostPage GoToMenus_FoodCost()
        {
            try
            {
                // On tente de cliquer sur le menu à partir de la barre des menus
                _menu = WaitForElementIsVisible(By.Id(MENUS_MENU));
                _menu.Click();
                var foodCost = WaitForElementIsVisible(By.Id(ME_FOODCOST));
                foodCost.Click();
            }
            catch
            {
                // Retour page d'accueil
                GoToWinrestHome();
                var foodCostlink = WaitForElementIsVisible(By.Id(ME_FOODCOST_LINK));
                foodCostlink.Click();
            }
            WaitForLoad();
            return new FoodCostPage(_webDriver, _testContext);
        }


        //______________________________________________________TODOLIST________________________________________________________________________
        public SchedulerPage GoToDoList_Scheduler()
        {
            try
            {

                // On tente de cliquer sur le menu à partir de la barre des menus
                _toDoList = WaitForElementIsVisible(By.Id(TODOLIST));
                _toDoList.Click();

                _schedulerTab = WaitForElementIsVisible(By.Id(SCHEDULER_TAB));
                _schedulerTab.Click();

            }
            catch
            {
                // Retour page d'accueil
                GoToWinrestHome();

                _schedulerLink = WaitForElementIsVisible(By.Id(SCHEDULER_LINK));
                _schedulerLink.Click();


            }
            WaitForLoad();

            return new SchedulerPage(_webDriver, _testContext);
        }

        public TasksPage GoToDoList_Tasks()
        {
            try
            {

                // On tente de cliquer sur le menu à partir de la barre des menus
                _toDoList = WaitForElementIsVisibleNew(By.Id(TODOLIST));
                _toDoList.Click();

                _taskTab = WaitForElementIsVisibleNew(By.Id(TASK_TAB));
                _taskTab.Click();

            }
            catch
            {
                // Retour page d'accueil
                GoToWinrestHome();

                _taskLink = WaitForElementIsVisible(By.Id(TASK_LINK));
                _taskLink.Click();
            }
            WaitForLoad();

            return new TasksPage(_webDriver, _testContext);
        }

        //______________________________________________________PRODUCTION________________________________________________________________________

        public CustomerOrderPage GoToProduction_CustomerOrderPage()
        {
            try
            {
                // On tente de cliquer sur le menu à partir de la barre des menus

                _production = WaitForElementIsVisible(By.Id(PRODUCTION_MENU));
                _production.Click();
                WaitForLoad();

                _production_Customer_Order = WaitForElementIsVisible(By.Id(PR_CUSTOMER_ORDER));
                _production_Customer_Order.Click();

            }
            catch
            {
                // Retour page d'accueil
                GoToWinrestHome();

                _production_Customer_Order_link = WaitForElementIsVisible(By.Id(PR_CUSTOMER_ORDER_LINK));
                _production_Customer_Order_link.Click();

            }
            WaitForLoad();

            return new CustomerOrderPage(_webDriver, _testContext);
        }

        public DispatchPage GoToProduction_DispatchPage()
        {
            try
            {

                _production = WaitForElementIsVisible(By.Id(PRODUCTION_MENU));
                _production.Click();
                WaitForLoad();

                _production_dipatch = WaitForElementIsVisible(By.Id(PR_DISPATCH));
                _production_dipatch.Click();
                WaitPageLoading();
            }
            catch
            {
                // Retour page d'accueil
                GoToWinrestHome();

                _production_dipatch_link = WaitForElementIsVisible(By.Id(PR_DISPATCH_LINK));
                _production_dipatch_link.Click();

            }
            WaitPageLoading();

            return new DispatchPage(_webDriver, _testContext);
        }

        public ProductionPage GoToProduction_ProductionPage()
        {
            try
            {
                // On tente de cliquer sur le menu à partir de la barre des menus

                _production = WaitForElementIsVisible(By.Id(PRODUCTION_MENU));
                _production.Click();
                WaitForLoad();

                _production_production = WaitForElementIsVisible(By.Id(PR_PRODUCTION));
                _production_production.Click();

            }
            catch
            {
                // Retour page d'accueil
                GoToWinrestHome();

                _production_production_link = WaitForElementIsVisible(By.Id(PR_PRODUCTION_LINK));
                _production_production_link.Click();

            }
            WaitForLoad();

            return new ProductionPage(_webDriver, _testContext);
        }

        public SetupFilterAndFavoritesPage GoToProduction_SetupPage()
        {
            try
            {
                // On tente de cliquer sur production à partir de la barre des menus

                _production = WaitForElementIsVisible(By.Id(PRODUCTION_MENU));
                _production.Click();
                WaitForLoad();

                _production_setup = WaitForElementIsVisible(By.Id(PR_SETUP));
                _production_setup.Click();
            }
            catch
            {
                // Retour page d'accueil
                GoToWinrestHome();

                _production_setup_link = WaitForElementIsVisible(By.Id(PR_SETUP_LINK));
                _production_setup_link.Click();
            }
            WaitForLoad();

            return new SetupFilterAndFavoritesPage(_webDriver, _testContext);
        }

        public FilterAndFavoritesPage GoToProduction_ProductionManagemenentPage()
        {
            try
            {

                // On tente de cliquer sur le menu à partir de la barre des menus

                _production = WaitForElementIsVisible(By.Id(PRODUCTION_MENU));
                _production.Click();
                WaitForLoad();

                _production_productionManagement = WaitForElementIsVisible(By.Id(PR_PRODUCTION_MANAGEMENT));
                _production_productionManagement.Click();

            }
            catch
            {
                // Retour page d'accueil
                GoToWinrestHome();

                _production_productionManagement_link = WaitForElementIsVisible(By.Id(PR_PRODUCTION_MANAGEMENT_LINK));
                _production_productionManagement_link.Click();

            }
            WaitForLoad();

            return new FilterAndFavoritesPage(_webDriver, _testContext);
        }

        public EarlyProductionPage GoToProduction_EarlyProduction()
        {
            try
            {
                // On tente de cliquer sur le menu à partir de la barre des menus

                _production = WaitForElementIsVisible(By.Id(PRODUCTION_MENU));
                _production.Click();
                WaitForLoad();

                _production_production = WaitForElementIsVisible(By.Id(PR_EARLY_PRODUCTION));
                _production_production.Click();

            }
            catch
            {
                // Retour page d'accueil
                GoToWinrestHome();

                _production_production_link = WaitForElementIsVisible(By.Id(PR_EARLY_PRODUCTION_LINK));
                _production_production_link.Click();

            }
            WaitForLoad();

            return new EarlyProductionPage(_webDriver, _testContext);
        }

        public ProductionCOPage GoToProduction_ProductionCOPageModified()
        {
            try
            {
                _production_Production_CO = WaitForElementIsVisible(By.XPath(PRODUCTION_MENU_LIST));
                _production_Production_CO.Click();

            }
            catch
            {
                // Retour page d'accueil
                GoToWinrestHome();

                _production_Production_CO_link = WaitForElementIsVisible(By.XPath(PR_PRODUCTION_CO_LINK));
                _production_Production_CO_link.Click();

            }
            WaitForLoad();

            return new ProductionCOPage(_webDriver, _testContext);
        }

        public ProductionCOPage GoToProduction_ProductionCOPage()
        {
            try
            {
                // On tente de cliquer sur le menu à partir de la barre des menus

                _production = WaitForElementIsVisible(By.Id(PRODUCTION_MENU));
                _production.Click();
                WaitForLoad();

                _production_Production_CO = WaitForElementIsVisible(By.Id(PR_PRODUCTION_CO));
                _production_Production_CO.Click();

            }
            catch
            {
                // Retour page d'accueil
                GoToWinrestHome();

                _production_Production_CO_link = WaitForElementIsVisible(By.XPath(PR_PRODUCTION_CO_LINK));
                _production_Production_CO_link.Click();

            }
            WaitForLoad();

            return new ProductionCOPage(_webDriver, _testContext);
        }

        //______________________________________________________WAREHOUSE________________________________________________________________________

        public ReceiptNotesPage GoToWarehouse_ReceiptNotesPage()
        {
            try
            {
                // On tente de cliquer sur le menu à partir de la barre des menus
                var wareHouse = WaitForElementIsVisibleNew(By.Id("TabWarehouse"));
                wareHouse.Click();

                var wareHouse_Receipt_Note = WaitForElementIsVisibleNew(By.Id("ReceiptNoteTabNav"));
                wareHouse_Receipt_Note.Click();
            }
            catch
            {
                // Retour page d'accueil
                GoToWinrestHome();

                _wareHouse_Receipt_Note_link = WaitForElementIsVisibleNew(By.Id("ReceiptNoteLinkDashBoard"));
                _wareHouse_Receipt_Note_link.Click();
            }
            WaitForLoad();

            return new ReceiptNotesPage(_webDriver, _testContext);
        }

        public ClaimsPage GoToWarehouse_ClaimsPage()
        {
            try
            {

                // On tente de cliquer sur le menu à partir de la barre des menus
                _wareHouse = WaitForElementIsVisibleNew(By.XPath(WAREHOUSE_MENU));
                _wareHouse.Click();

                _wareHouse_Claim = WaitForElementIsVisibleNew(By.XPath(WA_CLAIM));
                _wareHouse_Claim.Click();
            }
            catch
            {
                // Si le menu n'est pas accessible, clic sur le lien dans la page d'accueil
                GoToWinrestHome();

                var claim = WaitForElementIsVisibleNew(By.Id("ReceiptNoteClaimsLinkDashBoard"));
                claim.Click();

            }

            WaitForLoad();

            return new ClaimsPage(_webDriver, _testContext);
        }

        public InventoriesPage GoToWarehouse_InventoriesPage()
        {
            try
            {
                // On tente de cliquer sur le menu à partir de la barre des menus
                var wareHouse = WaitForElementIsVisibleNew(By.Id("TabWarehouse"));
                wareHouse.Click();

                var inventory = WaitForElementIsVisibleNew(By.Id("InventoryTabNav"));
                inventory.Click();

            }
            catch
            {
                // Si le menu n'est pas accessible, clic sur le lien dans la page d'accueil
                GoToWinrestHome();

                var inv = WaitForElementIsVisibleNew(By.Id("InventoryLinkDashBoard"));
                inv.Click();

            }

            WaitForLoad();

            return new InventoriesPage(_webDriver, _testContext);
        }

        public OutputFormPage GoToWarehouse_OutputFormPage()
        {
            WaitForLoad();
            try
            {
                // On tente de cliquer sur le menu à partir de la barre des menus
                var wareHouse = WaitForElementIsVisibleNew(By.CssSelector("#TabWarehouse"));
                wareHouse.Click();

                var Output_Form = WaitForElementIsVisibleNew(By.CssSelector("#OutputFormTabNav"));
                Output_Form.Click();

            }
            catch
            {
                // Si le menu n'est pas accessible, clic sur le lien dans la page d'accueil
                GoToWinrestHome();
                var output = WaitForElementIsVisibleNew(By.Id("OutputFormLinkDashBoard"));
                output.Click();

            }

            WaitForLoad();

            return new OutputFormPage(_webDriver, _testContext);
        }


        //______________________________________________________CUSTOMERS________________________________________________________________________

        public CustomerPage GoToCustomers_CustomerPage()
        {
            try
            {
                // On tente de cliquer sur le menu à partir de la barre des menus
                var customer = WaitForElementIsVisible(By.Id("TabCustomers"));
                customer.Click();

                var customer_Customer = WaitForElementIsVisible(By.Id("CustomerTabNav"));
                customer_Customer.Click();

            }
            catch
            {
                // Si le menu n'est pas accessible, clic sur le lien dans la page d'accueil
                GoToWinrestHome();
                var customer_Customer_link = WaitForElementIsVisible(By.Id("CustomerLinkDashBoard"));
                customer_Customer_link.Click();

            }

            WaitForLoad();

            return new CustomerPage(_webDriver, _testContext);
        }

        public ReinvoicePage GoToCustomers_ReinvoicePage()
        {
            try
            {
                // On tente de cliquer sur le menu à partir de la barre des menus
                var customer = WaitForElementIsVisible(By.Id("TabCustomers"));
                customer.Click();

                var customer_Reinvoice = WaitForElementIsVisible(By.Id("ReinvoiceTabNav"));
                customer_Reinvoice.Click();

            }
            catch
            {
                // Si le menu n'est pas accessible, clic sur le lien dans la page d'accueil
                GoToWinrestHome();
                var customer_Reinvoice_link = WaitForElementIsVisible(By.Id("ReinvoiceLinkDashBoard"));
                customer_Reinvoice_link.Click();

            }

            WaitForLoad();
            return new ReinvoicePage(_webDriver, _testContext);
        }

        public DeliveryPage GoToCustomers_DeliveryPage()
        {
            try
            {
                // On tente de cliquer sur le menu à partir de la barre des menus
                var customer = WaitForElementIsVisible(By.Id("TabCustomers"));
                customer.Click();

                var customer_Delivery = WaitForElementIsVisible(By.Id("FlightDeliveryTabNav"));
                customer_Delivery.Click();

            }
            catch
            {
                // Si le menu n'est pas accessible, clic sur le lien dans la page d'accueil
                GoToWinrestHome();
                var customer_Delivery_link = WaitForElementIsVisible(By.Id("FlightDeliveryLinkDashBoard"));
                customer_Delivery_link.Click();

            }

            WaitForLoad();

            return new DeliveryPage(_webDriver, _testContext);
        }

        public DeliveryRoundPage GoToCustomers_DeliveryRoundPage()
        {
            try
            {
                // On tente de cliquer sur le menu à partir de la barre des menus
                var customer = WaitForElementIsVisible(By.Id("TabCustomers"));
                customer.Click();

                var customer_Delivery_Round = WaitForElementIsVisible(By.Id("DeliveryRoundTabNav"));
                customer_Delivery_Round.Click();

            }
            catch
            {
                // Si le menu n'est pas accessible, clic sur le lien dans la page d'accueil
                GoToWinrestHome();
                _customer_Delivery_Round_link = WaitForElementIsVisible(By.Id("DeliveryRoundLinkDashBoard"));
                _customer_Delivery_Round_link.Click();

            }

            WaitForLoad();

            return new DeliveryRoundPage(_webDriver, _testContext);
        }

        public PriceListPage GoToCustomers_PriceListPage()
        {
            try
            {
                // On tente de cliquer sur le menu à partir de la barre des menus
                var customer = WaitForElementIsVisible(By.Id("TabCustomers"));
                customer.Click();

                var customer_Price_List = WaitForElementIsVisible(By.Id("PriceTabNav"));
                customer_Price_List.Click();

            }
            catch
            {
                // Si le menu n'est pas accessible, clic sur le lien dans la page d'accueil
                GoToWinrestHome();

                var customer_Price_List_link = WaitForElementIsVisible(By.Id("PriceLinkDashBoard"));
                customer_Price_List_link.Click();

            }

            WaitForLoad();

            return new PriceListPage(_webDriver, _testContext);
        }

        public FoodPacketPage GoToCustomers_FoodPacketsPage()
        {
            try
            {
                // On tente de cliquer sur le menu à partir de la barre des menus
                var customer = WaitForElementIsVisible(By.Id("TabCustomers"));
                customer.Click();

                var customer_foodPackets = WaitForElementIsVisible(By.Id("FoodPacketTabNav"));
                customer_foodPackets.Click();

            }
            catch
            {
                // Si le menu n'est pas accessible, clic sur le lien dans la page d'accueil
                GoToWinrestHome();
                var customer_FoodPacket_link = WaitForElementIsVisible(By.XPath(CU_FOOD_PACKET_LINK));
                customer_FoodPacket_link.Click();

            }

            WaitForLoad();

            return new FoodPacketPage(_webDriver, _testContext);
        }

        public ServicePage GoToCustomers_ServicePage()
        {
            try
            {
                // On tente de cliquer sur le menu à partir de la barre des menus
                var customer = WaitForElementIsVisibleNew(By.Id("TabCustomers"));
                customer.Click();

                var customer_Service = WaitForElementIsVisibleNew(By.Id("ServiceTabNav"));
                customer_Service.Click();

            }
            catch
            {
                // Si le menu n'est pas accessible, clic sur le lien dans la page d'accueil
                GoToWinrestHome();
                var customer_Service_link = WaitForElementIsVisibleNew(By.Id("ServiceLinkDashBoard"));
                customer_Service_link.Click();

            }

            WaitForLoad();

            return new ServicePage(_webDriver, _testContext);
        }

        public ReconciliationPage GoToCustomers_ReconPage()
        {
            try
            {
                // On tente de cliquer sur le menu à partir de la barre des menus
                var customer = WaitForElementIsVisible(By.Id("TabCustomers"));
                customer.Click();

                var customer_Service = WaitForElementIsVisible(By.Id("ReconciliationTabNav"));
                customer_Service.Click();

            }
            catch
            {
                // Si le menu n'est pas accessible, click sur le lien dans la page d'accueil
                GoToWinrestHome();
                var customer_recon = WaitForElementIsVisible(By.Id("ReconciliationLinkDashBoard"));
                customer_recon.Click();

            }

            WaitForLoad();

            return new ReconciliationPage(_webDriver, _testContext);
        }

        public CatalogPage GoToCustomers_CatalogPage()
        {
            try
            {
                // On tente de cliquer sur le menu à partir de la barre des menus
                var customer = WaitForElementIsVisible(By.Id("TabCustomers"));
                customer.Click();

                var customer_Catalog = WaitForElementIsVisible(By.Id("CatalogTabNav"));
                customer_Catalog.Click();

            }
            catch
            {

                // Si le menu n'est pas accessible, click sur le lien dans la page d'accueil
                GoToWinrestHome();
                var customer_catalog = WaitForElementIsVisible(By.Id("CatalogLinkDashBoard"));
                customer_catalog.Click();

            }

            WaitForLoad();

            return new CatalogPage(_webDriver, _testContext);

        }

        //______________________________________________________ACCOUNTING________________________________________________________________________

        public SupplierInvoicesPage GoToAccounting_SupplierInvoices()
        {
            try
            {
                // On tente de cliquer sur le menu à partir de la barre des menus

                _accounting = WaitForElementIsVisible(By.Id(ACCOUNTING_MENU));
                _accounting.Click();

                _accounting_Supplier_Invoice = WaitForElementIsVisible(By.Id(AC_SUPPLIER_INVOICE));
                _accounting_Supplier_Invoice.Click();
            }
            catch
            {
                // Si le menu n'est pas accessible, clic sur le lien dans la page d'accueil
                GoToWinrestHome();

                _accounting_Supplier_Invoice_link = WaitForElementIsVisible(By.Id(AC_SUPPLIER_INVOICE_LINK));
                _accounting_Supplier_Invoice_link.Click();
            }
            WaitForLoad();

            return new SupplierInvoicesPage(_webDriver, _testContext);
        }

        public InvoicesPage GoToAccounting_InvoicesPage()
        {
            try
            {
                // On tente de cliquer sur le menu à partir de la barre des menus
                _accounting = WaitForElementIsVisible(By.Id(ACCOUNTING_MENU));
                _accounting.Click();

                _accounting_Invoice = WaitForElementIsVisible(By.Id(AC_INVOICE));
                _accounting_Invoice.Click();
            }
            catch
            {
                // Si le menu n'est pas accessible, clic sur le lien dans la page d'accueil
                GoToWinrestHome();

                _accounting_Invoice_link = WaitForElementIsVisible(By.Id(AC_INVOICE_LINK));
                _accounting_Invoice_link.Click();
            }

            WaitForLoad();

            return new InvoicesPage(_webDriver, _testContext);
        }

        public FreePricePage GoToAccounting_FreePricePage()
        {
            try
            {
                // On tente de cliquer sur le menu à partir de la barre des menus
                _accounting = WaitForElementIsVisible(By.Id(ACCOUNTING_MENU));
                _accounting.Click();

                _accounting_FreePrice = WaitForElementIsVisible(By.Id(AC_FREE_PRICE));
                _accounting_FreePrice.Click();
            }
            catch
            {
                // Si le menu n'est pas accessible, clic sur le lien dans la page d'accueil
                GoToWinrestHome();

                _accounting_FreePrice_link.Click();
            }

            WaitForLoad();

            return new FreePricePage(_webDriver, _testContext);
        }

        public ReportingPage GoToAccounting_Reporting()
        {
            try
            {
                // On tente de cliquer sur le menu à partir de la barre des menus
                _accounting = WaitForElementIsVisible(By.Id(ACCOUNTING_MENU));
                _accounting.Click();

                _accounting_Reporting = WaitForElementIsVisible(By.Id(AC_REPORTING));
                _accounting_Reporting.Click();
            }
            catch
            {
                // Si le menu n'est pas accessible, clic sur le lien dans la page d'accueil
                GoToWinrestHome();

                _accounting_Reporting_link.Click();
            }

            WaitForLoad();

            return new ReportingPage(_webDriver, _testContext);

        }

        //______________________________________________________INTERIM________________________________________________________________________
        public InterimOrderPage GoToInterim_Orders()
        {
            // On tente de cliquer sur le menu à partir de la barre des menus
            var interimTab = WaitForElementIsVisible(By.Id("TabMenuInterim"));
            interimTab.Click();

            var interimOrderimTabNav = WaitForElementIsVisible(By.Id("InterimOrderTabNav"));
            interimOrderimTabNav.Click();

            return new InterimOrderPage(_webDriver, _testContext);
        }

        public InterimReceptionsPage GoToInterim_Receptions()
        {
            // On tente de cliquer sur le menu à partir de la barre des menus
            var interimTab = WaitForElementIsVisible(By.Id("TabMenuInterim"));
            WaitForLoad();
            interimTab.Click();

            var interimReceptionTabNav = WaitForElementIsVisible(By.Id("InterimReceptionTabNav"));
            interimReceptionTabNav.Click();

            return new InterimReceptionsPage(_webDriver, _testContext);
        }

        public InterimReceptionsPage GoToInterim_ReceptionsModified()
        {
            // On tente de cliquer sur le menu à partir de la barre des menus
            var interimTab = WaitForElementIsVisible(By.XPath("//*[@id=\"InterimReceptionLinkDashBoard\"]"));
            interimTab.Click();
            return new InterimReceptionsPage(_webDriver, _testContext);
        }

        //______________________________________________________CLINIC________________________________________________________________________

        public PatientsPage GoToClinic_PatientPage()
        {
            try
            {
                // On tente de cliquer sur le menu à partir de la barre des menus
                _clinicTab = WaitForElementIsVisible(By.Id(CLINIC_TAB));
                _clinicTab.Click();

                _clinicPatientTab = WaitForElementIsVisible(By.Id(CLINIC_PATIENT_TAB));
                _clinicPatientTab.Click();
            }
            catch
            {
                // Si le menu n'est pas accessible, clic sur le lien dans la page d'accueil
                GoToWinrestHome();
                _clinicPatientLink = WaitForElementIsVisible(By.Id(CLINIC_PATIENT_LINK));
                _clinicPatientLink.Click();
            }

            WaitForLoad();

            return new PatientsPage(_webDriver, _testContext);
        }

        //______________________________________________________PARAMETRES________________________________________________________________________

        public ParametersUser GoToParameters_User()
        {
            try
            {
                // On tente de cliquer sur le menu à partir de la barre des menus
                var parameters = WaitForElementIsVisible(By.Id("TabParameters"));
                parameters.Click();

                var parameters_User = WaitForElementIsVisible(By.Id("UserManagementTabNav"));
                parameters_User.Click();

            }
            catch
            {
                // Si le menu n'est pas accessible, clic sur le lien dans la page d'accueil
                GoToWinrestHome();

                var parameters_User_link = WaitForElementIsVisible(By.Id("UserManagementLinkDashBoard"));
                parameters_User_link.Click();

            }

            WaitPageLoading();
            WaitForLoad();

            return new ParametersUser(_webDriver, _testContext);
        }

        public ParametersPortal GoToParameters_PortalPage()
        {
            try
            {
                // On tente de cliquer sur le menu à partir de la barre des menus
                var parameters = WaitForElementIsVisible(By.Id("TabParameters"));
                parameters.Click();

                var parameters_Portal = WaitForElementIsVisible(By.Id("CustomerPortalTabNav"));
                parameters_Portal.Click();

            }
            catch
            {
                // Si le menu n'est pas accessible, clic sur le lien dans la page d'accueil
                GoToWinrestHome();
                var parameters_Portal_link = WaitForElementIsVisible(By.Id("CustomerPortalLinkDashBoard"));
                parameters_Portal_link.Click();
            }

            WaitForLoad();

            return new ParametersPortal(_webDriver, _testContext);
        }

        public ParametersProduction GoToParameters_ProductionPage()
        {
            try
            {
                // On tente de cliquer sur le menu à partir de la barre des menus
                var parameters = WaitForElementIsVisibleNew(By.Id("TabParameters"));
                parameters.Click();

                var parameters_Production = WaitForElementIsVisibleNew(By.Id("SettingsProductionTabNav"));
                parameters_Production.Click();

            }
            catch
            {
                // Si le menu n'est pas accessible, clic sur le lien dans la page d'accueil
                GoToWinrestHome();
                var parameters_Production_link = WaitForElementIsVisibleNew(By.Id("SettingsProductionLinkDashBoard"));
                parameters_Production_link.Click();

            }

            WaitForLoad();

            return new ParametersProduction(_webDriver, _testContext);
        }

        public ParametersSetup GoToParameters_SetupPage()
        {
            try
            {
                // On tente de cliquer sur le menu à partir de la barre des menus
                var parameters = WaitForElementIsVisible(By.Id("TabParameters"));
                parameters.Click();

                var parameters_Production = WaitForElementIsVisible(By.Id("SettingsSetupTabNav"));
                parameters_Production.Click();

            }
            catch
            {
                // Si le menu n'est pas accessible, clic sur le lien dans la page d'accueil
                GoToWinrestHome();
                var parameters_Production_link = WaitForElementIsVisible(By.Id("SettingsSetupLinkDashBoard"));
                parameters_Production_link.Click();

            }

            WaitForLoad();

            return new ParametersSetup(_webDriver, _testContext);
        }

        public ParametersPurchasing GoToParameters_PurchasingPage()
        {
            try
            {

                // On tente de cliquer sur le menu à partir de la barre des menus
                var parameters = WaitForElementIsVisible(By.Id("TabParameters"));
                parameters.Click();

                _parameters_Purchasing = WaitForElementIsVisible(By.Id("PurchasingTabNav"));
                _parameters_Purchasing.Click();
            }
            catch
            {
                // Si le menu n'est pas accessible, clic sur le lien dans la page d'accueil
                GoToWinrestHome();
                var parameters_Purchasing_link = WaitForElementIsVisible(By.Id("PurchasingLinkDashBoard"));
                parameters_Purchasing_link.Click();

            }

            WaitForLoad();

            return new ParametersPurchasing(_webDriver, _testContext);
        }

        public ParametersCustomer GoToParameters_CustomerPage()
        {
            try
            {
                // On tente de cliquer sur le menu à partir de la barre des menus
                var parameters = WaitForElementIsVisible(By.Id("TabParameters"));
                parameters.Click();

                var parameters_Customer = WaitForElementIsVisible(By.Id("CustomersTabNav"));
                parameters_Customer.Click();

            }
            catch
            {
                // Si le menu n'est pas accessible, clic sur le lien dans la page d'accueil
                GoToWinrestHome();
                var parameters_Customer_link = WaitForElementIsVisible(By.Id("CustomersLinkDashBoard"));
                parameters_Customer_link.Click();

            }
            WaitForLoad();

            return new ParametersCustomer(_webDriver, _testContext);
        }

        public ParametersAccounting GoToParameters_AccountingPage()
        {
            try
            {
                // On tente de cliquer sur le menu à partir de la barre des menus
                var parameters = WaitForElementIsVisible(By.Id("TabParameters"));
                parameters.Click();

                var parameters_Accounting = WaitForElementIsVisible(By.Id("SettingsAccountingTabNav"));
                parameters_Accounting.Click();

            }
            catch
            {
                // Si le menu n'est pas accessible, clic sur le lien dans la page d'accueil
                GoToWinrestHome();
                var parameters_Accounting_link = WaitForElementIsVisible(By.Id("SettingsAccountingLinkDashBoard"));
                parameters_Accounting_link.Click();

            }

            WaitForLoad();

            return new ParametersAccounting(_webDriver, _testContext);
        }

        public ParametersSites GoToParameters_Sites()
        {
            //if popup news
            CloseNewsPopup();

            try
            {

                // On tente de cliquer sur le menu à partir de la barre des menus
                var customer = WaitForElementIsVisible(By.Id("TabParameters"));
                customer.Click();
                WaitForLoad();


                var parametersSites = WaitForElementIsVisible(By.Id("SiteTabNav"));
                parametersSites.Click();

            }
            catch
            {
                //if popup news
                CloseNewsPopup();

                // Si le menu n'est pas accessible, clic sur le lien dans la page d'accueil
                GoToWinrestHome();
                _parameters_Sites_link = WaitForElementIsVisible(By.Id("SiteLinkDashBoard"));
                _parameters_Sites_link.Click();
            }

            WaitPageLoading();
            WaitForLoad();

            return new ParametersSites(_webDriver, _testContext);
        }

        public ParametersGlobalSettings GoToParameters_GlobalSettings()
        {
            //if popup news
            CloseNewsPopup();

            try
            {
                // On tente de cliquer sur le menu à partir de la barre des menus
                var parameters = WaitForElementIsVisible(By.Id("TabParameters"));
                parameters.Click();

                var parameters_GlobalSettings = WaitForElementIsVisible(By.Id("GlobalSettingsTabNav"));
                parameters_GlobalSettings.Click();

            }
            catch
            {
                CloseNewsPopup();
                // Si le menu n'est pas accessible, clic sur le lien dans la page d'accueil
                GoToWinrestHome();
                var parameters_GlobalSettings_link = WaitForElementIsVisible(By.Id("GlobalSettingsLinkDashBoard"));
                parameters_GlobalSettings_link.Click();

            }

            WaitForLoad();

            return new ParametersGlobalSettings(_webDriver, _testContext);
        }

        public ParametersClinic GoToParameters_Clinic()
        {
            try
            {
                // On tente de cliquer sur le menu à partir de la barre des menus
                _parameters = WaitForElementIsVisible(By.XPath(PARAMETERS_MENU));
                _parameters.Click();

                _parameters_Clinic = WaitForElementIsVisible(By.Id(PA_CLINIC));
                _parameters_Clinic.Click();
            }
            catch
            {
                // Si le menu n'est pas accessible, clic sur le lien dans la page d'accueil
                GoToWinrestHome();
                _parameters_Sites_link = WaitForElementIsVisible(By.Id(PA_CLINIC_LINK));
                _parameters_Sites_link.Click();
            }

            WaitForLoad();

            return new ParametersClinic(_webDriver, _testContext);
        }
        public ParametersTablet GoToParametres_Tablet()
        {
            try
            {
                // On tente de cliquer sur le menu à partir de la barre des menus
                _parameters = WaitForElementIsVisible(By.XPath(PARAMETERS_MENU));
                _parameters.Click();

                var tabletBtn = WaitForElementIsVisible(By.XPath("//*[@id=\"TabletSettingsTabNav\"]/a"));
                tabletBtn.Click();
            }
            catch
            {
                // Si le menu n'est pas accessible, clic sur le lien dans la page d'accueil
                GoToWinrestHome();
                _parameters_Sites_link = WaitForElementIsVisible(By.Id("TabletSettingsLinkDashBoard"));
                _parameters_Sites_link.Click();
            }

            WaitForLoad();

            return new ParametersTablet(_webDriver, _testContext);
        }

        public ParametresFlights GoToParametres_Flights()
        {
            // On tente de cliquer sur le menu à partir de la barre des menus
            _parameters = WaitForElementIsVisible(By.XPath(PARAMETERS_MENU));
            _parameters.Click();

            var flightsBtn = WaitForElementIsVisible(By.XPath("//*[@id=\"SettingsFlightTabNav\"]/a"));
            flightsBtn.Click();
            WaitForLoad();

            return new ParametresFlights(_webDriver, _testContext);
        }
        //_____________________________________________________________CUSTOMER PORTAL_________________________________________________

        public CustomerPortalLoginPage GoToCustomerPortal()
        {
            Actions action = new Actions(_webDriver);

            _customerPortal = WaitForElementExists(By.XPath(CUSTOMER_PORTAL));
            action.MoveToElement(_customerPortal).Perform();
            _customerPortal.Click();
            WaitForLoad();

            //Results are opened in a new tab, switch the driver to the newly created one
            var wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(60));
            wait.Until((driver) => driver.WindowHandles.Count > 1);
            _webDriver.SwitchTo().Window(_webDriver.WindowHandles.Last());

            WaitForLoad();

            return new CustomerPortalLoginPage(_webDriver, _testContext);
        }

        public TabletAppPage GotoTabletApp()
        {
            Actions action = new Actions(_webDriver);

            _tabletApp = WaitForElementExists(By.XPath(TABLET_APP));
            action.MoveToElement(_tabletApp).Perform();
            _tabletApp.Click();

            return new TabletAppPage(_webDriver, _testContext);
        }

        public EdiPage GoToAccounting_EdiPage()
        {
            Actions action = new Actions(_webDriver);

            _edi = WaitForElementExists(By.Id(EDI));
            action.MoveToElement(_edi).Perform();
            _edi.Click();

            WaitPageLoading();
            return new EdiPage(_webDriver, _testContext);
        }

        //______________________________________________________JOBS________________________________________________________________________
        public SettingsPage GoToJobs_Settings()
        {
            //if popup news
            CloseNewsPopup();

            try
            {

                // On tente de cliquer sur le menu à partir de la barre des menus
                _jobs = WaitForElementIsVisible(By.XPath(JOBS));
                _jobs.Click();

                _jobs_settings = WaitForElementIsVisible(By.XPath(JOBS_SETTINGS_TAB));
                _jobs_settings.Click();

            }
            catch
            {
                // Retour page d'accueil
                GoToWinrestHome();

                // On tente de cliquer sur le menu à partir de la barre des menus
                _settings = WaitForElementIsVisible(By.Id(JOBS_SETTINGS_LINK));
                _settings.Click();
            }
            WaitForLoad();

            return new SettingsPage(_webDriver, _testContext);
        }

        public ScheduledJobsPage GoToJobs_ScheduledJobs()
        {
            //if popup news
            CloseNewsPopup();

            try
            {

                // On tente de cliquer sur le menu à partir de la barre des menus
                _jobs = WaitForElementIsVisible(By.XPath(JOBS));
                _jobs.Click();

                _scheduledJobs = WaitForElementIsVisible(By.XPath(JOBS_SCHEDULED_JOBS_TAB));
                _scheduledJobs.Click();

            }
            catch
            {
                // Retour page d'accueil
                GoToWinrestHome();

                // On tente de cliquer sur le menu à partir de la barre des menus
                _jobs_scheduledJobs = WaitForElementIsVisible(By.Id(JOBS_SCHEDULED_JOBS_LINK));
                _jobs_scheduledJobs.Click();
            }
            WaitForLoad();

            return new ScheduledJobsPage(_webDriver, _testContext);
        }
        public JobsPage GoToJobs_Jobs()
        {
            //if popup news
            CloseNewsPopup();

            try
            {

                // On tente de cliquer sur le menu à partir de la barre des menus
                _jobs = WaitForElementIsVisible(By.XPath(JOBS));
                _jobs.Click();

                _jobs_jobs = WaitForElementIsVisible(By.Id(JOBS_JOBS));
                _jobs_jobs.Click();

            }
            catch
            {
                // Retour page d'accueil
                GoToWinrestHome();

                // On tente de cliquer sur le menu à partir de la barre des menus
                _jobs_jobs_link = WaitForElementIsVisible(By.Id(JOBS_JOBS_LINK));
                _jobs_jobs_link.Click();
            }
            WaitForLoad();

            return new JobsPage(_webDriver, _testContext);
        }

        //_____________________________________________________________DIVERS__________________________________________________________

        public void Close()
        {
            if (_webDriver.WindowHandles.Count > 1)
            {
                _webDriver.Close();
                _webDriver.SwitchTo().Window(_webDriver.WindowHandles.Last());
            }
        }
        public bool isPageSizeEqualsTo100()
        {
            if (CheckTotalNumber() == 0)
            {
                throw new Exception("Pas de données sur l'index, donc pas de NbPages visible");
            }
            var nbPages = WaitForElementExists(By.XPath("//option[text()='100']"));
            return nbPages.Selected;
        }
        public bool isPageSizeEqualsTo100WidhoutTotalNumber()
        {
            if (isElementVisible(By.XPath("//option[text()='100']")))
            {
                var nbPages = WaitForElementExists(By.XPath("//option[text()='100']"));
                return nbPages.Selected;
            }
            return false;
        }
        public bool PageSizeEqualsTo100()
        {
            var pageSize = WaitForElementExists(By.XPath(PAGE_SIZE_SELECTED));
            var size = pageSize.GetAttribute("value");
            if (size != "100")
            {
                return false;
            }
            return true;
        }
        public IWebElement WaitForElementIsVisible(By value, string elementName = null)
        {
            // Initialisation des variables
            int compteur = 1;
            bool isFound = false;
            IWebElement element = null;

            WebDriverWait wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(5));
            while (compteur <= 10 && !isFound)
            {
                try
                {
                    var elm = wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(value));
                    isFound = true;
                    element = elm;
                }
                catch
                {
                    compteur++;
                }
            }

            if (!isFound)
            {
                Console.WriteLine($"[{DateTime.UtcNow.ToShortDateString()} {DateTime.UtcNow.ToLongTimeString()}] [WaitForElementIsVisible] Element not visible {value}");

                if (elementName != null)
                {
                    throw new Exception($"L'element {elementName} n'est pas visible dans la page.");
                }
                else
                {
                    throw new Exception($"L'element {value} n'est pas visible dans la page.");
                }

            }
            return element;
        }
    
        //public IWebElement WaitForElementIsVisible(By value, string elementName = null)
        //{
        // Initialisation des variables
        //int compteur = 1;
        //bool isFound = false;
        //IWebElement element = null;
        //WebDriverWait wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(10)); // Augmentation du délai à 10 secondes
        //while (compteur <= 10 && !isFound)
        //{
        //    try
        //    {
        //        var elm = wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementExists(value));
        //        if (elm.Displayed) // Vérifie si l'élément est visible après avoir trouvé l'existence
        //        {
        //            isFound = true;
        //            element = elm;
        //        }
        //    }
        //    catch
        //    {
        //        compteur++;
        //        Thread.Sleep(500); // Ajoute une pause supplémentaire
        //    }
        //}

        //if (!isFound)
        //{
        //    Console.WriteLine($"[{DateTime.UtcNow.ToShortDateString()} {DateTime.UtcNow.ToLongTimeString()}] [WaitForElementIsVisible] Element not visible {value}");

        //    if (elementName != null)
        //    {
        //        throw new Exception($"L'element {elementName} n'est pas visible dans la page.");
        //    }
        //    else
        //    {
        //        throw new Exception($"L'element {value} n'est pas visible dans la page.");
        //    }
        //}

        //return element;
        //}

        public bool isElementVisible(By value)
        {
            try
            {
                var elm = _webDriver.FindElements(value);
                if (elm.Count() > 0 && !elm[0].Displayed)
                {
                    Console.WriteLine("isElementVisible not Displayed : " + value.ToString());
                }
                return (elm.Count() > 0 && elm[0].Displayed);
            }
            catch
            {
                return false;
            }
        }

        public bool isElementExists(By value)
        {
            try
            {
                var elm = _webDriver.FindElements(value);
                return (elm.Count() > 0);
            }
            catch
            {
                return false;
            }
        }

        public void TryToClickOnElement(IWebElement element)
        {
            if (element.Displayed)
                element.Click();
            else TryToClickOnInvisibleUsingJavascript(element);
        }

        private void TryToClickOnInvisibleUsingJavascript(IWebElement element, bool setVisible = false)
        {
            if (setVisible)
                _webDriver.ExecuteJavaScript("arguments[0].show();", element);

            _webDriver.ExecuteJavaScript("arguments[0].click();", element);
        }

        public IWebElement WaitForElementExists(By value)
        {
            // Initialisation des variables
            int compteur = 1;
            bool isFound = false;
            IWebElement element = null;

            WebDriverWait wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(3));
            while (compteur <= 10 && !isFound)
            {
                try
                {
                    var elm = wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementExists(value));
                    isFound = true;
                    element = elm;
                }
                catch
                {
                    compteur++;
                }
            }

            if (!isFound)
            {
                Console.WriteLine($"[{DateTime.UtcNow.ToShortDateString()} {DateTime.UtcNow.ToLongTimeString()}] [WaitForElementExists] Element does not exists {value}");
                throw new Exception($"L'element {value} n'existe pas dans la page.");
            }

            return element;
        }

        public IWebElement WaitForElementToBeClickable(By value)
        {
            // Initialisation des variables
            int compteur = 1;
            bool isFound = false;
            IWebElement element = null;

            WebDriverWait wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(3));
            while (compteur <= 10 && !isFound)
            {
                try
                {
                    var elm = wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(value));
                    isFound = true;
                    element = elm;
                }
                catch
                {
                    compteur++;
                }
            }

            if (!isFound)
            {
                Console.WriteLine($"[{DateTime.UtcNow.ToShortDateString()} {DateTime.UtcNow.ToLongTimeString()}] [WaitForElementToBeClickable] Element not clickable {value}");
                throw new Exception($"L'element {value} n'est pas clickable.");
            }

            return element;
        }


        // ______________________________________ Global Settings ________________________________________________

        public string GetDateFormatPickerValue()
        {
            var globalSettingsPage = GoToParameters_GlobalSettings();

            var globalSettingsModalPage = globalSettingsPage.GetDateFormatPickerPage();
            var dateFormatPickerValue = globalSettingsModalPage.GetDateFormatPickerValue();
            globalSettingsModalPage.Save();

            return dateFormatPickerValue;
        }

        public string GetDecimalSeparatorValue()
        {
            var globalSettingsPage = GoToParameters_GlobalSettings();

            var globalSettingsModalPage = globalSettingsPage.GetDecimalSeparatorPage();
            var decimalSeparatorValue = globalSettingsModalPage.GetDecimalValue();
            globalSettingsModalPage.Save();

            return decimalSeparatorValue;
        }

        public void SetSageAutoEnabled(string site, bool sageAutoEnabled, string area = null, bool sageAutoSite = true)
        {
            if (!sageAutoEnabled)
                sageAutoSite = false;

            var settingsSitesPage = GoToParameters_Sites();
            WaitForLoad();
            settingsSitesPage.Filter(ParametersSites.FilterType.SearchSite, site);
            settingsSitesPage.ClickOnFirstSite();

            settingsSitesPage.ClickToInformations();
            settingsSitesPage.SetSageAutoEnabledForSite(sageAutoSite);
            var statusSageAutoEnable = settingsSitesPage.GetSageAutoEnabledStatusForSite();
            Assert.AreEqual(sageAutoSite, statusSageAutoEnable, "Le status du paramètre sage auto enable sur le site n'est pas bien paramétré");
            var globalSettingsPage = GoToParameters_GlobalSettings();
            var globalSettingsModalPage = globalSettingsPage.IsSageAutoEnabled();
            globalSettingsModalPage.SetSageAutoEnabledValue(sageAutoEnabled, area);
            globalSettingsModalPage.Save();
        }

        public void SetNewVersionKeywordValue(bool newVersionKeyword)
        {
            var globalSettingsPage = GoToParameters_GlobalSettings();

            var globalSettingsModalPage = globalSettingsPage.GetNewKeywordGlobalPage();
            globalSettingsModalPage.SetKeywordVersionValue(newVersionKeyword);
            globalSettingsModalPage.Save();
        }

        public void SetNewVersionSearchValue(bool versionSearch)
        {
            var globalSettingsPage = GoToParameters_GlobalSettings();

            var globalSettingsModalPage = globalSettingsPage.GetIsNewSearchModePage();
            globalSettingsModalPage.SetNewSearchModeValue(versionSearch);
            globalSettingsModalPage.Save();
        }

        public void SetNewGroupDisplayValue(bool value)
        {
            var globalSettingsPage = GoToParameters_GlobalSettings();

            var globalSettingsModalPage = globalSettingsPage.GetIsNewGroupDisplayPage();
            globalSettingsModalPage.SetGroupDisplayValue(value);
            globalSettingsModalPage.Save();
        }

        public void SetSubGroupFunctionValue(bool value)
        {
            var globalSettingsPage = GoToParameters_GlobalSettings();

            var globalSettingsModalPage = globalSettingsPage.GetSubGroupGlobalPage();
            globalSettingsModalPage.SetSubGroupValue(value);
            globalSettingsModalPage.Save();
        }

        public bool GetSubGroupFunctionValue()
        {
            var globalSettingsPage = GoToParameters_GlobalSettings();

            var globalSettingsModalPage = globalSettingsPage.GetSubGroupGlobalPage();
            bool value = globalSettingsModalPage.GetSubGroupValue();
            globalSettingsModalPage.Save();

            return value;
        }

        public void SetIsCompareActiveValue(bool active)
        {
            var globalSettingsPage = GoToParameters_GlobalSettings();

            var globalSettingsModalPage = globalSettingsPage.GetIsCompareActiveGlobalPage();
            globalSettingsModalPage.SetCompareActiveValue(active);
            globalSettingsModalPage.Save();
        }

        public void SetNumberItemToCompareValue(string value)
        {
            var globalSettingsPage = GoToParameters_GlobalSettings();

            var globalSettingsModalPage = globalSettingsPage.GetNumberOfItemsToComparePage();
            globalSettingsModalPage.SetNumberItemCompareValue(value);
            globalSettingsModalPage.Save();
        }

        public string GetNumberItemToCompareValue()
        {
            var globalSettingsPage = GoToParameters_GlobalSettings();

            var globalSettingsModalPage = globalSettingsPage.GetNumberOfItemsToComparePage();
            string number = globalSettingsModalPage.GetNumberItemCompareValue();
            globalSettingsModalPage.Save();

            return number;
        }

        public void SetProductionSettingsRecipeExpiryDateValue(string value)
        {
            var globalSettingsPage = GoToParameters_GlobalSettings();

            var globalSettingsModalPage = globalSettingsPage.GetRecipeExpiryDatePage();
            globalSettingsModalPage.SetRecipeExpiryDateValue(value);
            globalSettingsModalPage.Save();
        }

        /**
         * Récupère le WebElement visible à travers la page
         */
        public IWebElement SolveVisible(string xpath, bool id = false)
        {
            WaitForLoad(); 
            if (id)
            {
                elementList = _webDriver.FindElements(By.Id(xpath));
            }
            else
            {
                elementList = _webDriver.FindElements(By.XPath(xpath));
            }

            for (int i = 0; i < elementList.Count; i++)
            {
                if (elementList[i].Displayed)
                {
                    return elementList[i];
                }
            }
            return null;
        }

        public void DropdownListSelectById(string ddlID, object valueToSet, PageBase pageContainingDDL = null)
        {
            IWebElement selectedDDL;

            if (pageContainingDDL != null)
            { selectedDDL = pageContainingDDL.WaitForElementIsVisible(By.Id(ddlID)); }
            else
            { selectedDDL = WaitForElementIsVisible(By.Id(ddlID)); }

            selectedDDL.SetValue(PageBase.ControlType.DropDownList, valueToSet);
        }

        /// <summary>
        /// Sélectionne un élement du multi select, avec les options complémentaires demandées
        /// </summary>
        /// <param name="cbOpt">L'ensemble des options de paramétrage du multi select</param>
        public void ComboBoxSelectById(ComboBoxOptions cbOpt)
        {
            IWebElement input = WaitForElementIsVisible(By.Id(cbOpt.XpathId));
            WaitForLoad();
            input.Click();
            WaitForLoad();

            if (cbOpt.ClickCheckAllAtStart)
            {
                var checkAllVisible = SolveVisible("//*/span[text()='Check all']");
                Assert.IsNotNull(checkAllVisible);
                checkAllVisible.Click();
            }
            else if (cbOpt.ClickUncheckAllAtStart)
            {
                var uncheckAllVisible = SolveVisible("//*/span[text()='Uncheck all']");
                Assert.IsNotNull(uncheckAllVisible);
                uncheckAllVisible.Click();
            }

            if (cbOpt.IsUsedInFilter)
            {
                WaitPageLoading();
                WaitForLoad();
            }
            else if (cbOpt.ClickUncheckAllAtStart)
            {
                WaitForLoad();
            }

            bool selectionWasModified = false;

            if (cbOpt.SelectionValue != null)
            {
                var searchVisible = SolveVisible("//*/input[@type='search']");
                Assert.IsNotNull(searchVisible);
                if (cbOpt.IsUsedInFilter)
                {
                    searchVisible.SetValue(ControlType.TextBox, cbOpt.SelectionValue);
                }
                else
                {
                    // 22/10/2024 : maj Chrome 130 bloque la zone de saisie
                    _webDriver.ExecuteJavaScript("arguments[0].value = arguments[1]; arguments[0].dispatchEvent(new Event('input'));", searchVisible, cbOpt.SelectionValue);
                    WaitLoading();
                }
                WaitPageLoading();
                WaitLoading();
                var select = SolveVisible("//*/label[contains(@for, 'ui-multiselect')]/span[contains(text(),'" + cbOpt.SelectionValue + "')]");
                Assert.IsNotNull(select, "Pas de sélection de " + cbOpt.SelectionValue);

                if (cbOpt.ClickCheckAllAfterSelection)
                {
                    var checkAllVisible = SolveVisible("//*/span[text()='Check all']");
                    Assert.IsNotNull(checkAllVisible);
                    checkAllVisible.Click();
                }
                else if (cbOpt.ClickUncheckAllAfterSelection)
                {
                    var uncheckAllVisible = SolveVisible("//*/span[text()='Uncheck all']");
                    Assert.IsNotNull(uncheckAllVisible);
                    uncheckAllVisible.Click();
                }
                else
                {
                    select.Click();
                }
                selectionWasModified = true;
            }

            if (selectionWasModified)
            {
                if (cbOpt.IsUsedInFilter)
                {
                    WaitPageLoading();
                    WaitForLoad();
                }
                else
                {
                    WaitForLoad();
                }
            }

            input = WaitForElementIsVisible(By.Id(cbOpt.XpathId));

            try
            {
                input.SendKeys(Keys.Enter);
            }
            catch
            {
                //Silent catch: sometimes there's no associated action with "enter" key
                input.Click();
            }

            WaitForLoad();
        }

        /**
         * A ne pas supprimer, delta entre dev et patch
         */
        public bool IsDev()
        {
            string url = _testContext.Properties["Winrest_URL"].ToString();
            return url.ToLower().Contains("dev");
        }

        public string TakeScreenshot()
        {
            WaitPageLoading();
            var imageName = "imageCaptured.png";
            // Take a screenshot
            Screenshot screenshot = ((ITakesScreenshot)_webDriver).GetScreenshot();
            screenshot.SaveAsFile(imageName, ScreenshotImageFormat.Png);
            return imageName;
        }

        public bool VerifyRGBExistInImg(string fileName, int R, int G, int B)
        {
            Bitmap image = new Bitmap(fileName);
            for (int x = 0; x < image.Width; x++)
            {
                for (int y = 0; y < image.Height; y++)
                {
                    System.Drawing.Color pixelColor = image.GetPixel(x, y);
                    if (pixelColor.R == R && pixelColor.G == G && pixelColor.B == B)
                    {
                        return true;
                    }
            }
                }
            return false;

        }

        public void ScrollToElement(IWebElement element)
        {
            _webDriver.ExecuteJavaScript("arguments[0].scrollIntoView(true);", element);
        }
        public DashboardUsagePage ClickDashboardUsage()
        {
            var dasboardUsageBtn = WaitForElementIsVisible(By.XPath("/html/body/div[2]/div/div[2]/div[6]/div/h2/a"));
            dasboardUsageBtn.Click();
            var wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(60));
            wait.Until((driver) => driver.WindowHandles.Count > 1);
            _webDriver.SwitchTo().Window(_webDriver.WindowHandles[1]);

            return new DashboardUsagePage(_webDriver, _testContext);
        }

        public void PageUp()
        {
            var html = _webDriver.FindElement(By.TagName("html"));
            html.SendKeys(Keys.PageUp);
            html.SendKeys(Keys.PageUp);
            html.SendKeys(Keys.PageUp);
            html.SendKeys(Keys.PageUp);
        }

        public void ScrollUntilElementIsInView(By by)
        {
            bool elementFound = false;
            int retryCount = 3; // Number of retries

            while (!elementFound && retryCount > 0)
            {
                try
                {
                    var element = WaitForElementExists(by);
                    var javaScriptExecutor = _webDriver as IJavaScriptExecutor;
                    javaScriptExecutor.ExecuteScript("arguments[0].scrollIntoView(true);", element);
                    elementFound = true; // Element found and scrolled successfully
                }
                catch (StaleElementReferenceException)
                {
                    retryCount--; // Retry if a stale element exception is thrown
                }
            }

            if (!elementFound)
            {
                throw new NoSuchElementException("Unable to find the element after multiple retries.");
            }
        }
        public InterimOrdersPage GoToInterim_InterimOrders()
        {
            // On tente de cliquer sur le menu à partir de la barre des menus
            var interimTab = WaitForElementIsVisible(By.Id("TabMenuInterim"));
            WaitForLoad();
            interimTab.Click();

            var interimOrderimTabNav = WaitForElementIsVisible(By.Id("InterimOrderTabNav"));
            interimOrderimTabNav.Click();

            return new InterimOrdersPage(_webDriver, _testContext);
        }

        public bool SearchPop()
        {
            return isElementExists(By.XPath(SEARCH_FIELD));
        }

        public void ZoomOut(decimal percentage)
        {
            ((IJavaScriptExecutor)_webDriver).ExecuteScript($"document.body.style.zoom='{percentage}';");
        }
        public List<string> GetGrossAmounts()
        {
            var grossAmountElements = _webDriver.FindElements(By.XPath("//*[@id='tableListMenu']/tbody/tr/td[12]"));
            List<string> grossAmountList = new List<string>();
            foreach (var element in grossAmountElements)
            {
                grossAmountList.Add(element.Text.Trim());
            }
            return grossAmountList;
        }
        public List<string> GetTotalInclTaxes()
        {
            var totalInclTaxesElements = _webDriver.FindElements(By.XPath("//*[@id='tableListMenu']/tbody/tr/td[13]"));
            List<string> totalInclTaxesList = new List<string>();
            foreach (var element in totalInclTaxesElements)
            {
                totalInclTaxesList.Add(element.Text.Trim());
            }
            return totalInclTaxesList;
        }
        public int GetGrossAmountXPosition(int rowIndex)
        {
            var grossElement = _webDriver.FindElement(By.XPath($"//*[@id='tableListMenu']/tbody/tr[{rowIndex}]/td[12]"));
            return grossElement.Location.X;
        }

        public int GetTotalInclTaxesXPosition(int rowIndex)
        {
            var totalInclTaxesElement = _webDriver.FindElement(By.XPath($"//*[@id='tableListMenu']/tbody/tr[{rowIndex}]/td[13]"));
            return totalInclTaxesElement.Location.X;
        }
        public void FileLoaded(By elementReceherche)
        {
            int maxAttempts = 30;
            int delayBetweenAttempts = 5000; 
            bool isFileDownloaded = false;

            for (int attempt = 1; attempt <= maxAttempts; attempt++)
            {
                ClickPrintButton();

                try
                {
                    var downloadFileBtn = _webDriver.FindElement(elementReceherche);
                    downloadFileBtn.Click();
                    isFileDownloaded = true;
                    WaitForLoad();
                    break; 
                }
                catch (NoSuchElementException)
                {
                    Console.WriteLine($"Attempt {attempt}/{maxAttempts}: Element not found. Retrying in {delayBetweenAttempts / 1000} seconds...");
                    ClosePrintButton(); 
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Attempt {attempt}/{maxAttempts}: An unexpected error occurred: {ex.Message}");
                    ClosePrintButton();
                }

                if (!isFileDownloaded)
                {
                    Thread.Sleep(delayBetweenAttempts);
                }
            }

            if (!isFileDownloaded)
            {
                Console.WriteLine($"[{DateTime.UtcNow.ToShortDateString()} {DateTime.UtcNow.ToLongTimeString()}] [IsFileLoaded] Element not found after {maxAttempts} attempts: {elementReceherche}");
                throw new Exception("Délai d'attente dépassé pour le téléchargement du fichier.");
            }
        }
        public ParametersLogs GoToParameters_Logs()
        {
            try
            {
                // On tente de cliquer sur le menu à partir de la barre des menus
                var customer = WaitForElementIsVisible(By.Id("TabParameters"));
                customer.Click();

                var customer_Catalog = WaitForElementIsVisible(By.Id("LogTabNav"));
                customer_Catalog.Click();

            }
            catch
            {

                // Si le menu n'est pas accessible, click sur le lien dans la page d'accueil
                GoToWinrestHome();
                var customer_catalog = WaitForElementIsVisible(By.Id("LogLinkDashBoard"));
                customer_catalog.Click();

            }

            WaitForLoad();

            return new ParametersLogs(_webDriver, _testContext);
         
        }
        public ToDoListTabletAppPage GotoTabletApp_ToDoList()
        {
            Actions action = new Actions(_webDriver);

            _tabletApp = WaitForElementExists(By.XPath("/html/body/app-root/mat-sidenav-container/mat-sidenav-content/main/home/div/div[2]/div/div[12]/div"));
            action.MoveToElement(_tabletApp).Perform();
            _tabletApp.Click();

            return new ToDoListTabletAppPage(_webDriver, _testContext);
        }

        // Cette méthode est provisoire pour les tests.
        public IWebElement WaitForElementIsVisibleNew(By value, string elementName = null)

        {
            int compteur = 1;
            bool isFound = false;
            IWebElement element = null;
            WebDriverWait wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(10)); // Augmentation du délai à 10 secondes
            while (compteur <= 10 && !isFound)
            {
                try
                {
                    var elm = wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementExists(value));
                    if (elm.Displayed) // Vérifie si l'élément est visible après avoir trouvé l'existence
                    {
                        isFound = true;
                        element = elm;
                    }
                }
                catch
                {
                    compteur++;
                    Thread.Sleep(500); // Ajoute une pause supplémentaire
                }
            }
            if (!isFound)
            {
                Console.WriteLine($"[{DateTime.UtcNow.ToShortDateString()} {DateTime.UtcNow.ToLongTimeString()}] [WaitForElementIsVisible] Element not visible {value}");
                if (elementName != null)
                {
                    throw new Exception($"L'element {elementName} n'est pas visible dans la page.");
                }
                else
                {
                    throw new Exception($"L'element {value} n'est pas visible dans la page.");
                }
            }
            return element;
        }
        public int GetGrossQty()
        {
            var grossQtyElement = _webDriver.FindElement(By.XPath(GROSS_QTY));
            string quantityText = grossQtyElement.GetAttribute("value").Trim();

            return int.Parse(quantityText);
        }
        public void SetGrossQty(int value)
        {
            var grossQtyElement = _webDriver.FindElement(By.XPath(GROSS_QTY));
            grossQtyElement.Clear();
            grossQtyElement.SendKeys(value.ToString());
            WaitLoading();
        }
        public double GetNetWeight()
        {
            var netWeight = _webDriver.FindElement(By.XPath(NET_WEIGHT));
            string weightText = netWeight.GetAttribute("value").Trim();

            // Parse the text as a double to handle fractional values
            return double.Parse(weightText);
        }


        public void SetNetWeight(int value)
        {
            var netWeightElement = _webDriver.FindElement(By.XPath(NET_WEIGHT));
            netWeightElement.Clear();
            netWeightElement.SendKeys(value.ToString());
            WaitLoading();
        }
    }
}