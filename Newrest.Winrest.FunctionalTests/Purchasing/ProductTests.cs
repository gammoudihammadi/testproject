using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Purchasing;
using Newrest.Winrest.FunctionalTests.PageObjects;
using Newrest.Winrest.FunctionalTests.PageObjects.Parameters.Production;
using Newrest.Winrest.FunctionalTests.PageObjects.Customers.Service;
using Newrest.Winrest.FunctionalTests.PageObjects.Parameters.Purchasing;
using Newrest.Winrest.FunctionalTests.PageObject.Parameters.Customer;
using Newrest.Winrest.FunctionalTests.PageObjects.Purchasing.Product;
using Newrest.Winrest.FunctionalTests.Utils;
using System.Security.Policy;
using System;
using Newrest.Winrest.FunctionalTests.PageObjects.Accounting.SupplierInvoices;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;
using OpenQA.Selenium;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using System.IO;
using System.Linq;
using System.Threading;

namespace Newrest.Winrest.FunctionalTests.Purchasing
{
    [TestClass]
    public class ProductTests : TestBase
    {
        private const int _timeout = 600000;
        string item = "";
        string service = "";
        string item2 = "";
        string service2 = "";

        /* Test Prep */
        [TestInitialize]
        public override void TestInitialize()
        {

            base.TestInitialize();

            var testMethod = TestContext.TestName;
            switch (testMethod)
            {

                  case nameof(PU_PROD_SearchService):
                      PU_PROD_SearchService_TestInitialize();
                      break;

                case nameof(PU_PROD_Delete):
                    PU_PROD_Delete_TestInitialize();
                    break;

                case nameof(PU_PROD_SearchItem):
                    PU_PROD_SearchItem_TestInitialize();
                    break;

                case nameof(PU_PROD_CreateAdd):
                    PU_PROD_CreateAdd_TestInitialize();
                    break;
                case nameof(PU_PROD_CreateAddNew):
                    PU_PROD_CreateAddNew_TestInitialize();
                    break;
                case nameof(PU_PROD_LinkItem):
                    PU_PROD_LinkItem_TestInitialize();
                    break;
                default:
                    break;
                  

            }
            base.TestInitialize();

        }

        [TestCleanup]
        public override void TestCleanup()
        {
            

            var testMethod = TestContext.TestName;
            switch (testMethod)
            {
                case nameof(PU_PROD_SearchService):
                    Product_Clean_up();
                    break;

                case nameof(PU_PROD_SearchItem):
                    Product_Clean_up();
                    break;

                case nameof(PU_PROD_CreateAdd):
                    Product_Clean_up();
                    break;
                case nameof(PU_PROD_CreateAddNew):
                    Product_Clean_up();
                    break;
                case nameof(PU_PROD_LinkItem):
                    Product_Clean_up();
                    break;
                default:
                    break;
            }
            base.TestCleanup();
        }

        private void Product_Clean_up()
        {
            HomePage homePage = LogInAsAdmin();
            var purchasingPage = homePage.GoToPurchasing_ProductPage();
            purchasingPage.ResetFilter();

            purchasingPage.Search(item, false);
            purchasingPage.DeleteFirstProduct();

        }
       
        private void PU_PROD_CreateAdd_TestInitialize()
        {
            //Prepare
            Random random = new Random();
            string category = "CATEGORY_PRODUCT"; //TestContext.Properties["Prodman_Needs_ServiceCategory2"].ToString(); 
            string group = "product_group"; //TestContext.Properties["IsProductGroup"].ToString();
            service = "ServiceForProduct" + DateTime.Now.ToString("dd/MM/yyyy/HHmmssFF") + random.Next(10,9999);
            item = "ItemForTest" + DateTime.Now.ToString("dd/MM/yyyy/HHmmssFF") + random.Next(10, 9999);
            string site = TestContext.Properties["Site"].ToString();
            string customer = TestContext.Properties["Customer"].ToString();
            string workshop = TestContext.Properties["Item_Workshop"].ToString();
            string taxType = TestContext.Properties["Item_TaxType"].ToString();
            string prodUnit = TestContext.Properties["Item_ProdUnit"].ToString();

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            /* Creat item*/
            ItemPage itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.Filter(ItemPage.FilterType.Search, item);
          
            var itemCreateModalPage = itemPage.ItemCreatePage();
            var itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(item, group, workshop, taxType, prodUnit);
           itemPage = itemGeneralInformationPage.BackToList();

            /*creat service */
            ServicePage servicePage = homePage.GoToCustomers_ServicePage();
            ServiceCreateModalPage creatService = servicePage.ServiceCreatePage();
            creatService.FillFields_CreateServiceModalPage(service, null, null, category);
            var serviceGeneralInformationsPage = creatService.Create();
       

            ProductPage productPage = homePage.GoToPurchasing_ProductPage();
            if (productPage.CheckTotalNumber() != 0)
            {
                if (productPage.GetFirstServiceProduct() == service && productPage.GetFirstItemProduct() == item)
                {
                    productPage.Search(item , false);
                    productPage.DeleteFirstProduct();
                }
            }

            //Act1
            ParametersProduction parametrePage = homePage.GoToParameters_ProductionPage();
            parametrePage.GoToTab_Group();
            if (parametrePage.IsGroupExist(group))
            {
                if (!parametrePage.isGroupProduct(group))
                {
                    parametrePage.EditGroup(group, true, false);

                }
                if (!parametrePage.isGroupFood(group))
                {
                    parametrePage.EditGroup(group, false, true);
                }
            }
           
            ParametersCustomer parametersCustomer = homePage.GoToParameters_CustomerPage();
            parametersCustomer.GoToTab_Category();
            if (parametersCustomer.isCategoryExist(category))
            {
                if (!parametersCustomer.isCategoryProduct(category))
                {
                    parametersCustomer.EditCategory(category, true, false);

                }
                if (!parametersCustomer.isCategoryFood(category))
                {
                    parametersCustomer.EditCategory(category, false, true);
                }
            }
            else
            {
                parametersCustomer.AddNewCategory("CATEGORY_PRODUCT", true, true);
            }
        }

        private void PU_PROD_CreateAddNew_TestInitialize()
        {
            //Prepare
            Random random = new Random();
            string category = "CATEGORY_PRODUCT"; //TestContext.Properties["Prodman_Needs_ServiceCategory2"].ToString(); 
            string group = "product_group"; //TestContext.Properties["IsProductGroup"].ToString();
            service = "ServiceForProduct" + DateTime.Now.ToString("dd/MM/yyyy/HHmmssFF") + random.Next(10, 9999);
            item = "ItemForTest" + DateTime.Now.ToString("dd/MM/yyyy/HHmmssFF") + random.Next(10, 9999);
            service2 = "ServiceForProduct" + DateTime.Now.ToString("dd/MM/yyyy/HHmmssFF") + random.Next(10, 9999) + "2";
            item2 = "ItemForTest" + DateTime.Now.ToString("dd/MM/yyyy/HHmmssFF") + random.Next(10, 9999) + "2";
            string site = TestContext.Properties["Site"].ToString();
            string customer = TestContext.Properties["Customer"].ToString();
            string workshop = TestContext.Properties["Item_Workshop"].ToString();
            string taxType = TestContext.Properties["Item_TaxType"].ToString();
            string prodUnit = TestContext.Properties["Item_ProdUnit"].ToString();

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            /* Creat item*/
            ItemPage itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.Filter(ItemPage.FilterType.Search, item);
            /* item1 */
            var itemCreateModalPage = itemPage.ItemCreatePage();
            var itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(item, group, workshop, taxType, prodUnit);
            itemPage = itemGeneralInformationPage.BackToList();
            /* item 2*/
            itemCreateModalPage = itemPage.ItemCreatePage();
            itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(item2 , group, workshop, taxType, prodUnit);
            itemPage = itemGeneralInformationPage.BackToList();
            
            
            /*creat service */
            ServicePage servicePage = homePage.GoToCustomers_ServicePage();
            /* service 1 */ 
            ServiceCreateModalPage creatService = servicePage.ServiceCreatePage();
            creatService.FillFields_CreateServiceModalPage(service, null, null, category);
            var serviceGeneralInformationsPage = creatService.Create();
            /* service 2*/
            servicePage = serviceGeneralInformationsPage.BackToList(); 
            creatService = servicePage.ServiceCreatePage();
            creatService.FillFields_CreateServiceModalPage(service2 ,null, null, category);
            serviceGeneralInformationsPage = creatService.Create();


            ProductPage productPage = homePage.GoToPurchasing_ProductPage();
            if (productPage.CheckTotalNumber() != 0)
            {
                if (productPage.GetFirstServiceProduct() == service && productPage.GetFirstItemProduct() == item)
                {
                    productPage.Search(item, false);
                    productPage.DeleteFirstProduct();
                }
            }

            //Act1
            ParametersProduction parametrePage = homePage.GoToParameters_ProductionPage();
            parametrePage.GoToTab_Group();
            if (parametrePage.IsGroupExist(group))
            {
                if (!parametrePage.isGroupProduct(group))
                {
                    parametrePage.EditGroup(group, true, false);

                }
                if (!parametrePage.isGroupFood(group))
                {
                    parametrePage.EditGroup(group, false, true);
                }
            }

            ParametersCustomer parametersCustomer = homePage.GoToParameters_CustomerPage();
            parametersCustomer.GoToTab_Category();
            if (parametersCustomer.isCategoryExist(category))
            {
                if (!parametersCustomer.isCategoryProduct(category))
                {
                    parametersCustomer.EditCategory(category, true, false);

                }
                if (!parametersCustomer.isCategoryFood(category))
                {
                    parametersCustomer.EditCategory(category, false, true);
                }
            }
            else
            {
                parametersCustomer.AddNewCategory("CATEGORY_PRODUCT", true, true);
            }
        }
        private void PU_PROD_Delete_TestInitialize()
        {
            //Prepare
            Random random = new Random();
            string category = "CATEGORY_PRODUCT"; //TestContext.Properties["Prodman_Needs_ServiceCategory2"].ToString(); 
            string group = "product_group"; //TestContext.Properties["IsProductGroup"].ToString();
            service = "ServiceForProduct" + DateTime.Now.ToString("dd/MM/yyyy/HHmmssFF") + random.Next(10, 9999);
            item = "ItemForTest" + DateTime.Now.ToString("dd/MM/yyyy/HHmmssFF") + random.Next(10, 9999);
            string site = TestContext.Properties["Site"].ToString();
            string customer = TestContext.Properties["Customer"].ToString();
            string workshop = TestContext.Properties["Item_Workshop"].ToString();
            string taxType = TestContext.Properties["Item_TaxType"].ToString();
            string prodUnit = TestContext.Properties["Item_ProdUnit"].ToString();

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            /* Creat item*/
            ItemPage itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.Filter(ItemPage.FilterType.Search, item);

            var itemCreateModalPage = itemPage.ItemCreatePage();
            var itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(item, group, workshop, taxType, prodUnit);
            itemPage = itemGeneralInformationPage.BackToList();

            /*creat service */
            ServicePage servicePage = homePage.GoToCustomers_ServicePage();
            ServiceCreateModalPage creatService = servicePage.ServiceCreatePage();
            creatService.FillFields_CreateServiceModalPage(service, null, null, category);
            var serviceGeneralInformationsPage = creatService.Create();


            ProductPage productPage = homePage.GoToPurchasing_ProductPage();
            if (productPage.CheckTotalNumber() != 0)
            {
                if (productPage.GetFirstServiceProduct() == service && productPage.GetFirstItemProduct() == item)
                {
                    productPage.Search(item, false);
                    productPage.DeleteFirstProduct();
                }
            }

            //Act1
            ParametersProduction parametrePage = homePage.GoToParameters_ProductionPage();
            parametrePage.GoToTab_Group();
            if (parametrePage.IsGroupExist(group))
            {
                if (!parametrePage.isGroupProduct(group))
                {
                    parametrePage.EditGroup(group, true, false);

                }
                if (!parametrePage.isGroupFood(group))
                {
                    parametrePage.EditGroup(group, false, true);
                }
            }

            ParametersCustomer parametersCustomer = homePage.GoToParameters_CustomerPage();
            parametersCustomer.GoToTab_Category();
            if (parametersCustomer.isCategoryExist(category))
            {
                if (!parametersCustomer.isCategoryProduct(category))
                {
                    parametersCustomer.EditCategory(category, true, false);

                }
                if (!parametersCustomer.isCategoryFood(category))
                {
                    parametersCustomer.EditCategory(category, false, true);
                }
            }
            else
            {
                parametersCustomer.AddNewCategory("CATEGORY_PRODUCT", true, true);
            }

            ProductPage purchasingPage = homePage.GoToPurchasing_ProductPage();

            int checkTotalNbreBefore = purchasingPage.CheckTotalNumber();
            purchasingPage.ResetFilter();
            purchasingPage.ClickAddNewProduct();
            purchasingPage.CreateNewProduct(item, service);
            purchasingPage.ClickAdd();
        }
      private  void PU_PROD_LinkItem_TestInitialize()
        {
            //Prepare
            Random random = new Random();
            string category = "CATEGORY_PRODUCT"; //TestContext.Properties["Prodman_Needs_ServiceCategory2"].ToString(); 
            string group = "product_group"; //TestContext.Properties["IsProductGroup"].ToString();
            service = "ServiceForProduct" + DateTime.Now.ToString("dd/MM/yyyy/HHmmssFF") + random.Next(10, 9999);
            item = "ItemForTest" + DateTime.Now.ToString("dd/MM/yyyy/HHmmssFF") + random.Next(10, 9999);
            string site = TestContext.Properties["Site"].ToString();
            string customer = TestContext.Properties["Customer"].ToString();
            string workshop = TestContext.Properties["Item_Workshop"].ToString();
            string taxType = TestContext.Properties["Item_TaxType"].ToString();
            string prodUnit = TestContext.Properties["Item_ProdUnit"].ToString();

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            /* Creat item*/
            ItemPage itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.Filter(ItemPage.FilterType.Search, item);

            var itemCreateModalPage = itemPage.ItemCreatePage();
            var itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(item, group, workshop, taxType, prodUnit);
            itemPage = itemGeneralInformationPage.BackToList();

            /*creat service */
            ServicePage servicePage = homePage.GoToCustomers_ServicePage();
            ServiceCreateModalPage creatService = servicePage.ServiceCreatePage();
            creatService.FillFields_CreateServiceModalPage(service, null, null, category);
            var serviceGeneralInformationsPage = creatService.Create();


            ProductPage productPage = homePage.GoToPurchasing_ProductPage();
            if (productPage.CheckTotalNumber() != 0)
            {
                if (productPage.GetFirstServiceProduct() == service && productPage.GetFirstItemProduct() == item)
                {
                    productPage.Search(item, false);
                    productPage.DeleteFirstProduct();
                }
            }
            productPage.ResetFilter();

            //Act1
            ParametersProduction parametrePage = homePage.GoToParameters_ProductionPage();
            parametrePage.GoToTab_Group();
            if (parametrePage.IsGroupExist(group))
            {
                if (!parametrePage.isGroupProduct(group))
                {
                    parametrePage.EditGroup(group, true, false);

                }
                if (!parametrePage.isGroupFood(group))
                {
                    parametrePage.EditGroup(group, false, true);
                }
            }

            ParametersCustomer parametersCustomer = homePage.GoToParameters_CustomerPage();
            parametersCustomer.GoToTab_Category();
            if (parametersCustomer.isCategoryExist(category))
            {
                if (!parametersCustomer.isCategoryProduct(category))
                {
                    parametersCustomer.EditCategory(category, true, false);

                }
                if (!parametersCustomer.isCategoryFood(category))
                {
                    parametersCustomer.EditCategory(category, false, true);
                }
            }
            else
            {
                parametersCustomer.AddNewCategory("CATEGORY_PRODUCT", true, true);
            }

            ProductPage purchasingPage = homePage.GoToPurchasing_ProductPage();

            int checkTotalNbreBefore = purchasingPage.CheckTotalNumber();
            purchasingPage.ResetFilter();
            purchasingPage.ClickAddNewProduct();
            purchasingPage.CreateNewProduct(item, service);
            purchasingPage.ClickAdd();
        }
        private  void PU_PROD_SearchItem_TestInitialize()
        {
            //Prepare
            Random random = new Random();
            string category = "CATEGORY_PRODUCT"; //TestContext.Properties["Prodman_Needs_ServiceCategory2"].ToString(); 
            string group = "product_group"; //TestContext.Properties["IsProductGroup"].ToString();
            service = "ServiceForProduct" + DateTime.Now.ToString("dd/MM/yyyy/HHmmssFF") + random.Next(10, 9999);
            item = random.Next(10, 999) + "_ItemForTest" + DateTime.Now.ToString("dd/MM/yyyy/HHmmssFF") + random.Next(10, 9999);
            string site = TestContext.Properties["Site"].ToString();
            string customer = TestContext.Properties["Customer"].ToString();
            string workshop = TestContext.Properties["Item_Workshop"].ToString();
            string taxType = TestContext.Properties["Item_TaxType"].ToString();
            string prodUnit = TestContext.Properties["Item_ProdUnit"].ToString();

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            /* Creat item*/
            ItemPage itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.Filter(ItemPage.FilterType.Search, item);

            var itemCreateModalPage = itemPage.ItemCreatePage();
            var itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(item, group, workshop, taxType, prodUnit);
            itemPage = itemGeneralInformationPage.BackToList();

            /*creat service */
            ServicePage servicePage = homePage.GoToCustomers_ServicePage();
            ServiceCreateModalPage creatService = servicePage.ServiceCreatePage();
            creatService.FillFields_CreateServiceModalPage(service, null, null, category);
            var serviceGeneralInformationsPage = creatService.Create();


            ProductPage productPage = homePage.GoToPurchasing_ProductPage();
            if (productPage.CheckTotalNumber() != 0)
            {
                if (productPage.GetFirstServiceProduct() == service && productPage.GetFirstItemProduct() == item)
                {
                    productPage.Search(item, false);
                    productPage.DeleteFirstProduct();
                }
            }

            //Act1
            ParametersProduction parametrePage = homePage.GoToParameters_ProductionPage();
            parametrePage.GoToTab_Group();
            if (parametrePage.IsGroupExist(group))
            {
                if (!parametrePage.isGroupProduct(group))
                {
                    parametrePage.EditGroup(group, true, false);

                }
                if (!parametrePage.isGroupFood(group))
                {
                    parametrePage.EditGroup(group, false, true);
                }
            }

            ParametersCustomer parametersCustomer = homePage.GoToParameters_CustomerPage();
            parametersCustomer.GoToTab_Category();
            if (parametersCustomer.isCategoryExist(category))
            {
                if (!parametersCustomer.isCategoryProduct(category))
                {
                    parametersCustomer.EditCategory(category, true, false);

                }
                if (!parametersCustomer.isCategoryFood(category))
                {
                    parametersCustomer.EditCategory(category, false, true);
                }
            }
            else
            {
                parametersCustomer.AddNewCategory("CATEGORY_PRODUCT", true, true);
            }

            ProductPage purchasingPage = homePage.GoToPurchasing_ProductPage();

            int checkTotalNbreBefore = purchasingPage.CheckTotalNumber();
            purchasingPage.ResetFilter();
            purchasingPage.ClickAddNewProduct();
            purchasingPage.CreateNewProduct(item, service);
            purchasingPage.ClickAdd();
        }
      private void PU_PROD_SearchService_TestInitialize()
        {
            //Prepare
            Random random = new Random();
            string category = "CATEGORY_PRODUCT"; //TestContext.Properties["Prodman_Needs_ServiceCategory2"].ToString(); 
            string group = "product_group"; //TestContext.Properties["IsProductGroup"].ToString();
            service = random.Next(10, 999) +  "_ServiceForProduct" + DateTime.Now.ToString("dd/MM/yyyy/HHmmssFF") + random.Next(10, 9999);
            item =  "ItemForTest" + DateTime.Now.ToString("dd/MM/yyyy/HHmmssFF") + random.Next(10, 9999);
            string site = TestContext.Properties["Site"].ToString();
            string customer = TestContext.Properties["Customer"].ToString();
            string workshop = TestContext.Properties["Item_Workshop"].ToString();
            string taxType = TestContext.Properties["Item_TaxType"].ToString();
            string prodUnit = TestContext.Properties["Item_ProdUnit"].ToString();

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            /* Creat item*/
            ItemPage itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.Filter(ItemPage.FilterType.Search, item);

            var itemCreateModalPage = itemPage.ItemCreatePage();
            var itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(item, group, workshop, taxType, prodUnit);
            itemPage = itemGeneralInformationPage.BackToList();

            /*creat service */
            ServicePage servicePage = homePage.GoToCustomers_ServicePage();
            ServiceCreateModalPage creatService = servicePage.ServiceCreatePage();
            creatService.FillFields_CreateServiceModalPage(service, null, null, category);
            var serviceGeneralInformationsPage = creatService.Create();


            ProductPage productPage = homePage.GoToPurchasing_ProductPage();
            if (productPage.CheckTotalNumber() != 0)
            {
                if (productPage.GetFirstServiceProduct() == service && productPage.GetFirstItemProduct() == item)
                {
                    productPage.Search(item, false);
                    productPage.DeleteFirstProduct();
                }
            }

            //Act1
            ParametersProduction parametrePage = homePage.GoToParameters_ProductionPage();
            parametrePage.GoToTab_Group();
            if (parametrePage.IsGroupExist(group))
            {
                if (!parametrePage.isGroupProduct(group))
                {
                    parametrePage.EditGroup(group, true, false);

                }
                if (!parametrePage.isGroupFood(group))
                {
                    parametrePage.EditGroup(group, false, true);
                }
            }

            ParametersCustomer parametersCustomer = homePage.GoToParameters_CustomerPage();
            parametersCustomer.GoToTab_Category();
            if (parametersCustomer.isCategoryExist(category))
            {
                if (!parametersCustomer.isCategoryProduct(category))
                {
                    parametersCustomer.EditCategory(category, true, false);

                }
                if (!parametersCustomer.isCategoryFood(category))
                {
                    parametersCustomer.EditCategory(category, false, true);
                }
            }
            else
            {
                parametersCustomer.AddNewCategory("CATEGORY_PRODUCT", true, true);
            }

            ProductPage purchasingPage = homePage.GoToPurchasing_ProductPage();

            int checkTotalNbreBefore = purchasingPage.CheckTotalNumber();
            purchasingPage.ResetFilter();
            purchasingPage.ClickAddNewProduct();
            purchasingPage.CreateNewProduct(item, service);
            purchasingPage.ClickAdd();
        }
      
        

        /*Test Methods */
        [TestMethod]
        [Timeout(_timeout)]
        public void PU_PROD_CreateCancel()
        {
            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            var purchasingPage = homePage.GoToPurchasing_ProductPage();
            purchasingPage.ResetFilter();
            int initialCount = purchasingPage.CheckTotalNumber();
            purchasingPage.ClickAddNewProduct();
            purchasingPage.CreateNewProduct("a", "a");

            purchasingPage.ClickClose();
            int finalCount = purchasingPage.CheckTotalNumber();

            //Assert
            Assert.IsTrue(initialCount.Equals(finalCount) && purchasingPage.IsNewProductFormClosed(), "Problème d'ajout sur new product");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void PU_PROD_SearchService()
        {
            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            var purchasingPage = homePage.GoToPurchasing_ProductPage();
            purchasingPage.ResetFilter();

            purchasingPage.Search(service.Substring(0, 5));
            string serviceName = purchasingPage.GetFirstServiceProduct();
            //Assert
            Assert.AreEqual(serviceName, service, "Search par service non fonctionnel");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void PU_PROD_Delete()
        {
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            var purchasingPage = homePage.GoToPurchasing_ProductPage();
            purchasingPage.ResetFilter();

            purchasingPage.Search(service, false ); 
            purchasingPage.ClickDeleteFirstProduct();
            purchasingPage.CancelDeleteFirstProduct();
            purchasingPage.DeleteFirstProduct();
            
            //Assert
            Assert.AreEqual(0, purchasingPage.CheckTotalNumber(), "Search par service non fonctionnel");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void PU_PROD_SearchItem()
        {
         
            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            var purchasingPage = homePage.GoToPurchasing_ProductPage();
            purchasingPage.ResetFilter();
            Assert.IsTrue(purchasingPage.CheckTotalNumber() > 0, "Aucune donnée de produit à traiter");
            purchasingPage.Search(item.Substring(0,5));
            string itemName = purchasingPage.GetFirstItemProduct();
            //Assert
            Assert.AreEqual(itemName, item, "La recherche par nom ne fontionne pas");
        }
        [TestMethod]
        [Timeout(_timeout)]
        public void PU_PROD_LinkService()
        {
            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            var purchasingPage = homePage.GoToPurchasing_ProductPage();
            purchasingPage.ResetFilter();
            Assert.IsTrue(purchasingPage.CheckTotalNumber() > 0, "Aucune donnée de produit à traiter");
            string serviceName = purchasingPage.GetFirstServiceProduct();
            var servicePricePage = purchasingPage.OpenEditService();
            servicePricePage.Go_To_New_Navigate();
            ServiceGeneralInformationPage serviceGeneralInformationPage =servicePricePage.ClickOnGeneralInformationTab();
            var serviceNameGeneralInformation = serviceGeneralInformationPage.GetServiceName();
            servicePricePage.Close();
            //Assert
            Assert.AreEqual(serviceNameGeneralInformation.Trim(), serviceName.Trim(), "L'onglet qui est ouvert n'est pas la page de service cible");
        }

     
        [TestMethod]
        [Timeout(_timeout)]
        public void PU_PROD_LinkItem()
        {
            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            var purchasingPage = homePage.GoToPurchasing_ProductPage();
            purchasingPage.ResetFilter();
            Assert.IsTrue(purchasingPage.CheckTotalNumber() > 0, "Aucune donnée de produit à traiter");
          
            purchasingPage.Search( item , false );
            var itemGeneralInformationPage = purchasingPage.OpenEditItem();
            itemGeneralInformationPage.Go_To_New_Navigate();
            itemGeneralInformationPage.ClickOnGeneralInformationPage();
            string ItemNameGeneralInformation = itemGeneralInformationPage.GetItemName();
            itemGeneralInformationPage.Close(); 
            //Assert
            Assert.AreEqual(ItemNameGeneralInformation.Trim(), item.Trim() , "L'onglet qui est ouvert n'est pas la page de l'item cible");
        }
     

        [TestMethod]
        [Timeout(_timeout)]
        public void PU_PROD_CreateAdd()
        {
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            ProductPage purchasingPage = homePage.GoToPurchasing_ProductPage();
           
            
            purchasingPage.ResetFilter();
            int checkTotalNbreBefore = purchasingPage.CheckTotalNumberProduct();
            purchasingPage.ClickAddNewProduct();
            purchasingPage.CreateNewProduct(item, service);
            purchasingPage.ClickAdd();
         
            int checkTotalNbreAfter = purchasingPage.CheckTotalNumberProduct(); 
            Assert.IsTrue(checkTotalNbreAfter  == checkTotalNbreBefore + 1, "le product n'est pas ajouté");
         
        }
        [TestMethod]
        [Timeout(_timeout)]
        public void PU_PROD_CreateAddNew()
        {
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            ProductPage purchasingPage = homePage.GoToPurchasing_ProductPage();

            purchasingPage.ResetFilter();
            purchasingPage.ClickAddNewProduct();
            purchasingPage.CreateNewProduct(item, service);
            purchasingPage.ClickAddNew();
            purchasingPage.CreateNewProduct(item2, service2);
            purchasingPage.ClickAdd();
            Assert.IsTrue(purchasingPage.HasLink(), "Le lien ne fonctionne pas!");

        }
        [TestMethod]
        [Timeout(_timeout)]
        public void PU_PROD_Export()
        {
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            var homePage = LogInAsAdmin();
            ItemPage itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();
            var item =itemPage.GetFirstItemName();
            itemPage.Filter(ItemPage.FilterType.Search, item);
            DeleteAllFileDownload();
            itemPage.ClearDownloads();
            itemPage.Export(true);
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();
            var correctDownloadedFile = itemPage.GetExportExcelFile(taskFiles);
            Assert.IsNotNull(correctDownloadedFile);
            var fileName = correctDownloadedFile.Name;
            var filePath = System.IO.Path.Combine(downloadsPath, fileName);
            int resultNumber = OpenXmlExcel.GetExportResultNumber("Sheet 1", filePath);
            var listResult = OpenXmlExcel.GetValuesInList("Item name", "Sheet 1", filePath);
            //Assert
            Assert.AreNotEqual(0, resultNumber, MessageErreur.EXCEL_PAS_DE_DONNEES);
            string Normalize(string input) => string.Join(" ", input.Trim().Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries));
            foreach (var i in listResult)
            {
                Assert.AreEqual(Normalize(item), Normalize(i), MessageErreur.EXCEL_DONNEES_KO);
            }
        }
    }
}
