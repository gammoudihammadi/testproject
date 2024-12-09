using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects;
using System;
using Newrest.Winrest.FunctionalTests.PageObjects.Customers.Catalogs;
using Newrest.Winrest.FunctionalTests.Utils;

namespace Newrest.Winrest.FunctionalTests.Customers.Catalogs
{
    [TestClass]
    public class CatalogTest : TestBase
    {

        private const int _timeout = 600000;
        private static Random random = new Random();
        private static DateTime fromDate = DateUtils.Now;
        private static DateTime toDate = DateUtils.Now.AddDays(+31);
        private static string CatalogName = "CATALOGWithoutProd" + DateTime.Now.Day + DateTime.Now.Month + DateTime.Now.Year + DateTime.Now.Hour + DateTime.Now.Minute + DateTime.Now.Second + random.Next(0, 999);
        private string CatalogName1 = "CATALOGWithProd" + DateTime.Now.Day + DateTime.Now.Month + DateTime.Now.Year + DateTime.Now.Hour + DateTime.Now.Minute + DateTime.Now.Second + random.Next(0, 999);
        private string ProdName = "PROD" + DateTime.Now.Day + DateTime.Now.Month + DateTime.Now.Year + DateTime.Now.Hour + DateTime.Now.Minute + DateTime.Now.Second + random.Next(0, 999);
        private string CategorieName = "CATEG" + DateTime.Now.Day + DateTime.Now.Month + DateTime.Now.Year + DateTime.Now.Hour + DateTime.Now.Minute + DateTime.Now.Second;
        [TestInitialize]
        public override void TestInitialize()
        {

            base.TestInitialize();

            var testMethod = TestContext.TestName;
            switch (testMethod)
            {
                case nameof(CU_CAT_DetailCatalogAvecOuSansProduct):

                    CU_CAT_Insertiondata_TestInitialize();
                    CU_CAT_Insertiondata_With_Product_TestInitialize();
                    break;

                default:
                    break;
            }
        }

        [TestCleanup]
        public override void TestCleanup()
        {
            var testMethod = TestContext.TestName;
            switch (testMethod)
            {
                case nameof(CU_CAT_DetailCatalogAvecOuSansProduct):

                    CU_CAT_deletedata_TestCleanup();
                    CU_CAT_deletedata_With_Prod_TestCleanup();
                    break;
                default:
                    break;
            }
            base.TestCleanup();
        }

        private void CU_CAT_Insertiondata_TestInitialize()
        {

            //arrange
            HomePage homePage = LogInAsAdmin();
            CatalogPage CatalogPage = homePage.GoToCustomers_CatalogPage();
            CatalogPage.ResetFilters();

            //create Catalog WithoutProd
            var CatalogCreateModalPage = CatalogPage.ClickCatalogCreatePage();
            CatalogPage = CatalogCreateModalPage.FillField_CreatNewCatalog(CatalogName, fromDate, toDate);


        }
        private void CU_CAT_Insertiondata_With_Product_TestInitialize()
        {

            //arrange
            HomePage homePage = LogInAsAdmin();
            CatalogPage CatalogPage = homePage.GoToCustomers_CatalogPage();
            CatalogPage.ResetFilters();

            //create prod 
            var ProductModalPage = CatalogPage.GoToProductPage();
            var ProductCreateModalPage = ProductModalPage.ClickProductCreatePage();
            CatalogPage = ProductCreateModalPage.FillField_CreatNewProduct(ProdName);

            //create Catalog WithProd
            homePage.GoToCustomers_CatalogPage();
            CatalogPage.ResetFilters();
            var CatalogCreateModalPageWithProd = CatalogPage.ClickCatalogCreatePage();
            CatalogPage = CatalogCreateModalPageWithProd.FillField_CreatNewCatalog(CatalogName1, fromDate, toDate);


        }
        private void CU_CAT_deletedata_TestCleanup()
        {
            //arrange
            HomePage homePage = LogInAsAdmin();
            CatalogPage CatalogPage = homePage.GoToCustomers_CatalogPage();
            CatalogPage.ResetFilters();

            CatalogPage.Filter(CatalogPage.FilterType.Search, CatalogName);
            CatalogPage.DeleteFirstCatalog();
            CatalogPage.ResetFilters();
        }
        private void CU_CAT_deletedata_With_Prod_TestCleanup()
        {
            //arrange
            HomePage homePage = LogInAsAdmin();
            CatalogPage CatalogPage = homePage.GoToCustomers_CatalogPage();
            CatalogPage.ResetFilters();

            CatalogPage.Filter(CatalogPage.FilterType.Search, CatalogName1);
            var CatalogItemWithoutProd = CatalogPage.SelectFirstItemCatalog();
            var DetailsPage1 = CatalogItemWithoutProd.GoToDetailPage();
            DetailsPage1.deleteItem();
            DetailsPage1.BackToList();
            DetailsPage1.WaitPageLoading();
            CatalogPage.DeleteFirstCatalog();
            CatalogPage.ResetFilters();

            // delete product 
            var ProductModalPage = homePage.GoToCustomers_CatalogPage().GoToProductPage();
            ProductModalPage.FilterProd(CatalogPage.FilterType.Search, ProdName);
            ProductModalPage.DeleteFirstProd();
            ProductModalPage.ResetFilters();
        }


        [Timeout(_timeout)]
		[TestMethod]
        public void CU_CAT_DetailCatalogAvecOuSansProduct()
        {

            //arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            CatalogPage CatalogPage = homePage.GoToCustomers_CatalogPage();
            CatalogPage.ResetFilters();

            //test
            CatalogPage.Filter(CatalogPage.FilterType.Search, CatalogName);
            var CatalogItemWithoutProd = CatalogPage.SelectFirstItemCatalog();
            var DetailsPage1 = CatalogItemWithoutProd.GoToDetailPage();
            bool ItemVerifiedShow = CatalogItemWithoutProd.ItemIsVerifiedShow();
            //Assert
            Assert.IsTrue(ItemVerifiedShow, "Le catalogue ne contient aucun product, alors tout doit s'ouvrir correctement");
            CatalogItemWithoutProd.BackToList();
            CatalogPage.ResetFilters();

            CatalogPage.Filter(CatalogPage.FilterType.Search, CatalogName1);
            var CatalogItem = CatalogPage.SelectFirstItemCatalog();
            var CategoriePage = CatalogItem.GoToCategoriePage();

            var CategorieCreateModalPage = CategoriePage.FillField_CreatNewCategorie(CategorieName);
            var DetailsPage = CatalogItem.GoToDetailPage();
            var CategProdCreateModalPage = DetailsPage.FillField_CreatNewCatalogProd(ProdName);
            DetailsPage.WaitPageLoading();
            bool ItemVerifiedShowWithProd = CatalogItemWithoutProd.ItemIsVerifiedShowWithProd();
            //Assert
            Assert.IsTrue(ItemVerifiedShowWithProd, " Le catalogue contient au moins un product, alors tout doit s'ouvrir correctement");



        }

        [Timeout(_timeout)]
		[TestMethod]
        public void CU_CAT_GeneralInformationProduct()
        {
            // Prepare
            string ProdName = "PROD" + DateTime.Now.Day + DateTime.Now.Month + DateTime.Now.Hour + DateTime.Now.Minute + DateTime.Now.Second;

            //arrange
            HomePage homePage = LogInAsAdmin();
            
            var CatalogPage = homePage.GoToCustomers_CatalogPage();
            try
            {
                //act
                CatalogPage.ResetFilters();
                var ProductModalPage = CatalogPage.GoToProductPage();
                var ProductCreateModalPage = ProductModalPage.ClickProductCreatePage();
                CatalogPage = ProductCreateModalPage.FillField_CreatNewProduct(ProdName);
                homePage.GoToCustomers_CatalogPage();
                CatalogPage.GoToProductPage();
                CatalogPage.FilterProd(CatalogPage.FilterType.Search, ProdName);
                CatalogPage.SelectFirstItemProduct();
                bool isGeneralInfoNotEmpty = CatalogPage.GeneralInfoNotEmpty(ProdName);
                Assert.IsTrue(isGeneralInfoNotEmpty, "L'onglet General information ne doit pas apparaître vide.");
            }

            finally
            {
                homePage.GoToCustomers_CatalogPage();
                var ProductModalPage = CatalogPage.GoToProductPage();
                if (!string.IsNullOrEmpty(ProdName))
                {

                    CatalogPage.FilterProd(PageObjects.Customers.Catalogs.CatalogPage.FilterType.Search, ProdName);
                    CatalogPage.DeleteFirstProd();
                }
            }

        }
        [Timeout(_timeout)]
		[TestMethod]
        public void CU_CAT_AddItemProduct_SpecialCaractere()
        {
            //Prepare
            string ProdName = "PROD" + DateTime.Now.Day + DateTime.Now.Month + DateTime.Now.Year + DateTime.Now.Hour + DateTime.Now.Minute + DateTime.Now.Second + random.Next(0, 999);
            string site = TestContext.Properties["Site"].ToString();
            string ingredient = "&" + "test" + new Random().Next().ToString();
            string group = TestContext.Properties["Item_Group"].ToString();
            string workshop = TestContext.Properties["Item_Workshop"].ToString();
            string taxType = TestContext.Properties["Item_TaxType"].ToString();
            string prodUnit = TestContext.Properties["Item_ProdUnit"].ToString();
            string packagingName = TestContext.Properties["Item_PackagingName"].ToString();
            string supplier = TestContext.Properties["SupplierPo"].ToString();
            string storageUnit = "KG";
            string storageQty = "10";
            string qty = "10";

            //Arrange
            HomePage homePage= LogInAsAdmin();

            //ACT
            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();
            var itemCreateModalPage = itemPage.ItemCreatePage();
            var itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(ingredient, group, workshop, taxType, prodUnit);
            var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
            itemGeneralInformationPage = itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty, supplier, "2");

            var CatalogPage = homePage.GoToCustomers_CatalogPage();
            var ProductModalPage = CatalogPage.GoToProductPage();
            var ProductCreateModalPage = ProductModalPage.ClickProductCreatePage();
            bool addedProduct = CatalogPage.SearchProductByName(ingredient);
            Assert.IsTrue(addedProduct, "Le nom du produit ne contient pas le caractère spécial attendu");
            CatalogPage = ProductCreateModalPage.FillField_CreatNewProduct_SpecialCaractere(ProdName, ingredient);
            CatalogPage.FilterProd(CatalogPage.FilterType.Search, ProdName);
            bool ingredientFound = ProductCreateModalPage.IsIngredientPresent(ingredient);
           
            //Assert
            Assert.IsTrue(ingredientFound, $"L'ingrédient '{ingredient}' n'a pas été trouvé dans le catalogue.");
        }
    }
}
