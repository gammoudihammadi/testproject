using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Newrest.Winrest.FunctionalTests.Parameters
{
    [TestClass]
    public class RedistributiveSupplierTaxTest : RedistributiveTaxTestBase
    {
        private const int Timeout = 600000;

        [TestMethod]
        [Timeout(Timeout)]
        public void PA_AC_RST_RedistributiveSupplierTaxTest()
        {
            var homePage = LogInAsAdmin();
            bool error = false;
            string message = string.Empty;
            int count_before = CheckRedistributiveSupplierTaxCount();
            CreateRedistributiveSupplierTax($"code_{count_before + 1}");
            int count_after = CheckRedistributiveSupplierTaxCount();
            Assert.AreEqual(count_before + 1, count_after, "La redistributive tax supplier n'a pas pu être créée");
            Assert.IsTrue(error == false, message);
        }
    }
}
