using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects;

namespace Newrest.Winrest.FunctionalTests.Parameters
{
    [TestClass]
    public class RedistributiveCustomerTaxTest : RedistributiveTaxTestBase
    {
        private const int Timeout = 600000;
        [TestMethod]
        [Timeout(Timeout)]
        public void PA_AC_RCT_RedistributiveCustomerTaxTest()
        {
            HomePage homePage = LogInAsAdmin();
            var error = false;
            var message = string.Empty;
            int count_before = CheckRedistributiveCustomerTaxCount();
            CreateRedistributiveCustomerTax($"code_{count_before + 1}");
            int count_after = CheckRedistributiveCustomerTaxCount();
            Assert.AreEqual(count_before + 1, count_after, "La redistributive tax customer n'a pas pu être créée");
            Assert.IsTrue(error == false, message);
        }
    }
}
