using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Warehouse.ReceiptNotes;
using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Warehouse.Inventory
{
    public class InventoryExpiry : ReceiptNoteExpiry
    {
        public InventoryExpiry(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {

        }
    }
}
