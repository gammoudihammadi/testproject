using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;
using System;
using System.Text.RegularExpressions;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Accounting.Invoice
{
    public class InvoiceGeneralInformations : PageBase
    {
        public InvoiceGeneralInformations(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }


        // _____________________________________________ Constantes ______________________________________________

        // Général
        private const string BACK_TO_LIST = "/html/body/div[2]/a/span[2]";

        // Onglets
        private const string DETAILS_TAB = "hrefTabContentItems";
        private const string FOOTER_TAB = "hrefTabInvoiceFooter";

        // Informations
        private const string COMMENT = "Comment";
        private const string DATE = "InvoiceDate";
        private const string ENGAGEMENT_NO = "EngagementNumber";
        private const string APPLY_VAT = "IsVATForced";
        private const string INVOICE_ID = "//*[@id='div-body']/div/div[1]/h1";

        // _____________________________________________ Variables _______________________________________________

        // Général
        [FindsBy(How = How.XPath, Using = BACK_TO_LIST)]
        private IWebElement _backToList;


        //INVOICE_ID
        [FindsBy(How = How.XPath, Using = INVOICE_ID)]
        private IWebElement _invoiceIdTitle;

        // Onglets
        [FindsBy(How = How.XPath, Using = DETAILS_TAB)]
        private IWebElement _detailsTab;

        [FindsBy(How = How.XPath, Using = FOOTER_TAB)]
        private IWebElement _footerTab;

        // Informations
        [FindsBy(How = How.Id, Using = COMMENT)]
        private IWebElement _comment;

        [FindsBy(How = How.Id, Using = DATE)]
        private IWebElement _date;

        [FindsBy(How = How.Id, Using = ENGAGEMENT_NO)]
        private IWebElement _engagementNo;

        // _____________________________________________ Méthodes___________________________________________________

        // General
        public InvoicesPage BackToList()
        {
            _backToList = WaitForElementIsVisible(By.XPath(BACK_TO_LIST));
            _backToList.Click();
            WaitForLoad();
            return new InvoicesPage(_webDriver, _testContext);
        }

        // Onglets
        public InvoiceDetailsPage ClickOnDetails()
        {
            _detailsTab = WaitForElementIsVisible(By.Id(DETAILS_TAB));
            _detailsTab.Click();
            WaitForLoad();

            return new InvoiceDetailsPage(_webDriver, _testContext);
        }

        public InvoiceFooterPage ClickOnFooter()
        {
            _footerTab = WaitForElementIsVisible(By.Id(FOOTER_TAB));
            _footerTab.Click();
            WaitForLoad();

            return new InvoiceFooterPage(_webDriver, _testContext);
        }

        // Informations
        public void SetComment(string comment)
        {
            _comment = WaitForElementIsVisible(By.Id(COMMENT));
            _comment.SetValue(ControlType.TextBox, comment);
            WaitPageLoading();
        }

        public string GetComment()
        {
            _comment = WaitForElementExists(By.Id(COMMENT));
            return _comment.Text;
        }

        public void SetDate(DateTime date)
        {
            _date = WaitForElementExists(By.Id(DATE));
            _date.SetValue(ControlType.DateTime, date);
            WaitPageLoading();
        }

        public string GetDate()
        {
            _date = WaitForElementExists(By.Id(DATE));
            return _date.GetAttribute("value");
        }


        public void SetEngagementNo(int no)
        {
            _engagementNo = WaitForElementExists(By.Id(ENGAGEMENT_NO));
            _engagementNo.SetValue(ControlType.TextBox, "" + no);
            WaitPageLoading();
        }

        public string GetEngagementNo()
        {
            _engagementNo = WaitForElementExists(By.Id(ENGAGEMENT_NO));
            return _engagementNo.GetAttribute("value");
        }

        public void SetApplyVAT(bool applyVAT)
        {
            _engagementNo = WaitForElementExists(By.Id(APPLY_VAT));
            _engagementNo.SetValue(ControlType.CheckBox, applyVAT);
            WaitPageLoading();
        }

        public bool GetApplyVAT()
        {
            _engagementNo = WaitForElementExists(By.Id(APPLY_VAT));
            return _engagementNo.Selected;
        }

        public string GetInvoiceId()
        {
            _invoiceIdTitle = WaitForElementIsVisible(By.XPath(INVOICE_ID));
            // INVOICE : 125178 - EMIRATES AIRLINES
            string pattern = ".*: ([0-9]+) -.*";
            Regex rg = new Regex(pattern);
            Match matchedId = rg.Match(_invoiceIdTitle.Text);
            Assert.AreEqual(2, matchedId.Groups.Count);
            string invoiceId = matchedId.Groups[1].Value;
            return invoiceId;
        }

    }
}
