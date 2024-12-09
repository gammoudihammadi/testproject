using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime;
using DocumentFormat.OpenXml.Drawing.Charts;
using System.Globalization;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Interim.InterimOrders
{
    public class InterimOrdersGeneralInformation : PageBase
    {

        private const string INTERIM_NUMBER = "tb-new-interim-reception-number";
        private const string COMMENT = "InterimOrder_Comment";
        private const string CREATION_DATE = "//*[@id=\"form-create-interim-order\"]/div/div[6]/div/div/div/input";
        private const string VALIDATION_DATE = "//*[@id=\"form-create-interim-order\"]/div/div[5]/div[1]/div/div/input";
        private const string VALIDATED_BY = "InterimOrder_UserValidator_FullName";
        private const string DELIVERY_DATE = "//*[@id=\"form-create-interim-order\"]/div/div[4]/div/div/div/input";


        //__________________________________ Variables _________________________________________________

        [FindsBy(How = How.Id, Using = COMMENT)]
        private IWebElement _comment;

        [FindsBy(How = How.XPath, Using = VALIDATION_DATE)]
        private IWebElement _validationDate;

        [FindsBy(How = How.XPath, Using = VALIDATED_BY)]
        private IWebElement _validationBy;

        [FindsBy(How = How.XPath, Using = CREATION_DATE)]
        private IWebElement _creationDate;
        [FindsBy(How = How.Id, Using = DELIVERY_DATE)]
        private IWebElement _deliveryDate;

        public InterimOrdersGeneralInformation(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        //__________________________________Methodes_______________________________________

        public string GetComment()
        {
            _comment = WaitForElementIsVisible(By.Id(COMMENT));
            return _comment.GetAttribute("value");
        }

        public string CreationDate()
        {
            _creationDate = WaitForElementIsVisible(By.XPath(CREATION_DATE));
            return _creationDate.GetAttribute("value");

        }

        public string ValidationDate()
        {
            _validationDate = WaitForElementIsVisible(By.XPath(VALIDATION_DATE));
            return _validationDate.GetAttribute("value");
        }

        public string ValidatedBy()
        {
            _validationBy = WaitForElementIsVisible(By.Id(VALIDATED_BY));
            return _validationBy.GetAttribute("value");
        }

        public bool IsValidDate(string dateString, string format, out DateTime dateTime)
        {
            return DateTime.TryParseExact(dateString, format,
                                          CultureInfo.InvariantCulture,
                                          DateTimeStyles.None,
                                          out dateTime);
        }


    }
}
