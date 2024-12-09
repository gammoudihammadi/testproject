using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using System;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Parameters.Tablet
{
    public class ParametersTablet : PageBase
    {
        private const string WORKSHOPS_TAB = "//*[@id=\"paramTabletTab\"]/li[3]/a";
        private const string CHECKBOXES = "//*[@id=\"workshop-type-table\"]/tbody/tr[*]/td[2]/input[contains(@value,'{0}')]/../../td[10]/div/div/input[@type=\"checkbox\"]";
        private const string CHECKBOXES_IS_TIME_BLOCK = "//*[@id=\"workshop-type-table\"]/tbody/tr[*]/td[2]/input[contains(@value,'{0}')]/../../td[16]/div/div/input[@type=\"checkbox\"]";
        private const string ORDERS = "//*[@id=\"workshop-type-table\"]/tbody/tr[*]/td[2]/input[contains(@value,'{0}')]/../../td[13]/input";
        private const string EDIT_BUTTON = "//*[@id=\"flight-type-table\"]/tbody/tr[*]/td[1]/input[@value='{0}']/../../td[7]/a/span";
        private const string LABELS_WORKSHOPS = "/html/body/div[2]/div/div/div/div[2]/div/table/tbody/tr[*]/td[2]/input";
        private const string SITE_WORKSHOP = "//*[@id=\"workshop-type-table\"]/tbody/tr[*]/td[2]/input[contains(@value,'{0}')]/../../td[15]/a/span";


        public ParametersTablet(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {

        }

        // ___________________________________ Constantes ______________________________________________



        // __________________________________ Méthodes __________________________________________________
        public void ClickFlightTypeTab()
        {
            var flightTypeBtn = WaitForElementIsVisible(By.XPath("//*[@id=\"tab-flight-type\"]/a"));
            flightTypeBtn.Click();
            WaitPageLoading();
            WaitForLoad();
        }
        public bool VerifyInternationalExistOnSite(string flightTypeString, string siteName)
        {
            var siteToSet = WaitForElementIsVisible(By.Id("cbSites"));
            siteToSet.SetValue(ControlType.DropDownList, siteName);
            try
            {
                var flightType = WaitForElementExists(By.XPath("//*[@id='flight-type-table']/tbody/tr[*]/td[1]/input[@value='"+flightTypeString+"']"));
                if (flightType.GetAttribute("value") == flightTypeString)
                {
                    return true;
                }
               
            }
            catch(Exception ex)
            {
                //no flight types
                return false;
            }
            return false;
            
        }

        public void CreateNewFlightType(string flightType, string color, string site)
        {
            var ShowPlus = WaitForElementIsVisible(By.XPath("//*/button[text()='+']"));
            new Actions(_webDriver).MoveToElement(ShowPlus).Perform();
            var newBtn = WaitForElementIsVisible(By.XPath("//*/a[text()='New']"));
            //var newBtn = WaitForElementIsVisible(By.XPath("//*[@id=\"tabContentItemContainer\"]/div[1]/a"));
            newBtn.Click();
            var selectSite = WaitForElementIsVisible(By.Id("selectSite"));
            selectSite.SetValue(ControlType.DropDownList, site);
            var name = WaitForElementIsVisible(By.Id("Name"));
            name.SendKeys(flightType);
            var colorInput = WaitForElementIsVisible(By.Id("colors"));
            colorInput.SetValue(ControlType.DropDownList, color);
            // Please choose at least one filter before saving
            ComboBoxSelectById(new ComboBoxOptions("SelectedCustomers_ms", "AIR FRANCE", false));
            var submit = WaitForElementIsVisible(By.Id("submit-button"));
            submit.Click();
        }
        public void EditFlightTypeColor(string flighType , string color, string site)
        {
            if (IsDev())
            {
                var editModal = WaitForElementIsVisible(By.XPath(String.Format("//*[@id='flight-type-table']/tbody/tr[*]/td[1]/input[@value='International']/../../td[7]/a", flighType)));
                editModal.Click();
            }
            else
            {
                var editModal = WaitForElementIsVisible(By.XPath(String.Format(EDIT_BUTTON, flighType)));
                editModal.Click();
            }
            WaitForLoad();

            var colorInput = WaitForElementIsVisible(By.Id("colors"));
            colorInput.SetValue(ControlType.DropDownList, color);
            WaitForLoad();

            ComboBoxSelectById(new ComboBoxOptions("SelectedRoutes_ms", site+"-"+ site, false));

            var submit = WaitForElementIsVisible(By.Id("submit-button"));
            submit.Click();
            WaitPageLoading();
            WaitForLoad();
        }
        public void WorkshopsTab()
        {
            var workshopsTab = WaitForElementIsVisible(By.XPath(WORKSHOPS_TAB));
            workshopsTab.Click();
            WaitForLoad();
        }
        public void ShowOnTabletByLabel(params string[] values)
        {
            foreach(var value in values)
            {
                var checkbox = WaitForElementExists(By.XPath(String.Format(CHECKBOXES,value)));
                if(!checkbox.Selected)
                {
                    checkbox.SetValue(ControlType.CheckBox, true);
                }
            }
        }
        public void DesactiverTabletByLabel(params string[] values)
        {
            var listVal = _webDriver.FindElements(By.XPath(LABELS_WORKSHOPS));
            foreach (var value in listVal)
            {
                bool ext = false;
                for(var i = 0; i< values.Length; i++)
                {
                    if(values[i] == value.GetAttribute("value"))
                    {
                        ext = true;
                        break;
                    }
                    else
                    {
                        ext = false;
                    }
                }
                if(!ext)
                {
                    var checkbox = WaitForElementExists(By.XPath(String.Format(CHECKBOXES, value.GetAttribute("value"))));
                    if (checkbox.Selected)
                    {
                        checkbox.SetValue(ControlType.CheckBox, false);
                    }
                }
            }
        }
        public void AffecterOrdre(string workshop , int ordre)
        {
            var input = WaitForElementExists(By.XPath(String.Format(ORDERS, workshop)));
            if(int.Parse(input.GetAttribute("value")) != ordre)
            {
                input.Clear();
                input.SendKeys(ordre.ToString());
            }
        }
        public void AllWorkshopIsTimeBlock(params string[] values)
        {
            WaitForLoad();
            var listVal = _webDriver.FindElements(By.XPath(LABELS_WORKSHOPS));
            foreach (var value in listVal)
            {
                bool ext = false;
                for (var i = 0; i < values.Length; i++)
                {
                    if (values[i] == value.GetAttribute("value"))
                    {
                        ext = true;
                        break;
                    }
                    else
                    {
                        ext = false;
                    }
                }
                if (ext)
                {
                    var checkbox = WaitForElementExists(By.XPath(String.Format(CHECKBOXES_IS_TIME_BLOCK, value.GetAttribute("value"))));
                    if (!checkbox.Selected)
                    {
                        checkbox.SetValue(ControlType.CheckBox, true);
                    }
                }
            }
        }
        public void WorkshopExistOnSite(string siteName, params string[] values)
        {
            for (var i = 0; i < values.Length; i++)
            {
                //open Add Sites Model
                var siteToSet = WaitForElementIsVisible(By.XPath(string.Format(SITE_WORKSHOP, values[i])));
                siteToSet.Click();
                WaitForLoad();

                ComboBoxSelectById(new ComboBoxOptions("Sites_ms", siteName, false));

                var saveButton = WaitForElementIsVisible(By.XPath("//*[@id=\"last\"]"));
                saveButton.Click();
                WaitPageLoading();
                WaitForLoad();
                _webDriver.Navigate().Refresh();
                WorkshopsTab();
            }
        }
    }
}

