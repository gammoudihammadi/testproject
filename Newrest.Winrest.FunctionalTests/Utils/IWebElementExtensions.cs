using System;
using DocumentFormat.OpenXml.Bibliography;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using static Newrest.Winrest.FunctionalTests.PageObjects.Shared.PageBase;

namespace Newrest.Winrest.FunctionalTests.Utils
{
    public static class IWebElementExtensions
    {       
        public static void SetValue(this IWebElement webElement, ControlType controlType, object value, string expectedFormat = null)
        {          
            switch (controlType)
            {
                case ControlType.DropDownList:
                    if (!(value is string)) throw new ArgumentException();

                    SelectElement select = new SelectElement(webElement);
                    //select.SelectByText((string)value, true);
                    // [YC] ou YC PREMIUM ?
                    select.SelectByText((string)value, false);
                    break;
                case ControlType.TextBox:                  
                    if (!(value is string)) throw new ArgumentException();            
                    webElement.ClearElement();
                    webElement.Click();
                    webElement.SendKeys((string)value);
                    break;
                case ControlType.CheckBox:
                    if (!(value is bool)) throw new ArgumentException();
                    if ((bool)value == false && webElement.Selected == true)
                    {
                        webElement.Click();
                    }
                    if ((bool)value == true && webElement.Selected == false)
                    {
                        webElement.Click();
                    }
                    break;
                case ControlType.RadioButton:
                    webElement.Click();
                    break;
                case ControlType.DateTime:

                    if (!(value is DateTime)) throw new ArgumentException();

                    if(expectedFormat != null)
                    {
                        webElement.Clear();
                        webElement.Click();
                        webElement.SendKeys(((DateTime)value).ToString(expectedFormat));
                    }
                    else if (webElement.GetAttribute("data_data_format")!=null)
                    {
                        var formattedDate = ((DateTime)value).ToString("yyyy-MM-dd");
                        webElement.Clear();
                        webElement.Click();
                        webElement.SendKeys(formattedDate);
                    }
                    else
                    {
                        var dateFormat = webElement.GetAttribute("data-date-format").Replace("mm", "MM").Replace(" (DD)", "");
                        webElement.Clear();
                        webElement.Click();
                        webElement.SendKeys(((DateTime)value).ToString(dateFormat));
                    }
                    break;
                default:
                    throw new ArgumentOutOfRangeException(nameof(controlType), controlType, null);
            }
        }

            public static void ClearElement(this IWebElement webElement)
        {
            webElement.Click();
            if (!webElement.GetAttribute("value").Equals(""))
            {
                //webElement.Clear();
                //webElement.SendKeys(Keys.End);
                webElement.SendKeys(Keys.Control + "a");
                webElement.SendKeys(Keys.Backspace);
            }
        }
            public static void ClickIfStale(this IWebElement webElement,By by,IWebDriver webDriver)
            {
                try
                {
                    webElement.Click();
                }
                catch (StaleElementReferenceException) 
                {
                    WaitForPageToLoad(webDriver,30);
                    webElement = webDriver.FindElement(by);
                    webElement.Click();
                }
            }
            private static void WaitForPageToLoad(IWebDriver driver, int timeoutInSeconds)
            {
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(timeoutInSeconds));

            // Wait for the document to be in a complete state and for all AJAX requests to complete
            wait.Until(_=> ((IJavaScriptExecutor)_).ExecuteScript(
                "return document.readyState === 'complete' && jQuery.active === 0").Equals(true));
            }
    }
}
