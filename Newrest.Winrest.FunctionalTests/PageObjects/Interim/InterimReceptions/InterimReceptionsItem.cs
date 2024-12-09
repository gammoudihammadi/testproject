using DocumentFormat.OpenXml.Office2013.Drawing.ChartStyle;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Purchasing;
using Newrest.Winrest.FunctionalTests.PageObjects.Interim.InterimOrders;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.PageObjects.Warehouse.Claims;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.PageObjects;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenQA.Selenium.Support.UI;
using DocumentFormat.OpenXml.Office2013.Word;
using System.Globalization;
using System.Text.RegularExpressions;
using System.Threading;
namespace Newrest.Winrest.FunctionalTests.PageObjects.Interim.InterimReceptions
{
    public class InterimReceptionsItem : PageBase
    {
        public InterimReceptionsItem(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        //___________________________________________ Constantes ____________________________________________________

        // General 
        private const string BACK_TO_LIST = "/html/body/div[2]/a/span[2]";
        private const string BACK_TO_LIST_PRINCIPAL = "/html/body/div[2]/a";
        private const string GO_TO_GENERAL_INFORMATION = "hrefTabContentInformations";

        private const string GROUPS_FILTER = "ItemIndexVMSelectedGroups_ms";
        private const string UNSELECT_ALL_GROUPS = "/html/body/div[11]/div/ul/li[2]/a/span[2]";
        private const string GROUPS_SEARCH = "/html/body/div[11]/div/div/label/input";
        private const string GROUPS_NAME = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[1]/div/div/span";
        private const string GROUPS_TO_CHECK = "//*[@id=\"ui-multiselect-0-ItemIndexVMSelectedGroups-option-0\"]";
        private const string KEYWORD = "tbSearchByKeyword";
        private const string ITEM_LIGNES = "//*[@id=\"itemForm_0\"]/div[2]";
        private const string ITEM_RECIVED = "item_IrdRowDto_NewReceivedQuantity";
        private const string CHECK_BUTTON = "//*[@id=\"div-body\"]/div/div[1]/div/div[2]/button";
        private const string VALIDATE = "btn-validate-interim-reception";
        private const string VALIDATE_POP_UP = "btn-popup-validate";
        private const string RECEPTION_ORDER = "//*[@id=\"InterimOrderTabNav\"]/a";
        private const string NAME_REF= "//*[@id=\"itemForm_0\"]/div/div[3]";
        private const string NAME_REF_FILTER= "tbSearchPatternWithAutocomplete";
        private const string NOMBRE_ROW_ITEM_LIGNES = "/html/body/div[3]/div/div[3]/div/div/div/div[2]/div/div/div/div[2]/div[2]/div/div/form/div";
        private const string FIRST_LINE_NOT_VALIDATE = "//*[@id=\"list-item-with-action\"]/table/tbody/tr[1]";
        private const string FIRST_ITEM = "//*[@id=\"itemForm_0\"]/div[2]";
        private const string QUANTITE = "item_IrdRowDto_NewReceivedQuantity";
        private const string IMG_EDIT = "//*[@id=\"itemForm_0\"]/div[1]/div[1]/span/img";
        private const string SUB_GROUPES_FILTER = "ItemIndexVMSelectedSubGroups_ms";
        private const string SUB_GROUPES_SEARCH = "/html/body/div[12]/div/div/label/input";
        private const string SUB_GROUPES_TOCHECK = "//*[@id=\"ui-multiselect-1-ItemIndexVMSelectedSubGroups-option-0\"]";
        private const string ELEMENT_ITEM = "//*[@id=\"div-body\"]/div/div[1]/h1";
        private const string CRAYON_ITEM = "//*[@id=\"itemForm_0\"]/div[1]/div[10]/div/a[3]";
        private const string RECEIVED = "//*[@id=\"item_IrdRowDto_NewReceivedQuantity\"]";
        private const string INTERIMNUMBER = "//*[@id=\"div-body\"]/div/div[1]";
        private const string TOTAL_PRICE_SPAN = "//*[@id=\"total-price-span\"]";



        private const string ORDERED_INTERIM_RECEPTION_ITEM = "/html/body/div[3]/div/div[3]/div/div/div/div[2]/div/div/div/div[2]/div[2]/div/div/form/div[2]/div[7]/span";
        private const string RECEIVED_INTERIM_RECEPTION_ITEM = "/html/body/div[3]/div/div[3]/div/div/div/div[2]/div/div/div/div[2]/div[2]/div/div/form/div[2]/div[8]/span";
        private const string PROD_QTY_INTERIM_ORDER_ITEM = "/html/body/div[3]/div/div[3]/div/div/div/div[2]/div/div/div/div[2]/div[2]/div/div/form/div/div[8]/span";
        private const string DELIVRED_QTE = "/html/body/div[3]/div/div[3]/div/div/div/div[2]/div/div/div/div[2]/div[2]/div/div/form/div/div[8]/span";
        private const string DELIVRED_QTE_AFTER = "//*[@id=\"item_IrdRowDto_NewReceivedQuantity\"]";
        private const string TOTAL_PRICE = "//*[@id=\"itemForm_0\"]/div[1]/div[9]/span";
        private const string NEW_TOTAL_VAT_INTERIM_RECEPTION_ITEM = "/html/body/div[3]/div/div[3]/div/div/div/div[2]/div/div/div/div[2]/div[2]/div/div/form/div[2]/div[9]/span";
        private const string SELECTED_RECEIVED_INTERIM_RECEPTION_ITEM = "/html/body/div[3]/div/div[3]/div/div/div/div[2]/div/div/div/div[2]/div[2]/div/div/form/div/div[8]/span";
        private const string SELECTED_TOTAL_VAT_INTERIM_RECEPTION_ITEM = "/html/body/div[3]/div/div[3]/div/div/div/div[2]/div/div/div/div[2]/div[2]/div/div/form/div/div[9]/span";
        private const string DELETE_ICON = "//*[@id=\"itemForm_0\"]/div[1]/div[10]/div/a[2]/span";



        //___________________________________________ Variables ________________________________________________________________


        [FindsBy(How = How.XPath, Using = NOMBRE_ROW_ITEM_LIGNES)]
        private IWebElement _nombre_row_item_ligne;

        [FindsBy(How = How.XPath, Using = BACK_TO_LIST)]
        private IWebElement _backToList;
        [FindsBy(How = How.XPath, Using = BACK_TO_LIST_PRINCIPAL)]
        private IWebElement _backToListPrincipal;

        [FindsBy(How = How.Id, Using = GROUPS_FILTER)]
        private IWebElement _groupsFilter;

        [FindsBy(How = How.XPath, Using = UNSELECT_ALL_GROUPS)]
        private IWebElement _unselectAllGroups;

        [FindsBy(How = How.XPath, Using = GROUPS_SEARCH)]
        private IWebElement _searchGroups;

        [FindsBy(How = How.XPath, Using = GROUPS_TO_CHECK)]
        private IWebElement _groupToCheck;

        [FindsBy(How = How.Id, Using = GO_TO_GENERAL_INFORMATION)]
        private IWebElement _goToGeneralInformation;

        [FindsBy(How = How.Id, Using = KEYWORD)]
        private IWebElement _keyword;

        [FindsBy(How = How.Id, Using = ITEM_RECIVED)]
        private IWebElement _itemRecevied;

        [FindsBy(How = How.Id, Using = VALIDATE)]
        private IWebElement _validate;

        [FindsBy(How = How.Id, Using = VALIDATE_POP_UP)]
        private IWebElement _validatePopUp;

        [FindsBy(How = How.XPath, Using = CHECK_BUTTON)]
        private IWebElement _checkButton;

        [FindsBy(How = How.XPath, Using = RECEPTION_ORDER)]
        private IWebElement _receptionOrder;
        [FindsBy(How = How.XPath, Using = NAME_REF)]
        private IWebElement _name_ref;
        [FindsBy(How = How.XPath, Using = NAME_REF_FILTER)] 
        private IWebElement _nameref_filter;
        [FindsBy(How = How.Id, Using = QUANTITE)]
        private IWebElement _quantite;

        [FindsBy(How = How.XPath, Using = SUB_GROUPES_FILTER)]
        private IWebElement _subGroupsFilter;

        [FindsBy(How = How.XPath, Using = SUB_GROUPES_SEARCH)]
        private IWebElement _subGroupsSearch;

        [FindsBy(How = How.XPath, Using = SUB_GROUPES_TOCHECK)]
        private IWebElement _subGroupsToCheck;


        [FindsBy(How = How.XPath, Using = RECEIVED)]
        private IWebElement _received;
        
        [FindsBy(How = How.XPath, Using = INTERIMNUMBER)]
        private IWebElement _interimNumber;

        [FindsBy(How = How.XPath, Using = NEW_TOTAL_VAT_INTERIM_RECEPTION_ITEM)]
        private IWebElement _newTotalVatInterimReceptionItem;

        [FindsBy(How = How.XPath, Using = SELECTED_TOTAL_VAT_INTERIM_RECEPTION_ITEM)]
        private IWebElement _selectedTotalVatInterimReceptionItem;

        //___________________________________________ Méthodes _____________________________________________________


        public string showTotalPriceToCopy()
        {
            var interimOrderPriceElement = WaitForElementExists(By.XPath(TOTAL_PRICE_SPAN));
            WaitForLoad();
            return interimOrderPriceElement.Text;

        }
        public float showTotalPriceToCopy2()
        {
            WaitPageLoading();
            var interimOrderPriceElement = WaitForElementExists(By.XPath(TOTAL_PRICE_SPAN));
            string price = interimOrderPriceElement.Text.Trim();
            string cleanedPrice = Regex.Replace(price, @"[^\d,]", "");

            cleanedPrice = cleanedPrice.Replace(',', '.');
            if (decimal.TryParse(cleanedPrice, NumberStyles.Any, CultureInfo.InvariantCulture, out decimal result))
            {
                WaitForLoad();
                return (float)result;
            }
            throw new FormatException("Total price is not Parsed ");
        }

            public string getPriceAfter()
        {
            var interimOrderPriceElement = WaitForElementExists(By.XPath(TOTAL_PRICE));
            return interimOrderPriceElement.Text;
        }
        //public string showDateToCopy()
        //{
        //    var interimOrderDateElement = WaitForElementExists(By.XPath("//*[@id=\"list-item-with-action\"]/div/div[2]/div[2]/div/div/form/div/div[4]/span"));
        //    return interimOrderDateElement.Text;
        //}
        public string showDateToCopy()
        {
            var interimOrderDateElement = WaitForElementExists(By.XPath("//*[@id=\"list-item-with-action\"]/div/div[2]/div[2]/div/div/form/div/div[4]/span"));
            DateTime interimOrderDate = DateTime.Parse(interimOrderDateElement.Text);
            return interimOrderDate.ToString("dd/MM/yyyy");
        }

        // General       
        public InterimReceptionsPage BackToList()
        {
            WaitForLoad();
            _backToList = WaitForElementIsVisible(By.XPath(BACK_TO_LIST));
            _backToList.Click();

            return new InterimReceptionsPage(_webDriver, _testContext);
        }

        public enum FilterType
        {
            Groups,
            Keyword,
            Name,
            SubGroups
        }
        public void Filter(FilterType filterType, object value)
        {
            Actions action = new Actions(_webDriver);

            switch (filterType)
            {
                case FilterType.Groups:
                    _groupsFilter = WaitForElementIsVisible(By.Id(GROUPS_FILTER));
                    _groupsFilter.Click();

                    _unselectAllGroups = WaitForElementIsVisible(By.XPath(UNSELECT_ALL_GROUPS));
                    _unselectAllGroups.Click();

                    _searchGroups = WaitForElementIsVisible(By.XPath(GROUPS_SEARCH));
                    _searchGroups.SetValue(ControlType.TextBox, value);
                    WaitForLoad();

                    _groupToCheck = WaitForElementIsVisible(By.XPath(GROUPS_TO_CHECK));
                    _groupToCheck.Click();

                    _groupsFilter.Click();
                    break;

                case FilterType.Keyword:
                    _keyword = WaitForElementIsVisible(By.Id(KEYWORD));
                    _keyword.SetValue(ControlType.TextBox, value);
                    WaitForLoad();
                    break;
                case FilterType.Name:
                    _nameref_filter = WaitForElementIsVisible(By.Id(NAME_REF_FILTER));
                    _nameref_filter.SetValue(ControlType.TextBox, value);
                    break;

                case FilterType.SubGroups:
                    _subGroupsFilter = WaitForElementIsVisible(By.Id(SUB_GROUPES_FILTER));
                    _subGroupsFilter.Click();


                    _subGroupsSearch = WaitForElementIsVisible(By.XPath(SUB_GROUPES_SEARCH));
                    _subGroupsSearch.SetValue(ControlType.TextBox, value);
                    WaitForLoad();
                    if (IsDev())
                    {
                        _subGroupsToCheck = WaitForElementIsVisible(By.XPath("//*[@id=\"ui-multiselect-1-ItemIndexVMSelectedSubGroups-option-0\"]"));
                        _subGroupsToCheck.Click();
                    }
                    else
                    {
                        _subGroupsToCheck = WaitForElementIsVisible(By.XPath("//*[@id=\"ui-multiselect-1-ItemIndexVMSelectedSubGroups-option-1\"]"));
                        _subGroupsToCheck.Click();
                    }
                    _subGroupsFilter.Click();
                    break;
            }

            WaitPageLoading();
            WaitForLoad();
        }
        public InterimReceptionsGeneralInformation GoToGeneralInformation()
        {
            _goToGeneralInformation = WaitForElementIsVisible(By.Id(GO_TO_GENERAL_INFORMATION));
            _goToGeneralInformation.Click();
            WaitForLoad();
            return new InterimReceptionsGeneralInformation(_webDriver, _testContext);
        }
        public int NombreRowItem()
        {
            var itemrow = _webDriver.FindElements(By.XPath(ITEM_LIGNES));
            return itemrow.Count;
        }

        
        public int GetNombreRowItem()
        {
            var itemrow = _webDriver.FindElements(By.XPath(NOMBRE_ROW_ITEM_LIGNES));
            return itemrow.Count;
        }

        public int getDelQty()
        {
            var deliveredQty = _webDriver.FindElements(By.XPath(DELIVRED_QTE));
            return deliveredQty.Count;
        }
       
        
        public string GetReference(int i)
        {
            if (isElementVisible(By.XPath(String.Format("//*[@id=\"itemForm_{0}\"]/div/div[3]/span", i))))
            {
                var reference = WaitForElementIsVisible(By.XPath(String.Format("//*[@id=\"itemForm_{0}\"]/div/div[3]/span", i)));
                return reference.Text;
            }
            else
            {
                return "";
            }
        }

        public string GetDateItem(int i)
        {
            if (isElementVisible(By.XPath(String.Format("//*[@id=\"itemForm_{0}\"]/div/div[4]/span", i))))
            {
                var dateitem = WaitForElementIsVisible(By.XPath(String.Format("//*[@id=\"itemForm_{0}\"]/div/div[4]/span", i)));
                return dateitem.Text;
            }
            else
            {
                return "";
            }
        }
        public string GetPackaging(int i)
        {
            if (isElementVisible(By.XPath(String.Format("//*[@id=\"itemForm_{0}\"]/div/div[5]/span", i))))
            {
                var packaging = WaitForElementIsVisible(By.XPath(String.Format("//*[@id=\"itemForm_{0}\"]/div/div[5]/span", i)));
                return packaging.Text;
            }
            else
            {
                return "";
            }
        }
        public string GetPackagingPrice(int i)
        {
            if (isElementVisible(By.XPath(String.Format("//*[@id=\"itemForm_{0}\"]/div/div[6]/span", i))))
            {
                var packagingprice = WaitForElementIsVisible(By.XPath(String.Format("//*[@id=\"itemForm_{0}\"]/div/div[6]/span", i)));
                return packagingprice.Text;
            }
            else
            {
                return "";
            }
        }
        public string GetOrderedQty(int i)
        {
            if (isElementVisible(By.XPath(String.Format("//*[@id=\"itemForm_{0}\"]/div/div[7]/span", i))))
            {
                var orderedQt = WaitForElementIsVisible(By.XPath(String.Format("//*[@id=\"itemForm_{0}\"]/div/div[7]/span", i)));
                return orderedQt.Text;
            }
            else
            {
                return "";
            }
        }
        public string GetDeliveredQty(int i)
        {
            if (isElementVisible(By.XPath(String.Format("//*[@id=\"itemForm_{0}\"]/div/div[8]/span", i))))
            {
                var deliveredQty = WaitForElementIsVisible(By.XPath(String.Format("//*[@id=\"itemForm_{0}\"]/div/div[8]/span", i)));
                return deliveredQty.Text;
            }
            else
            {
                return "";
            }
        }
        public string GetTotalVAT(int i)
        {
            if (isElementVisible(By.XPath(String.Format("//*[@id=\"itemForm_{0}\"]/div/div[9]/span", i))))
            {
                var deliveredQty = WaitForElementIsVisible(By.XPath(String.Format("//*[@id=\"itemForm_{0}\"]/div/div[9]/span", i)));
                return deliveredQty.Text;
            }
            else
            {
                return "";
            }
        }

        public string GetTotalWOVAT()
        {
           var deliveredQty = WaitForElementIsVisible(By.XPath(TOTAL_PRICE_SPAN));
           return deliveredQty.Text;                   
        }
        public float GetTotalWOVATinNumbers()
        {
            WaitPageLoading();
            var totalPrice = WaitForElementIsVisible(By.XPath(TOTAL_PRICE_SPAN));
            string priceText = totalPrice.Text.Trim();

            string cleanedPrice = Regex.Replace(priceText, @"[^\d,]", "");

            cleanedPrice = cleanedPrice.Replace(',', '.');
            WaitPageLoading();
            if (float.TryParse(cleanedPrice, NumberStyles.Any, CultureInfo.InvariantCulture, out float result))
            {
                WaitForLoad();
                WaitPageLoading();
                return result;
            }
            throw new FormatException("The total price is  not parsed");
        }
        public string ReturnInterimReceptionNumber()
        {
            _interimNumber = WaitForElementIsVisible(By.XPath(INTERIMNUMBER));
            string orderNumber = _interimNumber.Text.Trim();
            Match match = Regex.Match(orderNumber, @"\d+");

            // Check if a match was found and return the numeric part
            if (match.Success)
            {
                WaitForLoad();
              
                return match.Value;
            }
            return orderNumber;


        }


        public void SetReceptionItemReceived()
        {
            Random rnd = new Random();

            var edit  = WaitForElementIsVisible(By.XPath("//*[@id=\"itemForm_0\"]/div[2]"));
            edit.Click();
            _itemRecevied = WaitForElementIsVisible(By.Id(ITEM_RECIVED));
            _itemRecevied.SetValue(ControlType.TextBox, rnd.Next(1,10).ToString());
            _itemRecevied.SendKeys(Keys.Enter);
            WaitForLoad();
        }
        public void Validate()
        {
            WaitPageLoading();
            _checkButton = WaitForElementIsVisible(By.XPath(CHECK_BUTTON));
            _checkButton.Click();
            _validate = WaitForElementIsVisible(By.Id(VALIDATE));
            _validate.Click();
            _validatePopUp = WaitForElementIsVisible(By.Id(VALIDATE_POP_UP));
            _validatePopUp.Click();
            WaitForLoad();
            WaitPageLoading();


        }
        public InterimOrdersPage GoToReceptionOrder()
        {
            _receptionOrder = WaitForElementExists(By.XPath(RECEPTION_ORDER));
            _receptionOrder.Click();

            return new InterimOrdersPage(_webDriver, _testContext);
        }
        public string GetFirstInterimReceptionsNameRef()
        {
            _name_ref = WaitForElementIsVisible(By.XPath(NAME_REF));
            return _name_ref.Text;
        }
        public List<string> GetItemsForInterimFiltred()
        {
            var listElements = _webDriver.FindElements(By.XPath("//*[@id=\"itemForm_0\"]/div/div[3]"));
            if (listElements.Count == 0)
            {
                // No item found; return null
                return new List<string>();
            }
            var allOrders = new List<string>();
            foreach (var listElement in listElements)
            {
                var ordersText = listElement.Text.Trim();
                var ordersList = ordersText.Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries).Select(order => order.Trim()).ToList();
                allOrders.AddRange(ordersList);
            }
            return allOrders;
        }
        public void refresh()
        {
            WaitForLoad();
            var refresh = WaitForElementIsVisible(By.XPath("//*[@id=\"div-body\"]/div/div[1]/div/div[1]/button"));
            refresh.Click();
        }
        public ItemGeneralInformationPage EditItem()
        {
            var edit = WaitForElementIsVisible(By.XPath("//*[@id=\"itemForm_0\"]/div[2]"));
            edit.Click();
            WaitForLoad();
            if (isElementExists(By.XPath("//*[@id=\"itemForm_0\"]/div[1]/div[10]/div/a[3]/span")))
            {

                var item = WaitForElementIsVisible(By.XPath("//*[@id=\"itemForm_0\"]/div[1]/div[10]/div/a[3]/span"));
                item.Click();
            }
            else
            {
                return null;
            }
            return new ItemGeneralInformationPage(_webDriver, _testContext);
        }
        public void Go_To_New_Navigate()

        {

            var wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(60));

            wait.Until((driver) => driver.WindowHandles.Count > 1);

            _webDriver.SwitchTo().Window(_webDriver.WindowHandles[1]);

            WaitPageLoading();

        }
        public void Go_To_Old_Navigate()

        {

            var wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(60));

            wait.Until((driver) => driver.WindowHandles.Count > 1);

            _webDriver.SwitchTo().Window(_webDriver.WindowHandles[0]);

            WaitPageLoading();

        }

        public bool CheckNameExistSInList(string name)
        {
            var listOfNames = _webDriver.FindElements(By.XPath("//*[starts-with(@id,\"itemForm\")]/div[2]/div[3]/span"));
            List<string> names = new List<string>();
            foreach (var element in listOfNames) 
            {
                names.Add(element.Text.ToString());
            }
            return names.Contains(name);
        }


        public void ClickOnItem()
        {
            WaitForLoad();
            var firstLineElement = _webDriver.FindElement(By.XPath(FIRST_ITEM));
            firstLineElement.Click();
            WaitForLoad();
        }
        public void SetQty(string qty)
        {
            WaitForLoad();
            if (!isElementVisible(By.Id(QUANTITE)))
            {
                var edit = WaitForElementIsVisible(By.XPath("//*[@id=\"item_IrdRowDto_NewReceivedQuantity\"]"));
                edit.Click();
            }
            _quantite = WaitForElementIsVisible(By.Id(QUANTITE));
            _quantite.SetValue(PageBase.ControlType.TextBox, qty);
            WaitForLoad();
        }

        public bool IsVisible()
        {
            WaitPageLoading();
            WaitForLoad();
            if (isElementVisible(By.XPath(IMG_EDIT)))
            {
                return true;
            }
            else
            {
                return false;
            }
        }
        public void SetReceived(string qty)
        {

            if (!isElementVisible(By.XPath(RECEIVED)))
            {
                var edit = WaitForElementIsVisible(By.XPath("//*[@id=\"itemForm_0\"]/div[2]/div[4]"));
                edit.Click();
            }
            _received = WaitForElementIsVisible(By.XPath(RECEIVED));
            _received.SetValue(PageBase.ControlType.TextBox, qty);
            WaitForLoad();
            LoadingPage();
        }


        public string GetItemName()
        {
            
            var name =  WaitForElementExists(By.XPath("//*[@id=\"itemForm_0\"]/div[2]/div[3]/span"));
            return name.Text.ToString(); 
        }
        public void DeleteItem()
        {
            WaitPageLoading();
            var deleteIcon = WaitForElementIsVisible(By.XPath(DELETE_ICON));
            deleteIcon.Click();
            WaitPageLoading();
            WaitForLoad();
        }

        public float GetQuanity()
        {
            var edit = WaitForElementIsVisible(By.XPath("//*[@id=\"itemForm_0\"]/div[2]/div[4]"));
            edit.Click();
            WaitForLoad();
            _received = WaitForElementIsVisible(By.XPath(RECEIVED));
            return float.Parse(_received.GetProperty("value"));
        }


        
        public void ClickOnCrayon()
        {
            var crayon = _webDriver.FindElement(By.XPath(CRAYON_ITEM));
            crayon.Click();
            WaitForLoad();
        }
        public bool ElementIsVisible()
        {
            WaitForLoad();
            if (isElementVisible(By.XPath(ELEMENT_ITEM)))
            {
                return true;
            }
            else
            {
                return false;
            }
        }
        public int RowsNumber()
        {
            var row = _webDriver.FindElements(By.XPath("/html/body/div[3]/div/div[3]/div/div/div/div[2]/div/div/div/div[2]/div[*]/div/div/form/div[2]"));
            return row.Count;
        }
        public string GetReceptionNumber()
        {
            var element = _webDriver.FindElement(By.XPath("//*[@id=\"div-body\"]/div/div[1]/h1"));
            string text = element.Text;
            string pattern = @"NO°(\d+)";
            var match = System.Text.RegularExpressions.Regex.Match(text, pattern);

            if (match.Success)
            {
                WaitForLoad();
                WaitPageLoading();
                return match.Groups[1].Value;
            }
            else
            {
                throw new Exception("Reception number not found in the text.");
            }
        }

        public string GetOrderedInterimReceptionItem()
        {
            WaitPageLoading();
            if (isElementVisible(By.XPath(ORDERED_INTERIM_RECEPTION_ITEM)))
            {
                var ordered = WaitForElementIsVisible(By.XPath(ORDERED_INTERIM_RECEPTION_ITEM));
                return ordered.Text;
            }
            else
            {
                return "";
            }
        }
        public int GetReceivedInterimReceptionItem()
        {
            WaitPageLoading();
            var received = WaitForElementExists(By.XPath(RECEIVED_INTERIM_RECEPTION_ITEM));
                return int.Parse(received.Text);
        }
        public string GetProdQtyInterimOrderItem()
        {
            WaitPageLoading();
            if (isElementVisible(By.XPath(PROD_QTY_INTERIM_ORDER_ITEM)))
            {
                var ordered = WaitForElementIsVisible(By.XPath(PROD_QTY_INTERIM_ORDER_ITEM));
                return ordered.Text;
            }
            else
            {
                return "";
            }
        }

        public decimal GetTotalVatInterimReceptionItem()
        {
            WaitPageLoading();
            _newTotalVatInterimReceptionItem = WaitForElementExists(By.XPath(NEW_TOTAL_VAT_INTERIM_RECEPTION_ITEM));
            var totalVat = _newTotalVatInterimReceptionItem.Text;

            var cleanTotalVatText = totalVat.Replace("€", "").Trim();

            return decimal.Parse(cleanTotalVatText);
        }
        public int GetSelectedReceivedInterimReceptionItem()
        {
            var received = WaitForElementIsVisible(By.XPath(SELECTED_RECEIVED_INTERIM_RECEPTION_ITEM));
            WaitForLoad(); 
            return int.Parse(received.Text);
        }
        public decimal GetSelectedTotalVatInterimReceptionItem()
        {
            _selectedTotalVatInterimReceptionItem = WaitForElementIsVisible(By.XPath(SELECTED_TOTAL_VAT_INTERIM_RECEPTION_ITEM));
            var totalVat = _selectedTotalVatInterimReceptionItem.Text;
            WaitLoading();
            var cleanTotalVatText = totalVat.Replace("€", "").Trim();

            return decimal.Parse(cleanTotalVatText);
        }

    }
}