using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.CustomerPortal;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium.Support.UI;
using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Threading;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Shared
{
    public class MailPage : PageBase
    {

        public MailPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        // _________________________________________ Constantes _______________________________________________

        // Connexion
        private const string EMAIL_ADRESS = "//input[@name='loginfmt']";
        private const string EMAIL_NEXT_BTN = "idSIButton9";
        private const string PASSWORD = "passwordInput";
        private const string EMAIL_SUBMIT_BUTTON = "submitButton";
        private const string EMAIL_IS_SIGNED_IN = "//div[contains(text(), 'Stay signed in?')]";
        private const string DONT_SHOW_CHECKBOX = "KmsiCheckboxField";
        private const string YES_BTN = "idSIButton9";

        private const string FIRST_MAIL = "//*[@id=\"MailList\"]/div/div/div/div/div/div/div/div[2]";

        private const string IS_LOGGED_IN = "//*[@id=\"meInitialsButton\"]";

        private const string MAIL_PASSWORD_LINK = "//a[text()='{0}']";

        //private const string OUTLOOK_ATTACHMENT_TITLE = "//*[@id=\"ConversationReadingPaneContainer\"]/div[2]/div/div/div[1]/div/div/div/div/div[2]/div/div/div/div/div/div/div[1]/div/div/div[1]";

        private const string DELETE_CURRENT_MAIL_BUTTON = "//i[@data-icon-name='DeleteRegular']";
        private const string DELETE_CURRENT_MAIL_BUTTON_BIS = "//*/button[@aria-label='Supprimer' and contains(@class,'splitPrimaryButton')]";

        // plusieurs pièces attachée
        private const string OUTLOOK_ATTACHMENT_ZIP = "//span[contains(text(), 'Télécharger tout')]";
        // plusieurs pièces attachée (le premier)
        private const string OUTLOOK_ATTACHMENT_PDF = "//div[2]/div[1]/div/div/div/div/div[2]/div/div[2]/div/div/div/div/div[2]/button";
        // une seule pièce attachée donc pas de ZIP
        private const string OUTLOOK_ATTACHMENT_PDF_SINGLE = "//*[@id=\"focused\"]/div[2]/div/div/div/div/div/div/div[1]/div/div/div[1]";
        private const string OUTLOOK_DOWNLOAD_BTN = "//span[text()='Télécharger']";
                                           
        private const string CLOSE_BTN = "//i[@data-icon-name='Cancel']/..";
        private const string OUTLOOK_BODY = "//i[@data-icon-name='Cancel']/..";
        private const string MAIL_BODY = "UniqueMessageBody_2";
        // _________________________________________ Variables _________________________________________________

        // Connexion
        [FindsBy(How = How.XPath, Using = EMAIL_ADRESS)]
        private IWebElement _emailAdress;

        [FindsBy(How = How.Id, Using = EMAIL_NEXT_BTN)]
        private IWebElement _emailNextBtn;

        [FindsBy(How = How.Id, Using = PASSWORD)]
        private IWebElement _password;

        [FindsBy(How = How.Id, Using = EMAIL_SUBMIT_BUTTON)]
        private IWebElement _emailSubmitBtn;

        [FindsBy(How = How.XPath, Using = EMAIL_IS_SIGNED_IN)]
        private IWebElement _emailIsSignedIn;
        
        [FindsBy(How = How.Id, Using = DONT_SHOW_CHECKBOX)]
        private IWebElement _dontShowCheckbox;

        [FindsBy(How = How.Id, Using = YES_BTN)]
        private IWebElement _yesButton;

        [FindsBy(How = How.XPath, Using = FIRST_MAIL)]
        private IWebElement _firstMail;
        
        [FindsBy(How = How.XPath, Using = IS_LOGGED_IN)]
        private IWebElement _isLoggedIn;

        [FindsBy(How = How.XPath, Using = OUTLOOK_ATTACHMENT_PDF)]
        private IWebElement _attachmentTitle;

        [FindsBy(How = How.XPath, Using = DELETE_CURRENT_MAIL_BUTTON)]
        private IWebElement _deleteMailButton;

        [FindsBy(How = How.XPath, Using = DELETE_CURRENT_MAIL_BUTTON_BIS)]
        private IWebElement _deleteMailButtonBis;

        [FindsBy(How = How.XPath, Using = OUTLOOK_ATTACHMENT_ZIP)]
        private IWebElement _attachmentZIP;

        [FindsBy(How = How.XPath, Using = OUTLOOK_DOWNLOAD_BTN)]
        private IWebElement _downloadButton;

        [FindsBy(How = How.XPath, Using = CLOSE_BTN)]
        private IWebElement _closeButton;
        [FindsBy(How = How.XPath, Using = OUTLOOK_BODY)]
        private IWebElement _outlookbody;

        // _________________________________________ Méthodes __________________________________________________
        public void FillFields_LogToOutlookMailbox(string mail, string pwd = null)
        {
            int indexer = 0;
            while (!CheckIfConnectedToOutlook(mail) && indexer < 6){
                WaitLoading();
                indexer++;
                if (CheckIfConnectedToOutlook(mail))
                {
                    break;
                }
            }
            if (!CheckIfConnectedToOutlook(mail))
            {
                _emailAdress = WaitForElementIsVisible(By.XPath(EMAIL_ADRESS));
                _emailAdress.SetValue(ControlType.TextBox, mail);

                _emailNextBtn = WaitForElementIsVisible(By.Id(EMAIL_NEXT_BTN));
                _emailNextBtn.Click();
                WaitPageLoading();

                _password = WaitForElementIsVisible(By.Id(PASSWORD));
                _password.SetValue(ControlType.TextBox, pwd);

                _emailSubmitBtn = WaitForElementIsVisible(By.Id(EMAIL_SUBMIT_BUTTON));
                _emailSubmitBtn.Click();
                WaitPageLoading();

                //stay signed
                if (isElementVisible(By.XPath(EMAIL_IS_SIGNED_IN)))
                {
                    _dontShowCheckbox = WaitForElementIsVisible(By.Id(DONT_SHOW_CHECKBOX));
                    _dontShowCheckbox.SetValue(ControlType.CheckBox, true);

                    _yesButton = WaitForElementIsVisible(By.Id(YES_BTN));
                    _yesButton.Click();
                    WaitPageLoading();
                }
            }
            Assert.IsTrue(CheckIfConnectedToOutlook(mail), "L'utilisateur n'est pas connecté à la boite mail");
        }

        public bool CheckIfSpecifiedOutlookMailExist(string mailSubject)
        {
            Actions action = new Actions(_webDriver);
            bool valueBool = false;

            bool staleElement = true;
            int compteur = 0;

            while (staleElement && compteur < 5)
            {
                try
                {
                    _firstMail = WaitForElementToBeClickable(By.XPath(FIRST_MAIL));
                    if (_firstMail.Text.Contains(mailSubject))
                    {
                        action.MoveToElement(_firstMail).Perform();
                        _firstMail.Click();

                        WaitPageLoading();
                        staleElement = false;
                        valueBool = true;
                    }
                    else
                    {
                        WaitPageLoading();
                        throw new Exception("Aucun mail d'objet " + mailSubject + " n'a été reçu.");
                    }
                }
                catch
                {
                    staleElement = true;
                    compteur++;
                    _webDriver.Navigate().Refresh();
                    WaitPageLoading();
                }
            }

            return valueBool;
        }

        public void ClickOnSpecifiedOutlookMail(string mailSubject)
        {
            _firstMail = WaitForElementToBeClickable(By.XPath(FIRST_MAIL));
            if (_firstMail.Text.Contains(mailSubject))
            {
                Actions act = new Actions(_webDriver);
                act.MoveToElement(_firstMail).Perform();
                _firstMail.Click();
                WaitPageLoading();
                MailWaitForLoad();
            }

            void MailWaitForLoad()
            {
                var javaScriptExecutor = _webDriver as IJavaScriptExecutor;
                var wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(61));

                Func<IWebDriver, bool> readyCondition = webDriver =>
                    (bool)javaScriptExecutor.ExecuteScript("return (document.readyState == 'complete')");

                wait.Until(readyCondition);
            }
        }

        public bool CheckIfConnectedToOutlook(string emailAdress)
        {
           
            if (isElementVisible(By.XPath(IS_LOGGED_IN)))
            {
                _isLoggedIn = WaitForElementIsVisible(By.XPath(IS_LOGGED_IN));
                if (_isLoggedIn.Text != null && _isLoggedIn.Text == "WA")
                {
                    return true;
                }
            }
            return false;
        }

        public CustomerPortalUpdatePage ClickOnLinkForPassword(string mailLink)
        {
            var link = WaitForElementIsVisible(By.XPath(String.Format(MAIL_PASSWORD_LINK, mailLink)));
            link.Click();

            var wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(60));
            wait.Until((driver) => driver.WindowHandles.Count > 3);
            _webDriver.SwitchTo().Window(_webDriver.WindowHandles.Last());

            return new CustomerPortalUpdatePage(_webDriver, _testContext);
        }

        public bool IsOutlookPieceJointeOK(string pieceJointe)
        {
            while (!isElementVisible(By.XPath(OUTLOOK_ATTACHMENT_PDF_SINGLE)))
            {
                WaitPageLoading();

            }
            if (isElementVisible(By.XPath(OUTLOOK_ATTACHMENT_PDF_SINGLE)))
            {
                _attachmentTitle = WaitForElementIsVisible(By.XPath(OUTLOOK_ATTACHMENT_PDF_SINGLE));

                if (_attachmentTitle.Text.Equals(pieceJointe))
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            else
            {
                return false;
            }

        }

        public void DeleteCurrentOutlookMail()
        {
            Actions act = new Actions(_webDriver);

            _deleteMailButton = WaitForElementIsVisible(By.XPath(DELETE_CURRENT_MAIL_BUTTON));
            act.MoveToElement(_deleteMailButton).Perform();

            _deleteMailButtonBis = WaitForElementIsVisible(By.XPath(DELETE_CURRENT_MAIL_BUTTON_BIS));
            _deleteMailButtonBis.Click();
            WaitPageLoading();
        }

        public FileInfo DownloadOutlookAttachmentZIP(string downloadsPath, string attachment, bool unzip, string fileName)
        {
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo trouve = null;
            if (unzip)
            {
                _attachmentZIP = WaitForElementIsVisible(By.XPath(OUTLOOK_ATTACHMENT_ZIP));
                _attachmentZIP.Click();
                WaitPageLoading();
                WaitForDownload();
                //Close();

                taskDirectory.Refresh();
                foreach (FileInfo taskFile in taskDirectory.GetFiles())
                {
                    if (taskFile.Name.StartsWith(fileName) && taskFile.Name.EndsWith(".zip"))
                    {
                        trouve = taskFile;
                        break;
                    }
                }
                if (trouve != null)
                {
                    ZipFile.ExtractToDirectory(trouve.FullName, downloadsPath);
                    trouve.Delete();
                }
            }
            taskDirectory.Refresh();
            trouve = null;
            foreach (FileInfo taskFile in taskDirectory.GetFiles())
            {
                if (taskFile.Name.Equals(attachment))
                {
                    trouve = taskFile;
                    break;
                }
            }
            Assert.IsNotNull(trouve, "Fichier non téléchargé");
            return trouve;
        }

        public FileInfo DownloadOutlookAttachmentPDF(string downloadsPath, string attachment)
        {
            FileInfo fileDonwload = null;
            _attachmentTitle = WaitForElementIsVisible(By.XPath(OUTLOOK_ATTACHMENT_PDF_SINGLE));
            _attachmentTitle.Click();

            _downloadButton = WaitForElementIsVisible(By.XPath(OUTLOOK_DOWNLOAD_BTN));
            _downloadButton.Click();

            fileDonwload = new FileInfo(Path.Combine(downloadsPath, attachment));
            int counter = 0;

            while (!fileDonwload.Exists && counter < 10)
            {
                Thread.Sleep(2000);
                fileDonwload.Refresh();
                counter++;
            }

            _closeButton = WaitForElementIsVisible(By.XPath(CLOSE_BTN));
            _closeButton.Click();

            return fileDonwload;
        }
        public bool CheckIfEmailBodyContainsText(string bodyText)
        {
            Actions action = new Actions(_webDriver);
            bool valueBool = false;

            bool staleElement = true;
            int compteur = 0;

            while (staleElement && compteur < 5)
            {
                try
                {
                    // Wait for the first mail element to be clickable (assuming it is the email you want to check)
                    _firstMail = WaitForElementToBeClickable(By.XPath(FIRST_MAIL));
                    action.MoveToElement(_firstMail).Perform();
                    _firstMail.Click();
                    IWebElement emailBody = WaitForElementIsVisible(By.Id(MAIL_BODY));
                    if (emailBody.Text.Contains(bodyText))
                    {
                        valueBool = true;
                    }
                    else
                    {
                        throw new Exception("The email body does not contain the specified text: " + bodyText);
                    }
                    staleElement = false;
                }
                catch
                {
                    staleElement = true;
                    compteur++;
                    _webDriver.Navigate().Refresh();
                    WaitPageLoading();
                }
            }
            return valueBool;
        }
        public void FillFields_LogToOutlookMailbox_MoreThanOneMonth(string mail, string pwd = null)
        {
            int index = 10;
            while(!CheckIfConnectedToOutlook(mail) && index>0)
            {
                WaitPageLoading();
                index--;
                if (CheckIfConnectedToOutlook(mail))
                {
                    break;
                }
            }
            if (!CheckIfConnectedToOutlook(mail))
            {
                _emailAdress = WaitForElementIsVisible(By.XPath(EMAIL_ADRESS));
                _emailAdress.SetValue(ControlType.TextBox, mail);

                _emailNextBtn = WaitForElementIsVisible(By.Id(EMAIL_NEXT_BTN));
                _emailNextBtn.Click();
                WaitPageLoading();

                _password = WaitForElementIsVisible(By.Id(PASSWORD));
                _password.SetValue(ControlType.TextBox, pwd);

                _emailSubmitBtn = WaitForElementIsVisible(By.Id(EMAIL_SUBMIT_BUTTON));
                _emailSubmitBtn.Click();
                WaitPageLoading();

                //stay signed
                if (isElementVisible(By.XPath(EMAIL_IS_SIGNED_IN)))
                {
                    _dontShowCheckbox = WaitForElementIsVisible(By.Id(DONT_SHOW_CHECKBOX));
                    _dontShowCheckbox.SetValue(ControlType.CheckBox, true);

                    _yesButton = WaitForElementIsVisible(By.Id(YES_BTN));
                    _yesButton.Click();
                    WaitPageLoading();
                }
            }
            Assert.IsTrue(CheckIfConnectedToOutlook(mail), "L'utilisateur n'est pas connecté à la boite mail");
        }


    }
}

