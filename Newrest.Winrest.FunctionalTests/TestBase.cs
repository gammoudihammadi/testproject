using FluentAssertions;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects;
using Newrest.Winrest.FunctionalTests.PageObjects.CustomerPortal;
using Newtonsoft.Json;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.Extensions;
using S;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net.Http;

namespace Newrest.Winrest.FunctionalTests
{
    [TestClass]
    public abstract class TestBase
    {
        private StringWriter _writer;

        protected IWebDriver WebDriver
        {
            get => WebDriverFactory.Driver;
        }

        public TestContext TestContext { get; set; }

        [AssemblyInitialize()]
        public static void AssemblyInitialize(TestContext context)
        {
            var useLoginOnce = context.Properties["UseLoginOnce"].ToString();
            if (!string.IsNullOrEmpty(useLoginOnce) && useLoginOnce == "Y")
            {
                var downloadsPath = context.Properties["DownloadsPath"].ToString();
                WebDriverFactory.InitializeDriver(downloadsPath);
            }
        }

        [TestInitialize()]
        public virtual void TestInitialize()
        {
            var useLoginOnce = TestContext.Properties["UseLoginOnce"].ToString();
            if (string.IsNullOrEmpty(useLoginOnce) || useLoginOnce == "N")
            {
                var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
                WebDriverFactory.EnsureBrowserClosed();
                WebDriverFactory.InitializeDriver(downloadsPath);
            }

            _writer = new StringWriter();
            Console.SetOut(_writer);
        }

        public HomePage LogInAsAdmin() 
        {
            var options = WebDriver.Manage();
            if (options.Cookies?.AllCookies?.Count != 0)
                return new HomePage(WebDriver, TestContext);

            //Arrange

            var winrestUrl = TestContext.Properties["Winrest_URL"].ToString();
            var userName = TestContext.Properties["Admin_UserName"].ToString();
            var password = TestContext.Properties["Admin_Password"].ToString();

            //Act
            //WebDriver.Manage().Window.Maximize();
            WebDriver.Manage().Window.Size = new System.Drawing.Size(1366, 768);
            WebDriver.Navigate().GoToUrl(winrestUrl);
            var loginPage = new LoginPage(WebDriver, TestContext);
            var isLoggedIn = loginPage.Login(userName, password);
            ClearCache();
            //Assert
            isLoggedIn.Should().BeTrue();
            HomePage homePage = new HomePage(WebDriver, TestContext);
            homePage.CloseNewsPopup();
            return new HomePage(WebDriver, TestContext);
        }

        public CustomerPortalLoginPage LogInCustomerPortal(bool newTab = false)
        {
            if (newTab == true)
            {
                // Création d'un nouvel onglet pour ouvrir le customer portal
                var javaScriptExecutor = WebDriver as IJavaScriptExecutor;
                javaScriptExecutor.ExecuteScript("window.open('about:blank','_blank');");
                //Switch vers le nouvel onglet et navigation vers la page
                WebDriver.SwitchTo().Window(WebDriver.WindowHandles.Last());
            }

            var winrestUrl = TestContext.Properties["Winrest_URL"].ToString();

            WebDriver.Manage().Window.Size = new System.Drawing.Size(1366, 768);
            if (winrestUrl.EndsWith("/"))
            {
                WebDriver.Navigate().GoToUrl(winrestUrl + "CustomerPortal/");
            } else
            {
                WebDriver.Navigate().GoToUrl(winrestUrl + "/CustomerPortal/");
            }

            return new CustomerPortalLoginPage(WebDriver, TestContext);
        }

        public void ClearCache()
        {
            var urlCache = TestContext.Properties["Winrest_Cache"].ToString();

            // Aller jusqu'à la page clearcache
            WebDriver.Navigate().GoToUrl(urlCache);

            // Cliquer sur le lien de confirmation du ClearCache
            var clearCachePage = new ClearCachePage(WebDriver);
            clearCachePage.Clear();

            // Retour à la page précédente
            WebDriver.Navigate().Back();
            WebDriver.Navigate().Back();
        }

        //Supprime tous les fichiers dans le dossier "C:\ChromeDriverDownloads\"
        public void DeleteAllFileDownload()
        {
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();

            string[] filenames = Directory.GetFiles(downloadsPath, "*", SearchOption.TopDirectoryOnly);
            foreach (string fName in filenames)
            {
                try
                {
                    File.Delete(fName);
                }
                catch (IOException ioe)
                {
                    Console.WriteLine("cannot delete " + ioe.ToString());
                }
            }
        }

        public static string GenerateName(int len)
        {
            Random r = new Random();
            string[] consonants = { "b", "c", "d", "f", "g", "h", "j", "k", "l", "m", "l", "n", "p", "q", "r", "s", "sh", "zh", "t", "v", "w", "x" };
            string[] vowels = { "a", "e", "i", "o", "u", "ae", "y" };
            string name = "";
            name += consonants[r.Next(consonants.Length)].ToUpper();
            name += vowels[r.Next(vowels.Length)];
            int b = 2; //b tells how many times a new letter has been added. It's 2 right now because the first two letters are already in the name.
            while (b < len)
            {
                name += consonants[r.Next(consonants.Length)];
                b++;
                name += vowels[r.Next(vowels.Length)];
                b++;
            }

            return name;
        }


        [AssemblyCleanup()]
        public static void AssemblyCleanup()
        {
            WebDriverFactory.CloseDriver();
        }

        [TestCleanup]
        public virtual void TestCleanup()
            {
            if (TestContext.CurrentTestOutcome == UnitTestOutcome.Failed)
            {
                var screenshot = WebDriver.TakeScreenshot();
                screenshot.SaveAsFile($"{TestContext.TestName}.png");
                TestContext.AddResultFile($"{TestContext.TestName}.png");
            }

            try
            {
                //Adding Log file to the test results
                var logs = _writer.ToString();
                if (!String.IsNullOrWhiteSpace(logs))
                {
                    var logsFilePath = $"{TestContext.TestName}.txt";
                    File.AppendAllText(logsFilePath, logs);
                    TestContext.AddResultFile(logsFilePath);
                }
            }
            catch
            {
                TestContext.WriteLine("Error while writing logs file");
            }
            finally
            {
                _writer?.Dispose();
            }

            var useLoginOnce = TestContext.Properties["UseLoginOnce"].ToString();
            if (string.IsNullOrEmpty(useLoginOnce) || useLoginOnce == "N")
            {
                WebDriverFactory.CloseDriver();
            }
        }

        public double ArrangeTarif(string tarif, string decimalSeparator)
        {

            string tarifLinear = tarif.Replace("€", "").Replace(" ", "");
            double tarifDouble = Convert.ToDouble(tarifLinear, new NumberFormatInfo() { NumberDecimalSeparator = decimalSeparator });
            return Math.Round(tarifDouble, 3);
        }

        protected int ExecuteAndGetInt(string query, params KeyValuePair<string, object>[] parameters)
        {
            var sqlConnectionString = GetConnectionString();
            using (var sqlConnection = new SqlConnection(sqlConnectionString))
            using (SqlCommand sqlCommand = sqlConnection.CreateCommand())
            {
                foreach (var parameter in parameters)
                {
                    sqlCommand.Parameters.AddWithValue(parameter.Key, parameter.Value);
                }

                sqlCommand.CommandText = query;

                sqlConnection.Open();
                var r = sqlCommand.ExecuteScalar();
                return Convert.ToInt32(r);
            }
        }

        protected string ExecuteAndGetString(string query, params KeyValuePair<string, object>[] parameters)
        {
            var sqlConnectionString = GetConnectionString();
            using (var sqlConnection = new SqlConnection(sqlConnectionString))
            using (SqlCommand sqlCommand = sqlConnection.CreateCommand())
            {
                foreach (var parameter in parameters)
                {
                    sqlCommand.Parameters.AddWithValue(parameter.Key, parameter.Value);
                }

                sqlCommand.CommandText = query;

                sqlConnection.Open();
                var r = sqlCommand.ExecuteReader();
                if (r.Read())
                    return r[0].ToString();

                return null;
            }
        }

        protected void ExecuteNonQuery(string query, params KeyValuePair<string, object>[] parameters)
        {
            var sqlConnectionString = GetConnectionString();
            using (var sqlConnection = new SqlConnection(sqlConnectionString))
            using (SqlCommand sqlCommand = sqlConnection.CreateCommand())
            {
                foreach (var parameter in parameters)
                {
                    sqlCommand.Parameters.AddWithValue(parameter.Key, parameter.Value);
                }

                sqlCommand.CommandText = query;

                sqlConnection.Open();
                var r = sqlCommand.ExecuteNonQuery();
            }
        }

        private string GetConnectionString()
        {
            var client = new HttpClient();
            var urlAPI = $"{TestContext.Properties["Winrest_URL"]}/api/system";
            var response = client.GetStringAsync(urlAPI).Result;
            var processInfo = JsonConvert.DeserializeObject<ConnectionStringInfo>(response);
            return BuildConnectionString(processInfo);
        }
        private string BuildConnectionString(ConnectionStringInfo info)
        {
            string[] parts = info.Database.Split(new[] { " - " }, StringSplitOptions.None);
            var server = parts[0];
            var database = parts[1];
            return $"Server={server}.database.windows.net;" +
                $"Initial Catalog={database};Persist Security Info=False;" +
                $"User ID=sa_winrest;" +
                $"Password=N3wrest31!;" +
                $"MultipleActiveResultSets=True;" +
                $"Encrypt=True;" +
                $"TrustServerCertificate=False;" +
                $"Connection Timeout=30;" +
                $"Max Pool Size=200;";
        }
    }
    public class ConnectionStringInfo
    {
        public string Database { get; set; }
    }
}
