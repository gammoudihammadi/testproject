using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects;
using Newrest.Winrest.FunctionalTests.PageObjects.Jobs.Jobs;

namespace Newrest.Winrest.FunctionalTests.Jobs.Jobs
{
    [TestClass]
    public class JobsTest : TestBase
    {
        private const int _timeout = 600000;

        [Timeout(_timeout)]
        [TestMethod]
        public void JB_JB_Access()
        {
            HomePage homePage = LogInAsAdmin();
            JobsPage jobsPage = homePage.GoToJobs_Jobs();
            bool access = jobsPage.AccessPage();
            Assert.IsTrue(access, "Jobs n'est pas accessible");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void JB_JB_StateDisplay()
        {
            HomePage homePage = LogInAsAdmin();
            JobsPage jobsPage = homePage.GoToJobs_Jobs();
            CegidJobPage cegidJobPage = jobsPage.GoToCegidJobPage();
            //cegidJobPage.ResetFilter();
            //cegidJobPage.fil
            bool hasJob = cegidJobPage.HasJob();
            Assert.IsTrue(hasJob, "Aucun donnée disponible");
            CegidJobDetailPage cegidJobDetailPage = cegidJobPage.ClickFirstElement();
            bool isImportFileJobVisible = cegidJobDetailPage.IsImportFileJobVisible();
            Assert.IsTrue(isImportFileJobVisible, "le Statut IMPORT FILE JOB n'est pas afficher dans le cadre");
            bool isImportFileJobColored = cegidJobDetailPage.IsImportFileJobColored();
            Assert.IsTrue(isImportFileJobColored, "IMPORT FILE JOB n'est pas en couleur");
        }
    }
}
