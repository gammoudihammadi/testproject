using System.Collections.Generic;
using System.IO;

namespace Newrest.Winrest.AutomaticFunctionnalTestsReport
{
    public class AutomaticTestReportProduction
    {
        public static void Main(string[] args)
        {
            // Pré-requis : à renseigner dans un fichier de config si possible
            string CSVDirectory = "G:\\Drive partagés\\DH_CLIENTS\\Newrest\\Winrest\\TEST_MONITORING\\rapportAuto\\modeles CSV";
            int nbCSV = 10;
            
            // Recherche des fichier CSV à lire
            CSVFilesProcessing trtCSV = new CSVFilesProcessing();
            List<FileInfo> csvFiles = trtCSV.SearchForCSVFiles(CSVDirectory, nbCSV);

            // Récupération des données des fichiers CSV
            Dictionary<string, TestAuto> testsList = new Dictionary<string, TestAuto>();
            foreach (FileInfo csvFile in csvFiles)
            {
                trtCSV.ExtractInfosFromCSVFiles(csvFile.FullName, testsList);
            }
        }

        
    }
}
