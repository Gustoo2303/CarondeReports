using C1.Win.FlexReport;
using C1.Win.C1Document.Export;
using System;
using System.IO;
using System.Collections.Generic;

namespace CarondeReports
{
    /// <summary>
    /// Génère des rapports sous la forme de pdf à partir des paramètres pris en compte:  
    /// </summary>
    public class CarondeReports
    {
        /*********************************DECLARATION DES CONSTANTES ET DES VARIABLES***************************************/

        const string DATABASEFILE = @"\db\rondes.mdb"; //base de donnée Access
        const string ARCHIVESDIR = @"\archives\"; //dossier contenant les archives
        const string RETREPORT = @"\ret.flxr"; //modèle pour les rapports de rondes
        const string ANOREPORT = @"\ano.flxr"; //modèle pour les anomalies

        public string appdir; //chemin de l'application
        public string datadir; //chemin de la base de donnée dans l'application
        public string languageDir; //chemin du langage choisie dans l'application

        public string pathRet; //chemin couplant appdir + le dossier selon la langue choisie + le modèle des rondes
        public string pathAnoReport; //chemin couplant appdir + le dossier selon la langue choisie + le modèle anomalie
        public string pathDataBase; //chemin couplant datadir et la base de donnée


        /*********************************CONSTRUCTEUR***************************************/

        /// <summary>
        /// Constructeur prenant en paramètre les différents chemins de l'appli
        /// </summary>
        /// <param name="appdir"></param>
        /// <param name="datadir"></param>
        /// <param name="languageDir"></param>
        public CarondeReports(string appdir, string datadir, string languageDir)
        {
            this.appdir = appdir;
            this.datadir = datadir;
            this.languageDir = languageDir;

            pathDataBase = datadir + DATABASEFILE;

    
            pathRet = appdir + ChooseDir(languageDir) + RETREPORT;
            pathAnoReport = appdir + ChooseDir(languageDir) + ANOREPORT;
            
            
            //Vérification qu'aucun des paramètres n'est vide, sinon renvoie une exception
            string[] DirArgs = new string[] { appdir, datadir, languageDir }; 

            foreach (string arg in DirArgs)
            {
                if (appdir == "" || datadir == "" || languageDir == "")
                {
                    throw new ArgumentNullException(arg);
                }
            }
        }

        /*********************************METHODES***************************************/

        /// <summary>
        /// Retourne le chemin selon la langue choisie
        /// </summary>
        /// <param name="languageChoosen">""(=français par défaut), "fr", "us", "es"</param>
        /// <returns></returns>
        public string ChooseDir(string languageChoosen)
        {
            string languagePath = "";

            if (languageChoosen == "" || languageChoosen == "fr")
            {
                languagePath = @"\rpts";
            }
            else if (languageChoosen == "us")
            {
                languagePath = @"\rptsus";
            }
            else if (languagePath == "es")
            {
                languagePath = @"\rptses";
            }
            else
            {
                languagePath = @"\rpts" + languageChoosen;
            }

            if (!Directory.Exists(appdir + languagePath))
            {
                throw new Exception($"{languagePath} does not exist !");
            }

            return languagePath;
        }

        /// <summary>
        ///  Vérifie si tous les fichiers et dossiers dans un tableau existent sinon retourne une exception
        /// </summary>
        public void checkFilesExists(List<string> listToCheck)
        {
            foreach (string el in listToCheck)
            {
                if (!File.Exists(el))
                {
                    throw new FileNotFoundException($"{el} does not exist");
                }
            }
        }

        /// <summary>
        /// retourne la date actuelle
        /// </summary>
        /// <returns></returns>
        public static string GetDate()
        {
            //Ecrire la fonction sur une seule ligne
            DateTime localDate = DateTime.Now;
            return localDate.ToString("yyyymmddHHmmss");
        }

        /// <summary>
        /// Charge et exporte un rapport selon une requête et un contexte (anomalies, rondes) en s'appuyant sur les classes Connection.Connection_query et C1FlexReport
        /// </summary>
        /// <param name="nameModelReport"></param>
        /// <param name="pathReports"></param>
        /// <param name="sqlRequest"></param>
        /// <param name="pdfPathExport"></param>
        public void loadAndExportReport(string nameModelReport, string pathReports, string sqlRequest, string pdfPathExport)
        {
            //nouvel objet de classe Connection.Connection_query
            Connection.Connection_query objCon = new Connection.Connection_query();

            //On attribue une valeur à la variable ConnectionString grâce aux accesseurs initialisés dans la classe Connection
            objCon.ConnectionString = pathDataBase;
            objCon.Sql = sqlRequest;

            //Ouvre une nouvelle connexion
            try
            {
                objCon.OpenConnection();
            }
            catch (Exception e)
            {
                throw new Exception("objCon error :" +e.ToString());
            }

            //Créer un nouveau rapport qu'on nomme rep
            C1FlexReport rep = new C1FlexReport();

            //Charge le rapport avec en paramètre son chemin et son nom
            try
            {
                rep.Load(pathReports, nameModelReport);
            }
            catch (Exception e)
            {
                throw new Exception("Chargement FlexReport error : " + e.ToString());
            }

            //On injècte le DataSet qu'on a obtenu de la méthode MyDataSet dans le Recordset du DataSource du rapport
            try
            {
                rep.DataSource.Recordset = objCon.MyDataSet();

            }
            catch (Exception e)
            {
                throw new Exception("Dataset error : " + e.ToString());
            };

            //Création d'un objet PdfFilter 
            PdfFilter filter = new PdfFilter();

            //On donne un nom au fichier et on spécifie le chemin où il sera exporté
            filter.FileName = pdfPathExport;

            //On exporte le rapport en pdf
            try
            {
                rep.Export(filter);
            }
            catch (Exception e)
            {
                throw new Exception("Génération Flexreport error : " + e.ToString());
            }

            //On ferme la connexion
            objCon.CloseConnection();
        }

        /*********************************METHODES POUR LA GENERATION DES ANOMALIES PAR ID***************************************/

        /// <summary>
        /// Appelle la méthode pdfAnosPatrolById(idRound, pathAnoReport, pathDataBase, pdfPathExport) et retourne le nom du fichier créé si la génération du pdf a réussi
        /// </summary>
        /// <param name="idRound"></param>
        /// <param name="roundName"></param>
        /// <returns></returns>
        public string GetPatrolAnosById(int idRound, string roundName)
        {
            string pdfPathExport = appdir + ARCHIVESDIR + "anomalieN°" + GetDate() + "ID" +idRound + "ronde" +roundName + ".pdf";

            pdfAnosPatrolById(idRound, pathAnoReport, pdfPathExport);

            return pdfPathExport;
        }


        /// <summary>
        /// Génère sous la forme d'un fichier pdf les anomalies constatées selon l'id de la ronde fourni, s'appuie sur la méthode GetPatrolAnosById(int idRound, string roundName)
        /// </summary>
        /// <param name="idRound"></param>
        /// <param name="pathModelReport"></param>
        /// <param name="pathDataBase"></param>
        /// <param name="pdfPathExport"></param>
        private void pdfAnosPatrolById(int idRound, string pathAnoReport, string pdfPathExport)
        {
            List<string> array = new List<string> { pathAnoReport, pathDataBase };

            checkFilesExists(array);
            
            //nom du rapport (généré depuis l'application flexreport designer)
            string nameModelReport = "Liste des anomalies par id";

            //Requête SQL
            string sqlRequest = "SELECT * FROM `ANOS` INNER JOIN `ARCHRONDES` ON `ANOS`.`IdRonde` = `ARCHRONDES`.`IdArchRonde` where `ARCHRONDES`.`idArchRonde` =" + idRound + " ORDER BY DateAno desc, IDPOINT ASC;";

            loadAndExportReport(nameModelReport, pathAnoReport, sqlRequest, pdfPathExport);
        }





        /*********************************METHODES POUR LA GENERATION DES ANOMALIES PAR DATE***************************************/

        /// <summary>
        /// Appelle la méthode getPdfAnos(fromDate, toDate, pathAnoReport, pathDataBase, pdfPathExport) et retourne le nom du fichier créé Si la génération du pdf a réussi
        /// le nom du fichier pdf créé
        /// </summary>
        /// <param name="fromDate">date de début de la période</param>
        /// <param name="toDate">date de fin</param>
        /// <returns></returns>
        public string GetAnosByDate(DateTime fromDate, DateTime toDate)
        {

            string pdfPathExport = appdir + ARCHIVESDIR + "anomalie(s)" + GetDate()  + ".pdf";

            getPdfAnosByDate(fromDate, toDate, pathAnoReport, pathDataBase, pdfPathExport);

            return pdfPathExport;
        }


    /// <summary>
    /// Génère sous la forme d'un fichier pdf les anomalies constatées au cours d'une période donnée, s'appuie sur la méthode GetAnosByDate(string fromDate, string toDate)
    /// </summary>
    /// <param name="fromDate"></param>
    /// <param name="toDate"></param>
    /// <param name="pathModelReport"></param>
    /// <param name="pathDataBase"></param>
    /// <param name="pdfPathExport"></param>
    public void getPdfAnosByDate(DateTime fromDate, DateTime toDate, string pathAnoReport, string pathDataBase, string pdfPathExport) 
        {
            var fromDateString = ConvertDateTimeToString(fromDate); //objet fromDate converti en string fromDateString
            var toDateString = ConvertDateTimeToString(toDate); //objet toDate converti en string toDateString

            List<string> array = new List<string> { pathAnoReport, pathDataBase };

            checkFilesExists(array);

            //nom du rapport (généré depuis l'application flexreport designer)
            string nameModelReport = "Liste des anomalies par date";

            //Requête SQL
            string sqlRequest = "SELECT * FROM `ANOS` WHERE `ANOS`.`DateAno` >= #" + fromDateString + "# AND `ANOS`.`DateAno` < #" + toDateString + " #;";

            loadAndExportReport(nameModelReport, pathAnoReport, sqlRequest, pdfPathExport);
            
        }

        public string ConvertDateTimeToString(DateTime date)
        {
            string date_str = date.ToString("MM/dd/yyyy");
            return date_str;
        }


        /*********************************METHODES POUR LA GENERATION DES RONDES PAR ID***************************************/

        /// <summary>
        /// Méthode qui appelle la méthode pdfPatrol(idRound, pathRet, pathDataBase, pdfPathExport) et retourne le nom du fichier créé Si la génération du pdf a réussi
        /// </summary>
        /// <param name="idRound">Id de la ronde</param>
        /// <param name="roundName">Nom de la ronde</param>
        public string GetPatrolPoints(int idRound, string roundName)
        {
            string pdfPathExport = appdir + ARCHIVESDIR + GetDate() + roundName + ".pdf"; //chemin complet du fichier pdf exporté --> dans l'appli

            try
            {
                pdfPatrol(idRound, pathRet, pathDataBase, pdfPathExport);
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
                throw;
            }
            return pdfPathExport;
        }


        /// <summary>
        /// Génère sous la forme d'un fichier pdf les points d'une ronde sélectionnée, s'appuie sur la méthode pdfPatrol(int idRound, string pathRet, string pathDataBase, string pdfPathExport)
        /// </summary>
        /// <param name="idRound"></param>
        /// <param name="pathModelReport"></param>
        /// <param name="pathDataBase"></param>
        /// <param name="pdfPathExport"></param>
        private void pdfPatrol(int idRound, string pathRet, string pathDataBase, string pdfPathExport)
        {
            List<string> array = new List<string> { pathRet, pathDataBase };

            checkFilesExists(array);

            //nom du rapport (généré depuis l'application flexreport designer)
            string nameModelReport = "Compte-rendu de tournée";

            //Requête SQL
            string sqlRequest = "SELECT * FROM ARCHPOINTS INNER JOIN ARCHRONDES ON ARCHPOINTS.idronde = ARCHRONDES.idarchronde where idarchronde = " + idRound + " ORDER BY DATEPOINT desc, IDPOINT ASC;";

            loadAndExportReport(nameModelReport, pathRet, sqlRequest, pdfPathExport);
        }
    }
}
