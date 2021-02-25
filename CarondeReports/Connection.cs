using System.Data.OleDb;
using System.Data.Odbc;
using System.Data;

/*********************************CLASSE CONNECTION QUI RETOURNE UN DATASET POUR OLEDB ET ODBC***************************************/

/// <summary>
/// Retourne les données(dataset) associées à une table de donnée créée 
/// </summary>
namespace CarondeReports
{
    class Connection
    {
        public class Connection_query
        {
            //On protège l'accès à la valeur _connectionString
            private string _connectionString;
            //get et set permettent d'initialiser une valeur et de la retourner dans la classe Main
            public string ConnectionString
            {

                get
                {
                    return _connectionString;
                }

                set
                {
                    if (value.IndexOf("Provider=") > -1)
                        _connectionString = value;
                    else if (value.IndexOf("Driver=") > -1)
                        _connectionString = value;
                    else
                        _connectionString = "Driver={Microsoft Access Driver (*.mdb)};DBQ=" + value; //chemin du fichier seulement
                }

            }

            private string _sql;
            public string Sql
            {
                get
                {
                    return _sql;
                }

                set
                {
                    _sql = value;
                }
            }

            //On instancie un nouvel objet de la classe OleDbConnection
            //OleDbConnection con;
            OdbcConnection con;

            public void OpenConnection()
            {
                //Afin d'ouvrir la connection on lui passe en paramètre la valeur de la _connectionString puis on ouvre la connexion
                con = new OdbcConnection(_connectionString);
                con.Open();

                
            }

            //On ferme la connexion
            public void CloseConnection()
            {
                con.Close();
            }

            public object MyDataSet()
            {
                //On créer une table de donnée qui servira de source de donnée
                DataTable results = new DataTable();

                //On lance une nouvelle connexion qui prend en paramètre une requête sql et l'objet OleDbConnection qu'on a instancié plus haut sous le nom de con afin de pouvoir envoyer la requête à la base de donnée demandée
                OdbcCommand cmd = new OdbcCommand(_sql, con);

                //On instancie un nouvel objet adapter qui sert à lier la source de donnée à un dataset (ici pour rappel, l'objet cmd contient la requête à exécuter sur la base de donnée)
                OdbcDataAdapter adapter = new OdbcDataAdapter(cmd);

                //on envoie ensuie les données récoltées dans le dataset
                adapter.Fill(results);

                //et on retourne le dataset
                return results;
            }
        }
    }
}
