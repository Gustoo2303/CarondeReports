using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using CarondeReports;

/*********************************APPLICATION CONSOLE POUR TESTER LE PROGRAMME***************************************/

namespace ConsoleApp1
{
    class Program
    {
        static void Main(string[] args)
        {
            string GetPatrolPointsTest;
            string GetPatrolAnosByIdTest;
            string GetPatrolAnosByDateTest;

            CarondeReports.CarondeReports cr = new CarondeReports.CarondeReports(@"C:\Users\desou\Desktop\dirCaronde", @"C:\Users\desou\Desktop\dirCaronde", "fr");

            GetPatrolPointsTest = cr.GetPatrolPoints(14, "Bourgogne");
            Console.WriteLine(GetPatrolPointsTest);

            GetPatrolAnosByIdTest = cr.GetPatrolAnosById(14, "Bourgogne");
            Console.WriteLine(GetPatrolAnosByIdTest);

            GetPatrolAnosByDateTest = cr.GetAnosByDate(new DateTime(2018, 10, 11), new DateTime(2018, 10, 12));
            Console.WriteLine(GetPatrolAnosByDateTest);

            Console.ReadKey();
        }
    }
}
