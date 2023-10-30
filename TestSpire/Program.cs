using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenlabReport;
using NLog;

namespace TestSpire
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Logger log = LogManager.GetCurrentClassLogger();
            const string source = "d:\\data\\aemed\\assayPTerminado.xlsx";
            const string destination = "d:\\data\\aemed\\destination.xlsx";
            bool result;
            string salida;
            OpenlabToExcel testing = new OpenlabToExcel();
            result=testing.CreateWorkbook(source, destination);
            log.Info("Logging from testSpire");
            for (int i = 1; i < 30; i++)
            {
                salida= testing.NumberToExcell(destination, "data", "A"+i, i);
                Console.WriteLine("Datos: "+salida);
 //               salida= testing.NumberToExcell(destination, "Datos2", "A" + i, i);
 //               Console.WriteLine("Datos2 "+salida);
            }
            //           salida = testing.NumberToExcell(destination, "A6", 25);
            Console.WriteLine("press key to exit...");
            Console.ReadKey();
        }
    }
}
