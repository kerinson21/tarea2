using matriz.cls;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace matriz
{
    class Program
    {
        static void Main(string[] args)
        {
            var matriz = new Matriz();
            //matriz.imprimirNuevaMatriz();
            matriz.rellenar();
            Console.ReadKey();
        }
    }
}
