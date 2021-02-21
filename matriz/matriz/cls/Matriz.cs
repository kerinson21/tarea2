using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;


namespace matriz.cls
{
    class Matriz
    {
        public int rw { get; set; }
        public int cl { get; set; }
        public int x { get; set; }
        public int y { get; set; }
        public int[,] busqueda()
        {
            //ABRIR EXCEL
            Excel.Application xlApp;
            Excel.Workbook xlWorkbook;
            Excel.Worksheet xlWorkSheet;
            Excel.Range range;

            xlApp = new Excel.Application();
            xlWorkbook = xlApp.Workbooks.Open(@"C:\tmp\resources\matriz.xlsx");
            xlWorkSheet = (Excel.Worksheet)xlWorkbook.Worksheets.get_Item(1);//abre hoja1 del libro

            range = xlWorkSheet.UsedRange; //rango con datos usados, excluye null
            rw = range.Rows.Count; //lineas
            cl = range.Columns.Count; //columnas


            int[,] datos = new int[rw,cl];

            for (int i = 0; i < rw; i++)
            {
                for (int j = 0;  j < cl; j++)
                {
                    datos[i, j] = (int)(range.Cells[i+1, j+1] as Excel.Range).Value2;
                }
                
            }

            xlWorkbook.Close(true, null, null);
            xlApp.Quit();

            Marshal.ReleaseComObject(xlWorkSheet);//libera
            Marshal.ReleaseComObject(xlWorkbook);
            Marshal.ReleaseComObject(xlApp);

            return datos;
        }
        public void encontrarPosicion()
        {
            var matriz = busqueda();
            for (int i = 0; i < rw; i++)
            {
                for (int j = 0; j < cl; j++)
                {
                    if(matriz[i,j] == 0)
                    {
                        x = i;
                        y = j;
                        break;
                    }
                }
            }
        }
        public int[,] imprimirNuevaMatriz()
        {
            var matriz = busqueda();
            encontrarPosicion();
            for (int i = 0; i < rw; i++)
            {
                for (int j = 0; j < cl; j++)
                {
                    if (j == y || i == y)
                    {
                        matriz[i, j] = 0;
                    }
                   
                    /*Console.Write(matriz[i, j] + " ");*/
                }
                /*Console.WriteLine("");*/
            }
            /*Console.WriteLine(x + " " + y);*/
            return matriz;
        }

        public void rellenar()
        {
            //ABRIR EXCEL
            Excel.Application xlApp;
            Excel.Workbook xlWorkbook;
            Excel.Worksheet xlWorkSheet;
            Excel.Range range;

            xlApp = new Excel.Application();
            xlWorkbook = xlApp.Workbooks.Open(@"C:\tmp\resources\matriz.xlsx");
            xlWorkSheet = (Excel.Worksheet)xlWorkbook.Worksheets.get_Item(1);//abre hoja1 del libro

            xlApp.Visible = false;

            range = xlWorkSheet.UsedRange; //rango con datos usados, excluye null
            rw = range.Rows.Count; //lineas
            cl = range.Columns.Count; //columnas

            var datos = imprimirNuevaMatriz();

            for (int i = 0; i < rw; i++)
            {
                for (int j = 0; j < cl; j++)
                {
                    range.Cells[i + 1, j + 1] = datos[i, j];
                }
            }
            Console.WriteLine(@"El archivo se a creado de manera exitosa C:\tmp\resources\matriz.xlsx");
            
            xlWorkbook.Close(true, Type.Missing, Type.Missing);
            xlApp.Quit();

            Marshal.ReleaseComObject(xlWorkSheet);//libera
            Marshal.ReleaseComObject(xlWorkbook);
            Marshal.ReleaseComObject(xlApp);
        }
       
    }
}
