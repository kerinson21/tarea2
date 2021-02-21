using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace matriz.cls
{
    class Estructura
    {

        public string nombre { get; set; }
        public string direccion { get; set; }

    

        public List<Estructura> cargaDatosXLS()
        {
            //ABRIR EXCEL
            Excel.Application xlApp;
            Excel.Workbook xlWorkbook;
            Excel.Worksheet xlWorkSheet;
            Excel.Range range;

            int rCnt;
            int rw = 0;
            int cl = 0;

            xlApp = new Excel.Application();
            xlWorkbook = xlApp.Workbooks.Open(@"C:\tmp\resources\data.xlsx");
            xlWorkSheet = (Excel.Worksheet)xlWorkbook.Worksheets.get_Item(1);//abre hoja1 del libro

            range = xlWorkSheet.UsedRange; //rango con datos usados, excluye null
            rw = range.Rows.Count; //lineas
            cl = range.Columns.Count; //columnas

            List<Estructura> todos = new List<Estructura>();
            Estructura individual = new Estructura();

            for (rCnt = 1; rCnt <= rw; rCnt++)
            {
                individual.nombre = (string)(range.Cells[rCnt, 1] as Excel.Range).Value2;
                individual.direccion = (string)(range.Cells[rCnt, 2] as Excel.Range).Value2;

                todos.Add(individual);
                individual = new Estructura();

            }






            xlWorkbook.Close(true, null, null);
            xlApp.Quit();

            Marshal.ReleaseComObject(xlWorkSheet);//libera
            Marshal.ReleaseComObject(xlWorkbook);
            Marshal.ReleaseComObject(xlApp);

            return todos;
        }
    }
}
