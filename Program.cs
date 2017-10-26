using System;
using NetOffice.ExcelApi;

namespace Excel_Class
{
    class Program
    {
        static void Main(string[] args)
        {
            CriarExcel();
            LerExcel();
        }

        static void CriarExcel(){
            Application ex = new Application();
            ex.Workbooks.Add();
            ex.Cells[1,1].Value = "Ford";
            ex.ActiveWorkbook.SaveAs(@"C:\Users\31049529812\Desktop\Back_A\Project_1\Excel_Class\lista_nvo.xls");
            ex.Quit();
        }

        static void LerExcel(){
            Application ex = new Application();
            ex.Workbooks.Open(@"C:\Users\31049529812\Desktop\Back_A\Project_1\Excel_Class\lista_nvo.xls");
            string valor = ex.Cells[1,1].Value.ToString();
            
            Console.WriteLine(valor);
            ex.Quit();
        }
    }
}
