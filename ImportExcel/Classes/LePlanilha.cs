using Excel = Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Web;
using ImportExcel.Models;

namespace ImportExcel.Classes
{
    public class LePlanilha
    {
        static public List<Planilha> ObtemItems(HttpPostedFileBase fileItem, out string retorno)
        {
            List<Planilha> planilhaCtrl = new List<Planilha>();
            retorno = "ok";
            try
            {
                if (fileItem != null && fileItem.ContentLength > 0 && (fileItem.FileName.EndsWith("xls") || fileItem.FileName.EndsWith("xlsx")))
                {
                    
                    string fileLocation = System.Web.Hosting.HostingEnvironment.MapPath("~/Content" + fileItem.FileName);
                    if (System.IO.File.Exists(fileLocation))
                        System.IO.File.Delete(fileLocation);
                    fileItem.SaveAs(fileLocation);
                    Excel.Application excelApp = new Excel.Application();
                    // --
                    // Lê o arquivo
                    Excel.Workbook planCtrlBook = excelApp.Workbooks.Open(fileLocation, true);
                    Excel.Worksheet planCtrlSheet = (Excel.Worksheet)planCtrlBook.Worksheets.get_Item(1);
                    Excel.Range planCtrlRange = planCtrlSheet.UsedRange;
                    // --
                    // Instancia a lista de documentos
                    planilhaCtrl = new List<Planilha>();
                    // --
                    // Percorre todas as linhas do range utilizado no documento
                    // iniciamos na linha 2 para ignorar o título
                    for (int rCnt = 2; rCnt <= planCtrlRange.Rows.Count; rCnt++)
                    {
                        // Instancia
                        Planilha doc = new Planilha();
                        doc.nrDemanda = Convert.ToString((planCtrlRange.Cells[rCnt, 1] as Excel.Range).Value2);
                        doc.job = Convert.ToString((planCtrlRange.Cells[rCnt, 2] as Excel.Range).Value2);
                        doc.gerente = new profissional()
                        {
                            nome = Convert.ToString((planCtrlRange.Cells[rCnt, 3] as Excel.Range).Value2),
                            email = Convert.ToString((planCtrlRange.Cells[rCnt, 4] as Excel.Range).Value2),
                        };
                        doc.caminhoArquivo = Convert.ToString((planCtrlRange.Cells[rCnt, 5] as Excel.Range).Value2);
                        // --
                        // Adiciona na lista
                        planilhaCtrl.Add(doc);
                    }
                    // --
                    // Fecha a planilha
                    planCtrlBook.Close(true, null, null);
                    excelApp.Quit();
                }
                else
                {
                    retorno = "Planilha invalida ou não nenhum item inserido na planilha.";                    
                }
            }
            catch (Exception)
            {
                throw;
            }
            return planilhaCtrl;
        }
    }
}