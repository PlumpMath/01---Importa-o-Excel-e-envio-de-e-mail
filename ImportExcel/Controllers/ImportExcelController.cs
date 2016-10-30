using EnvioEmail.Classes;
using System.Collections.Generic;
using System.Web.Mvc;
using ImportExcel.Models;
using ImportExcel.Classes;
using System;
using System.IO;

namespace ImportExcel.Controllers
{
    public class ImportExcelController : Controller
    {
        public ActionResult ImportExcel()
        {
            return View();
        }

        [HttpPost]
        public ActionResult LoadFile(FormCollection itemFile)
        {
            try
            {
                var fileItem = Request.Files[0];
                List<Planilha> itensPlanilha = new List<Planilha>();
                string retornoPlanilha;
                itensPlanilha = LePlanilha.ObtemItems(fileItem, out retornoPlanilha);
                if (retornoPlanilha != "ok")
                {
                    @ViewBag.msgErro = retornoPlanilha;
                    return View("ImportExcel");
                }
                if (itensPlanilha.Count > 0)
                {
                    ViewBag.retornoEnvio = "Itens da Planilha foram importados, preparando envio de e-mail.";
                    enviaEmailToItensPlanilha(itensPlanilha);
                    ViewBag.retornoEnvio = $"Emails enviados para {itensPlanilha.Count} contatos.";
                    return View("ImportExcel");
                }
                else
                {
                    @ViewBag.msgErro = "Nenhuma planilha Selecionada.";
                    return View("ImportExcel");
                }
            }
            catch (Exception e)
            {
                throw;
            }
        }

        private void enviaEmailToItensPlanilha(List<Planilha> itensPlanilha)
        {
            foreach (var itemPlanilha in itensPlanilha)
            {
                try
                {
                    string mensagem = $"<h1 style='color:green'>mensagem enviada para :{itemPlanilha.gerente.nome}</h1>"
                         + $"<b>Cargo/Empresa:</b> {itemPlanilha.job}<br/>"
                        + $"<b>Identificação:</b> {itemPlanilha.nrDemanda}";
                    string retorno;
                    string[] files = new string[] { itemPlanilha.caminhoArquivo };
                    ServicoEmail.enviaEmail(itemPlanilha.gerente.email,
                       $"Primeiro teste de envio de e-mail via planilha{itemPlanilha.gerente.nome}",
                        mensagem,
                        out retorno, files);
                    if (System.IO.File.Exists(itemPlanilha.caminhoArquivo)) moveArquivo(itemPlanilha.caminhoArquivo);
                }
                catch (Exception e)
                {
                    throw;
                }
            }
        }

        private void moveArquivo(string file)
        {
            try
            {
                string[] nomearquivo = file.Split('\\');
                string destinationFile = "C:\\Users\\leand\\Desktop\\Teste\\Movidos\\" + nomearquivo[nomearquivo.Length - 1];
                Directory.Move(file, @destinationFile);
            }
            catch (Exception e)
            {
                throw e;
            }
        }
    }
}
    