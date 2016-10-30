using System.Web.Mvc;

namespace EnvioEmail.Controllers
{
    public class MailController : Controller
    {
        public ActionResult Form()
        {
            return View();
        }
        [HttpPost]
        public ActionResult Form(string receiverEmail, string subject, string message)
        {
            string exception;
            string[] files = new string[] { };
            Classes.ServicoEmail.enviaEmail(receiverEmail, subject, message, out exception,files);
            if (exception == "Sucesso")
                ViewBag.retornoEnvio = "E - mail enviado com sucesso";
            else ViewBag.retornoEnvio = exception;
            return View("Form");
        }
    }
}