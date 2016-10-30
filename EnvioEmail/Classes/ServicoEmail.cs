using System;
using System.IO;
using System.Net;
using System.Net.Mail;
using System.Net.Mime;
using static System.Net.WebRequestMethods;

namespace EnvioEmail.Classes
{
    public class ServicoEmail
    {
        static public bool enviaEmail(string destinatario, string assunto, string mensagem, out string exception, string[] files)
        {
            try
            {
                var sendermail = new MailAddress("leandrofire@live.com", "Leandro Teixeira");
                var receivermail = new MailAddress(destinatario, "Receiver");
                var password = "DlA685947";
                SmtpClient smtp = new SmtpClient
                {
                    Host = "smtp.live.com",
                    Port = 587,
                    EnableSsl = true,
                    DeliveryMethod = SmtpDeliveryMethod.Network,
                    UseDefaultCredentials = false,
                    Credentials = new NetworkCredential(sendermail.Address, password)
                };
                using (var mess = new MailMessage(sendermail.Address, destinatario)
                {
                    Subject = assunto,
                    Body = mensagem,
                    IsBodyHtml = true
                })
                {
                    foreach (var file in files)
                    {
                        bool enviaEmail = false;
                        if (System.IO.File.Exists(file)) enviaEmail = true;

                        if (enviaEmail)
                        {
                            // Create the file attachment for this e-mail message.
                            Attachment data = new Attachment(file, MediaTypeNames.Application.Octet);
                            // Add time stamp information for the file.
                            ContentDisposition disposition = data.ContentDisposition;
                            disposition.CreationDate = System.IO.File.GetCreationTime(file);
                            disposition.ModificationDate = System.IO.File.GetLastWriteTime(file);
                            disposition.ReadDate = System.IO.File.GetLastAccessTime(file);
                            // Add the file attachment to this e-mail message.
                            mess.Attachments.Add(data);
                            
                        }
                    }

                    smtp.Send(mess);
                }
                exception = "Sucesso";
                return true;
            }
            catch (Exception e)
            {
                exception = e.ToString();
                return false;
            }
        }


    }

}