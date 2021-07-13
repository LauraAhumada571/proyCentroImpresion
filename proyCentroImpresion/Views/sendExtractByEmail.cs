using proyCentroImpresion.Models;
using System;
using System.Collections.Generic;
using System.Net;
using System.Net.Mail;
using System.Text;
using System.Windows;

namespace proyCentroImpresion.Views
{
    public class sendExtractByEmail
    {
        //configuración y envio de los extractos a través de correo electrónico
        public void send(List<DataExtractCollection> dataextractcollection)
        {
            String from = "pruebastrabajo66@gmail.com";
            String fromname = "Carvajal Soluciones de Comunicación";
            String smtp_user = "pruebastrabajo66@gmail.com";
            String smtp_pass = "0uw(%KxQbhdU";
            String host = "smtp.gmail.com";
            int port = 587;

            try
            {
                MessageBox.Show("Enviando emails");
                for (int i = 0; i < dataextractcollection.Count; i++)
                {
                    if (dataextractcollection[i].output_type == "send by email")
                    {
                        var attachment = new System.Net.Mail.Attachment("E:\\enviarEmail\\" + dataextractcollection[i].id_invoice + ".pdf");
                        String to = dataextractcollection[i].email;

                        String subject = "Extracto de cuenta " + dataextractcollection[i].id_invoice;

                        String mess =
                        "<h1>CARVAJAL SOLUCIONES DE COMUNICACIÓN S.A.S.</h1>" +
                        "<p>Apreciado (a) " + dataextractcollection[i].name + "</p><br>" +
                        "<p>Adjunto enviamos el extracto correspodiente al cotracto No " + dataextractcollection[i].id_contract +
                        " para su verificación. </p> <br><br><br>" +
                        "<p> " + dataextractcollection[i].random_text + "</p>";

                        MailMessage message = new MailMessage();
                        message.IsBodyHtml = true;
                        message.From = new MailAddress(from, fromname);
                        message.To.Add(new MailAddress(to));
                        message.Subject = subject;
                        message.Body = mess;
                        message.Attachments.Add(attachment);

                        using (var client = new System.Net.Mail.SmtpClient(host, port))
                        {
                            client.Credentials =
                                new NetworkCredential(smtp_user, smtp_pass);

                            client.EnableSsl = true;

                            try
                            {
                                client.Send(message);
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show("No fue posible enviar el extracto" + dataextractcollection[i].id_invoice + " a " + dataextractcollection[i].name);
                                MessageBox.Show("Error "+ ex);
                            }
                        }
                    }
                    
                }
                
                MessageBox.Show("Correos enviados satisfactoriamente");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Aún no hay información para enviar por correo, por favor cargue la información " + ex);
            }
        }
    }
}
