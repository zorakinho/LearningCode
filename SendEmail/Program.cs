using System;
using System.Net;
using System.Net.Mail;
using System.Collections.Generic;
using System.IO;

namespace MassEmailSenderExample
{
    class Program
    {
        static void Main(string[] args)
        {
            // Configurações do remetente
            string senderEmail = "robert.alves@olxbr.com";
            string senderPassword = "ra02xbo0$TRK";

            // Configurações do servidor SMTP adequado para envio em massa
            string smtpHost = "smtp.gmail.com";
            int smtpPort = 587;

            List<string> recipientEmails = new List<string>
            {

                "bulpert@yahoo.com",
                // Adicione mais destinatários aqui
            };

            try
            {
                // Configurar o cliente SMTP
                SmtpClient smtpClient = new SmtpClient(smtpHost, smtpPort);
                smtpClient.EnableSsl = true;
                smtpClient.Credentials = new NetworkCredential(senderEmail, senderPassword);

                foreach (string recipientEmail in recipientEmails)
                {
                    try
                    {
                        // Criar o objeto de mensagem
                        MailMessage mail = new MailMessage(senderEmail, recipientEmail);
                        mail.Subject = "Assunto do email";
                        mail.Body = "Conteúdo do email";

                        // Incluir opção de cancelamento de inscrição (unsubscribe)
                        mail.Body += "\n\nClique aqui para cancelar a inscrição: https://seusite.com/unsubscribe";

                        // Anexar arquivos
                        string[] filesToAttach = new string[]
                        {
                            @"C:\Users\robert.alves\source\aulas\AulasEuCodo\ReadFiles\imobfile\1010.pdf",
                            // Adicione mais caminhos de arquivos aqui
                        };

                        foreach (string filePath in filesToAttach)
                        {
                            if (File.Exists(filePath))
                            {
                                mail.Attachments.Add(new Attachment(filePath));
                            }
                        }

                        // Enviar o email
                        smtpClient.Send(mail);

                        Console.WriteLine($"Email enviado para {recipientEmail}");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Erro ao enviar email para {recipientEmail}: {ex.Message}");
                    }
                }

                Console.WriteLine("Envio em massa concluído!");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Ocorreu um erro: " + ex.Message);
            }
        }
    }
}
