/*using System;
using System.Net;
using System.Net.Mail;
using System.Threading.Tasks;
using System.Collections.Generic;

class Program
{
    static async Task Main(string[] args)
    {
        // Lista de destinatários
        string[] destinatarios = { "robert.ads.anjos@gmail.com", "bulpert@yahoo.com", "robert_hk_@hotmail.com", "robert.ads.anjos@gmail.com", "bulpert@yahoo.com", "robert_hk_@hotmail.com", "robert.ads.anjos@gmail.com", "bulpert@yahoo.com", "robert_hk_@hotmail.com", "robert.ads.anjos@gmail.com", "bulpert@yahoo.com", "robert_hk_@hotmail.com", "robert.ads.anjos@gmail.com", "bulpert@yahoo.com", "robert_hk_@hotmail.com", "robert.ads.anjos@gmail.com", "bulpert@yahoo.com", "robert_hk_@hotmail.com", "robert.ads.anjos@gmail.com", "bulpert@yahoo.com", "robert_hk_@hotmail.com", "robert.ads.anjos@gmail.com", "bulpert@yahoo.com", "robert_hk_@hotmail.com", "robert.ads.anjos@gmail.com", "bulpert@yahoo.com", "robert_hk_@hotmail.com", };

        // Assunto e corpo do e-mail
        string assunto = "Assunto do e-mail";
        string corpo = "Corpo do e-mail";

        // Crie uma lista para armazenar tarefas de envio de e-mail
        List<Task> tasks = new List<Task>();

        foreach (var destinatario in destinatarios)
        {
            var smtpClient = new SmtpClient("smtp.gmail.com")
            {
                Port = 587,
                Credentials = new NetworkCredential("robert.alves@olxbr.com", "ra02xbo0$TRK"),
                EnableSsl = true,
            };

            var mensagem = new MailMessage("robert.alves@olxbr.com", destinatario, assunto, corpo);

            // Adicione a tarefa de envio de e-mail à lista
            tasks.Add(EnviarEmailAsync(smtpClient, mensagem));
        }

        // Aguarde até que todas as tarefas tenham sido concluídas
        await Task.WhenAll(tasks);

        Console.WriteLine("E-mails enviados com sucesso!");
    }

    static async Task EnviarEmailAsync(SmtpClient smtpClient, MailMessage mensagem)
    {
        try
        {
            await smtpClient.SendMailAsync(mensagem);
            Console.WriteLine($"E-mail enviado para: {mensagem.To[0].Address}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Erro ao enviar e-mail para {mensagem.To[0].Address}: {ex.Message}");
        }
        finally
        {
            mensagem.Dispose();
        }
    }
}
*/