using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Mail;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace Mensageria.Service
{
    public class EmailService
    {
        public static void EmailBot(string senderEmail, string senderPassword) {

            // Licença Excel EPPlus
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;


            // Configurações do servidor SMTP adequado para envio em massa
            string smtpHost = "smtp.gmail.com";
            int smtpPort = 587;



            //Caminho da Pasta onde está a planilha e os arquivos: recuando 3 pastas e acessando a pasta betti

            // Obtém o diretório do projeto
            string diretorioDoProjeto = Path.Combine(Directory.GetCurrentDirectory(), "..", "..", "..", "betti");


            // Acessando a planilha excel
            string excelFilePath = Path.Combine(diretorioDoProjeto, "email.xlsx");
            Console.WriteLine(excelFilePath);

            string saudacao = DateTime.Now.Hour switch
            {
                int horaAtual when horaAtual > 4 && horaAtual <= 11 => "Bom Dia,",
                int horaAtual when horaAtual >= 12 && horaAtual <= 18 => "Boa Tarde,",
                _ => "Boa Noite,"
            };


            try
            {
                using (var package = new ExcelPackage(new FileInfo(excelFilePath)))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[0]; // Assume que a planilha está na primeira guia
                    int rowCount = worksheet.Dimension.Rows;

                    // Configurar o cliente SMTP
                    SmtpClient smtpClient = new SmtpClient(smtpHost, smtpPort);
                    smtpClient.EnableSsl = true;
                    smtpClient.Credentials = new NetworkCredential(senderEmail, senderPassword);

                    for (int row = 2; row <= rowCount; row++)
                    {
                        // Abaixo estão as colunas na planilha
                        string empreendimento = worksheet.Cells[row, 1].Value?.ToString();
                        string torre = worksheet.Cells[row, 2].Value?.ToString();
                        string unidade = worksheet.Cells[row, 3].Value?.ToString();
                        string corretor = worksheet.Cells[row, 4].Value?.ToString();
                        string gerente = worksheet.Cells[row, 5].Value?.ToString();
                        string cliente = worksheet.Cells[row, 6].Value?.ToString();
                        string clienteEmail = worksheet.Cells[row, 7].Value?.ToString();
                        string emailsCC = worksheet.Cells[row, 8].Value?.ToString();


                        // Supondo que você tenha a classe TemplateEmail e o método GetHtmlTemplate definidos

                        // Instancie a classe TemplateEmail
                        TemplateEmail templateEmail = new TemplateEmail();

                        // Agora você pode usar a variável emailBody para compor o corpo do seu e-mail
                        // e enviá-lo usando a biblioteca ou método de envio de e-mails de sua escolha


                        MailMessage mail = new MailMessage(senderEmail, clienteEmail)
                        {
                            Subject = $@"PASTA {empreendimento.ToUpper()} | Corretor: {corretor}",
                            Body = templateEmail.GetHtmlTemplate(cliente, corretor, gerente, torre, unidade, empreendimento, saudacao),
                            IsBodyHtml = true
                        };

                        // Construir o caminho do arquivo com base no número da unidade
                        string filePath = Path.Combine(diretorioDoProjeto, $@"arquivos\{unidade}.pdf");

                        if (File.Exists(filePath))
                        {
                            mail.Attachments.Add(new Attachment(filePath));

                            try
                            {
                                string[] enderecosCC = emailsCC?.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);

                                if (enderecosCC != null && enderecosCC.Length > 0)
                                {
                                    // Adicione cada email em cópia (CC) separadamente
                                    foreach (string endereco in enderecosCC)
                                    {
                                        mail.CC.Add(endereco.Trim());
                                    }
                                }

                                smtpClient.Send(mail);
                                Console.WriteLine($"Email com arquivo enviado para {clienteEmail} | Unidade {unidade}");
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine($"Erro ao enviar email para {clienteEmail}: {ex.Message}");
                            }
                        }
                        else
                        {
                            Console.WriteLine($"Arquivo não encontrado para a unidade {unidade}. Destinatário {clienteEmail} não irá receber e-mail.");
                        }
                    }

                    Console.WriteLine("Envio em massa concluído!");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Ocorreu um erro: " + ex.Message);
            }

        }



        // ENVIO INSTANTÂNEO 
        public static async Task EnviarEmailsEmParalelo(string senderEmail, string senderPassword)
        {
            // Configurações do servidor SMTP adequado para envio em massa
            string smtpHost = "smtp.gmail.com";
            int smtpPort = 587;

            //smtplw.com.br (localweb)


            // Caminho da Pasta onde está a planilha e os arquivos: recuando 3 pastas e acessando a pasta betti
            string diretorioDoProjeto = Path.Combine(Directory.GetCurrentDirectory(), "..", "..", "..", "betti");

            // Acessando a planilha excel
            string excelFilePath = Path.Combine(diretorioDoProjeto, "email.xlsx");
            Console.WriteLine(excelFilePath);

            string saudacao = DateTime.Now.Hour switch
            {
                int horaAtual when horaAtual > 4 && horaAtual <= 11 => "Bom Dia,",
                int horaAtual when horaAtual >= 12 && horaAtual <= 18 => "Boa Tarde,",
                _ => "Boa Noite,"
            };

            try
            {
                using (var package = new ExcelPackage(new FileInfo(excelFilePath)))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[0]; // Assume que a planilha está na primeira guia
                    int rowCount = worksheet.Dimension.Rows;

                    List<Task> emailTasks = new List<Task>();

                    for (int row = 2; row <= rowCount; row++)
                    {
                        string empreendimento = worksheet.Cells[row, 1].Value?.ToString();
                        string torre = worksheet.Cells[row, 2].Value?.ToString();
                        string unidade = worksheet.Cells[row, 3].Value?.ToString();
                        string corretor = worksheet.Cells[row, 4].Value?.ToString();
                        string gerente = worksheet.Cells[row, 5].Value?.ToString();
                        string cliente = worksheet.Cells[row, 6].Value?.ToString();
                        string clienteEmail = worksheet.Cells[row, 7].Value?.ToString();
                        string emailsCC = worksheet.Cells[row, 8].Value?.ToString();

                        // Supondo que você tenha a classe TemplateEmail e o método GetHtmlTemplate definidos
                        TemplateEmail templateEmail = new TemplateEmail();

                        using (SmtpClient smtpClient = new SmtpClient(smtpHost, smtpPort))
                        {

                            smtpClient.EnableSsl = true;
                            smtpClient.Credentials = new NetworkCredential(senderEmail, senderPassword);
                            //smtpClient.Timeout = 3000;  Tempo limite em milissegundos (aumente conforme necessário)



                            MailMessage mail = new MailMessage(senderEmail, clienteEmail)
                            {
                                Subject = $@"PASTA {empreendimento.ToUpper()} | Corretor: {corretor}",
                                Body = templateEmail.GetHtmlTemplate(cliente, corretor, gerente, torre, unidade, empreendimento, saudacao),
                                IsBodyHtml = true
                            };

                            // Construir o caminho do arquivo com base no número da unidade
                            string filePath = Path.Combine(diretorioDoProjeto, $@"arquivos\{unidade}.pdf");

                            if (File.Exists(filePath))
                            {
                                mail.Attachments.Add(new Attachment(filePath));

                                try
                                {
                                    string[] enderecosCC = emailsCC.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);

                                    // Adicione cada email em cópia (CC) separadamente
                                    foreach (string endereco in enderecosCC)
                                    {
                                        mail.CC.Add(endereco.Trim());
                                    }

                                    emailTasks.Add(smtpClient.SendMailAsync(mail));
                                    Console.WriteLine($"Email com arquivo enviado para {clienteEmail} | Unidade {unidade}");
                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine($"Erro ao enviar email para {clienteEmail}: {ex.Message}");
                                }
                            }
                            else
                            {
                                Console.WriteLine($"Arquivo não encontrado para a unidade {unidade}. Destinatário {clienteEmail} não irá receber e-mail.");
                            }
                        }
                    }

                    try
                    {
                        await Task.WhenAll(emailTasks);
                        Console.WriteLine("E-mails enviados com sucesso!");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("Erro ao enviar e-mails: " + ex.Message);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Ocorreu um erro: " + ex.Message);
            }
        }




        // ENVIANDO PARA DIVERSOS DISTINATÁRIOS DE MANEIRA INSTANTÂNEA

       public static async Task SendInstantEmailsToRecipients(string senderEmail, string senderPassword)
    {


            string[] destinatarios = {
            "robert.ads.anjos@gmail.com", "bulpert@yahoo.com", "robert_hk_@hotmail.com",
            "robert.ads.anjos@gmail.com", "bulpert@yahoo.com", "robert_hk_@hotmail.com",
            "robert.ads.anjos@gmail.com", "bulpert@yahoo.com", "robert_hk_@hotmail.com",
            "robert.ads.anjos@gmail.com", "bulpert@yahoo.com", "robert_hk_@hotmail.com",
            "robert.ads.anjos@gmail.com", "bulpert@yahoo.com", "robert_hk_@hotmail.com",
            "robert.ads.anjos@gmail.com", "bulpert@yahoo.com", "robert_hk_@hotmail.com",
            "robert.ads.anjos@gmail.com", "bulpert@yahoo.com", "robert_hk_@hotmail.com",
        };

            // Assunto e corpo do e-mail
            string assunto = "DEV TESTE ASSYNC";
        string corpo = "TESTINHO";

        // Crie uma lista para armazenar tarefas de envio de e-mail
        List<Task> tasks = new List<Task>();

        foreach (var destinatario in destinatarios)
        {
            var smtpClient = new SmtpClient("smtp.gmail.com")
            {
                Port = 587,
                Credentials = new NetworkCredential(senderEmail, senderPassword),
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
}
