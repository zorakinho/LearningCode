﻿using OfficeOpenXml;
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
        /*
        public static void EmailBot(string remetente, string remetentePassword) {

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
                    smtpClient.Credentials = new NetworkCredential(remetente, remetentePassword);

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


                        MailMessage mail = new MailMessage(remetente, clienteEmail)
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
        */


        // ENVIO INSTANTÂNEO 
        public static async Task EnviarEmailsEmParalelo(string remetente, string remetentePassword, string emailDestino, string smtpHost, int smtpPort)
        {

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial; // Licença Excel
            string diretorioDoProjeto = Path.Combine(Directory.GetCurrentDirectory(), "..", "..", "..", "betti"); // Localizando diretórios com os arquivos
            string excelFilePath = Path.Combine(diretorioDoProjeto, "email.xlsx"); // Localizando planilha Excel


            
            
            // Assunto e corpo do e-mail
            string assunto = @$"BOT SENDMAIL {DateTime.Now}";
            // string corpo = "TESTINHO";


            string saudacao = DateTime.Now.Hour switch
            {
                int horaAtual when horaAtual > 4 && horaAtual <= 11 => "Bom Dia,",
                int horaAtual when horaAtual >= 12 && horaAtual <= 18 => "Boa Tarde,",
                _ => "Boa Noite,"
            };


            // Crie uma lista para armazenar tarefas de envio de e-mail
            List<Task> tasks = new();


            using (var package = new ExcelPackage(new FileInfo(excelFilePath)))
            {

                ExcelWorksheet worksheet = package.Workbook.Worksheets[0]; // Assume que a planilha está na primeira guia
                int rowCount = worksheet.Dimension.Rows;
                Console.WriteLine($"enviando {rowCount-1}");

                for (int row = 2; row <= rowCount; row++)
                {

                    string empreendimento = worksheet.Cells[row, 1].Value?.ToString();
                    string torre = worksheet.Cells[row, 2].Value?.ToString();
                    string unidade = worksheet.Cells[row, 3].Value?.ToString();
                    string corretor = worksheet.Cells[row, 4].Value?.ToString();
                    string gerente = worksheet.Cells[row, 5].Value?.ToString();
                    string cliente = worksheet.Cells[row, 6].Value?.ToString();
                    //string clienteEmail = worksheet.Cells[row, 7].Value?.ToString();
                    //string emailsCC = worksheet.Cells[row, 8].Value?.ToString();


                    // Construir o caminho do arquivo com base no número da unidade
                    string filePath = Path.Combine(diretorioDoProjeto, $@"arquivos\{unidade}.pdf");

                    string arquivo = Path.GetFileNameWithoutExtension(filePath);

                    Console.WriteLine($"ARQUIVO: {arquivo}");
                    if (File.Exists(filePath) && arquivo == unidade)
                    {

                        var smtpClient = new SmtpClient(smtpHost)
                        {
                            Port = smtpPort,
                            Credentials = new NetworkCredential(remetente, remetentePassword),
                            EnableSsl = true,
                        };

                        var mensagem = new MailMessage(remetente, emailDestino, assunto, "");


                        // Adicione o template HTML ao corpo do e-mail
                        string templateHtml = TemplateEmail.GetHtmlTemplate(cliente, corretor, gerente, torre, unidade, empreendimento, saudacao); // Suponha que esta função retorne o HTML do seu template
                        mensagem.Body = templateHtml;
                        mensagem.IsBodyHtml = true;

                        // Adicione a tarefa de envio de e-mail à lista
                        tasks.Add(EnviarEmailAsync(smtpClient, mensagem, filePath));

                        //await Task.Delay(200);


                    }
                    else
                    {
                        Console.WriteLine($"Unidade {unidade} não enviada, está faltando o anexo!");
                    }

                }
            }
            // Aguarde até que todas as tarefas tenham sido concluídas
            await Task.WhenAll(tasks);

            Console.WriteLine("E-mails enviados com sucesso!");
        }





        // ENVIANDO PARA DIVERSOS DISTINATÁRIOS DE MANEIRA INSTANTÂNEA


        /*
        public static async Task SendInstantEmailsToRecipients(string remetente, string remetentePassword)
    {

            
                        string[] destinatarios = { "robert.ads.anjos@gmail.com", "bulpert@yahoo.com", "robert_hk_@hotmail.com"};
            


        // Assunto e corpo do e-mail
        string assunto = "DEV 400 MILISEC ";
        string corpo = "TESTINHO";

        // Crie uma lista para armazenar tarefas de envio de e-mail
        List<Task> tasks = new List<Task>();

        for(int i = 0; i < 50; i++) 
        {
            var smtpClient = new SmtpClient("smtp.gmail.com")
            {
                Port = 587,
                Credentials = new NetworkCredential(remetente, remetentePassword),
                EnableSsl = true,
            };

            var mensagem = new MailMessage("robert.alves@olxbr.com", "robert.ads.anjos@gmail.com", assunto, corpo);

            // Adicione a tarefa de envio de e-mail à lista
            tasks.Add(EnviarEmailAsync(smtpClient, mensagem));

            await Task.Delay(400);
            }

        // Aguarde até que todas as tarefas tenham sido concluídas
        await Task.WhenAll(tasks);

        Console.WriteLine("E-mails enviados com sucesso!");
    }
*/



        static async Task EnviarEmailAsync(SmtpClient smtpClient, MailMessage mensagem, string filePath)
        {
            try
            {
                mensagem.Attachments.Add(new Attachment(filePath));
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
