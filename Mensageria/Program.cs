using System;
using System.Net;
using System.Net.Mail;
using System.Collections.Generic;
using System.IO;
using OfficeOpenXml;
using Mensageria;

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



            // Aqui deve ser informado o caminho da planilha excel, que será usada com base dos dados para enviar os emails
            string excelFilePath = @"C:\Users\robert.alves\source\aulas\AulasEuCodo\ReadFiles\dadostoemail.xlsx";


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
                        //abaixo estão as colunas na planilha
                        string bloco = worksheet.Cells[row, 1].Value?.ToString();
                        string unidade = worksheet.Cells[row, 2].Value?.ToString();
                        string corretor = worksheet.Cells[row, 3].Value?.ToString();
                        string cliente = worksheet.Cells[row, 4].Value?.ToString();
                        string gerente = worksheet.Cells[row, 5].Value?.ToString();
                        string clienteEmail = worksheet.Cells[row, 6].Value?.ToString();
                        string empreendimento = worksheet.Cells[row, 7].Value?.ToString();



                        // Supondo que você tenha a classe TemplateEmail e o método GetHtmlTemplate definidos

                        // Instancie a classe TemplateEmail
                        TemplateEmail templateEmail = new TemplateEmail();

                        // Agora você pode usar a variável emailBody para compor o corpo do seu e-mail
                        // e enviá-lo usando a biblioteca ou método de envio de e-mails de sua escolha


                        MailMessage mail = new MailMessage(senderEmail, clienteEmail)
                        {
                            Subject = $@"PASTA {empreendimento.ToUpper()} | Corretor: {corretor}",
                            Body = templateEmail.GetHtmlTemplate(cliente, corretor, gerente, bloco, unidade, empreendimento, saudacao),
                            IsBodyHtml = true
                        };

                        // Construir o caminho do arquivo com base no número da unidade
                        string filePath = $@"C:\Users\robert.alves\source\aulas\AulasEuCodo\ReadFiles\imobfile\{unidade}.pdf"; // Substitua pela extensão e caminho correto

                        if (File.Exists(filePath))
                        {
                            mail.Attachments.Add(new Attachment(filePath));

                            try
                            {
                                smtpClient.Send(mail);
                                Console.WriteLine($"Email enviado para {clienteEmail} com arquivo anexado");
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine($"Erro ao enviar email para {clienteEmail}: {ex.Message}");
                            }
                        }
                        else
                        {
                            Console.WriteLine($"Arquivo não encontrado para a unidade {unidade}. Email não enviado.");
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
    }
}
