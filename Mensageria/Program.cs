using System;
using System.Net;
using System.Net.Mail;
using System.Collections.Generic;
using System.IO;
using OfficeOpenXml;

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

            string excelFilePath = @"C:\Users\robert.alves\source\aulas\AulasEuCodo\ReadFiles\dadostoemail.xlsx"; // Caminho para o arquivo Excel

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
                        string bloco = worksheet.Cells[row, 1].Value?.ToString();
                        string unidade = worksheet.Cells[row, 2].Value?.ToString();
                        string corretor = worksheet.Cells[row, 3].Value?.ToString();
                        string cliente = worksheet.Cells[row, 4].Value?.ToString();
                        string gerente = worksheet.Cells[row, 5].Value?.ToString();
                        string clienteEmail = worksheet.Cells[row, 6].Value?.ToString();

                        string emailBody = $"Olá {cliente},\n\nSegue abaixo as informações:\n\n" +
                            $"Bloco: {bloco}\nUnidade: {unidade}\nCorretor: {corretor}\nGerente: {gerente}";

                        MailMessage mail = new MailMessage(senderEmail, clienteEmail)
                        {
                            Subject = "Informações da unidade",
                            Body = emailBody
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
