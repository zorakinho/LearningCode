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

        // ENVIO ASSYNC
        public static async Task BotEmail(string remetente, string remetentePassword, string smtpHost, int smtpPort, int time = 1000)
        {
            // Licença Excel
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            // Localizando diretórios com os enviar
            string diretorioDoProjeto = Path.Combine(Directory.GetCurrentDirectory(), "..", "..", "..", "pastas");

            // Localizando planilha Excel
            string excelcaminhoDaPasta = Path.Combine(diretorioDoProjeto, "email.xlsx"); 


            string saudacao = DateTime.Now.Hour switch
            {
                int horaAtual when horaAtual > 4 && horaAtual <= 11 => "Bom Dia,",
                int horaAtual when horaAtual >= 12 && horaAtual <= 18 => "Boa Tarde,",
                _ => "Boa Noite,"
            };


            // Crie uma lista para armazenar tarefas de envio de e-mail
            List<Task> tarefas = new();


            using (var package = new ExcelPackage(new FileInfo(excelcaminhoDaPasta)))
            {
                // Assume que a planilha está na primeira guia
                ExcelWorksheet planilhaExcel = package.Workbook.Worksheets[0]; 
                int linhasExistentes = planilhaExcel.Dimension.Rows;
                Console.WriteLine($"FORAM CONTABILIZADAS {linhasExistentes-1} PASTAS PARA O ENVIO:");

                var enviarEmailevent = new AutoResetEvent(false);
                var timer = new Timer(state =>
                {
                    enviarEmailevent.Set();
                }, null, 0, time); // Dispara o evento a cada segundo

                // Usando o Parallel para realizar uma iteração paralela 
                _ = Parallel.For(2, linhasExistentes + 1, row =>
                {

                    string empreendimento = planilhaExcel.Cells[row, 1].Value?.ToString();
                    string torre = planilhaExcel.Cells[row, 2].Value?.ToString();
                    string unidade = planilhaExcel.Cells[row, 3].Value?.ToString();
                    string corretor = planilhaExcel.Cells[row, 4].Value?.ToString();
                    string gerente = planilhaExcel.Cells[row, 5].Value?.ToString();
                    string cliente = planilhaExcel.Cells[row, 6].Value?.ToString();
                    string emailDestino = planilhaExcel.Cells[row, 7].Value?.ToString();
                    string emailsCC = planilhaExcel.Cells[row, 8].Value?.ToString();




                    // Construir o caminho do arquivo com base no número da unidade
                    string caminhoDaPasta = Path.Combine(diretorioDoProjeto, $@"enviar\{unidade}.pdf");

                    
                    if (File.Exists(caminhoDaPasta) && Path.GetFileNameWithoutExtension(caminhoDaPasta) == unidade)
                    {
                        enviarEmailevent.WaitOne();
                        try
                        {

                            // Assunto do email
                            string assunto = @$"INNOVA BR: PASTA {empreendimento.ToUpper()} | Unidade {unidade}";

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
                            mensagem.BodyEncoding = Encoding.UTF8;
                            mensagem.SubjectEncoding = Encoding.UTF8;
                            mensagem.Headers.Add("Content-Type", "text/html; charset=UTF-8");

                            // Adicione os e-mails em cópia (CC) ao e-mail
                            AdicionarEmailsCC(mensagem, emailsCC);
                            // Adicione a tarefa de envio de e-mail à lista
                            tarefas.Add(EnviarEmailAsync(smtpClient, mensagem, caminhoDaPasta, unidade, diretorioDoProjeto, emailsCC));

                           // await Task.Delay(5000); // milisegundos
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Erro ao processar a Unidade {unidade}: {ex.Message}");
                        }

                    }
                    else
                    {
                        Console.WriteLine($"Unidade {unidade} não enviada, está faltando o anexo!");
                    }

                });
                timer.Dispose(); // Liberar recursos do temporizador
            }
            // Aguarde até que todas as tarefas tenham sido concluídas
            await Task.WhenAll(tarefas);

            Console.WriteLine("Todos E-mails válidos enviados com sucesso!");
        }



        // RENOMEANDO enviar
        public static void RenomearArquivo(string caminhoDaPasta, string novoNome, string diretorioDoProjeto, string unidade)
        {
            try
            {
                string novoCaminhoArquivo = Path.Combine(diretorioDoProjeto, "enviar", novoNome);
                File.Move(caminhoDaPasta, novoCaminhoArquivo);
                Console.WriteLine($"E-mail da Unidade {unidade} enviada com sucesso!");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Erro ao renomear o arquivo: {ex.Message}");
            }
        }

        // Método para adicionar e-mails em cópia (CC) ao e-mail
        private static void AdicionarEmailsCC(MailMessage mensagem, string emailsCC)
        {
            if (!string.IsNullOrWhiteSpace(emailsCC))
            {
                string[] copiados = emailsCC.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
                foreach (string copiado in copiados)
                {
                    mensagem.CC.Add(copiado.Trim());
                }
            }
        }



        //ENVIANDO E-MAIL ASSINC E RENOMEANDO ARQUIVO ENVIADO
        static async Task EnviarEmailAsync(SmtpClient smtpClient, MailMessage mensagem, string caminhoDaPasta, string unidade, string diretorioDoProjeto, string emailsCC)
        {
            bool renomear;
            try
            {
                mensagem.Attachments.Add(new Attachment(caminhoDaPasta));
                await smtpClient.SendMailAsync(mensagem);
                renomear = true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Erro ao enviar e-mail para a Unidade: {unidade} - motivo:{ex.Message}");
                renomear = false;
            }
            finally
            {
                mensagem.Dispose();
            }
            if (renomear)
            {

            // Aguardar a conclusão do envio de e-mail antes de renomear o arquivo
            RenomearArquivo(caminhoDaPasta, $"enviado_{unidade}.pdf", diretorioDoProjeto, unidade);

            }
        }



        //RENOMEANDO ARQUIVOS EM ORDEM NÚMERICA
        public static void RenomearArquivosParaNumerosSequenciais()
        {
            string diretorio = Path.Combine(Directory.GetCurrentDirectory(), "..", "..", "..", "pastas", "enviar");
            try
            {
                string[] arquivos = Directory.GetFiles(diretorio, "*.pdf");

                int contador = 1;

                foreach (string arquivo in arquivos)
                {
                    string nomeArquivo = Path.GetFileNameWithoutExtension(arquivo);
                    string extensao = Path.GetExtension(arquivo);

                    string novoNome = $"{contador}{extensao}";

                    string novoCaminho = Path.Combine(diretorio, novoNome);

                    // Verifica se o novo nome já existe, e se existir, adiciona um número ao nome.
                    while (File.Exists(novoCaminho))
                    {
                        contador++;
                        novoNome = $"{contador}{extensao}";
                        novoCaminho = Path.Combine(diretorio, novoNome);
                    }

                    File.Move(arquivo, novoCaminho);

                    Console.WriteLine($"Renomeado: {arquivo} para {novoNome}");

                    contador++;
                }

                Console.WriteLine("Renomeação concluída.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ocorreu um erro: {ex.Message}");
            }
        }

    }
}
