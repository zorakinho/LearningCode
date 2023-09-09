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

        // ENVIO ASSYNC
        public static async Task EnviarEmailAssync(string remetente, string remetentePassword, string smtpHost, int smtpPort)
        {
            // Licença Excel
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            // Localizando diretórios com os arquivos
            string diretorioDoProjeto = Path.Combine(Directory.GetCurrentDirectory(), "..", "..", "..", "betti");

            // Localizando planilha Excel
            string excelFilePath = Path.Combine(diretorioDoProjeto, "email.xlsx"); 


            
            
            // Assunto e corpo do e-mail
            string assunto = @$"BOT BOB SENDMAIL - {DateTime.Now}";
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
                Console.WriteLine($"FORAM CONTABILIZADAS {rowCount-1} PASTAS PARA O ENVIO:");

                Parallel.For(2, rowCount + 1, row =>
                {

                    string empreendimento = worksheet.Cells[row, 1].Value?.ToString();
                    string torre = worksheet.Cells[row, 2].Value?.ToString();
                    string unidade = worksheet.Cells[row, 3].Value?.ToString();
                    string corretor = worksheet.Cells[row, 4].Value?.ToString();
                    string gerente = worksheet.Cells[row, 5].Value?.ToString();
                    string cliente = worksheet.Cells[row, 6].Value?.ToString();
                    string emailDestino = worksheet.Cells[row, 7].Value?.ToString();
                    string emailsCC = worksheet.Cells[row, 8].Value?.ToString();


                    // Construir o caminho do arquivo com base no número da unidade
                    string filePath = Path.Combine(diretorioDoProjeto, $@"arquivos\{unidade}.pdf");

                   // string arquivo = Path.GetFileNameWithoutExtension(filePath);

                    if (File.Exists(filePath) && Path.GetFileNameWithoutExtension(filePath) == unidade)
                    {
                        try {

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
                            tasks.Add(EnviarEmailAsync(smtpClient, mensagem, filePath, unidade, diretorioDoProjeto, emailsCC));

                            //await Task.Delay(200);

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
            }
            // Aguarde até que todas as tarefas tenham sido concluídas
            await Task.WhenAll(tasks);

            Console.WriteLine("Todos E-mails válidos enviados com sucesso!");
        }



        // RENOMEANDO ARQUIVOS
        public static async Task RenomearArquivoAsync(string filePath, string novoNome, string diretorioDoProjeto, string unidade)
        {
            try
            {
                string novoCaminhoArquivo = Path.Combine(diretorioDoProjeto, "arquivos", novoNome);
                File.Move(filePath, novoCaminhoArquivo);
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
        static async Task EnviarEmailAsync(SmtpClient smtpClient, MailMessage mensagem, string filePath, string unidade, string diretorioDoProjeto, string emailsCC)
        {
            bool renomear;
            try
            {
                mensagem.Attachments.Add(new Attachment(filePath));
                await smtpClient.SendMailAsync(mensagem);
                //Console.WriteLine($"E-mail enviado para: {mensagem.To[0].Address}");
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
            await RenomearArquivoAsync(filePath, $"enviado_{unidade}.pdf", diretorioDoProjeto, unidade);

            }
        }



        //RENOMEANDO ARQUIVOS EM ORDEM NÚMERICA
        public static void RenomearArquivosParaNumerosSequenciais()
        {
            string diretorio = Path.Combine(Directory.GetCurrentDirectory(), "..", "..", "..", "betti", "arquivos");
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
