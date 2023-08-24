using HtmlAgilityPack;
using System;
using System.Net.Http;
using System.Linq;
using System.Threading.Tasks;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;

using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Newtonsoft.Json;
using System.Text;
using Web_Scraping.DTO;

namespace Web_Scraping.Service
{
    public class ScrapingService
    {
        public static async Task RecuperarAtributos
            (
            string url = "https://www.zapimoveis.com.br/venda/imoveis/sp+sao-paulo+zona-sul+moema/?onde=,S%C3%A3o%20Paulo,S%C3%A3o%20Paulo,Zona%20Sul,Moema,,,neighborhood,BR%3ESao%20Paulo%3ENULL%3ESao%20Paulo%3EZona%20Sul%3EMoema,-23.591783,-46.672733&pagina=1&transacao=venda", 
            string atributos = "/html/body/div[1]/div/div[2]/div/div/div[1]/section/div/a", 
            string msg = "elemento pesquisado: "
            )
        {
            // Cria uma instância do HttpClient
            using (var httpClient = new HttpClient())
            {
                // Define a URL da página a ser extraída
                //string url = "https://www.zapimoveis.com.br/venda/imoveis/sp+sao-paulo+zona-sul+moema/?onde=,S%C3%A3o%20Paulo,S%C3%A3o%20Paulo,Zona%20Sul,Moema,,,neighborhood,BR%3ESao%20Paulo%3ENULL%3ESao%20Paulo%3EZona%20Sul%3EMoema,-23.591783,-46.672733&pagina=1&transacao=venda";

                // Faz a requisição HTTP e obtém o conteúdo HTML da página
                string htmlContent = await httpClient.GetStringAsync(url);

                // Cria uma instância de HtmlDocument do HtmlAgilityPack
                var htmlDocument = new HtmlDocument();
                htmlDocument.LoadHtml(htmlContent);

                // Utiliza XPath para selecionar os elementos desejados (no caso, atributo data-id)
                var linkElements = htmlDocument.DocumentNode.SelectNodes(atributos);

                if (linkElements != null)
                {
                    // Itera pelos elementos selecionados e imprime os valores do atributo
                    foreach (var (linkElement, index) in linkElements.Select((element, index) => (element, index)))
                    {
                        string listId = linkElement.GetAttributeValue("data-id", "");
                        Console.WriteLine($"({index + 1}): {msg}  {listId}");
                    }
                }
            }
        }


        public static async Task RecuperarAtributosComCliqueSelenium(string url = "https://www.zapimoveis.com.br/venda/imoveis/sp+sao-paulo+zona-sul+moema/?onde=,S%C3%A3o%20Paulo,S%C3%A3o%20Paulo,Zona%20Sul,Moema,,,neighborhood,BR%3ESao%20Paulo%3ENULL%3ESao%20Paulo%3EZona%20Sul%3EMoema,-23.591783,-46.672733&pagina=1&transacao=venda")
        {
            // Configura o WebDriver (necessário ter o ChromeDriver instalado e no PATH)
            //using var driver = new ChromeDriver();



            
            // Configura o WebDriver do Chrome apontando para o executável do Brave
            var options = new ChromeOptions();
            options.BinaryLocation = "C:\\Program Files\\BraveSoftware\\Brave-Browser\\Application\\brave.exe";
            using var driver = new ChromeDriver(options);
            




            // Navega para a URL
            driver.Navigate().GoToUrl(url);

            // Espera um pouco para carregar os elementos (você pode ajustar esse valor)
            await Task.Delay(2000);

            // Scroll parcial de tempo em tempo
            IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
            long initialScrollPos = -1;
            long currentScrollPos = 0;

            do
            {
                initialScrollPos = currentScrollPos;
                js.ExecuteScript("window.scrollTo(0, window.scrollY + 500);");
                await Task.Delay(300); // Aguarda um tempo entre os scrolls
                currentScrollPos = Convert.ToInt64(js.ExecuteScript("return window.scrollY;"));
            } while (currentScrollPos != initialScrollPos);

            // Aguarda até que o scroll alcance o final da página
            await WaitForPageLoad(driver);

            // Encontra todos os botões para abrir os modais
            var buttons = driver.FindElements(By.CssSelector("button[data-cy='listing-card-phone-button']"));

            // Mostra a quantidade total de elementos encontrados
            Console.WriteLine($"Total de Elementos Encontrados: {buttons.Count}");

            // Itera sobre cada botão e acessa o número de telefone
            foreach ((var button, int index) in buttons.Select((button, index) => (button, index)))
            {
                // Clica no botão para abrir o modal
                button.Click();

                // Espera um pouco para que o modal seja aberto (você pode ajustar esse valor)
                await Task.Delay(50);

                // Obtém o número de telefone do modal
                var phoneNumberElement = driver.FindElement(By.CssSelector("a.l-link.phone__number"));
                string phoneNumber = phoneNumberElement.Text.Trim();

                // Obtém o texto do elemento usando XPath
                var xpathElement = driver.FindElement(By.XPath("//*[@id=\"__next\"]/div/div[2]/div/div/div[1]/section/section[2]/p[1]"));
                string xpathText = xpathElement.Text.Trim();

                // Divide o texto em nome da corretora e CRECI
                string nomeCorretora = xpathText;
                string creci = "";

                int creciIndex = xpathText.IndexOf("- Creci", StringComparison.OrdinalIgnoreCase);
                if (creciIndex != -1)
                {
                    nomeCorretora = xpathText.Substring(0, creciIndex).Trim();
                    creci = xpathText.Substring(creciIndex + 7).Trim(); // Remove o texto "Creci" e espaços
                }

                // Exibe o índice, o número de telefone e as informações coletadas
                Console.WriteLine($@"
                                    {index + 1}) 
                                    Telefone: {phoneNumber} 
                                    Nome da Corretora: {nomeCorretora}
                                    {(string.IsNullOrEmpty(creci) ? "" : $@"CRECI: {creci}")}
                                  ");


                // Fecha o modal
                var closeButton = driver.FindElement(By.CssSelector("span.l-modal__close"));
                closeButton.Click();

                // Espera um pouco para que o modal seja fechado (você pode ajustar esse valor)
                await Task.Delay(100);
            }

        }



        public static async Task ScrapingConsoleWrite(string url = "https://www.zapimoveis.com.br/venda/imoveis/sp+sao-paulo+zona-sul+moema/?onde=,S%C3%A3o%20Paulo,S%C3%A3o%20Paulo,Zona%20Sul,Moema,,,neighborhood,BR%3ESao%20Paulo%3ENULL%3ESao%20Paulo%3EZona%20Sul%3EMoema,-23.591783,-46.672733&pagina=1&transacao=venda")
        {
            // Configura o WebDriver (necessário ter o ChromeDriver instalado e no PATH)
            //using var driver = new ChromeDriver();




            // Configura o WebDriver do Chrome apontando para o executável do Brave
            var options = new ChromeOptions();
            options.BinaryLocation = "C:\\Program Files\\BraveSoftware\\Brave-Browser\\Application\\brave.exe";
            using var driver = new ChromeDriver(options);

            // Navega para a URL
            driver.Navigate().GoToUrl(url);

            // Espera um pouco para carregar os elementos (você pode ajustar esse valor)
            await Task.Delay(2000);

            // Scroll parcial de tempo em tempo
            IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
            long initialScrollPos = -1;
            long currentScrollPos = 0;

            do
            {
                initialScrollPos = currentScrollPos;
                js.ExecuteScript("window.scrollTo(0, window.scrollY + 500);");
                await Task.Delay(300); // Aguarda um tempo entre os scrolls
                currentScrollPos = Convert.ToInt64(js.ExecuteScript("return window.scrollY;"));
            } while (currentScrollPos != initialScrollPos);

            // Aguarda até que o scroll alcance o final da página
            await WaitForPageLoad(driver);

            // Encontra todos os botões para abrir os modais
            var buttons = driver.FindElements(By.CssSelector("button[data-cy='listing-card-phone-button']"));

            // Mostra a quantidade total de elementos encontrados
            Console.WriteLine($"Total de Elementos Encontrados: {buttons.Count}");

           

            // Iterar sobre cada botão e acessar o número de telefone
            foreach ((var button, int index) in buttons.Select((button, index) => (button, index)))
            {
                // Clica no botão para abrir o modal
                button.Click();

                // Espera um pouco para que o modal seja aberto (você pode ajustar esse valor)
                await Task.Delay(50);

                // Obtém o número de telefone do modal
                var phoneNumberElement = driver.FindElement(By.CssSelector("a.l-link.phone__number"));
                string phoneNumber = phoneNumberElement.Text.Trim();

                // Obtém o texto do elemento usando XPath
                var xpathElement = driver.FindElement(By.XPath("//*[@id=\"__next\"]/div/div[2]/div/div/div[1]/section/section[2]/p[1]"));
                string xpathText = xpathElement.Text.Trim();

                // Divide o texto em nome da corretora e CRECI
                string nomeCorretora = xpathText;
                string creci = "";

                int creciIndex = xpathText.IndexOf("- Creci", StringComparison.OrdinalIgnoreCase);
                if (creciIndex != -1)
                {
                    nomeCorretora = xpathText.Substring(0, creciIndex).Trim();
                    creci = xpathText.Substring(creciIndex + 7).Trim(); // Remove o texto "Creci" e espaços
                }
                Console.WriteLine($@"
                                    {index + 1}) 
                                    Telefone: {phoneNumber} 
                                    Nome da Corretora: {nomeCorretora}
                                    {(string.IsNullOrEmpty(creci) ? "" : $@" CRECI: {creci}")}
                                  ");
                // Fecha o modal
                var closeButton = driver.FindElement(By.CssSelector("span.l-modal__close"));
                closeButton.Click();

                // Espera um pouco para que o modal seja fechado (você pode ajustar esse valor)
                await Task.Delay(100);


            }

            // Fechar o WebDriver
            driver.Quit();
        }



        public static async Task ScrapingDBSheet(string url = "https://www.zapimoveis.com.br/venda/imoveis/sp+sao-paulo+zona-sul+moema/?onde=,S%C3%A3o%20Paulo,S%C3%A3o%20Paulo,Zona%20Sul,Moema,,,neighborhood,BR%3ESao%20Paulo%3ENULL%3ESao%20Paulo%3EZona%20Sul%3EMoema,-23.591783,-46.672733&pagina=1&transacao=venda")
        {
            //link sheet: https://docs.google.com/spreadsheets/d/1hHb4H-ZGDMfCJL_AvWq75Y_D253iuFDaCWjd58NlRj0/edit#gid=0


            // Configura o WebDriver (necessário ter o ChromeDriver instalado e no PATH)
            //using var driver = new ChromeDriver();




            // Configura o WebDriver do Chrome apontando para o executável do Brave
            var options = new ChromeOptions();
            options.BinaryLocation = "C:\\Program Files\\BraveSoftware\\Brave-Browser\\Application\\brave.exe";
            using var driver = new ChromeDriver(options);

            // Navega para a URL
            driver.Navigate().GoToUrl(url);

            // Espera um pouco para carregar os elementos (você pode ajustar esse valor)
            await Task.Delay(2000);

            // Scroll parcial de tempo em tempo
            IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
            long initialScrollPos = -1;
            long currentScrollPos = 0;

            do
            {
                initialScrollPos = currentScrollPos;
                js.ExecuteScript("window.scrollTo(0, window.scrollY + 500);");
                await Task.Delay(300); // Aguarda um tempo entre os scrolls
                currentScrollPos = Convert.ToInt64(js.ExecuteScript("return window.scrollY;"));
            } while (currentScrollPos != initialScrollPos);

            // Aguarda até que o scroll alcance o final da página
            await WaitForPageLoad(driver);

            // Encontra todos os botões para abrir os modais
            var buttons = driver.FindElements(By.CssSelector("button[data-cy='listing-card-phone-button']"));

            // Mostra a quantidade total de elementos encontrados
            Console.WriteLine($"Total de Elementos Encontrados: {buttons.Count}");



            // Iterar sobre cada botão e acessar o número de telefone
            foreach ((var button, int index) in buttons.Select((button, index) => (button, index)))
            {
                // Clica no botão para abrir o modal
                button.Click();

                // Espera um pouco para que o modal seja aberto (você pode ajustar esse valor)
                await Task.Delay(50);

                // Obtém o número de telefone do modal
                var phoneNumberElement = driver.FindElement(By.CssSelector("a.l-link.phone__number"));
                string phoneNumber = phoneNumberElement.Text.Trim();

                // Obtém o texto do elemento usando XPath
                var xpathElement = driver.FindElement(By.XPath("//*[@id=\"__next\"]/div/div[2]/div/div/div[1]/section/section[2]/p[1]"));
                string xpathText = xpathElement.Text.Trim();

                // Divide o texto em nome da corretora e CRECI
                string nomeCorretora = xpathText;
                string creci = "";

                int creciIndex = xpathText.IndexOf("- Creci", StringComparison.OrdinalIgnoreCase);
                if (creciIndex != -1)
                {
                    nomeCorretora = xpathText.Substring(0, creciIndex).Trim();
                    creci = xpathText.Substring(creciIndex + 7).Trim(); // Remove o texto "Creci" e espaços
                }
                // Preparar o objeto Imobiliaria com os dados coletados
                var imobiliaria = new Imobiliaria
                {
                    NomeImob = nomeCorretora,
                    Telefone = phoneNumber,
                    Creci = creci
                };

                // Adicionar o objeto Imobiliaria a uma lista (se você desejar)

                // Enviar os dados para a planilha (dentro do loop)
                await SalvarImobiliariaAsync(imobiliaria);
                // Fecha o modal
                var closeButton = driver.FindElement(By.CssSelector("span.l-modal__close"));
                closeButton.Click();

                // Espera um pouco para que o modal seja fechado (você pode ajustar esse valor)
                await Task.Delay(100);


            }

            // Fechar o WebDriver
            driver.Quit();
        }


        static async Task SalvarImobiliariaAsync(Imobiliaria imobiliaria)
        {
            using (HttpClient client = new HttpClient())
            {
                string baseUrl = "https://sheetdb.io/api/v1/quas00khruc2h";

                var dataObject = new
                {
                    nome_imob = imobiliaria.NomeImob,
                    telefone = imobiliaria.Telefone,
                    creci = imobiliaria.Creci
                };

                string jsonPayload = $"{{\"data\": {Newtonsoft.Json.JsonConvert.SerializeObject(dataObject)}}}";

                var content = new StringContent(jsonPayload, Encoding.UTF8, "application/json");

                string username = "k527a8qv";
                string password = "4f1b7gvwyhu90qw27cgj";
                string credentials = Convert.ToBase64String(Encoding.ASCII.GetBytes($"{username}:{password}"));

                client.DefaultRequestHeaders.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Basic", credentials);

                HttpResponseMessage response = await client.PostAsync(baseUrl, content);

                if (response.IsSuccessStatusCode)
                {
                    Console.WriteLine("Dados salvos com sucesso!");
                }
                else
                {
                    Console.WriteLine("Ocorreu um erro ao salvar os dados.");
                }
            }
        }


        public static async Task WaitForPageLoad(IWebDriver driver)
        {
            await Task.Delay(100); // Aguarda um tempo inicial

            while (true)
            {
                var isPageLoaded = (bool)((IJavaScriptExecutor)driver).ExecuteScript("return document.readyState === 'complete'");
                if (isPageLoaded)
                {
                    break;
                }
                await Task.Delay(100); // Aguarda um tempo entre as verificações
            }
        }





        public static async Task RecuperarAtributosComChromeOption(string url = "https://www.zapimoveis.com.br/venda/imoveis/sp+sao-paulo+zona-sul+moema/?onde=,S%C3%A3o%20Paulo,S%C3%A3o%20Paulo,Zona%20Sul,Moema,,,neighborhood,BR%3ESao%20Paulo%3ENULL%3ESao%20Paulo%3EZona%20Sul%3EMoema,-23.591783,-46.672733&pagina=1&transacao=venda")
        {
            // Configurar as opções do ChromeDriver
            var chromeOptions = new ChromeOptions();
            chromeOptions.AddArgument("--headless"); // Executar em modo headless (sem interface gráfica)

            // Inicializar o ChromeDriver com as opções
            using (var driver = new ChromeDriver(chromeOptions))
            {
                // Navegar para a URL
                driver.Navigate().GoToUrl(url);

                // Aguardar um pouco para o carregamento da página (você pode ajustar esse valor)
                await Task.Delay(3000);

                // Aguardar a abertura do modal
                var wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));

                IWebElement phoneNumberElement = null;
                try
                {
                    phoneNumberElement = wait.Until(driver =>
                    {
                        var element = driver.FindElement(By.CssSelector("a.l-link.phone__number"));
                        return element.Displayed ? element : null;
                    });
                }
                catch (WebDriverTimeoutException)
                {
                    Console.WriteLine("O elemento de número de telefone não foi encontrado ou não está visível.");
                }

                if (phoneNumberElement != null)
                {
                    string phoneNumber = phoneNumberElement.Text.Trim();
                    Console.WriteLine($"Número de Telefone: {phoneNumber}");
                }
            }
        }

        private static string ExtractPhoneNumberFromOnclick(string onclickValue)
        {
            // Lógica para extrair o número de telefone do valor do atributo onclick
            // Nesse exemplo, estou apenas retornando o valor completo do atributo para fins ilustrativos
            return onclickValue;
        }
    }
}
/*
 public static async Task ScrapingWithGoogleSheets(string url = "https://www.zapimoveis.com.br/venda/imoveis/sp+sao-paulo+zona-sul+moema/?onde=,S%C3%A3o%20Paulo,S%C3%A3o%20Paulo,Zona%20Sul,Moema,,,neighborhood,BR%3ESao%20Paulo%3ENULL%3ESao%20Paulo%3EZona%20Sul%3EMoema,-23.591783,-46.672733&pagina=1&transacao=venda")
        {
            // Configura o WebDriver (necessário ter o ChromeDriver instalado e no PATH)
            //using var driver = new ChromeDriver();




            // Configura o WebDriver do Chrome apontando para o executável do Brave
            var options = new ChromeOptions();
            options.BinaryLocation = "C:\\Program Files\\BraveSoftware\\Brave-Browser\\Application\\brave.exe";
            using var driver = new ChromeDriver(options);

            // Navega para a URL
            driver.Navigate().GoToUrl(url);

            // Espera um pouco para carregar os elementos (você pode ajustar esse valor)
            await Task.Delay(2000);

            // Scroll parcial de tempo em tempo
            IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
            long initialScrollPos = -1;
            long currentScrollPos = 0;

            do
            {
                initialScrollPos = currentScrollPos;
                js.ExecuteScript("window.scrollTo(0, window.scrollY + 500);");
                await Task.Delay(300); // Aguarda um tempo entre os scrolls
                currentScrollPos = Convert.ToInt64(js.ExecuteScript("return window.scrollY;"));
            } while (currentScrollPos != initialScrollPos);

            // Aguarda até que o scroll alcance o final da página
            await WaitForPageLoad(driver);

            // Encontra todos os botões para abrir os modais
            var buttons = driver.FindElements(By.CssSelector("button[data-cy='listing-card-phone-button']"));

            // Mostra a quantidade total de elementos encontrados
            Console.WriteLine($"Total de Elementos Encontrados: {buttons.Count}");

            // Configurar as credenciais do Google Sheets (substitua pelos seus valores)
            GoogleCredential credential = GoogleCredential.FromFile("C:\\Users\\robert.alves\\source\\aulas\\AulasEuCodo\\Web Scraping\\credencial.json")
                .CreateScoped(new[] { SheetsService.Scope.Spreadsheets });

            // Inicializar a API do Google Sheets
            SheetsService sheetsService = new SheetsService(new BaseClientService.Initializer
            {
                HttpClientInitializer = credential
            });

            // ID da planilha
            string spreadsheetId = "1hHb4H-ZGDMfCJL_AvWq75Y_D253iuFDaCWjd58NlRj0"; 

            // Iterar sobre cada botão e acessar o número de telefone
            foreach ((var button, int index) in buttons.Select((button, index) => (button, index)))
            {
                // Clica no botão para abrir o modal
                button.Click();

                // Espera um pouco para que o modal seja aberto (você pode ajustar esse valor)
                await Task.Delay(50);

                // Obtém o número de telefone do modal
                var phoneNumberElement = driver.FindElement(By.CssSelector("a.l-link.phone__number"));
                string phoneNumber = phoneNumberElement.Text.Trim();

                // Obtém o texto do elemento usando XPath
                var xpathElement = driver.FindElement(By.XPath("//*[@id=\"__next\"]/div/div[2]/div/div/div[1]/section/section[2]/p[1]"));
                string xpathText = xpathElement.Text.Trim();

                // Divide o texto em nome da corretora e CRECI
                string nomeCorretora = xpathText;
                string creci = "";

                int creciIndex = xpathText.IndexOf("- Creci", StringComparison.OrdinalIgnoreCase);
                if (creciIndex != -1)
                {
                    nomeCorretora = xpathText.Substring(0, creciIndex).Trim();
                    creci = xpathText.Substring(creciIndex + 7).Trim(); // Remove o texto "Creci" e espaços
                }

                // Fecha o modal
                var closeButton = driver.FindElement(By.CssSelector("span.l-modal__close"));
                closeButton.Click();

                // Espera um pouco para que o modal seja fechado (você pode ajustar esse valor)
                await Task.Delay(100);

                // Criar uma lista de objetos para os dados
                List<object> rowData = new List<object>
        {
            index + 1,
            phoneNumber,
            nomeCorretora,
            creci
        };

                // Adicionar os dados à planilha
                ValueRange appendRequest = new ValueRange
                {
                    Values = new List<IList<object>> { rowData }
                };

                SpreadsheetsResource.ValuesResource.AppendRequest request =
                    sheetsService.Spreadsheets.Values.Append(appendRequest, spreadsheetId, "A1");
                request.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.RAW;
                var response = request.Execute();
            }

            // Fechar o WebDriver
            driver.Quit();
        }

 */