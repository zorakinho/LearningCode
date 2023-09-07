using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Mensageria
{
    public class TemplateEmail
    {

        public static string GetHtmlTemplate
            (
                string cliente,
                string corretor,
                string gerente,
                string torre,
                string unidade,
                string empreendimento,
                string saudacao
            )
        {
            return $@"
<!DOCTYPE html>
<html lang=""en"">
<head>
<meta charset=""UTF-8"">
<meta name=""viewport"" content=""width=device-width, initial-scale=1.0"">
<title>Template de Email</title>
<style>
  body {{
    margin: 0;
    padding: 0;
    font-family: Roboto;
    color: #333333; /* Cor cinza escura */
    background-color: #f4f4f4;
  }}
  .container {{
    width: 100%;
    max-width: 600px;
    margin: 0 auto;
    background-image: url('https://innovabr.com.br/wp-content/uploads/2022/11/apartamento-edge-cambui-campinas-sp-entrada-2.jpg');
    background-size: cover;
    background-repeat: no-repeat;
    border-radius: 5px;
    box-shadow: 0px 2px 4px rgba(0, 0, 0, 0.8);
    overflow: hidden;
  }}
  .header {{
    text-align: center;
    padding: 10px 20px;
    background-color: rgba(255, 255, 255, 0.8);
    border-radius: 5px 5px 0 0;
  }}
  .logo {{
    max-width: 200px;
    margin: 0 auto; /* Centralize a logo */
  }}
  .Gerente {{
    text-align: center; /* Centralize o texto do gerente */
    color: #2b4272; /* Cor específica */
  }}
  .content {{
    padding: 20px;
    background-color: rgba(255, 255, 255, 0.8);
    border-radius: 5px;
    margin-top: 10px; /* Ajuste conforme necessário */
    margin: 15%;
  }}
  .footer {{
    text-align: center;
    padding: 20px 0;
    font-size: 12px;
    color: #888888;
    background-color: rgba(255, 255, 255, 0.8);
  }}
p {{
    font-size: 15px !important;
  }}
</style>
</head>
<body>
  <div class=""container"">
    <div class=""header"">
      <img src=""https://innovabr.com.br/wp-content/uploads/2021/12/LOGO-PSD-N-Colorido-PARA-FUNDO-BRANCO-1.png"" alt=""innovabr"" class=""logo"">
    </div>
    <div class=""content"">
      <p>Empreendimento: <b>{empreendimento}</b> -  Torre: <b>{torre}</b> | Unidade: <b>{unidade}</b></p>
      <p> {saudacao} Sou o Corretor <b>{corretor}</b>, segue a pasta em <b class=""Gerente"">anexo</b> de nosso cliente <b>{cliente}</b>.</p>
     
    </div>

    <div class=""footer"">
      <p class=""Gerente"" >cc: Gerente Imobiliário <b>{gerente}</b></p>
      <p><b>©InnovaBR. Todos os direitos reservados.</b></p>
    </div>
  </div>
</body>
</html>
";

        }
    }
}
