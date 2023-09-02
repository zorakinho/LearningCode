using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Mensageria.DTO
{
    public class EmailData
    {
        public string Empreendimento { get; set; }
        public string Torre { get; set; }
        public string Unidade { get; set; }
        public string Corretor { get; set; }
        public string Gerente { get; set; }
        public string Cliente { get; set; }
        public string ClienteEmail { get; set; }
        public string EmailsCC { get; set; }
    }
}
