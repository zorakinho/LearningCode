using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace consultandoIp.Service
{
    public class ArquivoService
    {
        public void addLog(string ip)
        {
            string logtxt = DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss") + " - " + ip;

            Console.WriteLine("log adicionado: "+logtxt);

            System.IO.File.AppendAllText("log_ip.txt", logtxt + Environment.NewLine);
        }
    }
}
