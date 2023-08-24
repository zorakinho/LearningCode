using RestSharp;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace consultandoIp.Service
{
    public class IpService
    {
        public IP ConsultarIp()
        {
              
            var client = new RestClient("https://api.ipify.org");
            var request = new RestRequest("?format=json", Method.Get);
            return client.Execute<IP>(request).Data;
            
        }
    }
}
