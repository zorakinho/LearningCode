using RestSharp;
using System;

namespace ConsultandoCripto.Service
{
    public class CriptoService
    {
        public CriptoListResponse GetCriptos()
        {
            var client = new RestClient("https://api.coinbase.com");
            var request = new RestRequest("/v2/currencies", Method.Get);
            
            return client.Execute<CriptoListResponse>(request).Data;

        }
    }
}
