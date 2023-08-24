using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsultandoCripto.Service
{
    public class CriptoResponse
    {
        public string? Id { get; set; }

        public string? Name { get; set; }

        public decimal? min_size { get; set; }
 
    }

    public class CriptoListResponse
    {
        public  List<CriptoResponse>? Data { get; set; }
    }
}
