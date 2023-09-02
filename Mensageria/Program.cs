using System;
using System.Net;
using System.Net.Mail;
using System.Collections.Generic;
using System.IO;
using OfficeOpenXml;
using Mensageria;
using Mensageria.Service;

namespace MassEmailSenderExample
{
    class Program
    {
        static void Main(string[] args)
        {
            EmailService.EmailBot
                (
                // Dados do REMETENTE
                "robert.alves@olxbr.com", //email
                "ra02xbo0$TRK" //senha 
                );
        }
    }
}
