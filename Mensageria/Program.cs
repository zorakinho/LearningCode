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

            string senderEmail = "robert.alves@olxbr.com"; // Substitua pelo seu e-mail
            string senderPassword = "ra02xbo0$TRK"; // Substitua pela sua senha

            //EmailService.EnviarEmailsEmParalelo(senderEmail, senderPassword).Wait();

            EmailService.EmailBot(senderEmail, senderPassword);
        }
    }
}
