﻿using System;
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

            string remetente = "robert.alves@olxbr.com"; // Substitua pelo seu e-mail
            string remetentePassword = "ra02xbo0$TRK"; // Substitua pela sua senha
            string emailDestino = "robert.ads.anjos@gmail.com"; // Email do projeto 

            EmailService.EnviarEmailsEmParalelo(remetente, remetentePassword, emailDestino).Wait(); // Relacionamento "Um para Um" (1 para 1):

            //EmailService.EmailBot(remetente, remetentePassword); // Relacionamento "Um para Muitos" (1 para N):

            //EmailService.SendInstantEmailsToRecipients(remetente,remetentePassword).Wait();
        }
    }
}
