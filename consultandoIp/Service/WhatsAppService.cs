using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System;
using Twilio;
using Twilio.Rest.Api.V2010.Account;
using Twilio.Types;

namespace consultandoIp.Service
{
    public class WhatsAppService
    {
        private const string AccountSid = "YOUR_TWILIO_ACCOUNT_SID";
        private const string AuthToken = "YOUR_TWILIO_AUTH_TOKEN";
        private const string TwilioNumber = "YOUR_TWILIO_PHONE_NUMBER";
        private const string RecipientNumber = "RECIPIENT_PHONE_NUMBER"; // Número de telefone do destinatário

        public void SendIpViaWhatsApp(string ip)
        {
            TwilioClient.Init(AccountSid, AuthToken);

            var message = MessageResource.Create(
                body: $"Meu IP atual é: {ip}",
                from: new PhoneNumber(TwilioNumber),
                to: new PhoneNumber(RecipientNumber)
            );

            Console.WriteLine($"Mensagem enviada via WhatsApp: SID = {message.Sid}");
        }
    }
}
