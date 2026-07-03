using Azure.Identity;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Graph.Users.Item.SendMail;
using System.Security.Cryptography.X509Certificates;


class Program
{
    static async Task Main(string[] args)
    {
        string tenantId = "5d7022ea-fba6-4cbb-9c99-b9081b8abde4";
        string clientId = "43ba21e8-86f3-4054-8441-76b00beb9105";
        string certPath = "C:\\certs\\GraphMailSenderApp.pfx";   // update your file path
        string certPassword = "Graph@123";//cert password

        var certificate = new X509Certificate2(certPath, certPassword);

       // Console.WriteLine("Certificate Thumbprint: " + certificate.Thumbprint);


        var clientSecretCredential = new ClientCertificateCredential(
            tenantId,
            clientId,
            certificate
        );
        string senderEmail = "e207@minusculetechnologies.com";
        string recipientEmail = "e204@minusculetechnologies.com";

        var graphClient = new GraphServiceClient(clientSecretCredential);

        var email = new Message
        {
            Subject = "Email from App Permissions",
            Body = new ItemBody
            {
                ContentType = BodyType.Text,
                Content = "Hello! Test email"
            },
            ToRecipients = new List<Recipient>
    {
        new Recipient
        {
            EmailAddress = new EmailAddress
            {
                Address = recipientEmail
            }
        }
    }
        };

        var requestBody = new SendMailPostRequestBody
        {
            Message = email,
            SaveToSentItems = true
        };

        await graphClient.Users[senderEmail]
                         .SendMail
                         .PostAsync(requestBody);
        Console.WriteLine("Email sent via App Permissions!");
    }
}
