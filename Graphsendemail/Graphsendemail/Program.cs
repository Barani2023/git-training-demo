using Azure.Identity;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Graph.Users.Item.SendMail;
using System.Security.Cryptography.X509Certificates;



class Program
{
    static async Task Main(string[] args)
    {
        // Tenant ID of the Azuer AD Application
        string tenantId = "5d7022ea-fba6-4cbb-9c99-b9081b8abde4";
        // Client ID of the registered Azure AD Application
        string clientId = "43ba21e8-86f3-4054-8441-76b00beb9105";
        // Local path of the certificate file used for authentication
        string certPath = "C:\\certs\\GraphMailSenderApp.pfx";
        // Password for the certificate file
        string certPassword = "Graph@123";
        // Create certificate object using certificate path and password
        var certificate = new X509Certificate2(certPath, certPassword);

       Console.WriteLine("Certificate Thumbprint: " + certificate.Thumbprint);

       
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
