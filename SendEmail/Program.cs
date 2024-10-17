using System.Net;
using System.Net.Mail;
using Microsoft.Extensions.Configuration;

namespace SendEmail;

class Program
{
    static void Main(string[] args)
    {
        var builder = new ConfigurationBuilder()
            .SetBasePath(basePath: Directory.GetCurrentDirectory())
            .AddJsonFile(path: "appsettings.json", optional: true, reloadOnChange: true);

        IConfiguration configuration = builder.Build();

        string? mailSender = configuration.GetValue<string>("MailSender:mail");
        string? mailReceiver = configuration.GetValue<string>("MailReceiver:mail");
        string? mailSubject = configuration.GetValue<string>("MailSubject");
        string? mailBody = configuration.GetValue<string>("MailBody");
        string? passwordSender = configuration.GetValue<string>("MailSender:password");
        string? attachmentPath = configuration.GetValue<string>("AttachmentPath");

        Attachment attachment = new Attachment(fileName: attachmentPath);

        MailMessage mailMessage = new();

        mailMessage.From = new MailAddress(address: mailSender);
        mailMessage.To.Add(addresses: mailReceiver);
        mailMessage.Subject = mailSubject;
        mailMessage.Body = mailBody;
        mailMessage.Attachments.Add(attachment);

        SmtpClient smtpClient = new();

        smtpClient.Host = "smtp.office365.com";
        smtpClient.Port = 587;
        smtpClient.UseDefaultCredentials = false;
        smtpClient.Credentials = new NetworkCredential(userName: mailSender, password: passwordSender);
        smtpClient.EnableSsl = true;

        try
        {
            smtpClient.Send(mailMessage);
            WriteLine("Email sent successfully");
        }
        catch (Exception ex)
        {
            WriteLine($"Error: {ex.Message}");
        }
    }
}
