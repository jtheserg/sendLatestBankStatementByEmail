using MailKit;
using MailKit.Net.Imap;
using MailKit.Net.Smtp;
using MailKit.Search;
using MimeKit;
using SmartFormat;
using System.Diagnostics.CodeAnalysis;
using System.Security.Authentication;
using System.Text;
using System.Text.RegularExpressions;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Writer;

namespace sendLatestBankStatementByEmail;

internal static class Program
{
    [RequiresDynamicCode("Calls sendLatestBankStatementByEmail.Startup.Startup()")]
    [RequiresUnreferencedCode("Calls sendLatestBankStatementByEmail.Startup.Startup()")]
    private static void Main(string[] args)
    {
        Console.OutputEncoding = Encoding.UTF8;
        var startup = new Startup();
        if (startup.Settings.MailImapHost is null or "")
        {
            Console.WriteLine("Empty setings. Exiting...");
            throw new Exception("Empty settings");
        }
        Console.WriteLine("Hello, World!");

        var imapClient = ImapClient(startup.Settings.MailImapHost, startup.Settings.MailImapPort, startup.Settings.MailLogin, startup.Settings.MailPassword);
        var inbox = MailFolder(imapClient);
        var listIDs = inbox.Search(SearchQuery.And(SearchQuery.DeliveredAfter(new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1, 0, 0, 0)), SearchQuery.FromContains(startup.Settings.BankEmail)));
        if (listIDs.Count == 0)
        {
            imapClient.Disconnect(true);
            Console.WriteLine("(no match)");
            return;
        }
        var list = new List<MimeMessage>();
        foreach (var item in listIDs)
        {
            try
            {
                list.Add(inbox.GetMessage(item));
            }
            catch (Exception e)
            {
                Console.WriteLine(item.Id + ": " + e.Message);
            }
        }
        foreach (var message in list)
        {
            Console.WriteLine(message.Date + " : " + message.From + " : " + message.Subject);
            foreach (var attachment in message.Attachments)
            {
                var fileName = attachment.ContentDisposition?.FileName ?? attachment.ContentType.Name;
                var filePath = Path.Combine(Directory.GetCurrentDirectory(), "work");
                var fileNamePath = Path.Combine(filePath, fileName);
                if (!fileName.EndsWith(".pdf")) continue;
                var fileNamePathOut = Path.Combine(filePath, fileName.Replace(".pdf", startup.Settings.PdfPostFix + ".pdf"));
                Directory.CreateDirectory(filePath);
                using var stream = File.Create(fileNamePath);
                if (attachment is MessagePart rfc822)
                {
                    rfc822?.Message.WriteTo(stream);
                }
                else
                {
                    var part = (MimePart)attachment;

                    part.Content.DecodeTo(stream);
                }
                stream.Close();
                UnprotectPdf(fileNamePath, fileNamePathOut, startup.Settings.PdfPassword);

                var m = Regex.Match(fileName, @"(?<year>\d*)-(?<month>\d*)(?>.pdf)");
                MailInfo mailInfo = new()
                {
                    Month = new Month(m.Groups["month"].Value),
                    Year = m.Groups["year"].Value,
                    EmailDate = message.Date.ToString("dd.MM.yyyy"),
                };

                var smtpClient = new SmtpClient()
                {
                    SslProtocols = SslProtocols.None,
                    SslCipherSuitesPolicy = null,
                    ClientCertificates = null,
                    CheckCertificateRevocation = false,
                    ServerCertificateValidationCallback = null,
                    LocalEndPoint = null,
                    ProxyClient = null,
                    Capabilities = SmtpCapabilities.None,
                    Timeout = 60000,
                    LocalDomain = null,
                    DeliveryStatusNotificationType = DeliveryStatusNotificationType.Full
                };
                smtpClient.Connect(startup.Settings.MailSmtpHost, startup.Settings.MailSmtpPort);
                smtpClient.Authenticate(startup.Settings.MailLogin, startup.Settings.MailPassword);
                var newMessage = new MimeMessage();
                newMessage.From.Add(new MailboxAddress(startup.Settings.MailSenderName, startup.Settings.MailSenderEmail));
                newMessage.To.Add(new MailboxAddress(startup.Settings.MailRecieverName, startup.Settings.MailRecieverEmail));
                newMessage.Subject = Smart.Format(startup.Settings.MailSubject, mailInfo);

                newMessage.Body = new BodyBuilder
                {
                    TextBody = startup.Settings.MailTextBody,
                    Attachments = { fileNamePathOut }
                }.ToMessageBody();
                SendMail(smtpClient, newMessage, startup.Settings.DebugFakeSend);
                smtpClient.Disconnect(true);
                Console.WriteLine("Email send! Subject=\"" + newMessage.Subject + "\"; File=\"" + fileNamePathOut + "\"");
            }
        }

        imapClient.Disconnect(true);
        Console.WriteLine("Done.");
        Console.ReadLine();
    }

    private static void SendMail(SmtpClient smtpClient, MimeMessage newMessage, bool fake = false)
    {
        if (fake)
        {
            Console.WriteLine("\n(DEBUG) Would send this mail:\n\n" + newMessage + "\n\n<END OF EMAIL>");
        }
        else
        {
            smtpClient.Send(newMessage);
        }
    }

    private static IMailFolder MailFolder(IMailStore client)
    {
        //Select Inbox folder
        var inbox = client.Inbox;
        inbox.Open(FolderAccess.ReadOnly);
        return inbox;
    }

    private static ImapClient ImapClient(string mailHost, int mailPort, string mailUsername, string mailPassword)
    {
        // Authenticate
        var client = new ImapClient();
        client.Connect(mailHost, mailPort);
        client.Authenticate(mailUsername, mailPassword);
        return client;
    }

    private static void UnprotectPdf(string input, string output, string? password = "")
    {
        bool passwordProtected;
        try
        {
            passwordProtected = PdfDocument.Open(input).IsEncrypted;
        }
        catch (UglyToad.PdfPig.Exceptions.PdfDocumentEncryptedException)
        {
            passwordProtected = true;
        }
        catch (Exception e)
        {
            Console.WriteLine(e);
            throw;
        }
        if (passwordProtected)
        {
            var doc = PdfDocument.Open(input, new ParsingOptions { Password = password });
            var newDoc = new PdfDocumentBuilder();
            foreach (var page in doc.GetPages())
            {
                newDoc.AddPage(doc, page.Number);
            }

            var newDocBytes = newDoc.Build();
            File.WriteAllBytes(output, newDocBytes);
            Console.WriteLine("Decrypted PDF: " + input + " ==> " + output);
        }
        else
        {
            // no decryption is required
            File.Copy(input, output, true);
            Console.WriteLine("NOT Decrypted PDF: " + input + " ==> " + output);
        }
    }
}