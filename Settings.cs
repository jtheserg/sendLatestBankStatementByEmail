namespace sendLatestBankStatementByEmail;

internal class Settings
{
    public string BankEmail { get; set; }
    public string PdfPassword { get; set; }
    public string PdfPostFix { get; set; }
    public string MailLogin { get; set; }
    public string MailPassword { get; set; }
    public string MailImapHost { get; set; }
    public int MailImapPort { get; set; }
    public string MailSmtpHost { get; set; }
    public int MailSmtpPort { get; set; }
    public string MailSenderName { get; set; }
    public string MailSenderEmail { get; set; }
    public string MailRecieverName { get; set; }
    public string MailRecieverEmail { get; set; }
    public string MailTextBody { get; set; }
    public string MailSubject { get; set; }
    public bool DebugFakeSend { get; set; }
}