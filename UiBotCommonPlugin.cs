using Newtonsoft.Json.Linq;
using System;
using System.Text.RegularExpressions;

//建议把下面的namespace名字改为您的插件名字
namespace UiBotCommonPlugin
{
 public interface Plugin_Interface
 {   //定义一个插件函数时，必须先在这个interface里面声明
  DateTime GetFileCreationTime(string path);
  DateTime GetFileLastWriteTime(string path);
  void SmtpSendHtmlMail(string Host, int Port, string Account, string Password, string Subject, string Body, string From, string To, string Cc, string Bcc);
 }

 public class Plugin_Implement : Plugin_Interface
 {   //在这里实现插件函数
  public DateTime GetFileCreationTime(string path)
  {
   return System.IO.File.GetCreationTime(path);
  }

  public DateTime GetFileLastWriteTime(string path)
  {
   return System.IO.File.GetLastWriteTime(path);
  }
  public void SmtpSendHtmlMail(string Host, int Port, string Account, string Password, string Subject, string Body, string From, string To, string Cc, string Bcc)
  {
   try
   {
    System.Net.Mail.SmtpClient smtp = new System.Net.Mail.SmtpClient();
    smtp.Host = Host;
    smtp.Port = Port;
    smtp.EnableSsl = true;
    smtp.Credentials = new System.Net.NetworkCredential(Account, Password);
    System.Net.Mail.MailMessage msg = new System.Net.Mail.MailMessage();
    msg.IsBodyHtml = true;
    msg.Body = Body;
    msg.To.Add("sean@everbiz.com.tw");
    msg.From = new System.Net.Mail.MailAddress(From);

    foreach (Match item in new Regex(@"[\w-\.]+@([\w-]+\.)+[\w-]{2,4}", RegexOptions.IgnoreCase).Matches(To)) msg.To.Add(item.Value);
    if(Cc != null)
    foreach (Match item in new Regex(@"[\w-\.]+@([\w-]+\.)+[\w-]{2,4}", RegexOptions.IgnoreCase).Matches(Cc)) msg.CC.Add(item.Value);
    if(Bcc != null)
    foreach (Match item in new Regex(@"[\w-\.]+@([\w-]+\.)+[\w-]{2,4}", RegexOptions.IgnoreCase).Matches(Bcc)) msg.Bcc.Add(item.Value);
    msg.Subject = Subject;
    smtp.Send(msg);
   }
   catch
   {
   }
  }
 }
}
