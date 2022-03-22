
using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;

//建议把下面的namespace名字改为您的插件名字
namespace UiBotCommonPlugin
{
 public interface Plugin_Interface
 {   //定义一个插件函数时，必须先在这个interface里面声明
  string GetFileCreationTime(string path);
  string GetFileLastWriteTime(string path);
  void SmtpSendHtmlMail(string Host, int Port, string Account, string Password, string Subject, string Body, string From, string To, string Cc, string Bcc);

  void PdfExtractPages(string sourcePDFpath, string outputPDFpath, int startpage, int endpage);
  void PdfToXls(string source, string xlspath);

  string TextFromPage(string _filePath, int startPage, int endPage);
  void GrayScaleImage(string filepath);
  void ThumbnailImage(string filepath, int width, int height);
  int[] ImageSize(string filepath);
  string UrlEncode(string data);
 }

 public class Plugin_Implement : Plugin_Interface
 {   //在这里实现插件函数
  public string UrlEncode(string data)
  {
   return System.Web.HttpUtility.UrlEncode(data, Encoding.UTF8);
  }
  public string GetFileCreationTime(string path)
  {
   return System.IO.File.GetCreationTime(path).ToString();
  }

  public string GetFileLastWriteTime(string path)
  {
   return System.IO.File.GetLastWriteTime(path).ToString();
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
    if (Cc != null)
     foreach (Match item in new Regex(@"[\w-\.]+@([\w-]+\.)+[\w-]{2,4}", RegexOptions.IgnoreCase).Matches(Cc)) msg.CC.Add(item.Value);
    if (Bcc != null)
     foreach (Match item in new Regex(@"[\w-\.]+@([\w-]+\.)+[\w-]{2,4}", RegexOptions.IgnoreCase).Matches(Bcc)) msg.Bcc.Add(item.Value);
    msg.Subject = Subject;
    smtp.Send(msg);
   }
   catch
   {
   }
  }
  public void PdfExtractPages(string sourcePDFpath, string outputPDFpath, int startpage, int endpage)
  {
   PdfReader reader = null;
   Document sourceDocument = null;
   PdfCopy pdfCopyProvider = null;
   PdfImportedPage importedPage = null;

   reader = new PdfReader(sourcePDFpath);
   sourceDocument = new Document(reader.GetPageSizeWithRotation(startpage));
   pdfCopyProvider = new PdfCopy(sourceDocument, new System.IO.FileStream(outputPDFpath, System.IO.FileMode.Create));

   sourceDocument.Open();

   for (int i = startpage; i <= endpage; i++)
   {
    importedPage = pdfCopyProvider.GetImportedPage(reader, i);
    pdfCopyProvider.AddPage(importedPage);
   }
   sourceDocument.Close();
   reader.Close();
  }
  public void PdfToXls(string filename, string xlspath)
  {
   Spire.Pdf.PdfDocument doc = new Spire.Pdf.PdfDocument(filename);
   doc.SaveToFile(xlspath, Spire.Pdf.FileFormat.XLSX);
  }
  public string TextFromPage(string _filePath, int startPage, int endPage)
  {
   var pdfReader = new PdfReader(_filePath);


   var locationTextExtractionStrategy = new TextWithFontExtractionStategy();
   string textFromPage = "";
   for (int i = startPage; i <= (endPage != -1 ? endPage : startPage); i++)
    textFromPage += PdfTextExtractor.GetTextFromPage(pdfReader, i, locationTextExtractionStrategy);

   return Encoding.UTF8.GetString(Encoding.Convert(Encoding.Default, Encoding.UTF8, Encoding.Default.GetBytes(textFromPage)));


  }


  public void ThumbnailImage(string filepath, int width, int height)
  {
   MemoryStream stream = new MemoryStream(File.ReadAllBytes(filepath));
   System.Drawing.Image img = System.Drawing.Image.FromStream(stream);
   Bitmap img1 = new Bitmap(width, height);
   Graphics g = Graphics.FromImage(img1);
   g.CompositingQuality = System.Drawing.Drawing2D.CompositingQuality.HighQuality;
   g.CompositingMode = System.Drawing.Drawing2D.CompositingMode.SourceCopy;
   g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
   g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;
   g.DrawImage(img, new System.Drawing.Rectangle(0, 0, img1.Width, img1.Height), new System.Drawing.Rectangle(0, 0, img.Width, img.Height), GraphicsUnit.Pixel);
   g.Dispose();
   img.Dispose();
   stream.Close();
   img1.Save(filepath, filepath.EndsWith(".png", StringComparison.InvariantCultureIgnoreCase) ? ImageFormat.Png : ImageFormat.Jpeg);
  }


  public void GrayScaleImage(string filepath)
  {
   MemoryStream stream = new MemoryStream(File.ReadAllBytes(filepath));
   System.Drawing.Bitmap img = (System.Drawing.Bitmap ) System.Drawing.Image.FromStream(stream);
   /*
   Bitmap img1 = new Bitmap(img.Width, img.Height);
   //get a graphics object from the new image
   Graphics g = Graphics.FromImage(img);
   //create the grayscale ColorMatrix
   ColorMatrix colorMatrix = new ColorMatrix(
      new float[][]
     {
                 new float[] {.3f, .3f, .3f, 0, 0},
                 new float[] {.59f, .59f, .59f, 0, 0},
                 new float[] {.11f, .11f, .11f, 0, 0},
                 new float[] {0, 0, 0, 1, 0},
                 new float[] {0, 0, 0, 0, 1}
     });
   //create some image attributes
   ImageAttributes attributes = new ImageAttributes();
   //set the color matrix attribute
   attributes.SetColorMatrix(colorMatrix);
   //draw the original image on the new image
   //using the grayscale color matrix
   g.DrawImage(img, new System.Drawing.Rectangle(0, 0, img.Width, img.Height),
      0, 0, img.Width, img.Height, GraphicsUnit.Pixel, attributes);
   //dispose the Graphics object
   g.Dispose();
   */
   int rgb;
   Color c;
   for (int y = 0; y < img.Height; y++)
    for (int x = 0; x < img.Width; x++)
    {
     c = img.GetPixel(x, y);
     img.SetPixel(x, y, Color.FromArgb(c.R >> 1, c.G >> 1, c.B >> 1));
    }
   img.Save(filepath, filepath.EndsWith(".png", StringComparison.InvariantCultureIgnoreCase) ? ImageFormat.Png : ImageFormat.Jpeg);
  }


  public int[] ImageSize(string filepath)
  {
   System.Drawing.Image img = System.Drawing.Image.FromFile(filepath);
   int[] size = new int[] { img.Size.Width, img.Size.Height };
   img.Dispose();
   return size;
  }
  public class TextWithFontExtractionStategy : iTextSharp.text.pdf.parser.ITextExtractionStrategy
  {
   //HTML buffer
   private StringBuilder result = new StringBuilder();
   // 用來存放文字的矩形
   List<System.util.RectangleJ> rectText = new List<System.util.RectangleJ>();

   // 用來存放文字
   List<String> textList = new List<String>();

   // 用來存放文字的Y座標
   List<float> listY = new List<float>();
   List<float> listX = new List<float>();

   // 用來存放每一行文字的座標位置
   List<Dictionary<String, System.util.RectangleJ>> row_text_rect = new List<Dictionary<String, System.util.RectangleJ>>();

   Dictionary<String, String> text_position = new Dictionary<String, String>();
   // 圖片座標
   List<float[]> arrays = new List<float[]>();

   // 圖片
   List<byte[]> arraysByte = new List<byte[]>();

   //http://api.itextpdf.com/itext/com/itextpdf/text/pdf/parser/TextRenderInfo.html
   private enum TextRenderMode
   {
    FillText = 0,
    StrokeText = 1,
    FillThenStrokeText = 2,
    Invisible = 3,
    FillTextAndAddToPathForClipping = 4,
    StrokeTextAndAddToPathForClipping = 5,
    FillThenStrokeTextAndAddToPathForClipping = 6,
    AddTextToPaddForClipping = 7
   }



   public void RenderText(iTextSharp.text.pdf.parser.TextRenderInfo renderInfo)
   {
    String text = renderInfo.GetText().Trim();

    if (text.Length > 0)
    {
     System.util.RectangleJ rectBase = renderInfo.GetBaseline().GetBoundingRectange();
     // 獲取文字下面的矩形
     System.util.RectangleJ rectAscen = renderInfo.GetAscentLine().GetBoundingRectange();
     // 計算出文字的邊框矩形
     //float leftX = (float)rectBase.X;
     //float leftY = (float)(rectBase.Y - 1);
     //float rightX = (float)rectBase.Width;
     //float rightY = (float)(rectBase.Height - 1);
     //Rectangle r = rectBase.GetBounds();
     //// System.out.println("float:" + leftX + ":" + leftY + ":" + rightX
     //// + ":" + rightY);
     //Rectangle rect = new Rectangle(rectBase.X, leftY, rightX - leftX, rightY - leftY);
     //// System.out.println("text:" + text + "X:" + rect.x + "Y:" + rect.y
     //// + "width:" + rect.width + "height:"
     //// + rect.height);
     if (listY.Contains(rectBase.Y))
     {
      int index = listY.IndexOf(rectBase.Y);
      float tempx = rectBase.X > rectText[index].X ? rectText[index].X : rectBase.X;
      rectText[index] = new System.util.RectangleJ(tempx, rectBase.Y, rectBase.Width + rectText[index].Width, rectBase.Height);
      textList[index] = textList[index] + text;
     }
     else
     {
      rectText.Add(rectBase);
      textList.Add(text);
      listY.Add(rectBase.Y);
     }
     if (!listX.Contains(rectBase.X))
     {
      listX.Add(rectBase.X);
     }
     text_position[rectBase.X + "," + rectBase.Y] = text;


     Dictionary<String, System.util.RectangleJ> map = new Dictionary<String, System.util.RectangleJ>();
     map[text] = rectBase;
     row_text_rect.Add(map);
    }
   }

   public string GetResultantText()
   {
    string result = "";
    listY.Sort();
    listX.Sort();
    for (int j = listY.Count - 1; j >= 0; j--)
    {
     string line = "";
     bool first = true;
     float lastLeft = 0;
     for (int i = 0; i < listX.Count; i++)
     {
      if (text_position.ContainsKey(listX[i] + "," + listY[j]))
      {
       if (first)
       {
        for (int k = 0; k <= listX[i]; k += 20) line += " ";
       }
       else
       {
        for (float k = lastLeft; k <= listX[i] - 20; k += 20) line += " ";
       }
       line += text_position[listX[i] + "," + listY[j]];
       first = false;
       lastLeft = listX[i];
      }
     }
     if (line != "") result += line + "\n";
    }
    return result;
   }

   //Not needed
   public void BeginTextBlock() { }
   public void EndTextBlock() { }
   public void RenderImage(ImageRenderInfo renderInfo) { }
  }
 }

}