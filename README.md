# UiBotCommonPlugin

Uibot 內建函數可能不足，此Plugin 提供可能需要的功能

安裝方式：將 /Release/ 底下所有檔案複製到 RPA專案\extend\DotNet

使用方式：

1 查詢檔案建立日期
DateTime = UiBotCommonPlugin.GetFileCreationTime(string path)



2 查詢檔案最後更新日期
DateTime = UiBotCommonPlugin.GetFileLastWriteTime(string path)


3 寄HTML格式的信件
UiBotCommonPlugin.SmtpSendHtmlMail(string Host, int Port, string Account, string Password, string Subject, string Body, string From, string To, string Cc, string Bcc)


4 把PDF 拆開，每頁一個檔案
UiBotCommonPlugin.PdfExtractPages(string sourcePDFpath, string outputPDFpath, int startpage, int endpage)

5 把PDF轉成Excel
UiBotCommonPlugin.PdfToXls(string source, string xlspath)

6 根據文字在PDF內的座標整理出來正確順序的文字內容
UiBotCommonPlugin.TextFromPage(string _filePath, int startPage, int endPage)

7 調整圖片大小
UiBotCommonPlugin.(string filepath, int width, int height)
