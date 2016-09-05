# Microsoft Office 使用說明

## 目錄

- 使用前注意事項
- 安裝 nuget 套件
- 開始開發
- 備註
- Libre Office SDK 起源

## 使用前注意事項

- 此套件使用 Microsoft Office 作為轉檔工具，因此若伺服器**無法安裝任何本版本的 Office 就可以跳過此套件**了
- Microsoft Office 於 2007 版本以後有內建將文件轉換為 Pdf 之功能，因此 Office 版本必須**大於或等於 2007** (2007尚未測試過)

## 安裝 nuget 套件

You can either <a href="https://github.com/lettucebo/Creatidea.Library.Office.git">download</a> the source and build your own dll or, if you have the NuGet package manager installed, you can grab them automatically.

```
PM> Install-Package Creatidea.Library.MsOffice
```

Once you have the libraries properly referenced in your project, you can include calls to them in your code. 
For a sample implementation, check the [Example](https://github.com/lettucebo/Creatidea.Library.Office/tree/master/Creatidea.Library.Office.Example) folder.

Add the following namespaces to use the library:
```csharp
using Creatidea.Library.Office.MsOffice;
```

## 環境設定
根據不同的 Office 版本要參考不同版本號的 dll

- Office 2016 的版本號為`16` 
  - 加入 COM 類別的 Microsoft Office 16.0 Object Library
    - ![ 加入 COM 類別](http://i.imgur.com/A1fVGzK.png)
  - 加入組件參考
    - Microsoft.Office.Interrop.Excel
    - Microsoft.Office.Interrop.PowerPoint
    - Microsoft.Office.Interrop.Word
      - ![加入組件參考](http://i.imgur.com/r1MpXjY.png)
- Office 2013的版本號為`15` 
- Office 2010的版本號為`14` 
- Office 2007的版本號為`12` 

## 開始開發

#### Word To Pdf

直接呼叫 **LibreOffice.OfficeConverter.WordToPdf(filePath)** 傳入 Word 檔案完整路徑即可
可傳入`*.doc` 或 `*.docx`
```csharp
var docResult = MsOffice.OfficeConverter.WordToPdf(docPath);
var link = SaveFile(docResult, "msdoc.pdf");
Console.WriteLine("Show docResult: {0}", link);
```
```csharp
var docxResult = MsOffice.OfficeConverter.WordToPdf(docxPath);
var linkdocx = SaveFile(docxResult, "msdocx.pdf");
Console.WriteLine("Show docxResult: {0}", linkdocx);
```

#### Excel To Pdf

直接呼叫 **LibreOffice.OfficeConverter.ExcelToPdf(filePath)** 傳入 Excel 檔案完整路徑即可
可傳入`*.xls` 或 `*.xlsx`
```csharp
var xlsResult = MsOffice.OfficeConverter.ExcelToPdf(xlsPath);
var linkxls = SaveFile(xlsResult, "msxls.pdf");
Console.WriteLine("Show xlsResult: {0}", linkxls);
```
```csharp
Console.WriteLine("xlsx 轉為 pdf：");
// 一定使用輸出為整頁
var xlsxResult = MsOffice.OfficeConverter.ExcelToPdf(xlsxPath);
// 提供尺寸與方向選項
var xlsxResult2 = MsOffice.OfficeConverter.ExcelToPdf(
    xlsxPath,
    XlPaperSize.xlPaperB4,
    XlPageOrientation.xlPortrait);
var linkxlsx = SaveFile(xlsxResult, "msxlsx.pdf");
Console.WriteLine("Show xlsxResult: {0}", linkxlsx);
```

#### PowerPoint To Pdf

直接呼叫 **LibreOffice.OfficeConverter.PptToPdf(filePath)** 傳入 PowerPoint 檔案完整路徑即可
可傳入`*.doc` 或 `*.docx`
```csharp
var pptResult = MsOffice.OfficeConverter.PptToPdf(pptPath);
var linkppt = SaveFile(pptResult, "msppt.pdf");
Console.WriteLine("Show pptResult: {0}", linkppt);
```
```csharp
var pptxResult = MsOffice.OfficeConverter.PptToPdf(pptxPath);
var linkpptx = SaveFile(pptxResult, "mspptx.pdf");
Console.WriteLine("Show pptxResult: {0}", linkpptx);
```

## 相關問題處理步驟

- [[Office | IIS] 在IIS上存取(Access)使用Office COM元件(Interop Word Excel)的設定流程(Configuration)](<https://dotblogs.com.tw/v6610688/2015/02/19/iis_office_access_word_excel_com_interop_api_configuration>)
  - [文章備份](<https://github.com/lettucebo/Creatidea.Library.Office/blob/master/Creatidea.Library.Office.Example/ReadMe/%E5%9C%A8IIS%E4%B8%8A%E5%AD%98%E5%8F%96(Access)%E4%BD%BF%E7%94%A8Office%20COM%E5%85%83%E4%BB%B6(Interop%20Word%20Excel)%E7%9A%84%E8%A8%AD%E5%AE%9A%E6%B5%81%E7%A8%8B(Configuration).docx>) 
