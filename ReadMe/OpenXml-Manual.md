# Open Xml 套版使用說明

## 目錄

- 使用前注意事項
- 安裝 nuget 套件
- 開始開發
- 備註

## 使用前注意事項

- 此套件使用 [Open XML SDK 2.5](https://www.microsoft.com/en-us/download/details.aspx?id=30425) 作為套版工具，因此若伺服器**無法安裝 Open XML SDK 2.5 就可以跳過此套件**了
- 因 Word 2007 版本以後才使用 XML 作為儲存格式(.docx)，儲存的檔案一定為 *.docx

## 安裝 nuget 套件

You can either <a href="https://github.com/lettucebo/Creatidea.Library.Office.git">download</a> the source and build your own dll or, if you have the NuGet package manager installed, you can grab them automatically.

```
PM> Install-Package Creatidea.Library.OpenXml
```

Once you have the libraries properly referenced in your project, you can include calls to them in your code. 
For a sample implementation, check the [Example](https://github.com/lettucebo/Creatidea.Library.Office/tree/master/Creatidea.Library.Office.Example) folder.

Add the following namespaces to use the library:
```csharp
using Creatidea.Library.Office.OpenXml;
```

## 環境設定
必須安裝 [Open XML SDK 2.5](https://www.microsoft.com/en-us/download/details.aspx?id=30425) 

## 開始開發

#### 1. 建立樣板 Word 檔案
先建立 Word 樣板，作為之後套版之標準

使用 Word 之**內容控制項**建立動態欄位

並將每個內容控制項賦予唯一ID

範例：[TemplateExample.docx]()


#### 2. 建立套版參數列表

直接呼叫 **LibreOffice.OfficeConverter.ExcelToPdf(filePath)** 傳入 Excel 檔案完整路徑即可
可傳入`*.xls` 或 `*.xlsx`
```csharp

```

#### 3. 呼叫 DocxMaker 進行處理

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

## 備註
- 應用程式要擁有暫存目錄寫入權限

## 相關問題處理步驟
