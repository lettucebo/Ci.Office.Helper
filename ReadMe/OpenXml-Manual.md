# Open Xml 套版使用說明

## 目錄

- 使用前注意事項
- 安裝 nuget 套件
- 開始開發
- 備註

## 使用前注意事項

- 此套件使用 [Open XML SDK 2.5](https://www.microsoft.com/en-us/download/details.aspx?id=30425) 作為套版工具
  - 原則上應須安裝在伺服器上，但因現在已可直接從 Nuget 取得套件，因此可不必安裝 SDK
- 因 Word 2007 版本以後才使用 XML 作為儲存格式(.docx)，開啟、編輯與儲存的檔案一定為 *.docx

## 安裝 nuget 套件

You can either <a href="https://github.com/lettucebo/Ci.Office.Helper.git">download</a> the source and build your own dll or, if you have the NuGet package manager installed, you can install them automatically.

```
PM> Install-Package Ci.Office.Helper.OpenXml
```

Once you have the libraries properly referenced in your project, you can include calls to them in your code. 
For a sample implementation, check the [Example](https://github.com/lettucebo/Creatidea.Library.Office/tree/master/Creatidea.Library.Office.Example) folder.

Add the following namespaces to use the library:
```csharp
using Ci.Office.Helper.OpenXml;
```

## 開始開發

#### 1. 建立樣板 Word 檔案
先建立 Word 樣板，作為之後套版之標準

使用 Word 之**內容控制項**建立動態欄位

並將每個內容控制項賦予唯一ID

範例：[WordTemplate.docx](https://github.com/lettucebo/Ci.Office.Helper/blob/master/Ci.Office.Helper.Example/Demo/Word/Template.docx)

備註1. 圖片的內容控制項需先內置圖片，否則會發生未知原因的找不到控制項問題

#### 2. 建立套版參數列表

依照需要套版的不同資料型態可區分為三種類型：文字、圖片、表格
```csharp
// 欲塞入資料之文字類型列表
var textDict = new Dictionary<string, OpenXmlTextInfo>();

// 欲塞入資料之圖片類型列表
var imgDict = new Dictionary<string, MemoryStream>();

// 欲塞入資料之表格類型列表
var tableDict = new Dictionary<string, DocumentFormat.OpenXml.Wordprocessing.Table>();
```

詳細範例可參考：Ci.Office.Helper.Example/[Program.cs](https://github.com/lettucebo/Ci.Office.Helper/blob/master/Ci.Office.Helper.Example/Program.cs#L80-L213)

#### 3. 呼叫 Template.DocxMaker 進行處理

先初始化 Template 類別，然後呼叫 DocxMaker 方法進行套版動作
```csharp
// create template engine
var template = new Template();

// call DocxMaker to template the file
var filePath = template.DocxMaker(docxTemplatePath, textDict, imageDict, tableDict);
```

## 備註
- 應用程式要擁有暫存目錄寫入權限，否則無法產生暫存檔進行套版
