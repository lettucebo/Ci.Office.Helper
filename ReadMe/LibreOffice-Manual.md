# Libre Office 使用說明

## 目錄

- SDK安裝說明
- 安裝 nuget 套件
- 開始開發
- 備註
- Libre Office SDK 起源

## SDK安裝說明

1. 安裝 Libre Office

    ![安裝 Libre Office](http://i.imgur.com/fZlP9QA.png)

2. 安裝 Libre Office**相對應版本之 SDK**

    ![安裝 Libre Office相對應版本之 SDK](http://i.imgur.com/iSM7uey.png)
    
3. 開發專案請參考以下路徑之 dll (若在安裝時選了不同路徑，請自行變更)
  - C:\Program Files (x86)\LibreOffice_5.1_SDK\sdk\cli 
  - ![開發專案請參考以下路徑之 dll](http://i.imgur.com/aRMd68y.png)

## 安裝 nuget 套件

You can either <a href="https://github.com/lettucebo/Creatidea.Library.Office.git">download</a> the source and build your own dll or, if you have the NuGet package manager installed, you can grab them automatically.

```
PM> Install-Package Creatidea.Library.LibreOffice
```

Once you have the libraries properly referenced in your project, you can include calls to them in your code. 
For a sample implementation, check the [Example](https://github.com/lettucebo/Creatidea.Library.Office/tree/master/Creatidea.Library.Office.Example) folder.

Add the following namespaces to use the library:
```csharp
using Creatidea.Library.Office.LibreOffice;
```

## 環境設定
至 `CiLiberOffice.json` 設定 SDK 路徑(BinPath)
```json
{
    "CiLibreOffice": {
        "Version": "1.0.0",
        "BinPath": "C:\\Program Files (x86)\\LibreOffice 5\\program\\soffice.exe"
    }
}
```

## 開始開發

直接呼叫 **LibreOffice.OfficeConverter.WordToPdf(filePath)** 傳入 Word 完整檔案路徑即可
可傳入`*.doc` 或 `*.docx`
```csharp
var docResult = LibreOffice.OfficeConverter.WordToPdf(docPath);
var link = SaveFile(docResult.Data, "doc.pdf");
Console.WriteLine("Show docResult: {0}", link);
```

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

## 備註

1. 佈署在 IIS 上需注意 LibreOffice 安裝路徑權限
2. 有些 OpenOffice 沒有的功能轉換過去會無法使用(畫布)
3. 確定 IIS 有執行過該程式

## 參考資料

- [Framework/Article/Filter/FilterList OOo 2 1](<https://wiki.openoffice.org/wiki/Framework/Article/Filter/FilterList_OOo_2_1>)

## Libre Office SDK 起源

在2010年9月28日，幾個 OpenOffice.org 的開發成員成立了檔案基金會（The Document Foundation），希望接手開發被甲骨文公司（Oracle）放棄的 OpenOffice.org 計畫；最初檔案基金會也邀請甲骨文成為其中的一員，但是甲骨文拒絕提供 OpenOffice.org 的品牌給檔案基金會使用。隨後甲骨文要求參與檔案基金會的成員退出 OpenOffice.org 的開發社群，因為甲骨文認為參與檔案基金會和該公司的利益發生衝突。

隨後 OpenOffice.org 其中一個分支 Go-oo 決定支援 LibreOffice，宣布停止開發[18]，並且把開發成果合併到 LibreOffice 中。
最終，隨著開發者轉移到 LibreOffice 上，甲骨文宣布停止 OpenOffice.org 的商業支援。2011年6月，甲骨文宣布將 OpenOffice.org 捐贈給 Apache 軟體基金會，未來 OpenOffice.org 的發展將由 Apache 軟體基金會主導。
