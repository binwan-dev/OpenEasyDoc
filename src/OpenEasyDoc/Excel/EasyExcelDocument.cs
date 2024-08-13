using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OpenEasyDoc.Excel;

public class EasyExcelDocument:IDisposable
{
    private bool _disposed;
    private SpreadsheetDocument _document;
    private OpenXmlWriter _writer;

    internal EasyExcelDocument(string path)
    {
        FilePath = path;
        _document = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook);
        _document.AddWorkbookPart();

        _writer = OpenXmlWriter.Create(_document.WorkbookPart ?? throw new ArgumentNullException(nameof(_document.WorkbookPart)));
        _writer.WriteStartElement(new Workbook());
        _writer.WriteStartElement(new Sheets());
    }
    
    public string FilePath { get; set; }

    public EasyWorkSheet CreateSheet(uint index,string name)
    {
        ArgumentNullException.ThrowIfNull(_document.WorkbookPart);
        var worksheetPart = _document.WorkbookPart.AddNewPart<WorksheetPart>();
        _writer.WriteElement(new Sheet()
        {
            Name = name,
            SheetId = (uint)index,
            Id = _document.WorkbookPart.GetIdOfPart(worksheetPart)
        });
        return new EasyWorkSheet(index, name, worksheetPart);
    }

    public EasyExcelDocument Close()
    {
        _writer.WriteEndElement();
        _writer.WriteEndElement();
        _writer.Close();
        return this;
    }
    
    public void Dispose()
    {
        if(_disposed) return;
        
        _document.Dispose();
        _disposed = true;
    }
}
