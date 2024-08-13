using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OpenEasyDoc.Excel;

public class EasyWorkSheet
{
    private OpenXmlWriter _writer;
    private int _rowIndex=1;
    private int _cellIndex=1;

    internal EasyWorkSheet(uint index,string name, WorksheetPart worksheetPart)
    {
        Index = index;
        Name = name;
        _writer = OpenXmlWriter.Create(worksheetPart);
        _writer.WriteStartElement(new Worksheet());
        _writer.WriteStartElement(new SheetData());
    }
    
    public uint Index { get; private set; }
    
    public string Name { get; private set; }

    public EasyWorkSheet NewRow()
    {
        var attributes = new List<OpenXmlAttribute>() { new OpenXmlAttribute("r", "", _rowIndex.ToString()) };
        _writer.WriteStartElement(new Row(), attributes);
        _cellIndex = 1;
        return this;
    }

    public EasyWorkSheet EndRow()
    {
        _rowIndex++;
        _writer.WriteEndElement();
        return this;
    }

    public EasyWorkSheet WriteCellData<T>(T val)
    {
        var data = val?.ToString() ?? string.Empty;
        var cleanedValue = data.Replace("\0", string.Empty); 
        if (data.Contains('\0')) data = cleanedValue;
        
        var attributes = new List<OpenXmlAttribute>();
        attributes.Add(new OpenXmlAttribute("t", "", "str"));
        attributes.Add(new OpenXmlAttribute("r", "", $"{GetColumnName(_cellIndex++)}{_rowIndex}"));
        
        _writer.WriteStartElement(new Cell(), attributes);
        _writer.WriteElement(new CellValue(data));
        _writer.WriteEndElement();
        return this;
    }

    public EasyWorkSheet Close()
    {
        _writer.WriteEndElement();
        _writer.WriteEndElement();
        _writer.Close();
        return this;
    }
    
    private string GetColumnName(int columnIndex)
    {
        int dividend = columnIndex;
        string columnName = String.Empty;
        int modifier;

        while (dividend > 0)
        {
            modifier = (dividend - 1) % 26;
            columnName = Convert.ToChar(65 + modifier).ToString() + columnName;
            dividend = (int)((dividend - modifier) / 26);
        }

        return columnName;
    }
}
