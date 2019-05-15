// Geoff Overfield
// 05/14/2019
// Scripts to transfer and merge client data from a collection
// of query-based CSV files to a single file for upload to new client
// Completed at request of Management of Peace of Mind Massage

#region Namespaces
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
#endregion

namespace POMDataTransfer
{
    class ExcelReader
    {
        public string Path { get; private set; }
        _Application m_pExcel;
        Workbook m_pWorkbook;
        Worksheet m_pWorksheet;

        public ExcelReader(string sPath, int iSheet)
        {
            Path = sPath;
            m_pExcel = new Excel.Application();
            m_pWorkbook = m_pExcel.Workbooks.Open(Path);
            m_pWorksheet = m_pWorkbook.Worksheets[iSheet];
        }

        public string ReadCell(int iRow, int iColumn)
        {
            iRow++;
            iColumn++;

            if (m_pWorksheet.Cells[iRow, iColumn].Value2 != null)
                return m_pWorksheet.Cells[iRow, iColumn].Value2.ToString();
            else return string.Empty;
        }

        public void WriteToCell(int iRow, int iColumn, string sValue)
        {
            iRow++;
            iColumn++;

            m_pWorksheet.Cells[iRow, iColumn].Value2 = sValue;
        }

        public void Save()
        {
            m_pWorkbook.Save();
        }

        public void SaveAs(string sPath)
        {
            m_pWorkbook.SaveAs(sPath);
        }

        public void Dispose()
        {
            m_pWorksheet = null;
            m_pWorkbook = null;
            m_pExcel.Quit();

        }
    }
}
