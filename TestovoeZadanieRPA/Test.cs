using System;
using System.Collections.Generic;
using System.Text;

namespace TestovoeZadanieRPA
{
    class Test
    {
        //void ExportToExcel(bool visible)
        //{
        //    Excel.Application excel = new Excel.Application();
        //    excel.Visible = visible;

        //    Excel.Workbook wb;
        //    Excel.Worksheet ws;

        //    object misValue = System.Reflection.Missing.Value;
        //    wb = excel.Workbooks.Add(misValue);
        //    ws = (Excel.Worksheet)wb.Worksheets.get_Item(1);

        //    ws.Cells[1, 1] = "Дата";
        //    ws.Cells[1, 2] = "Курс";
        //    ws.Cells[1, 3] = "Изменения";
        //    ws.Cells[1, 4] = "Дата";
        //    ws.Cells[1, 5] = "Курс";
        //    ws.Cells[1, 6] = "Изменения";
        //    ws.Cells[1, 7] = "Частное";

        //    Invoke((MethodInvoker)(() =>
        //    {
        //        for (int column = 0; column < rateTable.ColumnCount; column++)
        //        {
        //            for (int row = 1; row < rateTable.RowCount; row++)
        //            {
        //                ws.Cells[row + 1, column + 1] = rateTable.Rows[row].Cells[column].Value;
        //                if (column == 1 || column == 4)
        //                {
        //                    ws.Cells[row + 1, column + 1].NumberFormat = "DD/MM/YYYY";
        //                }
        //            }
        //        }
        //    }));
        //}
    }
}
