using OfficeOpenXml;
using System;
using System.Data;
using System.IO;

namespace AlfaBank_TestCase_Proshunin
{
    class DataSheet
    {

        public static DataTable GetDTFromExcel(string path)
        {

            DataTable dt = new DataTable();
            FileInfo fi = new FileInfo(path);

            // Check if the file exists
            if (!fi.Exists)
                throw new Exception("File " + path + " Does Not Exists");

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (ExcelPackage xlPackage = new ExcelPackage(fi))
            {
                // get the first worksheet in the workbook
                ExcelWorksheet worksheet = xlPackage.Workbook.Worksheets[0];

                // Fetch the WorkSheet size
                ExcelCellAddress startCell = worksheet.Dimension.Start;
                ExcelCellAddress endCell = worksheet.Dimension.End;

                // create all the needed DataColumn
                for (int col = startCell.Column; col <= endCell.Column; col++)
                    dt.Columns.Add(col.ToString());

                // place all the data into DataTable
                for (int row  = startCell.Row + 1; row < endCell.Row; row++)
                {
                    DataRow dr = dt.NewRow();
                    int x = 0;
                    for (int col = startCell.Column; col <= endCell.Column; col++)
                    {
                        dr[x++] = worksheet.Cells[row, col].Value;
                    }
                    dt.Rows.Add(dr);
                }
            }
            return dt;
        }   

    }
}
