using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.IO;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
namespace YXSummaryTool
{
    class NPOIHelp
    {
        public static DataSet GetDataTableFromExcelFile(string fileName, int sheetnmbr)
        {
            FileStream fs = null;
            DataSet ds = new DataSet();
            int TotalSheetNeedToParse = 0;
            try
            {
                IWorkbook wb = null;
                fs = File.Open(fileName, FileMode.Open, FileAccess.Read);
                switch (Path.GetExtension(fileName).ToUpper())
                {
                    case ".XLS":
                        {
                            wb = new HSSFWorkbook(fs);
                        }
                        break;
                    case ".XLSX":
                        {
                            wb = new XSSFWorkbook(fs);
                        }
                        break;
                }
                if (sheetnmbr == 0)
                {
                    TotalSheetNeedToParse = wb.NumberOfSheets;
                }
                else
                {
                    TotalSheetNeedToParse = sheetnmbr;
                }

                if (wb.NumberOfSheets > 0)
                {
                    for (int k = 0; k < wb.NumberOfSheets; k++)
                    {
                        if (k + 1 > TotalSheetNeedToParse)
                        {
                            //Logging.Write("Last table read: {0}", TotalSheetNeedToParse);
                            break;
                        }
                        ISheet sheet;
                        if (wb.IsSheetHidden(k))
                        {
                            sheet = wb.GetSheetAt(k);
                            //Logging.Write("Hidden sheet skipped: {0}", sheet.SheetName);
                            TotalSheetNeedToParse++;
                            continue;
                        }
                        else
                        {
                            sheet = wb.GetSheetAt(k);
                        }
                        IRow headerRow = sheet.GetRow(0);
                        
                        // 取得最右邊的位置
                        int RightmostNmbr = 0;
                        for (int ii = sheet.FirstRowNum; ii < sheet.LastRowNum; ii++)
                        {
                            IRow row = sheet.GetRow(ii);
                            if (RightmostNmbr < row.LastCellNum)
                            {
                                RightmostNmbr = row.LastCellNum;
                            }
                        }

                        if ((!(headerRow == null)))
                        {
                            DataTable dt = new DataTable();
                            dt.TableName = sheet.SheetName;
                            //處理標題列
                            for (int i = headerRow.FirstCellNum; i < RightmostNmbr; i++)
                            {
                                if (!(headerRow.GetCell(i) == null))
                                {
                                    dt.Columns.Add(headerRow.GetCell(i).StringCellValue.Trim());
                                }
                                else
                                {
                                    dt.Columns.Add("");
                                }
                            }

                            IRow row = null;
                            DataRow dr = null;
                            CellType ct = CellType.Blank;
                            //標題列之後的資料
                            for (int i = sheet.FirstRowNum + 1; i <= sheet.LastRowNum; i++)
                            {
                                dr = dt.NewRow();
                                row = sheet.GetRow(i);
                                if (row == null) continue;
                                for (int j = row.FirstCellNum; j < row.LastCellNum; j++)
                                {
                                    if (!(row.GetCell(j) == null))
                                    {
                                        ct = row.GetCell(j).CellType;
                                    }

                                    //如果此欄位格式為公式 則去取得CachedFormulaResultType
                                    if (ct == CellType.Formula)
                                    {
                                        ct = row.GetCell(j).CachedFormulaResultType;
                                    }
                                    if (ct == CellType.Numeric)
                                    {
                                        if (row.GetCell(j) == null)
                                        {
                                            dr[j] = "";
                                        }
                                        else
                                        {
                                            dr[j] = row.GetCell(j).NumericCellValue;
                                        }
                                    }
                                    else
                                    {
                                        if (row.GetCell(j) == null)
                                        {
                                            dr[j] = "";
                                        }
                                        else
                                        {
                                            dr[j] = row.GetCell(j).ToString().Replace("$", "");
                                        }
                                    }
                                }
                                dt.Rows.Add(dr);
                            }
                            ds.Tables.Add(dt);
                        }

                    }
                }
                fs.Close();
            }
            finally
            {
                if (fs != null) fs.Dispose();
            }
            return ds;
        }
    }
}
