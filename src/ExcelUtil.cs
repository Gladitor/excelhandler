using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.POIFS.FileSystem;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using crl = System.Runtime.InteropServices.Marshal;
using System.Runtime.InteropServices;

namespace ExcelHandler
{
    public class ExcelUtil
    {
        static DataRow temprow;
        static int temi;
        /// <summary>
        /// 将excel中的数据导入到DataTable中
        /// </summary>
        /// <param name="sheetName">excel工作薄sheet的名称</param>
        /// <param name="isFirstRowColumn">第一行是否是DataTable的列名</param>
        /// <param name="fileName">文件路径</param>
        /// <param name="isFirstSheetDefault">没找到指定sheet时是否返回第一个</param>
        /// <returns>返回的DataTable</returns>
        public static DataTable ExcelToDataTable(string sheetName, bool isFirstRowColumn, string fileName, bool isFirstSheetDefault)
        {
            if (string.IsNullOrEmpty(sheetName))
            {
                throw new ArgumentNullException(sheetName);
            }
            if (string.IsNullOrEmpty(fileName))
            {
                throw new ArgumentNullException(fileName);
            }
            var data = new DataTable();
            IWorkbook workbook = null;
            FileStream fs = null;
            try
            {
                fs = new FileStream(fileName, FileMode.Open, FileAccess.Read);
                if (fileName.IndexOf(".xlsx", StringComparison.Ordinal) > 0)
                {
                    workbook = new XSSFWorkbook(fs);
                }
                else if (fileName.IndexOf(".xls", StringComparison.Ordinal) > 0)
                {
                    workbook = new HSSFWorkbook(fs);
                }

                ISheet sheet = null;
                if (workbook != null)
                {
                    if(isFirstSheetDefault)
                    {
                        //如果没有找到指定的sheetName对应的sheet，则尝试获取第一个sheet
                        sheet = workbook.GetSheet(sheetName) ?? workbook.GetSheetAt(0);
                    }
                    else
                    {
                        sheet = workbook.GetSheet(sheetName);
                    }
                }
                //表格内容为空
                if (sheet == null || sheet.LastRowNum == 0) return data;
                var firstRow = sheet.GetRow(0);
                //一行最后一个cell的编号 即总的列数
                //int cellCount = firstRow.LastCellNum;
                int cellCount = getColumnCount(sheet);
                int startRow;
                if (isFirstRowColumn)
                {
                    for (int i = 0; i < cellCount; ++i)
                    {
                        var cell = firstRow.GetCell(i);
                        var cellValue = "";
                        if(cell!=null)
                        {
                            cellValue = cell.ToString();
                        }
                        var column = new DataColumn(cellValue);
                        data.Columns.Add(column);
                    }
                    startRow = sheet.FirstRowNum + 1;
                }
                else
                {
                    for (int i = 0; i < cellCount; ++i)
                    {
                        var column = new DataColumn("");
                        data.Columns.Add(column);
                    }
                    startRow = sheet.FirstRowNum;
                }
                //最后一列的标号
                var rowCount = sheet.LastRowNum;
                for (var i = startRow; i <= rowCount; ++i)
                {
                    temi = i;
                    var row = sheet.GetRow(i);
                    //没有数据的行默认是null
                    if (row == null || row.Cells.Count==0) continue;
                    var dataRow = data.NewRow();
                    int cellscount = row.Cells.Count;
                    //if (row.Cells.Count > cellCount)//若行列数大于第一行列数则使用第一行列数
                    //    cellscount = cellCount;
                    for (int j = row.FirstCellNum; j <= cellscount; ++j)
                    {
                        //同理，没有数据的单元格都默认是null
                        if (row.GetCell(j) != null) {
                            //单元格的类型为公式，返回公式的值
                            if (row.GetCell(j).CellType == CellType.Formula)
                            {
                                row.GetCell(j).SetCellType(CellType.String);
                                //是日期型
                                //if (HSSFDateUtil.IsCellDateFormatted(row.GetCell(j)))
                                //{
                                //    dataRow[j] = row.GetCell(j).DateCellValue.ToString("yyyy-MM-dd HH:mm:ss");
                                //}
                                ////不是日期型
                                //else
                                //{
                                //    dataRow[j] = row.GetCell(j).NumericCellValue.ToString();
                                //}
                                dataRow[j] = row.GetCell(j).StringCellValue.Trim();
                            }
                            else
                            {
                                dataRow[j] = row.GetCell(j).ToString().Trim();
                            }
                        }
                    }
                    data.Rows.Add(dataRow);
                    temprow = dataRow;
                }
                workbook.Close();
                return data;
            }
            catch (IOException ioex)
            {
                throw new IOException(ioex.Message);
            }
            catch (Exception ex)
            {
                int i = temi;
                DataRow dr = temprow;
                throw new Exception(ex.Message);
            }
            finally
            {
                if (fs != null)
                {
                    fs.Close();
                }
            }
        }
        //获取最大有效列数
        static int getColumnCount(ISheet sheet)
        {
            int cloNum = sheet.GetRow(sheet.FirstRowNum).LastCellNum;
            for (int rowCnt = sheet.FirstRowNum; rowCnt <= sheet.LastRowNum; rowCnt++)//迭代所有行
            {
                IRow row = sheet.GetRow(rowCnt);
                if (row != null && row.LastCellNum > cloNum)
                {
                    cloNum = row.LastCellNum;
                }
            }
            return cloNum;
        }
        //获取Sheet,不存在则创建
        public static ISheet getSheet(string sheetName, string fileName)
        {
            if (string.IsNullOrEmpty(sheetName))
            {
                throw new ArgumentNullException(sheetName);
            }
            if (string.IsNullOrEmpty(fileName))
            {
                throw new ArgumentNullException(fileName);
            }
            IWorkbook workbook = null;
            FileStream fs = null;
            try
            {
                fs = new FileStream(fileName, FileMode.Open, FileAccess.ReadWrite);
                if (fileName.IndexOf(".xlsx", StringComparison.Ordinal) > 0)
                {
                    workbook = new XSSFWorkbook(fs);
                }
                else if (fileName.IndexOf(".xls", StringComparison.Ordinal) > 0)
                {
                    workbook = new HSSFWorkbook(fs);
                }

                ISheet sheet = null;
                if (workbook != null)
                {
                    sheet = workbook.GetSheet(sheetName);
                    if(sheet == null)
                    {
                        sheet = workbook.CreateSheet(sheetName);
                        workbook.Write(fs);
                    }
                }
                workbook.Close();
                return sheet;
            }
            catch(Exception e)
            {
                throw new Exception(e.Message);
            }
            finally
            {
                if (fs != null)
                {
                    fs.Close();
                }
            }
        }

        /// <summary>
        /// 将DataTable数据导入到excel中
        /// </summary>
        /// <param name="data">要导入的数据</param>
        /// <param name="isColumnWritten">DataTable的列名是否要导入</param>
        /// <param name="sheetName">要导入的excel的sheet的名称</param>
        /// <param name="fileName">文件夹路径</param>
        /// <param name="sheetNames">其它需要创建的sheet</param>
        /// <returns>导入数据行数(包含列名那一行)</returns>
        public static int DataTableToExcel(DataTable data, string sheetName, bool isColumnWritten, string fileName, string[] sheetNames)
        {
            if (data == null)
            {
                throw new ArgumentNullException("data");
            }
            if (string.IsNullOrEmpty(sheetName))
            {
                throw new ArgumentNullException(sheetName);
            }
            if (string.IsNullOrEmpty(fileName))
            {
                throw new ArgumentNullException(fileName);
            }
            IWorkbook workbook = null;
            if (fileName.IndexOf(".xlsx", StringComparison.Ordinal) > 0)
            {
                workbook = new XSSFWorkbook();
            }
            else if (fileName.IndexOf(".xls", StringComparison.Ordinal) > 0)
            {
                workbook = new HSSFWorkbook();
            }

            FileStream fs = null;
            try
            {
                fs = new FileStream(fileName, FileMode.OpenOrCreate, FileAccess.ReadWrite);
                ISheet sheet;
                if (workbook != null)
                {
                    sheet = workbook.CreateSheet(sheetName);
                    if (sheetNames != null)
                    {
                        for (int s = 0; s < sheetNames.Length; s++)
                        {
                            workbook.CreateSheet(sheetNames[s]);
                        }
                    }
                }
                else
                {
                    return -1;
                }

                int j;
                int count;
                //写入DataTable的列名，写入单元格中
                if (isColumnWritten)
                {
                    var row = sheet.CreateRow(0);
                    for (j = 0; j < data.Columns.Count; ++j)
                    {
                        row.CreateCell(j).SetCellValue(data.Columns[j].ColumnName);
                    }
                    count = 1;
                }
                else
                {
                    count = 0;
                }
                //遍历循环datatable具体数据项
                int i;
                for (i = 0; i < data.Rows.Count; ++i)
                {
                    var row = sheet.CreateRow(count);
                    for (j = 0; j < data.Columns.Count; ++j)
                    {
                        row.CreateCell(j).SetCellValue(data.Rows[i][j].ToString());
                    }
                    ++count;
                }
                //将文件流写入到excel
                workbook.Write(fs);
                workbook.Close();
                return count;
            }
            catch (IOException ioex)
            {
                throw new IOException(ioex.Message);
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            finally
            {
                if (fs != null)
                {
                    fs.Close();
                }
            }
        }
        /// <summary>
        /// 将DataTable数据导入到excel中,sheetname存在则追加,不存在则创建
        /// </summary>
        /// <param name="dt"></param>
        /// <param name="filename"></param>
        /// <param name="sheetname"></param>
        /// <returns></returns>
        public static int DataTableToExcel(DataTable dt,string filename,string sheetname)
        {
            XSSFWorkbook workbook = new XSSFWorkbook();
            XSSFCellStyle styleHeader = (XSSFCellStyle)workbook.CreateCellStyle();
            styleHeader.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
            //NPOI.SS.UserModel.IFont font = workbook.CreateFont();
            //styleHeader.SetFont(font);
            XSSFCellStyle style = (XSSFCellStyle)workbook.CreateCellStyle();
            style.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;

            if (!File.Exists(filename)) //判断文件是否存在,不存在时则创建并添加表头
            {
                using (FileStream fs = new FileStream(filename, FileMode.OpenOrCreate))//读取流
                {

                    //创建sheet
                    ISheet sheet = workbook.CreateSheet(sheetname);
                    IRow rowHeader = sheet.CreateRow(0);
                    //添加表头
                    for (int col = 0; col < dt.Columns.Count; col++)
                    {
                        ICell cellHeader = rowHeader.CreateCell(col);
                        cellHeader.SetCellValue(dt.Columns[col].ColumnName);
                        sheet.SetColumnWidth(col, 15 * 256);
                        //cellHeader.CellStyle = styleHeader;
                    }
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        IRow row = sheet.CreateRow(i + 1);//如果无数据则+1留出表头位置
                        for (int j = 0; j < dt.Columns.Count; j++)
                        {
                            ICell cell = row.CreateCell(j);
                            cell.SetCellValue(dt.Rows[i][j].ToString());
                        }
                    }
                    workbook.Write(fs);
                    fs.Close();
                }
            }
            else
            {   //如果文件已存在则首先打开，并且获取到最大的行数
                using (FileStream fs = new FileStream(filename, FileMode.OpenOrCreate))//读取流
                {
                    workbook = new XSSFWorkbook(fs);
                    ISheet sheet1 = workbook.GetSheet(sheetname);
                    if(sheet1==null)
                    {
                        sheet1 = workbook.CreateSheet(sheetname);
                    }
                    //设置表头
                    IRow rowHeader = sheet1.GetRow(0)??sheet1.CreateRow(0);
                    for (int col = 0; col < dt.Columns.Count; col++)
                    {
                        ICell cellHeader = rowHeader.CreateCell(col);
                        cellHeader.SetCellValue(dt.Columns[col].ColumnName);
                        sheet1.SetColumnWidth(col, 15 * 256);
                        //cellHeader.CellStyle = styleHeader;
                    }

                    int num = sheet1.LastRowNum + 1;//获取最大行数
                    FileStream fout = new FileStream(filename, FileMode.Open, FileAccess.Write, FileShare.ReadWrite);//写入流

                    //获取到已存在的sheet
                    ISheet sheet = workbook.GetSheet(sheetname);

                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        //创建行数时+num
                        IRow row = sheet.CreateRow(i + num);
                        for (int j = 0; j < dt.Columns.Count; j++)
                        {
                            ICell cell = row.CreateCell(j);
                            cell.SetCellValue(dt.Rows[i][j].ToString());
                        }
                    }
                    workbook.Write(fout);
                    fs.Close();
                    fout.Close();
                }
            }
            return dt.Rows.Count;
        }

        /// <summary>
        /// 向已存在的excel追加数据
        /// </summary>
        //打开已经存在的一个excel，并且找到最后一行，然后按行，继续添加内容。
        public static void appendInfoToFile_old(string fullFilename,string sheetname,DataTable datadt)//AmazonProductInfo productInfo1
        {
            Excel.Application xlApp = null;
            Excel.Workbook xlWorkBook = null;
            Excel.Worksheet xlWorkSheet = null;
            try
            {
                object missingVal = System.Reflection.Missing.Value;

                xlApp = new Microsoft.Office.Interop.Excel.Application();
                //xlApp.Visible = true;
                //xlApp.DisplayAlerts = false;

                //http://msdn.microsoft.com/zh-cn/library/microsoft.office.interop.excel.workbooks.open%28v=office.11%29.aspx
                xlWorkBook = xlApp.Workbooks.Open(
            Filename: fullFilename,
                    //UpdateLinks:3,
                    ReadOnly: false,
                    //Format : 2, //use Commas as delimiter when open text file
                    //Password : missingVal,
                    //WriteResPassword : missingVal,
                    //IgnoreReadOnlyRecommended: false, //when save to readonly, will notice you
                    Origin: Excel.XlPlatform.xlWindows, //xlMacintosh/xlWindows/xlMSDOS
                                                        //Delimiter: ",",  // usefule when is text file
                    Editable: true,
            Notify: false,
                    //Converter: missingVal,
                    AddToMru: true, //True to add this workbook to the list of recently used files
                    Local: true,
            CorruptLoad: missingVal //xlNormalLoad/xlRepairFile/xlExtractData
                    );

                //Get the first sheet
                //xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1); //also can get by sheet name
                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(sheetname);
                Excel.Range range = xlWorkSheet.UsedRange;
                //int usedColCount = range.Columns.Count;
                int usedRowCount = range.Rows.Count;

                const int excelRowHeader = 1;
                const int excelColumnHeader = 1;

                //int curColumnIdx = usedColCount + excelColumnHeader;
                int curColumnIdx = 0 + excelColumnHeader; //start from column begin
                int curRrowIdx = usedRowCount + excelRowHeader; // !!! here must added buildin excelRowHeader=1, otherwise will overwrite previous (added title or whole row value)

                //curRrowIdx = curRrowIdx + 1;//空行间隔
                                            //xlWorkSheet.Cells[curRrowIdx, curColumnIdx] = "222";//productInfo.title;
                                            //xlWorkSheet.Cells[curRrowIdx, ++curColumnIdx] = "333";
                                            //xlWorkSheet.Cells[curRrowIdx, ++curColumnIdx] = "444";
                for (int i = 0; i < datadt.Rows.Count; i++)
                {
                    for (int j = 0; j < datadt.Columns.Count; j++)
                    {
                        Microsoft.Office.Interop.Excel.Range range1 = (Excel.Range)xlWorkSheet.Cells[curRrowIdx + i, curColumnIdx + j];
                        range1.NumberFormat = "@";//文本格式
                        xlWorkSheet.Cells[curRrowIdx + i, curColumnIdx + j] = datadt.Rows[i][j].ToString();
                    }
                }

                /*
                const int constBullerLen = 5;
                int bulletListLen = 0;
                if (productInfo.bulletArr.Length > constBullerLen)
                {
                    bulletListLen = constBullerLen;
                }
                else
                {
                    bulletListLen = productInfo.bulletArr.Length;
                }
                for (int bulletIdx = 0; bulletIdx < bulletListLen; bulletIdx++)
                {
                    xlWorkSheet.Cells[curRrowIdx, curColumnIdx + bulletIdx] = productInfo.bulletArr[bulletIdx];
                }
                curColumnIdx = curColumnIdx + bulletListLen;

                const int constImgNameListLen = 5;
                int imgNameListLen = 0;
                if (productInfo.imgFullnameArr.Length > constImgNameListLen)
                {
                    imgNameListLen = constImgNameListLen;
                }
                else
                {
                    imgNameListLen = productInfo.imgFullnameArr.Length;
                }
                for (int imgIdx = 0; imgIdx < imgNameListLen; imgIdx++)
                {
                    xlWorkSheet.Cells[curRrowIdx, curColumnIdx + imgIdx] = productInfo.imgFullnameArr[imgIdx];
                }
                curColumnIdx = curColumnIdx + imgNameListLen;

                xlWorkSheet.Cells[curRrowIdx, curColumnIdx++] = productInfo.highestPrice;
                xlWorkSheet.Cells[curRrowIdx, curColumnIdx++] = productInfo.isOneSellerIsAmazon;
                xlWorkSheet.Cells[curRrowIdx, curColumnIdx++] = productInfo.reviewNumber;
                xlWorkSheet.Cells[curRrowIdx, curColumnIdx++] = productInfo.isBestSeller;
                */

                ////http://msdn.microsoft.com/query/dev10.query?appId=Dev10IDEF1&l=ZH-CN&k=k%28MICROSOFT.OFFICE.INTEROP.EXCEL._WORKBOOK.SAVEAS%29;k%28SAVEAS%29;k%28TargetFrameworkMoniker-%22.NETFRAMEWORK%2cVERSION%3dV3.5%22%29;k%28DevLang-CSHARP%29&rd=true
                //xlWorkBook.SaveAs(
                //    Filename: fullFilename,
                //    ConflictResolution: XlSaveConflictResolution.xlLocalSessionChanges //The local user's changes are always accepted.
                //    //FileFormat : Excel.XlFileFormat.xlWorkbookNormal
                //);

                //if use above SaveAs -> will popup a window ask you overwrite it or not, even if you have set the ConflictResolution to xlLocalSessionChanges, which should not ask, should directly save
                xlWorkBook.Save();

                //http://msdn.microsoft.com/query/dev10.query?appId=Dev10IDEF1&l=ZH-CN&k=k%28MICROSOFT.OFFICE.INTEROP.EXCEL._WORKBOOK.CLOSE%29;k%28CLOSE%29;k%28TargetFrameworkMoniker-%22.NETFRAMEWORK%2cVERSION%3dV3.5%22%29;k%28DevLang-CSHARP%29&rd=true
                xlWorkBook.Close(SaveChanges: true);
            }
            catch (Exception e)
            {
                
            }
            finally
            {
                if (xlWorkSheet != null)
                    crl.ReleaseComObject(xlWorkSheet);
                if (xlWorkBook != null)
                    crl.ReleaseComObject(xlWorkBook);
                //if (xlApp != null)
                //    crl.ReleaseComObject(xlApp);//releaseObject
                CloseExcel(xlApp, xlWorkBook);
            }
            
        }
        /// <summary>
        /// 向已存在的excel追加数据
        /// </summary>
        //打开已经存在的一个excel，并且找到最后一行，然后按行，继续添加内容。
        public static void appendInfoToFile(string fullFilename, string sheetname, DataTable datadt)//AmazonProductInfo productInfo1
        {
            if (datadt.Rows.Count <= 0) return;
            Excel.Application xlApp = null;
            Excel.Workbook xlWorkBook = null;
            Excel.Worksheet xlWorkSheet = null;
            try
            {
                object missingVal = System.Reflection.Missing.Value;

                xlApp = new Microsoft.Office.Interop.Excel.Application();
                //xlApp.Visible = true;
                //xlApp.DisplayAlerts = false;

                //http://msdn.microsoft.com/zh-cn/library/microsoft.office.interop.excel.workbooks.open%28v=office.11%29.aspx
                xlWorkBook = xlApp.Workbooks.Open(
            Filename: fullFilename,
                    //UpdateLinks:3,
                    ReadOnly: false,
                    //Format : 2, //use Commas as delimiter when open text file
                    //Password : missingVal,
                    //WriteResPassword : missingVal,
                    //IgnoreReadOnlyRecommended: false, //when save to readonly, will notice you
                    Origin: Excel.XlPlatform.xlWindows, //xlMacintosh/xlWindows/xlMSDOS
                                                        //Delimiter: ",",  // usefule when is text file
                    Editable: true,
            Notify: false,
                    //Converter: missingVal,
                    AddToMru: true, //True to add this workbook to the list of recently used files
                    Local: true,
            CorruptLoad: missingVal //xlNormalLoad/xlRepairFile/xlExtractData
                    );

                //Get the first sheet
                //xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1); //also can get by sheet name
                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(sheetname);
                Excel.Range range = xlWorkSheet.UsedRange;
                //int usedColCount = range.Columns.Count;
                int usedRowCount = range.Rows.Count;

                const int excelRowHeader = 1;
                const int excelColumnHeader = 1;

                //int curColumnIdx = usedColCount + excelColumnHeader;
                int curColumnIdx = 0 + excelColumnHeader; //start from column begin
                int curRrowIdx = usedRowCount + excelRowHeader; // !!! here must added buildin excelRowHeader=1, otherwise will overwrite previous (added title or whole row value)

                //curRrowIdx = curRrowIdx + 1;//空行间隔
                //xlWorkSheet.Cells[curRrowIdx, curColumnIdx] = "222";//productInfo.title;
                //xlWorkSheet.Cells[curRrowIdx, ++curColumnIdx] = "333";
                //xlWorkSheet.Cells[curRrowIdx, ++curColumnIdx] = "444";
                //for (int i = 0; i < datadt.Rows.Count; i++)
                //{
                //    for (int j = 0; j < datadt.Columns.Count; j++)
                //    {
                //        Microsoft.Office.Interop.Excel.Range range1 = (Excel.Range)xlWorkSheet.Cells[curRrowIdx + i, curColumnIdx + j];
                //        range1.NumberFormat = "@";//文本格式
                //        xlWorkSheet.Cells[curRrowIdx + i, curColumnIdx + j] = datadt.Rows[i][j].ToString();
                //    }
                //}

                string[,] strarry = new string[datadt.Rows.Count,datadt.Columns.Count];
                for (int i = 0; i < datadt.Rows.Count; i++)
                {
                    for (int j = 0; j < datadt.Columns.Count; j++)
                    {
                        strarry[i, j] = datadt.Rows[i][j].ToString();
                    }
                }
                Microsoft.Office.Interop.Excel.Range rangestart = (Excel.Range)xlWorkSheet.Cells[curRrowIdx, curColumnIdx];
                rangestart = rangestart.get_Resize(datadt.Rows.Count, datadt.Columns.Count);
                rangestart.set_Value(Excel.XlRangeValueDataType.xlRangeValueDefault, strarry);

                /*
                const int constBullerLen = 5;
                int bulletListLen = 0;
                if (productInfo.bulletArr.Length > constBullerLen)
                {
                    bulletListLen = constBullerLen;
                }
                else
                {
                    bulletListLen = productInfo.bulletArr.Length;
                }
                for (int bulletIdx = 0; bulletIdx < bulletListLen; bulletIdx++)
                {
                    xlWorkSheet.Cells[curRrowIdx, curColumnIdx + bulletIdx] = productInfo.bulletArr[bulletIdx];
                }
                curColumnIdx = curColumnIdx + bulletListLen;

                const int constImgNameListLen = 5;
                int imgNameListLen = 0;
                if (productInfo.imgFullnameArr.Length > constImgNameListLen)
                {
                    imgNameListLen = constImgNameListLen;
                }
                else
                {
                    imgNameListLen = productInfo.imgFullnameArr.Length;
                }
                for (int imgIdx = 0; imgIdx < imgNameListLen; imgIdx++)
                {
                    xlWorkSheet.Cells[curRrowIdx, curColumnIdx + imgIdx] = productInfo.imgFullnameArr[imgIdx];
                }
                curColumnIdx = curColumnIdx + imgNameListLen;

                xlWorkSheet.Cells[curRrowIdx, curColumnIdx++] = productInfo.highestPrice;
                xlWorkSheet.Cells[curRrowIdx, curColumnIdx++] = productInfo.isOneSellerIsAmazon;
                xlWorkSheet.Cells[curRrowIdx, curColumnIdx++] = productInfo.reviewNumber;
                xlWorkSheet.Cells[curRrowIdx, curColumnIdx++] = productInfo.isBestSeller;
                */

                ////http://msdn.microsoft.com/query/dev10.query?appId=Dev10IDEF1&l=ZH-CN&k=k%28MICROSOFT.OFFICE.INTEROP.EXCEL._WORKBOOK.SAVEAS%29;k%28SAVEAS%29;k%28TargetFrameworkMoniker-%22.NETFRAMEWORK%2cVERSION%3dV3.5%22%29;k%28DevLang-CSHARP%29&rd=true
                //xlWorkBook.SaveAs(
                //    Filename: fullFilename,
                //    ConflictResolution: XlSaveConflictResolution.xlLocalSessionChanges //The local user's changes are always accepted.
                //    //FileFormat : Excel.XlFileFormat.xlWorkbookNormal
                //);

                //if use above SaveAs -> will popup a window ask you overwrite it or not, even if you have set the ConflictResolution to xlLocalSessionChanges, which should not ask, should directly save
                xlWorkBook.Save();

                //http://msdn.microsoft.com/query/dev10.query?appId=Dev10IDEF1&l=ZH-CN&k=k%28MICROSOFT.OFFICE.INTEROP.EXCEL._WORKBOOK.CLOSE%29;k%28CLOSE%29;k%28TargetFrameworkMoniker-%22.NETFRAMEWORK%2cVERSION%3dV3.5%22%29;k%28DevLang-CSHARP%29&rd=true
                xlWorkBook.Close(SaveChanges: true);
            }
            catch (Exception e)
            {
                throw new Exception(e.Message);
            }
            finally
            {
                if (xlWorkSheet != null)
                    crl.ReleaseComObject(xlWorkSheet);
                if (xlWorkBook != null)
                    crl.ReleaseComObject(xlWorkBook);
                //if (xlApp != null)
                //    crl.ReleaseComObject(xlApp);//releaseObject
                CloseExcel(xlApp, xlWorkBook);
            }

        }


        /// <summary>
        /// 关闭Excel进程
        /// </summary>
        public class KeyMyExcelProcess
        {
            [DllImport("User32.dll", CharSet = CharSet.Auto)]
            public static extern int GetWindowThreadProcessId(IntPtr hwnd, out int ID);
            public static void Kill(Microsoft.Office.Interop.Excel.Application excel)
            {
                try
                {
                    IntPtr t = new IntPtr(excel.Hwnd);   //得到这个句柄，具体作用是得到这块内存入口
                    int k = 0;
                    GetWindowThreadProcessId(t, out k);   //得到本进程唯一标志k
                    System.Diagnostics.Process p = System.Diagnostics.Process.GetProcessById(k);   //得到对进程k的引用
                    p.Kill();     //关闭进程k
                }
                catch (System.Exception ex)
                {
                    throw ex;
                }
            }
        }


        //关闭打开的Excel方法
        public static void CloseExcel(Microsoft.Office.Interop.Excel.Application ExcelApplication, Microsoft.Office.Interop.Excel.Workbook ExcelWorkbook)
        {
            //ExcelWorkbook.Close(false, Type.Missing, Type.Missing);
            //ExcelWorkbook = null;
            ExcelApplication.Quit();
            GC.Collect();
            KeyMyExcelProcess.Kill(ExcelApplication);
        }


    }
}
