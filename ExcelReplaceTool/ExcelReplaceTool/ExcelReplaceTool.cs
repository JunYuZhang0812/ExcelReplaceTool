using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace ExcelReplaceTool
{
    public partial class ExcelReplaceTool : Form
    {
        public ExcelReplaceTool()
        {
            InitializeComponent();
        }
        private void m_textPath_DragEnter(object sender, DragEventArgs e)
        {
            // 对文件拖拽事件做处理 
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
                e.Effect = DragDropEffects.Link;
            else e.Effect = DragDropEffects.None;
        }
        private void m_textPath_DragDrop(object sender, DragEventArgs e)
        {
            var filePath = (string[])e.Data.GetData(DataFormats.FileDrop);
            if (filePath.Length > 0)
            {
                var file = filePath[0];
                var ex = Path.GetExtension(file);
                if (ex.Equals(".xlsm") || ex.Equals(".xlsx") || ex.Equals(".xls"))
                {
                    m_textPath.Text = file;
                }
            }
        }
        private void m_textPath_TextChanged(object sender, EventArgs e)
        {

        }

        private void m_culName_TextChanged(object sender, EventArgs e)
        {

        }

        private void m_output_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void m_btnReplace_Click(object sender, EventArgs e)
        {
            BeginReplace();
        }

        Regex m_regexStr;
        bool m_isAllCase = false;
        string m_targetStr;
        string m_excelPath;
        List<string> m_culNameArr = new List<string>();
        IWorkbook workbook;
        IWorkbook newWorkbook;
        FileStream file;
        FileStream newFile;
        private void BeginReplace()
        {
            m_excelPath = m_textPath.Text;
            if (!File.Exists(m_excelPath))
            {
                MessageBox.Show("文件不存在！");
                return;
            }
            if( m_notIgnoreCase.Checked )
            {
                m_regexStr = new Regex(m_textSource.Text);
            }
            else
            {
                m_regexStr = new Regex(m_textSource.Text, RegexOptions.IgnoreCase);
            }
            m_isAllCase = m_caseAll.Checked;
            m_targetStr = m_textTarget.Text;
            var arr = m_culName.Text.Split(';');
            m_culNameArr.Clear();
            for (int i = 0; i < arr.Length; i++)
            {
                if (!string.IsNullOrEmpty(arr[i]))
                {
                    m_culNameArr.Add(arr[i]);
                }
            }
            //FileInfo fileInfo = new FileInfo(m_excelPath);
            //WriteExcel(fileInfo);
            m_textState.Text = "";
            m_output.Items.Clear();
            m_asyncWorker.RunWorkerAsync();
        }
        /// <summary>
        /// 是要查找的列
        /// </summary>
        /// <returns></returns>
        private bool IsSearchCul(string culName)
        {
            for (int i = 0; i < m_culNameArr.Count; i++)
            {
                if (m_culNameArr[i].Equals(culName))
                {
                    return true;
                }
            }
            return false;
        }
        /*private void WriteExcel(FileInfo fileInfo)
        {
            try
            {
                using (ExcelPackage package = new ExcelPackage(fileInfo))
                {
                    for (int i = 1; i <= package.Workbook.Worksheets.Count; i++)
                    {
                        var sheet = package.Workbook.Worksheets[i];
                        var sheetName = sheet.Name;
                        for (int culIndex = 1; culIndex <= sheet.Cells.Columns; culIndex++)
                        {
                            var cell0 = sheet.Cells[0, culIndex];
                            if (cell0 == null || cell0.Value == null || string.IsNullOrEmpty(cell0.Value.ToString()) )
                            {
                                var culName = cell0.Value.ToString();
                                if (IsSearchCul(culName))
                                {
                                    for (int rowIndex = 1; rowIndex < sheet.Cells.Rows; rowIndex++)
                                    {
                                        var cell = sheet.Cells[rowIndex, culIndex];
                                        if (cell == null || cell.Value == null || string.IsNullOrEmpty(cell.Value.ToString()))
                                        {
                                            var value = cell.Value.ToString();
                                            if( m_regexStr.IsMatch(value) )
                                            {
                                                var newValue = m_regexStr.Replace(value, m_targetStr);
                                                cell.Value = newValue;
                                                m_output.Items.Add(string.Format("设置成功 >>> Sheet:{0} 行:{1} 列:{2}  {3} -> {4}", sheetName, rowIndex, culName, value, newValue));
                                            }
                                        }
                                    }
                                    
                                }
                            }
                        }
                    }
                    package.Save();
                }
            }
            catch(Exception e)
            {
                MessageBox.Show(e.ToString())
            }
        }
    }*/
        private int GetCellNum(IWorkbook workbook , int sheetIndex)
        {
            var sheet = workbook.GetSheetAt(sheetIndex);
            if (sheet == null) return 0;
            int count = 0;
            for (int rowIndex = 0; rowIndex <= sheet.LastRowNum; rowIndex++)
            {
                try
                {
                    var row = sheet.GetRow(rowIndex);
                    if( row != null)
                    {
                        count += row.LastCellNum;
                    }
                }
                catch(Exception e)
                {
                    MessageBox.Show("Sheet:" + sheet.SheetName + "  Row:" + rowIndex + "  TotalRow:" + sheet.LastRowNum + "\r\n" + e.ToString());
                }
            }
            return count;
        }
        private string m_outputStr = "";
        private string m_stateStr = "";
        private void ReadExcel()
        {
            //IWorkbook workbook = null;
            try
            {
                m_outputStr = "开始读取Excel...";
                m_asyncWorker.ReportProgress(1);
                var newFilePath = Path.GetDirectoryName(m_excelPath) + "//" + Path.GetFileNameWithoutExtension(m_excelPath) + "-备份.xlsx";
                file = new FileStream(m_excelPath, FileMode.Open, FileAccess.Read);
                newFile = new FileStream(newFilePath, FileMode.Create, FileAccess.ReadWrite);
                //IWorkbook newWorkbook = null;
                if (m_excelPath.EndsWith(".xls"))
                {
                    workbook = new HSSFWorkbook(file);
                    newWorkbook = new HSSFWorkbook();
                }
                else
                {
                    workbook = new XSSFWorkbook(file);
                    newWorkbook = new XSSFWorkbook();
                }
                if (workbook != null)
                {
                    for (int _sheetIndex = 0; _sheetIndex < workbook.NumberOfSheets; _sheetIndex++)
                    {
                        var sheetIndex = _sheetIndex;
                        var sheet = workbook.GetSheetAt(sheetIndex);
                        if (sheet != null)
                        {
                            var sheetName = sheet.SheetName;
                            var newSheet = newWorkbook.CreateSheet(sheetName);
                            Dictionary<int, string> culNameDic = new Dictionary<int, string>();
                            int culNum = sheet.GetRow(0).LastCellNum;
                            for (int _rowIndex = 0; _rowIndex <= sheet.LastRowNum; _rowIndex++)
                            {
                                var rowIndex = _rowIndex;
                                try
                                {
                                    IRow currentRow = sheet.GetRow(rowIndex);
                                    if(currentRow != null)
                                    {
                                        m_stateStr = "开始检查:" + sheetName + "    " + Math.Floor((float)rowIndex / sheet.LastRowNum * 100) + "%";
                                        m_asyncWorker.ReportProgress(2);
                                        var newRow = newSheet.CreateRow(rowIndex);
                                        for (int _colIndex = 0; _colIndex < culNum; _colIndex++)
                                        {
                                            var colIndex = _colIndex;
                                            ICell cell = currentRow.GetCell(colIndex);
                                            var newCell = newRow.CreateCell(colIndex);
                                            if (cell != null)
                                            {
                                                var value = cell.ToString();
                                                if (rowIndex == 0)
                                                {
                                                    culNameDic.Add(colIndex, value);
                                                    newCell.SetCellValue(value);
                                                }
                                                else
                                                {
                                                    var culName = culNameDic[colIndex];
                                                    bool can = false;
                                                    if( m_isAllCase )
                                                    {
                                                        can = m_regexStr.ToString().Equals(value);
                                                    }
                                                    else
                                                    {
                                                        can = m_regexStr.IsMatch(value);
                                                    }
                                                    if (IsSearchCul(culName) && can )
                                                    {
                                                        var newValue = m_regexStr.Replace(value, m_targetStr);
                                                        cell.SetCellValue(newValue);
                                                        newCell.SetCellValue(newValue);
                                                        m_outputStr = string.Format("设置成功 >>> Sheet:{0} 行:{1} 列:{2}  {3} -> {4}", sheetName, rowIndex + 1, culName, value, newValue);
                                                        m_asyncWorker.ReportProgress(1);
                                                    }
                                                    else
                                                    {
                                                        newCell.SetCellValue(value);
                                                    }
                                                }
                                            }
                                            else
                                            {
                                                if (rowIndex == 0)
                                                {
                                                    culNameDic.Add(colIndex, "");
                                                }
                                            }
                                        }
                                    }
                                }
                                catch(Exception e)
                                {
                                    MessageBox.Show("Sheet:" + sheetName + "  行:" + rowIndex +"\r\n"+e.ToString());
                                }
                            }
                        }
                    }
                    newWorkbook.Write(newFile);
                    newWorkbook.Close();
                    workbook.Close();
                    file.Close();
                    newFile.Close();
                    m_outputStr = "替换完成";
                    m_asyncWorker.ReportProgress(1);
                }
            }
            catch (Exception e)
            {
                MessageBox.Show("[" + m_excelPath + "]:" + e.ToString());
                return;
            }
        }

        
        private void m_asyncWorker_DoWork(object sender, System.ComponentModel.DoWorkEventArgs e)
        {
            ReadExcel();
        }
        private void m_asyncWorker_ProgressChanged(object sender, System.ComponentModel.ProgressChangedEventArgs e)
        {
            if(e.ProgressPercentage == 1)
            {
                m_output.Items.Add(m_outputStr);
            }
            else
            {
                m_textState.Text = m_stateStr;
            }
        }
    }
}
