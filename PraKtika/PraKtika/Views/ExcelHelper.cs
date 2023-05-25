using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;
using System.Windows;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Data.Common;
using System.Windows.Markup;
using System.Windows.Media.Imaging;
using static System.Net.Mime.MediaTypeNames;
using Microsoft.Win32;
using Microsoft.Office.Core;
using System.Runtime.InteropServices;
using System.Windows.Media;
using System.Windows.Data;
using Image = System.Windows.Controls.Image;
using System.Reflection;
using System.Runtime.Remoting.Messaging;
using System.Windows.Documents;

namespace PraKtika.Views
{
    internal class ExcelHelper : IDisposable 
    {
        private Excel.Application _excel;
        private Workbook _workbook;
        private string _filePath;
        public ExcelHelper()
        {
            _excel = new Excel.Application();
        }
        internal bool Open(string filePath)
        {
            try
            {
                if (File.Exists(filePath))
                {
                    _workbook = _excel.Workbooks.Open(filePath);
                }
                else
                {
                    _workbook = _excel.Workbooks.Add();
                    _filePath = filePath;
                }
                return true;
            }
            catch (Exception ex) { MessageBox.Show(ex.Message); }
            return false;
        }
        public void Dispose()
        {
            try
            {
                _excel.Workbooks.Close();
            }
            catch (Exception ex) { MessageBox.Show(ex.Message); }
        }
        internal bool Set(string column, int row, object data)
        {
            try
            {
                ((Excel.Worksheet)_excel.ActiveSheet).Cells[row, column] = data;
                return true;
            }
            catch (Exception ex) { MessageBox.Show(ex.Message); }
            return false;
        }
        internal void Save()
        {
            if (!string.IsNullOrEmpty(_filePath))
            {
                _workbook.SaveAs(_filePath);
                _filePath = null;
            }
            else { _workbook.Save(); }
        }
        internal string[,] Export(string[,] list, int lastColumn, int lastrow)
        {

            for (int j = 0; j < lastColumn; j++)
            {
                for (int i = 0; i < lastrow; i++)
                {
                    list[i, j] = ((Excel.Worksheet)_excel.ActiveSheet).Cells[i + 1, j + 1].Text.ToString();
                }
            }
            return list;
        }
        internal int Heigthq()
        {
            var lastCell = ((Excel.Worksheet)_excel.ActiveSheet).Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);
            int lastRow = (int)lastCell.Row;
            if (lastCell.Text == "")
            {
                return 0;
            }
            else 
            {
                return lastRow;
            }
        }

        internal int Length()
        {
            var lastCell = ((Excel.Worksheet)_excel.ActiveSheet).Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);
            int lastColumn = (int)lastCell.Column;
            if (lastCell.Text == "") 
            {
                return 0;
            }
            else
            {
                return lastColumn;
            }
            
        }
        internal void CreateSpisok(string[,] list, ComboBox ComboExcibits , ExcelHelper helper)
        {
            try
            {    
                    int height = helper.Heigthq();
                    int length = helper.Length();
                    list = helper.Export(list, length, height);
                    string s = "";
                    for (int i = 0; i < height; i++)
                    {
                        s = "";
                        for (int j = 0; j < 1; j++)
                            s += "  " + list[i, j];
                        ComboExcibits.Items.Add(s);
                    }   
            }
            catch (Exception ex) { MessageBox.Show(ex.Message); }
        }
        internal void CreateInfo(string[,] list,int index, System.Windows.Controls.ListView ComboExcibits, ExcelHelper helper) 
        {
            try
            {
                int height = helper.Heigthq();
                int length = helper.Length(); 
                list = helper.Export(list, length, height);
               

                for (int i = 0; i < length - 1; i++)
                {
                    var s = "";
                    s += "  " + list[index, i];
                    ComboExcibits.Items.Add(s);  
                }
                
            }
            catch (Exception ex) { MessageBox.Show(ex.Message); }
        }
        internal void DeleteInfo(ExcelHelper helper, int index) 
        {
            try
            {
                for (int i = 1; i <= helper.Heigthq(); i++)
                {
                    if (i == index+1)
                    {
                        ((Excel.Worksheet)_excel.ActiveSheet).Rows[i].Delete();
                        break;
                    }
                }
            }
            catch (Exception ex) { MessageBox.Show(ex.Message); }
        
        }
    }
}
