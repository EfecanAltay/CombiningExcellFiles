using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows.Forms;
using Application = Microsoft.Office.Interop.Excel.Application;
using DataTable = System.Data.DataTable;

namespace CombiningExcellSheets
{
    public partial class Form1 : Form
    {
        public class ExcellFile
        {
            public string FileName = "";
            public string Path = "";
            public int RowSize = 0;

            public override string ToString()
            {
                if (RowSize != 0)
                    return FileName + " (rw_size : " + RowSize + ")";
                else
                    return FileName;
            }
        }

        public ExcellFile outFile;
        public int Row_Count = 0;

        ObservableCollection<ExcellFile> excellFilePaths = new ObservableCollection<ExcellFile>();
        public Form1()
        {
            InitializeComponent();
            excellFilePaths = new ObservableCollection<ExcellFile>();
        }

        private void Btn_loadExcells_Click(object sender, EventArgs e)
        {
            GetOpenAndFileType("Excell Dosyası |*.xlsx");
        }

        private void Btn_join_Click(object sender, EventArgs e)
        {
            ClearConsole();
            
            if (outFile == null)
            {
                MessageBox.Show("Çıkış Konumu Seçmediniz !!", "Eksik Giriş", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (excellFilePaths.Count > 0)
            {
                ReadAndJoinExcells();
            }
            else
            {
                MessageBox.Show("Hiç Bir Excell Girişi yapılmadı!", "Eksik Giriş", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private async Task Run()
        {
            await File.AppendText("temp.dat").WriteAsync("a");
            label1.Text = "test";
        }

        private void UpdatePathList(string[] fileNames)
        {
            excellFilePaths.Clear();
            listbox_excellPaths.Items.Clear();
            foreach (var path in fileNames)
            {
                excellFilePaths.Add(new ExcellFile() { Path = path, FileName = Path.GetFileName(path) });
            }
            listbox_excellPaths.Items.AddRange(excellFilePaths.ToArray());
        }

        private void AddPathList(string[] fileNames)
        {
            listbox_excellPaths.Items.Clear();
            foreach (var path in fileNames)
            {
                excellFilePaths.Add(new ExcellFile() { Path = path, FileName = Path.GetFileName(path) });
            }
            listbox_excellPaths.Items.AddRange(excellFilePaths.ToArray());
        }

        private void GetOpenAndFileType(string fileFilter , bool addExcell = false)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Multiselect = true;
            openFileDialog.Filter = fileFilter;
            if(addExcell)
                openFileDialog.FileOk += OpenFileDialog_FileOk2;
            else
                openFileDialog.FileOk += OpenFileDialog_FileOk;
            openFileDialog.ShowDialog();
        }

        private void OpenFileDialog_FileOk(object sender, CancelEventArgs e)
        {
            if (!e.Cancel)
            {
                OpenFileDialog opd = sender as OpenFileDialog;
                UpdatePathList(opd.FileNames);
            }
        }

        private void OpenFileDialog_FileOk2(object sender, CancelEventArgs e)
        {
            if (!e.Cancel)
            {
                OpenFileDialog opd = sender as OpenFileDialog;
                AddPathList(opd.FileNames);
            }
        }

        private async void ReadAndJoinExcells()
        {
            int startingRowIndex = 1;
            bool isFirst = true;
            int FileSizeIndex = 0;
            int maxfileSize = excellFilePaths.Count;

            Application application = null;
            Workbooks wbooks = null;
            Worksheet destinationSheet = null;
            Workbook destinationWorkbook = null;
            try
            {
                WriteConsole("Kopyalama Başlıyor...");
                await Task.Run(() =>
                {
                    //Instantiate the application object
                    application = new Application();
                    application.CutCopyMode = (Microsoft.Office.Interop.Excel.XlCutCopyMode)0;
                    wbooks = application.Workbooks;
                    object misValue = System.Reflection.Missing.Value;

                    destinationWorkbook = wbooks.Add(misValue);
                    destinationSheet = (Worksheet)destinationWorkbook.Worksheets.get_Item(1);
                    //Copy Excel worksheet from source workbook to the destination workbook
                });
                foreach (var item in excellFilePaths)
                {
                    FileSizeIndex++;
                    WriteConsole($" {FileSizeIndex}/{maxfileSize} {item.FileName} Kopyalanıyor...");
                    await Task.Run(() =>
                    {
                        Workbook sourceWorkbook = wbooks.Open(item.Path);
                        Worksheet wsheet = (Worksheet)sourceWorkbook.Worksheets.get_Item(1);
                        if (isFirst)
                        {
                            startingRowIndex = 1;
                        }
                        else
                        {
                            startingRowIndex = 2;
                        }
                        var usedRange = wsheet.UsedRange;
                        usedRange.Cells[1,1].Copy();
                        usedRange.Range[usedRange[startingRowIndex, 1], usedRange[usedRange.Rows.Count, usedRange.Columns.Count]].Copy();
                        if (isFirst)
                        {
                            destinationSheet.Paste(destinationSheet.UsedRange.Cells[1, 1]);
                            isFirst = false;
                        }
                        else
                        {
                            destinationSheet.Paste(destinationSheet.UsedRange.Cells[destinationSheet.UsedRange.Rows.Count + 1, 1]);
                        }
                        sourceWorkbook?.Close();
                        Marshal.ReleaseComObject(sourceWorkbook);
                    });
                }

                WriteConsole("Çıktı oluşturuluyor...");
                //Save the file

                await Task.Run(() =>
                {
                    destinationWorkbook.SaveAs(outFile.Path, AccessMode: XlSaveAsAccessMode.xlShared);
                });

            }
            catch
            {

            }
            finally
            {
                WriteConsole("Çıktın Hazır :) Güle güle kullan...");
                try
                {
                    destinationWorkbook?.Close();
                    Marshal.ReleaseComObject(destinationWorkbook);
                    wbooks?.Close();
                    Marshal.ReleaseComObject(wbooks);
                    application?.Quit();
                    Marshal.FinalReleaseComObject(application);
                }
                catch (Exception) { }
                application = null;
            }
        }

        private void Btn_select_out_Click(object sender, EventArgs e)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Excell Dosyası |*.xlsx";
            saveFileDialog.FileOk += SaveFileDialog_FileOk; ;
            saveFileDialog.ShowDialog();
        }

        private void WriteConsole(string text)
        {
            lstBox_output.Items.Add(text);
            lstBox_output.SelectedIndex = lstBox_output.Items.Count - 1;
        }
        private void ClearConsole()
        {
            lstBox_output.Items.Clear();
        }

        private void SaveFileDialog_FileOk(object sender, CancelEventArgs e)
        {
            if (!e.Cancel)
            {
                SaveFileDialog sfd = sender as SaveFileDialog;
                outFile = new ExcellFile()
                {
                    Path = sfd.FileName
                };
                lstbox_outPath.Items.Clear();
                lstbox_outPath.Items.Add(sfd.FileName);
            }
        }

        private void Btn_show_totalRowSize_Click(object sender, EventArgs e)
        {
            CalculateRowSize();
        }

        private async void CalculateRowSize()
        {
            int row_size = 0;

            Application application = null;
            Workbooks wbooks = null;
            Workbook sourceWorkbook = null;
            try
            {
                application = new Application();
                wbooks = application.Workbooks;
                listbox_excellPaths.Enabled = false;
                foreach (ExcellFile item in excellFilePaths)
                {
                    if (item.RowSize == 0)
                    {
                        WriteConsole($"{item.FileName} Boyutu Okunuyor...");
                        listbox_excellPaths.SelectedItem = item;
                        await Task.Run(() =>
                        {
                            sourceWorkbook = wbooks.Open(item.Path);
                            Worksheet wsheet = (Worksheet)sourceWorkbook.Worksheets.get_Item(1);
                            row_size += wsheet.UsedRange.Rows.Count;
                            item.RowSize = wsheet.UsedRange.Rows.Count;
                            sourceWorkbook?.Close();
                            Marshal.ReleaseComObject(sourceWorkbook);
                        });
                        int index = listbox_excellPaths.Items.IndexOf(item);
                        listbox_excellPaths.Items.Remove(item);
                        listbox_excellPaths.Items.Insert(index, item);
                    }
                    else {
                        row_size += item.RowSize;
                    }
                    SetRowCount(row_size);
                }
                listbox_excellPaths.Enabled = true;
                wbooks?.Close();
                Marshal.ReleaseComObject(wbooks);
                application?.Quit();
                Marshal.ReleaseComObject(application);
                WriteConsole($"Toplam Satır Boyutu {row_size} olarak okundu.");
                MessageBox.Show($"{row_size}", "Toplam Satır Sayısı");
            }
            catch { }
            finally
            {
                try
                {
                    //sourceWorkbook?.Close();
                    //Marshal.ReleaseComObject(sourceWorkbook);

                }
                catch (Exception) { }
                application = null;
            }
        }

        private void SetRowCount(int row_Count)
        {
            string row_count_str = row_Count.ToString();

         
            if(row_Count >= 1000000)
            {
                lbl_rowCount.ForeColor = Color.Red;
            }
            else
            {
                lbl_rowCount.ForeColor = Color.Green;
            }

            lbl_rowCount.Text = row_count_str;
            Row_Count = row_Count;
        }

        private void Listbox_excellPaths_KeyDown(object sender, KeyEventArgs e)
        {
            System.Windows.Forms.ListBox pathListBox = (System.Windows.Forms.ListBox)sender;
            ExcellFile selectedItem = pathListBox.SelectedItem as ExcellFile;
            if (e.KeyCode == Keys.Delete && selectedItem != null)
            {
                listbox_excellPaths.Items.Remove(selectedItem);
                excellFilePaths.Remove(selectedItem);
                if(selectedItem.RowSize != 0 && Row_Count != 0 )
                    SetRowCount(Row_Count - selectedItem.RowSize);
            }
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            GetOpenAndFileType("Excell Dosyası |*.xlsx" , true);
        }
    }
}
