using DMFolderWatcher.Class;
using Microsoft.Office.Interop.Excel;
using System;
using System.IO;
using System.Security.Permissions;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace DMFolderWatcher
{
    public partial class Form1 : Form
    {
        CFolder cFolder = new CFolder();
        FileSystemWatcher watcher = new FileSystemWatcher();
        Excel.Application excel = new Excel.Application();
        int counter_hoja1 = 0;

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            DialogResult result = fbd.ShowDialog();

            if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(fbd.SelectedPath))
            {
                cFolder.Path = fbd.SelectedPath;
                cFolder.Parent = Directory.GetParent(fbd.SelectedPath).FullName;
            }

            Watcher();
        }


        [PermissionSet(SecurityAction.Demand, Name = "FullTrust")]
        private void Watcher() {
           
            watcher.Path = cFolder.Path;          
            watcher.NotifyFilter = NotifyFilters.Attributes
                                 | NotifyFilters.CreationTime
                                 | NotifyFilters.DirectoryName
                                 | NotifyFilters.FileName
                                 | NotifyFilters.LastAccess
                                 | NotifyFilters.LastWrite
                                 | NotifyFilters.Security
                                 | NotifyFilters.Size;
            watcher.Filter = "*.*";         
            watcher.Created += OnCreated;           
            watcher.EnableRaisingEvents = true;
        }

        private void OnCreated(object sender, FileSystemEventArgs e)
        {
            if (e.ChangeType != WatcherChangeTypes.Created)
            {
                return;
            }


            try
            {
                if (e.Name.Split('.')[1] == "xlsx")
                {
                    this.CopyToFile(e);
                }
                else
                {
                    File.Move(e.FullPath, cFolder.NotApplicable_folder + "\\" + e.Name);
                }               
            }
            catch(Exception ex_lop)
            {
                Console.WriteLine($"Error at : {ex_lop.Message}");
            }
        }

        private void CopyToFile(FileSystemEventArgs e)
        {
            Workbook wb = null;
            Workbook wbMaster = null;

            try
            {
                wb = excel.Workbooks.Open(e.FullPath);
                wbMaster = excel.Workbooks.Open(cFolder.Master_file);

                int numberSheets = wb.Worksheets.Count;

                for (int i = 1; i < numberSheets + 1; i++)
                {
                    Worksheet wsWorkbookToCopy = (Worksheet)wb.Worksheets[i];
                    Worksheet wsMasterCopy = (Worksheet)wbMaster.Worksheets.Add();

                    if (wsWorkbookToCopy.Name.Equals("Hoja1"))
                    {
                        wsMasterCopy.Name = "Hoja1_" + counter_hoja1;
                        counter_hoja1++;
                    }
                    else
                    {
                        wsMasterCopy.Name = wsWorkbookToCopy.Name;
                    }
                    wsWorkbookToCopy.UsedRange.Copy(wsMasterCopy.Range["A1"]);
                }
            }
            catch (IOException ex)
            {
                Console.WriteLine("Error open file {0}", ex.Message);
            }
            catch (Exception genError)
            {
                Console.WriteLine("General error {0}", genError.Message);
            }
            finally
            {
                if (wb != null && wbMaster != null)
                {
                    wb.Close(true);
                    wbMaster.Close(true);
                    File.Move(e.FullPath, cFolder.Processed_folder + "\\" + e.Name);
                }
            }
        }

    }


    
}
