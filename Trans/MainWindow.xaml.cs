using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace Trans
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            if ((bool)this.isImportCB.IsChecked)
            {//import to xaml

                doImportExcelToXaml();
            }
            else
            {//export to excel
                doExportXamlDirToExcel();
            }
        }

        void doImportExcelToXaml()
        {
            Microsoft.Win32.OpenFileDialog openFile = new Microsoft.Win32.OpenFileDialog();
            // openFile.Filter = "WMV (*.wmv)|*.wmv|MP4 (*.mp4)|*.mp4|AVI (*.avi)|*.avi|All files|*.*";
            openFile.Filter = "Excel Files(*.xls;*.xlsx;*.xml)|*.xls;*.xlsx;*.xml|All files (*.*)|*.*";
            openFile.Title = "find language excel to import";
            if (openFile.ShowDialog() == true)
            {
                String fileName=openFile.FileName;
                this.tv_path.Text = fileName;

                String outXamlDirPath = "abcd";
                Boolean ret=BackWorker.generateXamlResourceFileFromExcel(fileName, outXamlDirPath, String.Empty);

                if (ret)
                {
                    System.Diagnostics.Process.Start("Explorer", System.Environment.CurrentDirectory + "\\" + outXamlDirPath);
                }
                else
                {
                    System.Windows.MessageBox.Show(this, "error happen", "error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        void doExportXamlDirToExcel()
        {
            System.Windows.Forms.FolderBrowserDialog foldBrowser = new System.Windows.Forms.FolderBrowserDialog();
            foldBrowser.Reset();
            foldBrowser.Description = "Open Xaml Dir";
            foldBrowser.RootFolder = Environment.SpecialFolder.DesktopDirectory;
            foldBrowser.ShowNewFolderButton = false;

            //foldBrowser.ShowDialog();
            if (foldBrowser.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                Boolean ret = false;
                string outputFileName = "outputTrans.xls";
                try
                {
                    String path=foldBrowser.SelectedPath;
                    this.tv_path.Text = path;

                    ret = BackWorker.exportToExcel(path, outputFileName);
                    //System.Environment.CurrentDirectory +
                }
                catch (Exception e)
                {
                    ret = true;
                }

                if (!ret)
                {
                    System.Windows.MessageBox.Show(this,"error happen","error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
                else
                {
                    System.Diagnostics.Process.Start(System.Environment.CurrentDirectory+"\\"+outputFileName);
                }

            }
        }

        private void isImportCB_Checked(object sender, RoutedEventArgs e)
        {

            io_tab.Header = this.FindResource("import");
            this.tv_path.Text = "";
        }

        private void isImportCB_Unchecked(object sender, RoutedEventArgs e)
        {
            io_tab.Header = this.FindResource("export");
            this.tv_path.Text = "";
        }
    }
}
