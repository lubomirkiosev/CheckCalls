using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Diagnostics;
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

namespace CheckCalls1._1
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        OpenFileDialog openFileDialog = new OpenFileDialog();

        private void btnOpenTxtFiles_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                openFileDialog = new OpenFileDialog();
                openFileDialog.Multiselect = false;
                openFileDialog.Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*";
                openFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                if (openFileDialog.ShowDialog() == true)
                    txtFile.Text = openFileDialog.FileName;
            }
            catch (Exception)
            {
                MessageBox.Show("Error!");
            }

        }

        private void btnOpenXlxsFiles_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                openFileDialog = new OpenFileDialog();
                openFileDialog.Multiselect = false;
                openFileDialog.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
                openFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                if (openFileDialog.ShowDialog() == true)
                    xlxsFile.Text = openFileDialog.FileName;
            }
            catch (Exception)
            {
                MessageBox.Show("Error!");
            }

        }

        private async void btnSubmitFiles_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Stopwatch stopWatch = new Stopwatch();
                stopWatch.Start();
                string[] paths = new string[2] { txtFile.Text, xlxsFile.Text, };
                await Function.Main(paths);

                MessageBoxResult result = MessageBox.Show($"Time: {stopWatch.Elapsed.ToString(@"mm\:ss\.ff")}min. Ready. Do you want to close this window?",
                                          "Confirmation",
                                          MessageBoxButton.YesNo,
                                          MessageBoxImage.Question);
                if (result == MessageBoxResult.Yes)
                {
                    Application.Current.Shutdown();
                }

            }
            catch (Exception)
            {
                MessageBox.Show("Error!");
            }
        }
    }
}
