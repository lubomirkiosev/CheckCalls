namespace CheckCalls1._1
{
    using Microsoft.Win32;
    using System;
    using System.Diagnostics;
    using System.IO;
    using System.Threading.Tasks;
    using System.Windows;

    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        OpenFileDialog openFileDialog = new OpenFileDialog();

        string path = Environment.CurrentDirectory + @"\path.txt";
        string target = string.Empty;

        private void BtnOpenXlxsFiles_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                target = StreamReader(path);

                openFileDialog = new OpenFileDialog();
                openFileDialog.Multiselect = false;
                openFileDialog.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
                openFileDialog.RestoreDirectory = true;
                openFileDialog.InitialDirectory = target;
                if (openFileDialog.ShowDialog() == true)
                    xlxsFile.Text = openFileDialog.FileName;
            }
            catch (Exception)
            {
                MessageBox.Show("Error!");
            }

        }

        private async void BtnSubmitFiles_Click(object sender, RoutedEventArgs e)
        {

            btnSubmitFile.IsEnabled = false;

            Stopwatch stopWatch = new Stopwatch();
            stopWatch.Start();

            try
            {
                var inputLine = input.Text.ToString();
                string[] paths = new string[2] { inputLine, xlxsFile.Text };

                var fullFilePath = paths[1];

                CreatePathFile(path, fullFilePath);

                await Task.Run(() => Function.Main(paths));

                MessageBoxResult result = MessageBox.Show($"Time: {stopWatch.Elapsed.ToString(@"mm\:ss\.ff")}min. Ready. Do you want to close this window?",
                                      "Confirmation",
                                      MessageBoxButton.YesNo,
                                      MessageBoxImage.Question);
                if (result == MessageBoxResult.Yes)
                {
                    Application.Current.Shutdown();
                }

            }
            catch (IOException)
            {
                MessageBox.Show("The .xlxs file is open!");
                btnSubmitFile.IsEnabled = true;
            }
            catch (Exception)
            {
                MessageBox.Show("Error!");
                btnSubmitFile.IsEnabled = true;
            }


        }

        private string StreamReader(string path)
        {
            if (File.Exists(path))
            {
                using (StreamReader sr = File.OpenText(path))
                {
                    return sr.ReadToEnd();
                }
            }
            else
            {
                return Environment.SpecialFolder.Desktop.ToString();
            }
        }

        private void CreatePathFile(string path, string fullFilePath)
        {

            var lastIndex = fullFilePath.LastIndexOf('\\');
            var folderPath = fullFilePath.Substring(0, lastIndex);

            if (!File.Exists(path))
            {
                using (StreamWriter sw = File.CreateText(path))
                {
                    sw.Write(folderPath);
                }
            }
            else
            {
                using (StreamReader sr = File.OpenText(path))
                {
                    target = sr.ReadToEnd();
                }
                if (target != folderPath)
                {
                    using (StreamWriter sw = File.CreateText(path))
                    {
                        sw.Write(folderPath);
                    }
                }
            }
        }
    }
}

