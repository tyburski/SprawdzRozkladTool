using Microsoft.Win32;
using System.ComponentModel.DataAnnotations.Schema;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Windows;
using System.Windows.Documents;
using System.Windows.Media;


namespace SprawdzRozklad
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        string zipFilePath = "";
        string firma = "";
        private void OpenFile_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Title = "Wybierz plik bazy";
            openFileDialog.Filter = "Plik bazy (*.zip)|*.zip";

            if (openFileDialog.ShowDialog() == true)
            {
                string[] parts = openFileDialog.FileName.Split("_");
                string[] withoutExt = parts[2].Split(".");
                zipFilePath = openFileDialog.FileName;
                string archiwumInfo = $"ArchiwumInfo_{withoutExt[0]}.rtf";

                using (ZipArchive archive = ZipFile.OpenRead(zipFilePath))
                {
                    var entry = archive.GetEntry(archiwumInfo);

                    if (entry != null)
                    {
                        using (StreamReader reader = new StreamReader(entry.Open(), Encoding.UTF8))
                        {
                            var rtb = new System.Windows.Controls.RichTextBox();
                            rtb.Selection.Load(entry.Open(), DataFormats.Rtf);
                            string plainText = new TextRange(rtb.Document.ContentStart, rtb.Document.ContentEnd).Text;                          
                            string[] lines = plainText.Split("\n");
                            string dane = lines.FirstOrDefault(x=>x.Contains("Dane firmy:"));

                            string[] line = dane.Split(":");
                            firma = line[1];
                            string lic = lines.FirstOrDefault(x => x.Contains("Licencja:"));

                            companyNr.Text = $"{dane}\n{lic}";
                        }
                    }
                    else
                    {
                        MessageBox.Show($"Nie znaleziono pliku '{archiwumInfo}' w archiwum.");
                    }
                    textBox.Text = "";
                }
            }
        }

        private async void Start_Click(object sender, RoutedEventArgs e)
        {
            textBox.Text = "";
            await CheckLines();
            int charCount = (int)(textBox.ActualWidth / 5);
            textBox.Text += ("\n" + new string('_', charCount) +"\n\n");
            //await CheckRoutes();
        }

        private async Task CheckLines()
        {
            string tempPathTrasy = "";
            string trasyFile = "trasylinii.dbf";
            string tempPathLinie = "";
            string linieFile = "linie.dbf";

            using (ZipArchive archive = ZipFile.OpenRead(zipFilePath))
            {

                var entryTrasy = archive.Entries
                    .FirstOrDefault(e => e.Name.Equals(trasyFile, StringComparison.OrdinalIgnoreCase));
                var entryLinie = archive.Entries
                   .FirstOrDefault(e => e.Name.Equals(linieFile, StringComparison.OrdinalIgnoreCase));
                if (entryTrasy != null)
                {
                    tempPathTrasy = Path.Combine(Path.GetTempPath(), entryTrasy.Name.ToLower());
                    entryTrasy.ExtractToFile(tempPathTrasy, true);
                    textBox.Text += $"Plik TRASYLINII.DBF zapisany tymczasowo: {tempPathTrasy}\n";
                }
                if (entryLinie != null)
                {
                    tempPathLinie = Path.Combine(Path.GetTempPath(), entryLinie.Name.ToLower());
                    entryLinie.ExtractToFile(tempPathLinie, true);
                    textBox.Text += $"Plik LINIE.DBF zapisany tymczasowo: {tempPathLinie}\n";
                }
            }
            List<DataRow> wszystkieTrasyLinii;
            string trasyDbPath = Path.GetDirectoryName(tempPathTrasy);
            string connStrTrasy = $@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={trasyDbPath};Extended Properties=dBASE IV;";
            using (OleDbConnection conn = new OleDbConnection(connStrTrasy))
            {
                conn.Open();
                string query = $"SELECT * FROM [{trasyFile}]";
                OleDbDataAdapter adapter = new OleDbDataAdapter(query, conn);
                DataTable dt = new DataTable();
                adapter.Fill(dt);

                wszystkieTrasyLinii = dt.AsEnumerable().ToList();

            }

            List<DataRow> filteredRows = wszystkieTrasyLinii
                .Where(r => wszystkieTrasyLinii.Any(s=>s["NRF"].ToString().Trim() == firma))
                    .OrderBy(r => int.Parse(r["NRLINII"].ToString().Trim()))
                    .ThenBy(r => int.Parse(r["WARLINII"].ToString().Trim()))
                    .ThenBy(r => int.Parse(r["NRPRZYST"].ToString().Trim()))
                    .ToList();


            textBox.Text += $"\nRozpoczynam sprawdzanie pliku TRASYLINII.DBF  filtered:{filteredRows.Count}\n";
            textBox.Text += "Sprawdzanie powtórzeń przystanków w kolejnych rekordach:\n";

            int lastNr = -1;
            string lastStopName = "";
            string lastKod2 = "";
            int lastStopNr = -1;
            int checkedLines = 0;
            int errors = 0;

            foreach (DataRow row in filteredRows)
            {
                int currentNr = int.Parse(row["NRLINII"].ToString().Trim());
                int currentStop = int.Parse(row["NRPRZYST"].ToString().Trim());
                string currentKod2 = row["Kod2"].ToString().Trim();
                string currentStopName = row["NAZPRZYST"].ToString().Trim();
                string wariant = row["WARLINII"].ToString().Trim();
                string[] dateFrom = row["WAZNAOD"].ToString().Split(" ");

                bool newLine = currentNr != lastNr || currentStop <= lastStopNr;
                if (newLine)
                {
                    checkedLines++;
                    lastKod2 = "";
                    lastStopName = "";
                }

                if (!newLine && currentKod2 == lastKod2 && currentStopName == lastStopName)
                {
                  
                    textBox.Text += $"\nUWAGA Powtarzający się przystanek: Linia: {currentNr}, Wariant: {wariant}, Ważna Od: {dateFrom[0]}, Nazwa Przystanku: {currentStopName}, Kod2: {currentKod2}\n";
                    errors++;
                    
                    Application.Current.Dispatcher.Invoke(
                        System.Windows.Threading.DispatcherPriority.Background,
                        new Action(delegate { }));                
                }
                lastNr = currentNr;
                lastStopNr = currentStop;
                lastKod2 = currentKod2;
                lastStopName = currentStopName;
            }
            if (errors == 0)
            {             
                textBox.Text += $"\nNie znaleziono błędów.\n";              
            }
            else
            {
                textBox.Text += $"\nZnaleziono błędów: {errors}.\n";
            }
                textBox.Text += $"\nProcedura zakończona. Sprawdzono {checkedLines} linii.\n";          
        }

        private async Task CheckRoutes()
        {
            string tempPathTrasy = "";
            string trasyFile = "trasykursow.dbf";

            using (ZipArchive archive = ZipFile.OpenRead(zipFilePath))
            {

                var entryTrasy = archive.Entries
                    .FirstOrDefault(e => e.Name.Equals(trasyFile, StringComparison.OrdinalIgnoreCase));
                if (entryTrasy != null)
                {
                    tempPathTrasy = Path.Combine(Path.GetTempPath(), entryTrasy.Name.ToLower());
                    entryTrasy.ExtractToFile(tempPathTrasy, true);
                    textBox.Text += $"Plik TRASYKURSOW.DBF zapisany tymczasowo: {tempPathTrasy}\n";
                }
            }
            List<DataRow> filteredRows;
            string trasyDbPath = Path.GetDirectoryName(tempPathTrasy);
            string connStrTrasy = $@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={trasyDbPath};Extended Properties=dBASE IV;";
            using (OleDbConnection conn = new OleDbConnection(connStrTrasy))
            {
                conn.Open();
                string query = $"SELECT * FROM [{trasyFile}]";
                OleDbDataAdapter adapter = new OleDbDataAdapter(query, conn);
                DataTable dt = new DataTable();
                adapter.Fill(dt);


                filteredRows = dt.AsEnumerable()
                    .OrderBy(r => int.Parse(r["NRKURSU"].ToString().Trim()))
                    .ThenBy(r => int.Parse(r["WARIANT"].ToString().Trim()))
                    .ThenBy(r => int.Parse(r["NRPRZYST"].ToString().Trim()))
                    .ToList();
            }

            textBox.Text += "\nRozpoczynam sprawdzanie pliku TRASYKURSOW.DBF\n";
            textBox.Text += "Sprawdzanie powtórzeń przystanków w kolejnych rekordach:\n";

            int lastNr = -1;
            string lastStopName = "";
            string lastKod2 = "";
            int lastStopNr = -1;
            int checkedLines = 0;
            int errors = 0;

            foreach (DataRow row in filteredRows)
            {
                int currentNr = int.Parse(row["NRKURSU"].ToString().Trim());
                int currentStop = int.Parse(row["NRPRZYST"].ToString().Trim());
                string currentKod2 = row["KOD2"].ToString().Trim();
                string currentStopName = row["NAZPRZYST"].ToString().Trim();
                string wariant = row["WARIANT"].ToString().Trim();
                string[] dateFrom = row["WAZNYOD"].ToString().Split(" ");

                bool newLine = currentNr != lastNr || currentStop <= lastStopNr;
                if (newLine)
                {
                    checkedLines++;
                    lastKod2 = "";
                    lastStopName = "";
                }

                if (!newLine && currentKod2 == lastKod2 && currentStopName == lastStopName)
                {

                    textBox.Text += $"\nUWAGA Powtarzający się przystanek: Kurs: {currentNr}, Wariant: {wariant}, Ważny Od: {dateFrom[0]}, Nazwa Przystanku: {currentStopName}, Kod2: {currentKod2}\n";
                    errors++;

                    Application.Current.Dispatcher.Invoke(
                        System.Windows.Threading.DispatcherPriority.Background,
                        new Action(delegate { }));
                }
                lastNr = currentNr;
                lastStopNr = currentStop;
                lastKod2 = currentKod2;
                lastStopName = currentStopName;
            }
            if (errors == 0)
            {
                textBox.Text += $"\nNie znaleziono błędów.\n";
            }
            else
            {
                textBox.Text += $"\nZnaleziono błędów: {errors}.\n";
            }
            textBox.Text += $"\nProcedura zakończona. Sprawdzono {checkedLines} kursów.\n";
        }
        private void Clear_Click(object sender, RoutedEventArgs e)
        {
            textBox.Text = "";
        }
        
    }
}