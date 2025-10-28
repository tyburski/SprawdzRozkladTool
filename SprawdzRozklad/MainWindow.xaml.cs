using Microsoft.Win32;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Windows;
using System.Windows.Documents;


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

        string zipFilePath = String.Empty;
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
                            string lic = lines.FirstOrDefault(x => x.Contains("Licencja:"));

                            companyNr.Text = $"{dane}\n{lic}";
                        }
                    }
                    else
                    {
                        MessageBox.Show($"Nie znaleziono pliku '{archiwumInfo}' w archiwum.");
                    }
                }
            }
        }

        private void Start_Click(object sender, RoutedEventArgs e)
        {
            textBox.Text = "";
            CheckLines();
        }

        private void CheckLines()
        {
            string tempPathTrasy = "";
            string tempPathLinie = "";
            string trasyFile = "trasylinii.dbf";
            string linieFile = "linie.dbf";

            using (ZipArchive archive = ZipFile.OpenRead(zipFilePath))
            {
                var entryLinie = archive.Entries
                    .FirstOrDefault(e => e.Name.Equals(linieFile, StringComparison.OrdinalIgnoreCase));
                if (entryLinie != null)
                {
                    tempPathLinie = Path.Combine(Path.GetTempPath(), entryLinie.Name.ToLower());
                    entryLinie.ExtractToFile(tempPathLinie, true);
                    textBox.Text += $"Plik Linie.DBF zapisany tymczasowo: {tempPathLinie}\n";
                }

                var entryTrasy = archive.Entries
                    .FirstOrDefault(e => e.Name.Equals(trasyFile, StringComparison.OrdinalIgnoreCase));
                if (entryTrasy != null)
                {
                    tempPathTrasy = Path.Combine(Path.GetTempPath(), entryTrasy.Name.ToLower());
                    entryTrasy.ExtractToFile(tempPathTrasy, true);
                    textBox.Text += $"Plik TRASYLINII.DBF zapisany tymczasowo: {tempPathTrasy}\n";
                }
            }


            HashSet<string> activeLines = new HashSet<string>();
            string linieDbPath = Path.GetDirectoryName(tempPathLinie);
            string connStrLinie = $@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={linieDbPath};Extended Properties=dBASE IV;";

            using (OleDbConnection conn = new OleDbConnection(connStrLinie))
            {
                conn.Open();
                string query = $"SELECT * FROM [{linieFile}]";
                OleDbDataAdapter adapter = new OleDbDataAdapter(query, conn);
                DataTable dtLinie = new DataTable();
                adapter.Fill(dtLinie);

                activeLines = dtLinie.AsEnumerable()
                    .Where(r =>
                    {
                        if (r["STATLINII"].ToString().Trim() != "2")
                            return false;

                        if (r["NRF"].ToString().Trim() != "1211")
                            return false;

                        string dateStr = r["WAZNADO"].ToString().Trim();
                        if (string.IsNullOrEmpty(dateStr))
                            return true;
                        if (DateTime.TryParse(dateStr, out DateTime dtWaznado))
                            return dtWaznado > new DateTime(2023, 1, 1);
                        return false;
                    })
                    .Select(r => r["NRLINII"].ToString().Trim())
                    .Where(nr => !string.IsNullOrEmpty(nr))
                    .ToHashSet();
            }

            List<DataRow> filteredRows;
            string trasyDbPath = Path.GetDirectoryName(tempPathTrasy);
            string connStrTrasy = $@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={trasyDbPath};Extended Properties=dBASE IV;";
            using (OleDbConnection conn = new OleDbConnection(connStrTrasy))
            {
                conn.Open();
                string query = $"SELECT NRLINII, NRPRZYST, NAZPRZYST, WAZNAOD, Kod2 FROM [{trasyFile}]";
                OleDbDataAdapter adapter = new OleDbDataAdapter(query, conn);
                DataTable dt = new DataTable();
                adapter.Fill(dt);

                filteredRows = dt.AsEnumerable()
                    .Where(r => activeLines.Contains(r["NRLINII"].ToString().Trim()))
                    .OrderBy(r => int.Parse(r["NRLINII"].ToString().Trim()))
                    .ThenBy(r =>
                    {
                        string dateStr = r["WAZNAOD"].ToString().Trim();
                        if (DateTime.TryParse(dateStr, out DateTime dt)) return dt;
                        return DateTime.MinValue;
                    })
                    .ThenBy(r => int.Parse(r["NRPRZYST"].ToString().Trim()))
                    .ToList();
            }

            
            textBox.Text += "Rozpoczynam sprawdzanie pliku TRASYLINII.DBF\n";
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

                bool newLine = currentNr != lastNr || currentStop <= lastStopNr;
                if (newLine)
                {
                    checkedLines++;
                    lastKod2 = "";
                    lastStopName = "";
                }

                if (!newLine && currentKod2 == lastKod2 && currentStopName == lastStopName)
                {
                    textBox.Text += $"\nUWAGA: NRLINII {currentNr}, przystanek {currentStopName} powtarza Kod2: {currentKod2}\n";
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
            if (errors == 0) textBox.Text += $"\nZnaleziono 0 błędów.\n";
            textBox.Text += $"\nProcedura zakończona. Sprawdzono {checkedLines} linii.\n";
            
        }






        private void Clear_Click(object sender, RoutedEventArgs e)
        {
            
        }
        
    }
}