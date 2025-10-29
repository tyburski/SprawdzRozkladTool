using Microsoft.Win32;
using System.Data;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Documents;


namespace SprawdzRozklad
{
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
                string path = openFileDialog.FileName;
                string[] parts = Path.GetFileName(path).Split("_");
                string[] withoutExt = parts[2].Split(".");
                zipFilePath = path;
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
                            string dane = lines.FirstOrDefault(x => x.Contains("Dane firmy:"));
                            string[] pp = dane.Split(":");
                            firma = pp[1].Trim();
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

        private async void Start_Click(object sender, RoutedEventArgs e)
        {
            if (companyNr.Text == "")
            {
                MessageBox.Show($"Najpierw wczytaj bazę.");
                return;
            }

            if (!validDatePicker.SelectedDate.HasValue)
            {
                MessageBox.Show("Ustaw datę!");
                return;
            }

            DateTime selectedDate = validDatePicker.SelectedDate.Value;

            if (selectedDate >= DateTime.Today)
            {
                MessageBox.Show("Data nie może być dzisiejsza ani przyszła!");
                return;
            }

            
            var loading = new LoadingWindow();
            loading.Owner = this;
            loading.Show();

            try
            {
                textBox.Text = "";


                await Task.Run(() =>
                {
                    CheckLines(selectedDate);

                    double totalWidth = textBox.ActualWidth;

                    int charCount = (int)(totalWidth / 5);

                    Dispatcher.Invoke(() => textBox.Text += $"\n{new string('_', charCount)}\n");

                    CheckRoutes(selectedDate);
                });
                
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Błąd: {ex.Message}\n\n{ex.StackTrace}");
            }
            finally
            {
                loading.Close();
            }
        }
        private void CheckLines(DateTime selectedDate)
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
                    Dispatcher.Invoke(() => textBox.Text += $"Plik Linie.DBF zapisany tymczasowo: {tempPathLinie}\n");
                }

                var entryTrasy = archive.Entries
                    .FirstOrDefault(e => e.Name.Equals(trasyFile, StringComparison.OrdinalIgnoreCase));
                if (entryTrasy != null)
                {
                    tempPathTrasy = Path.Combine(Path.GetTempPath(), entryTrasy.Name.ToLower());
                    entryTrasy.ExtractToFile(tempPathTrasy, true);
                    Dispatcher.Invoke(() => textBox.Text += $"Plik TRASYLINII.DBF zapisany tymczasowo: {tempPathTrasy}\n");
                }
            }

            List<(string NrLinii, string wariant)> activeLines = new List<(string NrLinii, string wariant)>();
            using (var table = NDbfReader.Table.Open(tempPathLinie))
            {
                var reader = table.OpenReader();

                while (reader.Read())
                {
                    string stat = reader.GetValue("STATLINII")?.ToString().Trim();
                    string nrf = reader.GetValue("NRF")?.ToString().Trim();
                    object dateObj = reader.GetValue("WAZNADO");

                    if (stat != "2") continue;
                    if (nrf != firma) continue;

                    if (dateObj != null && DateTime.TryParse(dateObj.ToString(), out DateTime dtWaznado))
                    {
                        if (dtWaznado <= selectedDate) continue;
                    }

                    string nr = reader.GetValue("NRLINII")?.ToString().Trim();
                    string wariant = reader.GetValue("WARLINII")?.ToString().Trim();
                    activeLines.Add((nr, wariant));
                }
            }
            Dispatcher.Invoke(() => textBox.Text += $"\nLinie: aktywne, data ważności do > {selectedDate.ToString("dd/MM/yyyy")} lub NULL: {activeLines.Count}\n");

            List<Dictionary<string, object>> filteredRows = new List<Dictionary<string, object>>();
            using (var table = NDbfReader.Table.Open(tempPathTrasy))
            {
                var reader = table.OpenReader();
                while (reader.Read())
                {
                    string nrl = reader.GetValue("NRLINII")?.ToString().Trim();
                    string war = reader.GetValue("WARLINII")?.ToString().Trim();

                    if (activeLines.Contains((nrl, war)))
                    {
                        var rowDict = new Dictionary<string, object>();
                        foreach (var col in table.Columns)
                        {
                            rowDict[col.Name] = reader.GetValue(col);
                        }
                        filteredRows.Add(rowDict);
                    }
                }
            }
            var groupedRows = filteredRows
                .OrderBy(r => int.Parse(r["NRLINII"].ToString().Trim()))
                .ThenBy(r => int.Parse(r["WARLINII"].ToString().Trim()))
                .ThenBy(r =>
                {
                    object val = r["WAZNAOD"];
                    if (val != null && DateTime.TryParse(val.ToString(), out var dt)) return dt;
                    return DateTime.MinValue;
                })
                .ThenBy(r => int.Parse(r["NRPRZYST"].ToString().Trim()))
                .ToList();

            Dispatcher.Invoke(() => textBox.Text += "\nSprawdzanie powtórzeń przystanków w kolejnych rekordach:\n");

            int lastNr = -1;
            int lastWariant = -1;
            DateTime lastWaznaOd = DateTime.MinValue;
            int lastStopNr = -1;
            int lastKod2 = -1;
            string lastStopName = "";
            int errors = 0;

            foreach (var row in groupedRows)
            {
                int currentNr = int.Parse(row["NRLINII"].ToString().Trim());
                int currentWariant = int.Parse(row["WARLINII"].ToString().Trim());
                DateTime currentWaznaOd = DateTime.TryParse(row["WAZNAOD"]?.ToString().Trim(), out var dt) ? dt : DateTime.MinValue;
                int currentStop = int.Parse(row["NRPRZYST"].ToString().Trim());
                int currentKod2 = int.Parse(row["KOD2"].ToString().Trim());
                string currentStopName = row["NAZPRZYST"].ToString().Trim();

                bool newGroup = false;

                if (currentNr != lastNr || currentWariant != lastWariant || currentWaznaOd != lastWaznaOd || currentStop <= lastStopNr)
                {
                    newGroup = true;
                    lastKod2 = -1;
                }

                if (!newGroup && currentKod2 == lastKod2)
                {
                    textBox.Text += $"\nUWAGA: Linia {currentNr}, wariant {currentWariant}, ważna od {currentWaznaOd:yyyy-MM-dd}, przystanek [ {lastStopName} ], kod2: [ {currentKod2} ] następuje jeden po drugim.\n";
                    errors++;
                    Application.Current.Dispatcher.Invoke(
                        System.Windows.Threading.DispatcherPriority.Background,
                        new Action(delegate { }));
                }

                lastNr = currentNr;
                lastWariant = currentWariant;
                lastWaznaOd = currentWaznaOd;
                lastStopNr = currentStop;
                lastKod2 = currentKod2;
                lastStopName = currentStopName;
            }

            Dispatcher.Invoke(() => textBox.Text += $"\nZnaleziono błędów: {errors}.\n");
        }


        private void CheckRoutes(DateTime selectedDate)
        {
            string tempPathTrasy = "";
            string tempPathKursy = "";
            string trasyFile = "trasykursow.dbf";
            string kursyFile = "kursy.dbf";

            using (ZipArchive archive = ZipFile.OpenRead(zipFilePath))
            {
                var entryKursy = archive.Entries
                    .FirstOrDefault(e => e.Name.Equals(kursyFile, StringComparison.OrdinalIgnoreCase));
                if (entryKursy != null)
                {
                    tempPathKursy = Path.Combine(Path.GetTempPath(), entryKursy.Name.ToLower());
                    entryKursy.ExtractToFile(tempPathKursy, true);
                    Dispatcher.Invoke(() => textBox.Text += $"\nPlik Kursy.DBF zapisany tymczasowo: {tempPathKursy}\n");
                }

                var entryTrasy = archive.Entries
                    .FirstOrDefault(e => e.Name.Equals(trasyFile, StringComparison.OrdinalIgnoreCase));
                if (entryTrasy != null)
                {
                    tempPathTrasy = Path.Combine(Path.GetTempPath(), entryTrasy.Name.ToLower());
                    entryTrasy.ExtractToFile(tempPathTrasy, true);
                    Dispatcher.Invoke(() => textBox.Text += $"Plik TRASYKURSOW.DBF zapisany tymczasowo: {tempPathTrasy}\n");
                }
            }

            List<(string NrKursu, string wariant)> activeRoutes = new List<(string NrKursu, string wariant)>();
            using (var table = NDbfReader.Table.Open(tempPathKursy))
            {
                var reader = table.OpenReader();
                while (reader.Read())
                {
                    string stat = reader.GetValue("STATKURSU")?.ToString().Trim();
                    string nrf = reader.GetValue("NRF")?.ToString().Trim();
                    object dateObj = reader.GetValue("WAZNYDO");

                    if (stat != "2") continue;
                    if (nrf != firma) continue;

                    if (dateObj != null && DateTime.TryParse(dateObj.ToString(), out DateTime dtWaznado))
                    {
                        if (dtWaznado <= selectedDate) continue;
                    }

                    string nr = reader.GetValue("NRKURSU")?.ToString().Trim();
                    string wariant = reader.GetValue("WARIANT")?.ToString().Trim();
                    activeRoutes.Add((nr, wariant));
                }
            }
            Dispatcher.Invoke(() => textBox.Text += $"\nKursy: aktywne, data ważności do > {selectedDate.ToString("dd/MM/yyyy")} lub NULL: {activeRoutes.Count}\n");

            List<Dictionary<string, object>> filteredRows = new List<Dictionary<string, object>>();
            using (var table = NDbfReader.Table.Open(tempPathTrasy))
            {
                var reader = table.OpenReader();
                while (reader.Read())
                {
                    string nrk = reader.GetValue("NRKURSU")?.ToString().Trim();
                    string war = reader.GetValue("WARIANT")?.ToString().Trim();

                    if (activeRoutes.Contains((nrk, war)))
                    {
                        var rowDict = new Dictionary<string, object>();
                        foreach (var col in table.Columns)
                        {
                            rowDict[col.Name] = reader.GetValue(col);
                        }
                        filteredRows.Add(rowDict);
                    }
                }
            }

            var groupedRows = filteredRows
                .OrderBy(r => int.Parse(r["NRKURSU"].ToString().Trim()))
                .ThenBy(r => int.Parse(r["WARIANT"].ToString().Trim()))
                .ThenBy(r =>
                {
                    object dataOp = r.TryGetValue("DATAOP", out var d) ? d : null;
                    object godzOp = r.TryGetValue("GODZOP", out var g) ? g : null;
                    if (dataOp != null && godzOp != null && DateTime.TryParse($"{dataOp} {godzOp}", out var dt)) return dt;
                    return DateTime.MinValue;
                })
                .ThenBy(r => int.Parse(r["NRPRZYST"].ToString().Trim()))
                .ToList();

            Dispatcher.Invoke(() => textBox.Text += "\nSprawdzanie powtórzeń przystanków w kolejnych rekordach:\n");

            int lastNr = -1;
            int lastWariant = -1;
            DateTime lastWaznaOd = DateTime.MinValue;
            int lastStopNr = -1;
            int lastKod2 = -1;
            string lastStopName = "";
            int checkedLines = 0;
            int errors = 0;

            foreach (var row in groupedRows)
            {
                int currentNr = int.Parse(row["NRKURSU"].ToString().Trim());
                int currentWariant = int.Parse(row["WARIANT"].ToString().Trim());
                DateTime currentWaznyOd = row.TryGetValue("WAZNYOD", out var w) && w != null
                    ? DateTime.TryParse(w.ToString().Trim(), out var dt) ? dt : DateTime.MinValue
                    : DateTime.MinValue;
                int currentStop = int.Parse(row["NRPRZYST"].ToString().Trim());
                int currentKod2 = row.TryGetValue("KOD2", out var k) && k != null
                    ? int.Parse(k.ToString().Trim())
                    : -1;
                string currentStopName = row.TryGetValue("NAZPRZYST", out var n) && n != null ? n.ToString().Trim() : "";

                bool newGroup = false;

                if (currentNr != lastNr || currentWariant != lastWariant || currentWaznyOd != lastWaznaOd || currentStop <= lastStopNr)
                {
                    newGroup = true;
                    checkedLines++;
                    lastKod2 = -1;
                }

                if (!newGroup && currentKod2 == lastKod2)
                {
                    Dispatcher.Invoke(() => textBox.Text += $"\nUWAGA: Kurs {currentNr}, wariant {currentWariant}, ważny od {currentWaznyOd:yyyy-MM-dd}, przystanek [ {lastStopName} ], kod2: [ {currentKod2} ] następuje jeden po drugim.\n");
                    errors++;
                    Application.Current.Dispatcher.Invoke(
                        System.Windows.Threading.DispatcherPriority.Background,
                        new Action(delegate { }));
                }

                lastNr = currentNr;
                lastWariant = currentWariant;
                lastWaznaOd = currentWaznyOd;
                lastStopNr = currentStop;
                lastKod2 = currentKod2;
                lastStopName = currentStopName;
            }

            Dispatcher.Invoke(() => textBox.Text += $"\nZnaleziono błędów: {errors}.\n");
        }

        private void Clear_Click(object sender, RoutedEventArgs e)
        {
            companyNr.Text = "";
            textBox.Text = "";
        }
    }
}