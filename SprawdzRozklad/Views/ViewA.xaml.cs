using Microsoft.Win32;
using NDbfReader;
using System.Data;
using System.IO;
using System.IO.Compression;
using System.Reflection.PortableExecutable;
using System.Security.Policy;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using static System.Runtime.InteropServices.JavaScript.JSType;


namespace SprawdzRozklad.Views
{
    public partial class ViewA : UserControl
    {
        public ViewA()
        {
            InitializeComponent();
            validDatePicker.SelectedDate = new DateTime(2018, 1, 1);
            coordsCheckBox.IsChecked = false;
            admCheckBox.IsChecked = true;
            firmyCheckBox.IsChecked = true;
            statusCheckBox.IsChecked = true;
            companyNr.Text = "";
        }

        string zipFilePath = "";
        string firma = "";
        DateTime date = DateTime.Now;
        private void OpenFile_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Title = "Wybierz plik bazy";
            openFileDialog.Filter = "Plik bazy (*.zip)|*.zip";

            if (openFileDialog.ShowDialog() == true)
            {
                string path = openFileDialog.FileName;
                zipFilePath = path;

                using (ZipArchive archive = ZipFile.OpenRead(zipFilePath))
                {
                    var entry = archive.Entries.FirstOrDefault(e =>
                                e.Name.StartsWith("ArchiwumInfo_", StringComparison.OrdinalIgnoreCase) &&
                                e.Name.EndsWith(".rtf", StringComparison.OrdinalIgnoreCase));

                    if (entry != null)
                    {
                        using (StreamReader reader = new StreamReader(entry.Open(), Encoding.UTF8))
                        {
                            Dispatcher.Invoke(() => textBox.Text = "");
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
                        MessageBox.Show($"Nie znaleziono pliku 'ArchiwumInfo_*.rtf'");
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

            Window.GetWindow(this).IsEnabled = false;
            var loading = new LoadingWindow();
            loading.Owner = Window.GetWindow(this);
            loading.Show();

            try
            {
                Dispatcher.Invoke(() => textBox.Text = "");


                await Task.Run(() =>
                {
                    string statusWsp = Dispatcher.Invoke(() => coordsCheckBox.IsChecked) is true ? "tak" : "nie";
                    string statusAdm = Dispatcher.Invoke(() => admCheckBox.IsChecked) is true ? "tak" : "nie";
                    string statusFirm = Dispatcher.Invoke(() => firmyCheckBox.IsChecked) is true ? "tak" : "nie";
                    string statusSts = Dispatcher.Invoke(() => statusCheckBox.IsChecked) is true ? "tylko aktywne" : "wszystkie";
                    DateTime newDate = DateTime.Now;
                    date = newDate;
                    Dispatcher.Invoke(() => textBox.Text += $"{date.ToString("dd/MM/yyyy HH:mm")}\n" +
                    $"Parametry weryfikacji:\n" +
                    $"\n\tFirma: {firma}\n"+
                    $"\tSprawdź współrzędne przystanków: {statusWsp}\n" +
                    $"\tSprawdź administrację przystanków: {statusAdm}\n" +
                    $"\tSprawdź linie/kursy firm obcych: {statusFirm}\n"+
                    $"\tLinie/kursy: {statusSts}, data ważności do > {selectedDate.ToString("dd/MM/yyyy")} lub pole puste\n");

                    CheckStops();

                    CheckLines(selectedDate);

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
                Window.GetWindow(this).IsEnabled = true;
            }
        }

        private void CheckStops()
        {
            string tempPathTrasy = "";
            string tempPathPrzystanki = "";
            string przystankiFile = "przystanki.dbf";

            using (ZipArchive archive = ZipFile.OpenRead(zipFilePath))
            {
                var entryPrzystanki = archive.Entries
                    .FirstOrDefault(e => e.Name.Equals(przystankiFile, StringComparison.OrdinalIgnoreCase));
                if (entryPrzystanki != null)
                {
                    tempPathPrzystanki = Path.Combine(Path.GetTempPath(), entryPrzystanki.Name.ToLower());
                    entryPrzystanki.ExtractToFile(tempPathPrzystanki, true);
                }
            }

            List<(string NAZWA, string KOD2, string SYM, string X_KRAJ, string Y_KRAJ)> stops = new List<(string NAZWA, string KOD2, string SYM, string X_KRAJ, string Y_KRAJ)>();
            using (var table = NDbfReader.Table.Open(tempPathPrzystanki))
            {
                var encoding = Encoding.GetEncoding(852);
                var reader = table.OpenReader(encoding);
                
                while (reader.Read())
                {
                    string nazwa = reader.GetValue("NAZWA")?.ToString().Trim();
                    string kod2 = reader.GetValue("KOD2")?.ToString().Trim();
                    string sym = reader.GetValue("SYM")?.ToString().Trim();
                    string x_kraj = reader.GetValue("X_KRAJ")?.ToString().Trim();
                    string y_kraj = reader.GetValue("Y_KRAJ")?.ToString().Trim();

                    stops.Add((nazwa, kod2, sym, x_kraj, y_kraj));
                }
            }
            Dispatcher.Invoke(() => textBox.Text += $"\n\nPrzystanki: {stops.Count}\n");

            int errors = 0;

            var nazwaUnique = stops
                .Select((r, i) => new { Row = r, Index = i + 1 })
                .OrderBy(x => x.Row.NAZWA.ToUpper())
                .ToList();

            string last = null;
            (string NAZWA, string KOD2)? lastRecord = null;
            List<(string NAZWA, string KOD2)> duplicates = new();

            foreach (var nazwaRow in nazwaUnique)
            {
                string nazwa = nazwaRow.Row.NAZWA?.ToUpperInvariant();

                if (nazwa == last)
                {
                    if (duplicates.Count == 0 && lastRecord.HasValue)
                        duplicates.Add(lastRecord.Value);

                    duplicates.Add((nazwaRow.Row.NAZWA, nazwaRow.Row.KOD2));

                    string msg = "mają taką samą nazwę.";
                    if (duplicates[0].KOD2 == duplicates[1].KOD2)
                    {
                        msg = "mają taką samą nazwę i kod.";
                        foreach (var dup in duplicates)
                        {
                            var remove = stops.FirstOrDefault(s => s.NAZWA == dup.NAZWA && s.KOD2 == dup.KOD2);
                            if (remove != default)
                                stops.Remove(remove);
                        }
                    }

                    Dispatcher.Invoke(() =>
                        textBox.Text += $"\n\tPrzystanki {string.Join(", ", duplicates.Select(x => $"[{x.NAZWA}] [{x.KOD2}]"))} {msg}\n");

                    errors++;
                }
                else
                {
                    last = nazwa;
                    duplicates.Clear();
                }

                lastRecord = (nazwaRow.Row.NAZWA, nazwaRow.Row.KOD2);
            }

            var kodUnique = stops
                .Select((r, i) => new { Row = r, Index = i + 1 })
                .OrderBy(x => x.Row.KOD2)
                .ToList();

            string lastKod = null;
            (string NAZWA, string KOD2)? lastRecordKod = null;
            List<(string NAZWA, string KOD2)> duplicatesKody = new();

            foreach (var kodRow in kodUnique)
            {
                string kod = kodRow.Row.KOD2;

                if (kod == lastKod)
                {
                    if (duplicates.Count == 0 && lastRecord.HasValue)
                        duplicates.Add(lastRecord.Value);

                    duplicates.Add((kodRow.Row.NAZWA, kodRow.Row.KOD2));

                    string msg = "mają taki sam kod.";

                    Dispatcher.Invoke(() =>
                        textBox.Text += $"\n\tPrzystanki {string.Join(", ", duplicates.Select(x => $"[{x.NAZWA}] [{x.KOD2}]"))} {msg}\n");

                    errors++;
                }
                else
                {
                    lastKod = kod;
                    duplicates.Clear();
                }

                lastRecord = (kodRow.Row.NAZWA, kodRow.Row.KOD2);
            }

            if (Dispatcher.Invoke(() => admCheckBox.IsChecked) is true)
            {
                var emptySym = stops.Where(s => string.IsNullOrWhiteSpace(s.SYM));

                foreach (var symRow in emptySym)
                {
                    Dispatcher.Invoke(() => textBox.Text += $"\n\tPrzystanek [{symRow.NAZWA}] [{symRow.KOD2}] nie ma przypisanej administracji.\n");
                    errors++;
                }
            }

            if (Dispatcher.Invoke(() => coordsCheckBox.IsChecked) is true)
            {
                var emptyCoordinates = stops.Where(s => string.IsNullOrWhiteSpace(s.X_KRAJ) || string.IsNullOrWhiteSpace(s.Y_KRAJ));

                foreach (var coorRow in emptyCoordinates)
                {
                    Dispatcher.Invoke(() => textBox.Text += $"\n\tPrzystanek [{coorRow.NAZWA}] [{coorRow.KOD2}] nie ma przypisanych współrzędnych.\n");
                    errors++;
                }
            }
            

            if(errors == 0) Dispatcher.Invoke(() => textBox.Text += $"\n\tNie znaleziono błędów.\n");
            else Dispatcher.Invoke(() => textBox.Text += $"\n\n\tZnaleziono błędów: {errors}\n");
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
                }

                var entryTrasy = archive.Entries
                    .FirstOrDefault(e => e.Name.Equals(trasyFile, StringComparison.OrdinalIgnoreCase));
                if (entryTrasy != null)
                {
                    tempPathTrasy = Path.Combine(Path.GetTempPath(), entryTrasy.Name.ToLower());
                    entryTrasy.ExtractToFile(tempPathTrasy, true);
                }
            }

            List<(string nrf, string NrLinii, string wariant, object dateFrom, string stat)> activeLines = new List<(string nrf, string NrLinii, string wariant, object dateFrom, string stat)>();
            using (var table = NDbfReader.Table.Open(tempPathLinie))
            {
                var encoding = Encoding.GetEncoding(852);
                var reader = table.OpenReader(encoding);

                while (reader.Read())
                {
                    string stat = reader.GetValue("STATLINII")?.ToString().Trim();
                    string nrf = reader.GetValue("NRF")?.ToString().Trim();
                    object dateObj = reader.GetValue("WAZNADO");

                    if (Dispatcher.Invoke(() => statusCheckBox.IsChecked) is true)
                    {
                        if (stat != "2") continue;
                    }                 
                    if (Dispatcher.Invoke(() => firmyCheckBox.IsChecked) is false)
                    {
                        if (nrf != firma) continue;
                    }
                                        
                    if (dateObj != null && DateTime.TryParse(dateObj.ToString(), out DateTime dtWaznado))
                    {
                        if (dtWaznado <= selectedDate) continue;
                    }

                    string nr = reader.GetValue("NRLINII")?.ToString().Trim();
                    string wariant = reader.GetValue("WARLINII")?.ToString().Trim();
                    object dateFrom = reader.GetValue("WAZNAOD");
                    activeLines.Add((nrf, nr, wariant, dateFrom, stat));
                }
            }
            Dispatcher.Invoke(() => textBox.Text += $"\n\nLinie: {activeLines.Count}\n");

            List<Dictionary<string, object>> filteredRows = new List<Dictionary<string, object>>();
            using (var table = NDbfReader.Table.Open(tempPathTrasy))
            {
                var encoding = Encoding.GetEncoding(852);
                var reader = table.OpenReader(encoding);

                while (reader.Read())
                {
                    string nrf = reader.GetValue("NRF")?.ToString().Trim();
                    string nrl = reader.GetValue("NRLINII")?.ToString().Trim();
                    string war = reader.GetValue("WARLINII")?.ToString().Trim();
                    object dateFrom = reader.GetValue("WAZNAOD");

                    var rowDict = new Dictionary<string, object>();

                    var line = activeLines.FirstOrDefault(x => x.nrf == nrf && x.NrLinii == nrl && x.wariant == war && x.dateFrom?.ToString() == dateFrom?.ToString());

                    if(line != default)
                    {
                        
                        foreach (var col in table.Columns)
                        {
                            rowDict[col.Name] = reader.GetValue(col);
                            rowDict["STATLINII"] = line.stat;
                        }
                        filteredRows.Add(rowDict);
                    }                 
                }
            }
            var groupedRows = filteredRows
                .OrderBy(r => int.Parse(r["NRF"].ToString().Trim()))
                .ThenBy(r => int.Parse(r["NRLINII"].ToString().Trim()))
                .ThenBy(r => int.Parse(r["WARLINII"].ToString().Trim()))
                .ThenBy(r =>
                {
                    object val = r["WAZNAOD"];
                    if (val != null && DateTime.TryParse(val.ToString(), out var dt)) return dt;
                    return DateTime.MinValue;
                })
                .ThenBy(r => int.Parse(r["NRPRZYST"].ToString().Trim()))
                .ToList();


            int lastNr = -1;
            int lastWariant = -1;
            DateTime lastWaznaOd = DateTime.MinValue;
            int lastStopNr = -1;
            int lastKod2 = -1;
            string lastStopName = "";
            int errors = 0;

            foreach (var row in groupedRows)
            {
                int nrf = int.Parse(row["NRF"].ToString().Trim());
                int currentNr = int.Parse(row["NRLINII"].ToString().Trim());
                int stat = int.Parse(row["STATLINII"].ToString().Trim());
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
                    string statName = "";
                    if (stat == 1) statName = "projektowana";
                    else if (stat == 2) statName = "aktywna";
                    else if (stat == 3) statName = "zawieszona";
                    else if (stat == 4) statName = "zlikwidowana";
                    else statName = "bd";

                    Dispatcher.Invoke(() => textBox.Text += $"\n\t[{nrf}] Linia {currentNr} ({statName}), wariant {currentWariant}, ważna od {currentWaznaOd:yyyy-MM-dd}, przystanek [ {lastStopName} ], kod2: [ {currentKod2} ] następuje jeden po drugim.\n");
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

            if (errors == 0) Dispatcher.Invoke(() => textBox.Text += $"\n\tNie znaleziono błędów.\n");
            else Dispatcher.Invoke(() => textBox.Text += $"\n\n\tZnaleziono błędów: {errors}\n");
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
                }

                var entryTrasy = archive.Entries
                    .FirstOrDefault(e => e.Name.Equals(trasyFile, StringComparison.OrdinalIgnoreCase));
                if (entryTrasy != null)
                {
                    tempPathTrasy = Path.Combine(Path.GetTempPath(), entryTrasy.Name.ToLower());
                    entryTrasy.ExtractToFile(tempPathTrasy, true);
                }
            }

            List<(string nrf, string NrKursu, string wariant, object dateFrom, string stat)> activeRoutes = new List<(string nrf, string NrKursu, string wariant, object dateFrom, string stat)>();
            using (var table = NDbfReader.Table.Open(tempPathKursy))
            {
                var encoding = Encoding.GetEncoding(852);
                var reader = table.OpenReader(encoding);

                while (reader.Read())
                {
                    string stat = reader.GetValue("STATKURSU")?.ToString().Trim();
                    string nrf = reader.GetValue("NRF")?.ToString().Trim();
                    object dateObj = reader.GetValue("WAZNYDO");
                    object dateFrom = reader.GetValue("WAZNYOD");

                    if (Dispatcher.Invoke(() => statusCheckBox.IsChecked) is true)
                    {
                        if (stat != "2") continue;
                    }
                    
                    if (Dispatcher.Invoke(() => firmyCheckBox.IsChecked) is false)
                    {
                        if (nrf != firma) continue;
                    }

                    if (dateObj != null && DateTime.TryParse(dateObj.ToString(), out DateTime dtWaznado))
                    {
                        if (dtWaznado <= selectedDate) continue;
                    }

                    string nr = reader.GetValue("NRKURSU")?.ToString().Trim();
                    string wariant = reader.GetValue("WARIANT")?.ToString().Trim();
                    activeRoutes.Add((nrf, nr, wariant, dateFrom, stat));
                }
            }
            Dispatcher.Invoke(() => textBox.Text += $"\n\nKursy: {activeRoutes.Count}\n");

            List<Dictionary<string, object>> filteredRows = new List<Dictionary<string, object>>();
            using (var table = NDbfReader.Table.Open(tempPathTrasy))
            {
                var encoding = Encoding.GetEncoding(852);
                var reader = table.OpenReader(encoding);

                while (reader.Read())
                {
                    string nrfk = reader.GetValue("NRFK")?.ToString().Trim();
                    string nrk = reader.GetValue("NRKURSU")?.ToString().Trim();
                    string war = reader.GetValue("WARIANT")?.ToString().Trim();
                    object dateFrom = reader.GetValue("WAZNYOD");

                    var rowDict = new Dictionary<string, object>();

                    var line = activeRoutes.FirstOrDefault(x => x.nrf == nrfk && x.NrKursu == nrk && x.wariant == war && x.dateFrom?.ToString() == dateFrom?.ToString());

                    if (line != default)
                    {

                        foreach (var col in table.Columns)
                        {
                            rowDict[col.Name] = reader.GetValue(col);
                            rowDict["STATKURSU"] = line.stat;
                        }
                        filteredRows.Add(rowDict);
                    }
                }
            }

            var groupedRows = filteredRows
                .OrderBy(r => int.Parse(r["NRFK"].ToString().Trim()))
                .ThenBy(r => int.Parse(r["NRKURSU"].ToString().Trim()))
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
                int nrf = int.Parse(row["NRFK"].ToString().Trim());
                int stat = int.Parse(row["STATKURSU"].ToString().Trim());
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
                    string statName = "";
                    if (stat == 1) statName = "projektowany";
                    else if (stat == 2) statName = "aktywny";
                    else if (stat == 3) statName = "zawieszony";
                    else if (stat == 4) statName = "zlikwidowany";
                    else statName = "bd";
                    Dispatcher.Invoke(() => textBox.Text += $"\n\t[{nrf}] Kurs {currentNr} ({statName}), wariant {currentWariant}, ważny od {currentWaznyOd:yyyy-MM-dd}, przystanek [ {lastStopName} ], kod2: [ {currentKod2} ] następuje jeden po drugim.\n");
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

            if (errors == 0) Dispatcher.Invoke(() => textBox.Text += $"\n\tNie znaleziono błędów.\n");
            else Dispatcher.Invoke(() => textBox.Text += $"\n\n\tZnaleziono błędów: {errors}\n");
        }


        private void Clear_Click(object sender, RoutedEventArgs e)
        {
            companyNr.Text = "";
            textBox.Text = "";
        }
        private void Copy_Click(object sender, RoutedEventArgs e)
        {
            ToolTip tool = new ToolTip { Content = "Treść komunikatu została skopiowana do schowka.", StaysOpen = false };

            if (string.IsNullOrWhiteSpace(Dispatcher.Invoke(()=>textBox.Text)))
            {
                tool.Content = "Brak tekstu do zapisania.";
                tool.IsOpen = true;
                return;
            }
            Clipboard.SetText(Dispatcher.Invoke(() => textBox.Text));
            tool.IsOpen = true;

        }
        private void Save_Click(object sender, RoutedEventArgs e)
        {        
           

            if (string.IsNullOrWhiteSpace(Dispatcher.Invoke(() => textBox.Text)))
            {
                ToolTip tool = new ToolTip { Content = "Brak tekstu do zapisania.", StaysOpen = false };
                tool.IsOpen = true;
                return;
            }

            SaveFileDialog dialog = new SaveFileDialog
            {
                Title = $"Zapisz plik",
                Filter = "Plik tekstowy (*.txt)|*.txt|Wszystkie pliki (*.*)|*.*",
                FileName = $"{firma}_{date.ToString("ddMMyyyHHmm")}.txt"
            };

            if(dialog.ShowDialog() == true)
            {
                try
                {
                    File.WriteAllText(dialog.FileName, Dispatcher.Invoke(() => textBox.Text));
                    MessageBox.Show("Pomyślnie zapisano do pliku.");
                }
                catch(Exception ex)
                {
                    MessageBox.Show("Błąd:" + ex.Message, "Błąd",
                        MessageBoxButton.OK);
                }
            }
        }
    }
}