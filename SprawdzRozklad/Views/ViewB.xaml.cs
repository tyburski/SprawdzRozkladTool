using Microsoft.Win32;
using Microsoft.WindowsAPICodePack.Dialogs;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;



namespace SprawdzRozklad.Views
{
    
    public partial class ViewB : UserControl
    {
        List<string> insertKasy = new List<string>();
        List<string> insertPracownicy = new List<string>();
        List<string> insertKarty = new List<string>();

        public ViewB()
        {
            InitializeComponent();
            saveSQLBtn.IsEnabled = true;
        }

        string zipFilePath = "";
        string dane = "";
        private void OpenFile_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Title = "Wybierz plik rozliczeń";
            openFileDialog.Filter = "Plik rozliczeń (*.zip)|*.zip";

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
                                dane = lines.FirstOrDefault(x => x.Contains("Dane firmy:"));
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
                MessageBox.Show($"Najpierw wczytaj archiwum rozliczeń.");
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
                    GenerateBileterki();
                    GeneratePracownicy();
                    GenerateKartyPamieci();
                });

            }
            catch (Exception ex)
            {
                MessageBox.Show($"Błąd: {ex.Message}");
            }
            finally
            {
                loading.Close();
                Window.GetWindow(this).IsEnabled = true;
            }
        }
        private void Clear_Click(object sender, RoutedEventArgs e)
        {
            companyNr.Text = "";
            textBox.Text = "";
            zipFilePath = "";
            dane = "";

            insertKasy = [];
            insertPracownicy = [];
            insertKarty = [];
        }

        string FormatDate(object value)
        {
            if (value == null || value == DBNull.Value) return null;
            if (DateTime.TryParse(value.ToString(), out DateTime dt))
                return dt.ToString("dd.MM.yyyy");
            return null;
        }
        string SqlValue(object value)
        {
            if (value == null || value == DBNull.Value)
                return "null";

            if (value is byte[])
                return "null";

            if (value is float fl)
            {
                if (fl % 1 == 0)
                    return $"'{fl.ToString("0")}'";
                return $"'{fl.ToString()}'";
            }
            if (value is decimal dc)
            {
                if (dc % 1 == 0)
                    return $"'{dc.ToString("0")}'";
                return $"'{dc.ToString()}'";
            }
            if (value is double db)
            {
                if (db % 1 == 0)
                    return $"'{db.ToString("0")}'";
                return $"'{db.ToString()}'";
            }

            return $"'{value.ToString().Trim()}'";
        }
        string SqlDate(object value)
        {
            var formatted = FormatDate(value);
            return formatted == null ? "null" : $"'{formatted}'";
        }

        private void GenerateBileterki()
        {
            string tempPathBileterki = "";
            string bileterkiFile = "KASYFISKALNE.DBF";
            string bileterkiPath = $"bazy/";

            using (ZipArchive archive = ZipFile.OpenRead(zipFilePath))
            {
                var entryBileterki = archive.Entries
                    .FirstOrDefault(e => e.Name.Equals(bileterkiFile, StringComparison.OrdinalIgnoreCase));
                if (entryBileterki != null)
                {
                    tempPathBileterki = Path.Combine(Path.GetTempPath(), entryBileterki.Name.ToLower());
                    entryBileterki.ExtractToFile(tempPathBileterki, true);
                }
            }

            int bileterki = 0;
            bool wrongVer = false;
            using (var table = NDbfReader.Table.Open(tempPathBileterki))
            {
                Dispatcher.Invoke(() => textBox.Text += $"Pomijam bileterki gdzie DATA TERAZ - DATAZORF > 1 ROK \n");
                Dispatcher.Invoke(() => textBox.Text += $"Pomijam bileterki gdzie LOGO jest puste \n");
                Dispatcher.Invoke(() => textBox.Text += $"Sprawdzam czy są bileterki EMAR-105 gdzie WERPSB < 1.5");
                var encoding = Encoding.GetEncoding(852);
                var reader = table.OpenReader(encoding);


                
                while (reader.Read())
                {
                    

                    string logo = reader.GetValue("LOGO")?.ToString().Trim();
                    string typbil = reader.GetValue("TYPBIL")?.ToString().Trim();
                    object dateObj = reader.GetValue("DATAZORF");
                    string werpsb = reader.GetValue("WERPSB")?.ToString().Trim();

                    if (dateObj == null) continue;
                    
                    if (DateTime.TryParse(dateObj.ToString(), out DateTime dtZorf))
                    {
                        if (dtZorf < DateTime.Today.AddYears(-1)) continue;
                    }

                    if (logo == string.Empty && logo == "") continue;
                    
                    if (typbil is not null && typbil.Equals("EMAR-105"))
                    {
                        if (string.IsNullOrEmpty(werpsb)) Dispatcher.Invoke(() => textBox.Text += $"Bileterka EMAR-105 {logo} nie ma wersji oprogramowania!\n");
                        else
                        {
                            var emarVersion = float.Parse(werpsb, CultureInfo.InvariantCulture);
                            if (emarVersion < 1.5)
                            {
                                if(wrongVer == false) Dispatcher.Invoke(() => textBox.Text += "\n\n");
                                wrongVer = true;
                                Dispatcher.Invoke(() => textBox.Text += $"Bileterka EMAR-105 {logo} ma wersję oprogramowania {werpsb}!\n");
                            }
                        }
                            
                        
                    }

                    string insert = "INSERT INTO IMP_KasyFiskalne (LOGO, NRINW, NRFABR, NREWUS, DATAPROD, TYPBIL, DATAPRZ, NRPRF, DATAOPRF, NAZRFZPRF, LRZPRF, DATAZPRF, NAZRZZPRF, NRRZZPRF, NRKPPRF, NRPRZPRF, IMIEZPRF, NAZWZPRF, STATZPRF, NRORF, DATAOORF, NAZRFZORF, LRZORF, NAZRZORF, NRRZORF, DATAZORF, NRKPORF, NRPRZORF, IMIEZORF, NAZWZORF, STATZORF, LRF, DATAWBF, LOGOPOWBF, DATAKON, DATAFISK, DATAPRZEGL, ADRES, KODP, MIEJSC, NRDECUS, DATADECUS, DATAZAK, LMIESGWAR, UWAGI, BEZOPSERW, WERPSB, DATAOP, GODZOP, NR_SLUZBOP) VALUES (" +
                                        SqlValue(reader.GetValue("LOGO")) + "," +
                                        SqlValue(reader.GetValue("NRINW")) + "," +
                                        SqlValue(reader.GetValue("NRFABR")) + "," +
                                        SqlValue(reader.GetValue("NREWUS")) + "," +
                                        SqlDate(reader.GetValue("DATAPROD")) + "," +
                                        SqlValue(reader.GetValue("TYPBIL")) + "," +
                                        SqlDate(reader.GetValue("DATAPRZ")) + "," +
                                        SqlValue(reader.GetValue("NRPRF")) + "," +
                                        SqlDate(reader.GetValue("DATAOPRF")) + "," +
                                        SqlValue(reader.GetValue("NAZRFZPRF")) + "," +
                                        SqlValue(reader.GetValue("LRZPRF")) + "," +
                                        SqlDate(reader.GetValue("DATAZPRF")) + "," +
                                        SqlValue(reader.GetValue("NAZRZZPRF")) + "," +
                                        SqlValue(reader.GetValue("NRRZZPRF")) + "," +
                                        SqlValue(reader.GetValue("NRKPPRF")) + "," +
                                        SqlValue(reader.GetValue("NRPRZPRF")) + "," +
                                        SqlValue(reader.GetValue("IMIEZPRF")) + "," +
                                        SqlValue(reader.GetValue("NAZWZPRF")) + "," +
                                        SqlValue(reader.GetValue("STATZPRF")) + "," +
                                        SqlValue(reader.GetValue("NRORF")) + "," +
                                        SqlDate(reader.GetValue("DATAOORF")) + "," +
                                        SqlValue(reader.GetValue("NAZRFZORF")) + "," +
                                        SqlValue(reader.GetValue("LRZORF")) + "," +
                                        SqlValue(reader.GetValue("NAZRZORF")) + "," +
                                        SqlValue(reader.GetValue("NRRZORF")) + "," +
                                        SqlDate(reader.GetValue("DATAZORF")) + "," +
                                        SqlValue(reader.GetValue("NRKPORF")) + "," +
                                        SqlValue(reader.GetValue("NRPRZORF")) + "," +
                                        SqlValue(reader.GetValue("IMIEZORF")) + "," +
                                        SqlValue(reader.GetValue("NAZWZORF")) + "," +
                                        SqlValue(reader.GetValue("STATZORF")) + "," +
                                        SqlValue(reader.GetValue("LRF")) + "," +
                                        SqlDate(reader.GetValue("DATAWBF")) + "," +
                                        SqlValue(reader.GetValue("LOGOPOWBF")) + "," +
                                        SqlDate(reader.GetValue("DATAKON")) + "," +
                                        SqlDate(reader.GetValue("DATAFISK")) + "," +
                                        SqlDate(reader.GetValue("DATAPRZEGL")) + "," +
                                        SqlValue(reader.GetValue("ADRES")) + "," +
                                        SqlValue(reader.GetValue("KODP")) + "," +
                                        SqlValue(reader.GetValue("MIEJSC")) + "," +
                                        SqlValue(reader.GetValue("NRDECUS")) + "," +
                                        SqlDate(reader.GetValue("DATADECUS")) + "," +
                                        SqlDate(reader.GetValue("DATAZAK")) + "," +
                                        SqlValue(reader.GetValue("LMIESGWAR")) + "," +
                                        SqlValue(reader.GetValue("UWAGI")) + "," +
                                        SqlValue(reader.GetValue("BEZOPSERW")) + "," +
                                        SqlValue(reader.GetValue("WERPSB")) + "," +
                                        SqlDate(reader.GetValue("DATAOP")) + "," +
                                        SqlValue(reader.GetValue("GODZOP")) + "," +
                                        SqlValue(reader.GetValue("NR_SLUZBOP")) + ");";


                    insertKasy.Add(insert);
                    bileterki++;
                }
            }
            if (wrongVer == false)
            {
                Dispatcher.Invoke(() => textBox.Text += " -- NIE\n");
            }
               
                Dispatcher.Invoke(() => textBox.Text += $"\nBileterki: {bileterki} | ");           
        }

        private void GeneratePracownicy()
        {
            string tempPathPracownicy = "";
            string PracownicyFile = "Pracownicy.DBF";
            string PracownicyPath = $"bazy/";

            using (ZipArchive archive = ZipFile.OpenRead(zipFilePath))
            {
                var entryPracownicy = archive.Entries
                    .FirstOrDefault(e => e.Name.Equals(PracownicyFile, StringComparison.OrdinalIgnoreCase));
                if (entryPracownicy != null)
                {
                    tempPathPracownicy = Path.Combine(Path.GetTempPath(), entryPracownicy.Name.ToLower());
                    entryPracownicy.ExtractToFile(tempPathPracownicy, true);
                }
            }
            int pracownicy = 0;
            using (var table = NDbfReader.Table.Open(tempPathPracownicy))
            {

                var encoding = Encoding.GetEncoding(852);
                var reader = table.OpenReader(encoding);

                while (reader.Read())
                {
                    string status = reader.GetValue("STATUS")?.ToString().Trim();
                    if (status != "7" && status != "8") status = "8";


                    string insert = "INSERT INTO IMP_Pracownicy (NR_SLUZB, IMIE, NAZWISKO, STATUS, ODDZIAL, DATA_PRZYJ, DATA_ZWOL, TRYBROZWUM, ADRES, KODP, MIEJSC, SYM, GMINA, NAZKR, KOD2, KIER, TEL, TELKOM, ROKROZL, PLANDNIURL, DNIURLDO, DATAURL1, LBDNIROB1, DATAURL2, LBDNIROB2, DATAURL3, LBDNIROB3, NRREJAUT, DATPRZYAUT, NRINWBIL, LOGOPB, DATPRZYBIL, NRKARPAM, DATPRZKP, NAZPRZ, NRPRZ, DATAREJP, GODZREJP, DATAPOCZP, DATAZAKP, NRPRF, LOGOPRF, NRPZAD, NRPKARDR, NRKPPZAD, NAZORZ, NRORZ, DATAREJO, GODZREJO, DATAPOCZO, DATAZAKO, NRORF, LOGOORF, NRZAD, NRKARDR, NRKPOZAD, NRSPLAC, WYSTNOTY, WYSTFAKT, DRUKDOK, ETAT, NUMERPRJ, DATAWYDPRJ, DATAWAZPRJ, PRAWOJ, DWAZLEK, DWAZPSYCH, DATAPKRT, DATAKKRT, NUMERKRT, WYDKRT, DATA_UR, NR_DOWOD, NR_PASZ, PESEL, DATABLEK, DATABPSYCH, NRZASKDPO, DATKDPO, NRZASKDPR, DATKDPR, NRTELKOM, DTELKOM, DATAOP, GODZOP, NR_SLUZBOP) VALUES (" +
                                        SqlValue(reader.GetValue("NR_SLUZB")) + "," +
                                        SqlValue(reader.GetValue("IMIE")) + "," +
                                        SqlValue(reader.GetValue("NAZWISKO")) + "," +
                                        SqlValue(status) + "," +
                                        SqlValue(reader.GetValue("ODDZIAL")) + "," +
                                        SqlDate(reader.GetValue("DATA_PRZYJ")) + "," +
                                        SqlDate(reader.GetValue("DATA_ZWOL")) + "," +
                                        SqlValue(reader.GetValue("TRYBROZWUM")) + "," +
                                        SqlValue(reader.GetValue("ADRES")) + "," +
                                        SqlValue(reader.GetValue("KODP")) + "," +
                                        SqlValue(reader.GetValue("MIEJSC")) + "," +
                                        SqlValue(reader.GetValue("SYM")) + "," +
                                        SqlValue(reader.GetValue("GMINA")) + "," +
                                        SqlValue(reader.GetValue("NAZKR")) + "," +
                                        SqlValue(reader.GetValue("KOD2")) + "," +
                                        SqlValue(reader.GetValue("KIER")) + "," +
                                        SqlValue(reader.GetValue("TEL")) + "," +
                                        SqlValue(reader.GetValue("TELKOM")) + "," +
                                        SqlValue(reader.GetValue("ROKROZL")) + "," +
                                        SqlValue(reader.GetValue("PLANDNIURL")) + "," +
                                        SqlValue(reader.GetValue("DNIURLDO")) + "," +
                                        SqlDate(reader.GetValue("DATAURL1")) + "," +
                                        SqlValue(reader.GetValue("LBDNIROB1")) + "," +
                                        SqlDate(reader.GetValue("DATAURL2")) + "," +
                                        SqlValue(reader.GetValue("LBDNIROB2")) + "," +
                                        SqlDate(reader.GetValue("DATAURL3")) + "," +
                                        SqlValue(reader.GetValue("LBDNIROB3")) + "," +
                                        SqlValue(reader.GetValue("NRREJAUT")) + "," +
                                        SqlDate(reader.GetValue("DATPRZYAUT")) + "," +
                                        SqlValue(reader.GetValue("NRINWBIL")) + "," +
                                        SqlValue(reader.GetValue("LOGOPB")) + "," +
                                        SqlDate(reader.GetValue("DATPRZYBIL")) + "," +
                                        SqlValue(reader.GetValue("NRKARPAM")) + "," +
                                        SqlDate(reader.GetValue("DATPRZKP")) + "," +
                                        SqlValue(reader.GetValue("NAZPRZ")) + "," +
                                        SqlValue(reader.GetValue("NRPRZ")) + "," +
                                        SqlDate(reader.GetValue("DATAREJP")) + "," +
                                        SqlValue(reader.GetValue("GODZREJP")) + "," +
                                        SqlDate(reader.GetValue("DATAPOCZP")) + "," +
                                        SqlDate(reader.GetValue("DATAZAKP")) + "," +
                                        SqlValue(reader.GetValue("NRPRF")) + "," +
                                        SqlValue(reader.GetValue("LOGOPRF")) + "," +
                                        SqlValue(reader.GetValue("NRPZAD")) + "," +
                                        SqlValue(reader.GetValue("NRPKARDR")) + "," +
                                        SqlValue(reader.GetValue("NRKPPZAD")) + "," +
                                        SqlValue(reader.GetValue("NAZORZ")) + "," +
                                        SqlValue(reader.GetValue("NRORZ")) + "," +
                                        SqlDate(reader.GetValue("DATAREJO")) + "," +
                                        SqlValue(reader.GetValue("GODZREJO")) + "," +
                                        SqlDate(reader.GetValue("DATAPOCZO")) + "," +
                                        SqlDate(reader.GetValue("DATAZAKO")) + "," +
                                        SqlValue(reader.GetValue("NRORF")) + "," +
                                        SqlValue(reader.GetValue("LOGOORF")) + "," +
                                        SqlValue(reader.GetValue("NRZAD")) + "," +
                                        SqlValue(reader.GetValue("NRKARDR")) + "," +
                                        SqlValue(reader.GetValue("NRKPOZAD")) + "," +
                                        SqlValue(reader.GetValue("NRSPLAC")) + "," +
                                        SqlValue(reader.GetValue("WYSTNOTY")) + "," +
                                        SqlValue(reader.GetValue("WYSTFAKT")) + "," +
                                        SqlValue(reader.GetValue("DRUKDOK")) + "," +
                                        SqlValue(reader.GetValue("ETAT")) + "," +
                                        SqlValue(reader.GetValue("NUMERPRJ")) + "," +
                                        SqlDate(reader.GetValue("DATAWYDPRJ")) + "," +
                                        SqlDate(reader.GetValue("DATAWAZPRJ")) + "," +
                                        SqlValue(reader.GetValue("PRAWOJ")) + "," +
                                        SqlValue(reader.GetValue("DWAZLEK")) + "," +
                                        SqlValue(reader.GetValue("DWAZPSYCH")) + "," +
                                        SqlDate(reader.GetValue("DATAPKRT")) + "," +
                                        SqlDate(reader.GetValue("DATAKKRT")) + "," +
                                        SqlValue(reader.GetValue("NUMERKRT")) + "," +
                                        SqlValue(reader.GetValue("WYDKRT")) + "," +
                                        SqlDate(reader.GetValue("DATA_UR")) + "," +
                                        SqlValue(reader.GetValue("NR_DOWOD")) + "," +
                                        SqlValue(reader.GetValue("NR_PASZ")) + "," +
                                        SqlValue(reader.GetValue("PESEL")) + "," +
                                        SqlDate(reader.GetValue("DATABLEK")) + "," +
                                        SqlDate(reader.GetValue("DATABPSYCH")) + "," +
                                        SqlValue(reader.GetValue("NRZASKDPO")) + "," +
                                        SqlDate(reader.GetValue("DATKDPO")) + "," +
                                        SqlValue(reader.GetValue("NRZASKDPR")) + "," +
                                        SqlDate(reader.GetValue("DATKDPR")) + "," +
                                        SqlValue(reader.GetValue("NRTELKOM")) + "," +
                                        SqlValue(reader.GetValue("DTELKOM")) + "," +
                                        SqlDate(reader.GetValue("DATAOP")) + "," +
                                        SqlValue(reader.GetValue("GODZOP")) + "," +
                                        SqlValue(reader.GetValue("NR_SLUZBOP")) + ");";

                    insertPracownicy.Add(insert);
                    pracownicy++;
                }
            }
            Dispatcher.Invoke(() => textBox.Text += $"Pracownicy: {pracownicy} | ");
        }

        private void GenerateKartyPamieci()
        {
            string tempPathKarty = "";
            string KartyFile = "Kartypamieci.DBF";
            string KartyPath = $"bazy/";

            using (ZipArchive archive = ZipFile.OpenRead(zipFilePath))
            {
                var entryKarty = archive.Entries
                    .FirstOrDefault(e => e.Name.Equals(KartyFile, StringComparison.OrdinalIgnoreCase));
                if (entryKarty != null)
                {
                    tempPathKarty = Path.Combine(Path.GetTempPath(), entryKarty.Name.ToLower());
                    entryKarty.ExtractToFile(tempPathKarty, true);
                }
            }
            int karty = 0;
            using (var table = NDbfReader.Table.Open(tempPathKarty))
            {

                var encoding = Encoding.GetEncoding(852);
                var reader = table.OpenReader(encoding);

                while (reader.Read())
                {
                    string insert = "INSERT INTO IMP_KartyPamieci (NRKP, DATAPROD, POJEMN, TYP, BILETERKA, TYPBIL, ODDZ, DATAREJ, GODZREJ, NRPRACREJ, PINKARTY, NRSLUZBP, DATAZMPIN, NRKIER, IMIE, NAZWISKO, STATUS, STKP1, STKP2, MAXLPAS, NRWSIECI, KASA, LDNIREJ, DATAZMST, DATAWST, NROPST, DATAZAPST, GODZZAPST, BLOKADA, DATAWBLOK, NROPWB, DATAZBLOK, NROPZB, LZZBAKP, NAZWAZBA, DATAZZBA, GODZZZBA, NROPZZBA, NRPZAD, LZAD, PDZAD, ODZAD, NRKARDR, DATAZZAD, GODZZZAD, NROPZZAD, LZZADKP, DATAZIN, GODZZIN, NROPZIN, LZINKP, NAZPRZ, NRPRZ, DATAREJP, GODZREJP, DATAPOCZP, DATAZAKP, NRPRF, LOGOPRF, NRPZADKK, NRPKARDR, NRPRACPZ, IMIEPRPZ, NAZWPRPZ, STATPRPZ, NAZORZ, NRORZ, DATAREJO, GODZREJO, DATAPOCZO, DATAZAKO, NRORF, LOGOORF, NROZAD, NROKARDR, NRPRACOZ, IMIEPROZ, NAZWPROZ, STATPROZ, STKP3, LDNIBLOK, STKP4, INFOK1, INFOK2, MINDWOD, MAXPRZ, DATAWYCOF, DATALIKW, ST205_1, DATAOP, GODZOP, NR_SLUZBOP) VALUES (" +
                                        SqlValue(reader.GetValue("NRKP")) + "," +
                                        SqlDate(reader.GetValue("DATAPROD")) + "," +
                                        SqlValue(reader.GetValue("POJEMN")) + "," +
                                        SqlValue(reader.GetValue("TYP")) + "," +
                                        SqlValue(reader.GetValue("BILETERKA")) + "," +
                                        SqlValue(reader.GetValue("TYPBIL")) + "," +
                                        SqlValue(reader.GetValue("ODDZ")) + "," +
                                        SqlDate(reader.GetValue("DATAREJ")) + "," +
                                        SqlValue(reader.GetValue("GODZREJ")) + "," +
                                        SqlValue(reader.GetValue("NRPRACREJ")) + "," +
                                        SqlValue(reader.GetValue("PINKARTY")) + "," +
                                        SqlValue(reader.GetValue("NRSLUZBP")) + "," +
                                        SqlDate(reader.GetValue("DATAZMPIN")) + "," +
                                        SqlValue(reader.GetValue("NRKIER")) + "," +
                                        SqlValue(reader.GetValue("IMIE")) + "," +
                                        SqlValue(reader.GetValue("NAZWISKO")) + "," +
                                        SqlValue(reader.GetValue("STATUS")) + "," +
                                        SqlValue(reader.GetValue("STKP1")) + "," +
                                        SqlValue(reader.GetValue("STKP2")) + "," +
                                        SqlValue(reader.GetValue("MAXLPAS")) + "," +
                                        SqlValue(reader.GetValue("NRWSIECI")) + "," +
                                        SqlValue(reader.GetValue("KASA")) + "," +
                                        SqlValue(reader.GetValue("LDNIREJ")) + "," +
                                        SqlDate(reader.GetValue("DATAZMST")) + "," +
                                        SqlDate(reader.GetValue("DATAWST")) + "," +
                                        SqlValue(reader.GetValue("NROPST")) + "," +
                                        SqlDate(reader.GetValue("DATAZAPST")) + "," +
                                        SqlValue(reader.GetValue("GODZZAPST")) + "," +
                                        SqlValue(reader.GetValue("BLOKADA")) + "," +
                                        SqlDate(reader.GetValue("DATAWBLOK")) + "," +
                                        SqlValue(reader.GetValue("NROPWB")) + "," +
                                        SqlDate(reader.GetValue("DATAZBLOK")) + "," +
                                        SqlValue(reader.GetValue("NROPZB")) + "," +
                                        SqlValue(reader.GetValue("LZZBAKP")) + "," +
                                        SqlValue(reader.GetValue("NAZWAZBA")) + "," +
                                        SqlDate(reader.GetValue("DATAZZBA")) + "," +
                                        SqlValue(reader.GetValue("GODZZZBA")) + "," +
                                        SqlValue(reader.GetValue("NROPZZBA")) + "," +
                                        SqlValue(reader.GetValue("NRPZAD")) + "," +
                                        SqlValue(reader.GetValue("LZAD")) + "," +
                                        SqlValue(reader.GetValue("PDZAD")) + "," +
                                        SqlValue(reader.GetValue("ODZAD")) + "," +
                                        SqlValue(reader.GetValue("NRKARDR")) + "," +
                                        SqlDate(reader.GetValue("DATAZZAD")) + "," +
                                        SqlValue(reader.GetValue("GODZZZAD")) + "," +
                                        SqlValue(reader.GetValue("NROPZZAD")) + "," +
                                        SqlValue(reader.GetValue("LZZADKP")) + "," +
                                        SqlDate(reader.GetValue("DATAZIN")) + "," +
                                        SqlValue(reader.GetValue("GODZZIN")) + "," +
                                        SqlValue(reader.GetValue("NROPZIN")) + "," +
                                        SqlValue(reader.GetValue("LZINKP")) + "," +
                                        SqlValue(reader.GetValue("NAZPRZ")) + "," +
                                        SqlValue(reader.GetValue("NRPRZ")) + "," +
                                        SqlDate(reader.GetValue("DATAREJP")) + "," +
                                        SqlValue(reader.GetValue("GODZREJP")) + "," +
                                        SqlDate(reader.GetValue("DATAPOCZP")) + "," +
                                        SqlDate(reader.GetValue("DATAZAKP")) + "," +
                                        SqlValue(reader.GetValue("NRPRF")) + "," +
                                        SqlValue(reader.GetValue("LOGOPRF")) + "," +
                                        SqlValue(reader.GetValue("NRPZADKK")) + "," +
                                        SqlValue(reader.GetValue("NRPKARDR")) + "," +
                                        SqlValue(reader.GetValue("NRPRACPZ")) + "," +
                                        SqlValue(reader.GetValue("IMIEPRPZ")) + "," +
                                        SqlValue(reader.GetValue("NAZWPRPZ")) + "," +
                                        SqlValue(reader.GetValue("STATPRPZ")) + "," +
                                        SqlValue(reader.GetValue("NAZORZ")) + "," +
                                        SqlValue(reader.GetValue("NRORZ")) + "," +
                                        SqlDate(reader.GetValue("DATAREJO")) + "," +
                                        SqlValue(reader.GetValue("GODZREJO")) + "," +
                                        SqlDate(reader.GetValue("DATAPOCZO")) + "," +
                                        SqlDate(reader.GetValue("DATAZAKO")) + "," +
                                        SqlValue(reader.GetValue("NRORF")) + "," +
                                        SqlValue(reader.GetValue("LOGOORF")) + "," +
                                        SqlValue(reader.GetValue("NROZAD")) + "," +
                                        SqlValue(reader.GetValue("NROKARDR")) + "," +
                                        SqlValue(reader.GetValue("NRPRACOZ")) + "," +
                                        SqlValue(reader.GetValue("IMIEPROZ")) + "," +
                                        SqlValue(reader.GetValue("NAZWPROZ")) + "," +
                                        SqlValue(reader.GetValue("STATPROZ")) + "," +
                                        SqlValue(reader.GetValue("STKP3")) + "," +
                                        SqlValue(reader.GetValue("LDNIBLOK")) + "," +
                                        SqlValue(reader.GetValue("STKP4")) + "," +
                                        SqlValue(reader.GetValue("INFOK1")) + "," +
                                        SqlValue(reader.GetValue("INFOK2")) + "," +
                                        SqlValue(reader.GetValue("MINDWOD")) + "," +
                                        SqlValue(reader.GetValue("MAXPRZ")) + "," +
                                        SqlDate(reader.GetValue("DATAWYCOF")) + "," +
                                        SqlDate(reader.GetValue("DATALIKW")) + "," +
                                        SqlValue(reader.GetValue("ST205_1")) + "," +
                                        SqlDate(reader.GetValue("DATAOP")) + "," +
                                        SqlValue(reader.GetValue("GODZOP")) + "," +
                                        SqlValue(reader.GetValue("NR_SLUZBOP")) + ");";

                    insertKarty.Add(insert);
                    karty++;
                }
            }
            Dispatcher.Invoke(() => textBox.Text += $"Karty Pamięci: {karty}");
        }

        private void SaveSQLBtn_Click(object sender, RoutedEventArgs e)
        {
            if (insertKarty.Count > 0 || insertPracownicy.Count > 0 || insertKarty.Count > 0 )
            {
                SaveAll();
            }
            else
            {
                ToolTip tool = new ToolTip { Content = "Nie ma nic do zapisania.", StaysOpen = false };
                tool.IsOpen = true;
            }
        }

        public void SaveAll()
        {          
            var folderDialog = new CommonOpenFileDialog
            {
                IsFolderPicker = true,
                Title = "Wybierz folder do zapisania skryptów."
            };

            if (folderDialog.ShowDialog() == CommonFileDialogResult.Ok)
            {
                string folder = folderDialog.FileName;

                Dictionary<string, List<string>> scripts = new Dictionary<string, List<string>>
                {
                    { "KasyFiskalne", insertKasy },
                    { "Pracownicy", insertPracownicy },
                    { "KartyPamieci", insertKarty }
                };
                var newScriptFolder = Path.Combine(folder, $"{dane.Split(':')[1].Trim()}_import_danych_startowych");
                Directory.CreateDirectory(newScriptFolder);
                folder = newScriptFolder;

                var scriptGenerator = new ScriptGenerator();

                foreach(var el in scripts)
                {
                    if (el.Value.Count < 1) continue;

                    
                    string destFile = Path.Combine(folder, $"import_{el.Key}.sql");
                    var scriptTempPath = scriptGenerator.Generate(el.Key, el.Value);

                    File.WriteAllText(destFile, File.ReadAllText(scriptTempPath, Encoding.GetEncoding(1250)), Encoding.GetEncoding(1250));
                    File.Delete(scriptTempPath);
                }

                MessageBox.Show("Skrypty zostały zapisane w wybranym folderze.");
                Process.Start("explorer.exe", folder);
            }
        }
    }   
}