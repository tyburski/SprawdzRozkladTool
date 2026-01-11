using Microsoft.Win32;
using NDbfReader;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Reflection;
using System.Reflection.PortableExecutable;
using System.Runtime.Intrinsics.X86;
using System.Security.Policy;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using Microsoft.WindowsAPICodePack.Dialogs;



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
                                string dane = lines.FirstOrDefault(x => x.Contains("Dane firmy:"));
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
                MessageBox.Show($"Błąd: {ex.Message}\n\n{ex.StackTrace}");
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
            using (var table = NDbfReader.Table.Open(tempPathBileterki))
            {
                Dispatcher.Invoke(() => textBox.Text += $"Pomijam bileterki gdzie DATA TERAZ - DATAZORF > 1 ROK \n");
                Dispatcher.Invoke(() => textBox.Text += $"Pomijam bileterki gdzie LOGO jest puste \n");
                Dispatcher.Invoke(() => textBox.Text += $"Sprawdzam czy są bileterki EMAR-105 gdzie WERPSB < 1.5\n");
                var encoding = Encoding.GetEncoding(1250);
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
                    
                    if (typbil == "EMAR-105" && !string.IsNullOrWhiteSpace(werpsb) && float.TryParse(werpsb, out float werpsbValue) && werpsbValue < 1.5f)
                    {
                        Dispatcher.Invoke(() => textBox.Text += $"Bileterka EMAR-105 {logo} ma wersję oprogramowania {werpsb}!\n");
                    }

                    string insert = "INSERT INTO IMP_KasyFiskalne (LOGO, NRINW, NRFABR, NREWUS, DATAPROD, TYPBIL, DATAPRZ, NRPRF, DATAOPRF, NAZRFZPRF, LRZPRF, DATAZPRF, NAZRZZPRF, NRRZZPRF, NRKPPRF, NRPRZPRF, IMIEZPRF, NAZWZPRF, STATZPRF, NRORF, DATAOORF, NAZRFZORF, LRZORF, NAZRZORF, NRRZORF, DATAZORF, NRKPORF, NRPRZORF, IMIEZORF, NAZWZORF, STATZORF, LRF, DATAWBF, LOGOPOWBF, DATAKON, DATAFISK, DATAPRZEGL, ADRES, KODP, MIEJSC, NRDECUS, DATADECUS, DATAZAK, LMIESGWAR, UWAGI, BEZOPSERW, WERPSB, DATAOP, GODZOP, NR_SLUZBOP) VALUES (" +
                                        $"'{reader.GetValue("LOGO")?.ToString().Trim() ?? "NULL"}'," +
                                        $"'{reader.GetValue("NRINW")?.ToString().Trim() ?? "NULL"}'," +
                                        $"'{reader.GetValue("NRFABR")?.ToString().Trim() ?? "NULL"}'," +
                                        $"'{reader.GetValue("NREWUS")?.ToString().Trim() ?? "NULL"}'," +
                                        $"'{reader.GetValue("DATAPROD")?.ToString().Trim() ?? "NULL"}'," +
                                        $"'{reader.GetValue("TYPBIL")?.ToString().Trim() ?? "NULL"}'," +
                                        $"'{reader.GetValue("DATAPRZ")?.ToString().Trim() ?? "NULL"}'," +
                                        $"'{reader.GetValue("NRPRF")?.ToString().Trim() ?? "NULL"}'," +
                                        $"'{reader.GetValue("DATAOPRF")?.ToString().Trim() ?? "NULL"}'," +
                                        $"'{reader.GetValue("NAZRFZPRF")?.ToString().Trim() ?? "NULL"}'," +
                                        $"'{reader.GetValue("LRZPRF")?.ToString().Trim() ?? "NULL"}'," +
                                        $"'{reader.GetValue("DATAZPRF")?.ToString().Trim() ?? "NULL"}'," +
                                        $"'{reader.GetValue("NAZRZZPRF")?.ToString().Trim() ?? "NULL"}'," +
                                        $"'{reader.GetValue("NRRZZPRF")?.ToString().Trim() ?? "NULL"}'," +
                                        $"'{reader.GetValue("NRKPPRF")?.ToString().Trim() ?? "NULL"}'," +
                                        $"'{reader.GetValue("NRPRZPRF")?.ToString().Trim() ?? "NULL"}'," +
                                        $"'{reader.GetValue("IMIEZPRF")?.ToString().Trim() ?? "NULL"}'," +
                                        $"'{reader.GetValue("NAZWZPRF")?.ToString().Trim() ?? "NULL"}'," +
                                        $"'{reader.GetValue("STATZPRF")?.ToString().Trim() ?? "NULL"}'," +
                                        $"'{reader.GetValue("NRORF")?.ToString().Trim() ?? "NULL"}'," +
                                        $"'{reader.GetValue("DATAOORF")?.ToString().Trim() ?? "NULL"}'," +
                                        $"'{reader.GetValue("NAZRFZORF")?.ToString().Trim() ?? "NULL"}'," +
                                        $"'{reader.GetValue("LRZORF")?.ToString().Trim() ?? "NULL"}'," +
                                        $"'{reader.GetValue("NAZRZORF")?.ToString().Trim() ?? "NULL"}'," +
                                        $"'{reader.GetValue("NRRZORF")?.ToString().Trim() ?? "NULL"}'," +
                                        $"'{reader.GetValue("DATAZORF")?.ToString().Trim() ?? "NULL"}'," +
                                        $"'{reader.GetValue("NRKPORF")?.ToString().Trim() ?? "NULL"}'," +
                                        $"'{reader.GetValue("NRPRZORF")?.ToString().Trim() ?? "NULL"}'," +
                                        $"'{reader.GetValue("IMIEZORF")?.ToString().Trim() ?? "NULL"}'," +
                                        $"'{reader.GetValue("NAZWZORF")?.ToString().Trim() ?? "NULL"}'," +
                                        $"'{reader.GetValue("STATZORF")?.ToString().Trim() ?? "NULL"}'," +
                                        $"'{reader.GetValue("LRF")?.ToString().Trim() ?? "NULL"}'," +
                                        $"'{reader.GetValue("DATAWBF")?.ToString().Trim() ?? "NULL"}'," +
                                        $"'{reader.GetValue("LOGOPOWBF")?.ToString().Trim() ?? "NULL"}'," +
                                        $"'{reader.GetValue("DATAKON")?.ToString().Trim() ?? "NULL"}'," +
                                        $"'{reader.GetValue("DATAFISK")?.ToString().Trim() ?? "NULL"}'," +
                                        $"'{reader.GetValue("DATAPRZEGL")?.ToString().Trim() ?? "NULL"}'," +
                                        $"'{reader.GetValue("ADRES")?.ToString().Trim() ?? "NULL"}'," +
                                        $"'{reader.GetValue("KODP")?.ToString().Trim() ?? "NULL"}'," +
                                        $"'{reader.GetValue("MIEJSC")?.ToString().Trim() ?? "NULL"}'," +
                                        $"'{reader.GetValue("NRDECUS")?.ToString().Trim() ?? "NULL"}'," +
                                        $"'{reader.GetValue("DATADECUS")?.ToString().Trim() ?? "NULL"}'," +
                                        $"'{reader.GetValue("DATAZAK")?.ToString().Trim() ?? "NULL"}'," +
                                        $"'{reader.GetValue("LMIESGWAR")?.ToString().Trim() ?? "NULL"}'," +
                                        $"'{reader.GetValue("UWAGI")?.ToString().Trim() ?? "NULL"}'," +
                                        $"'{reader.GetValue("BEZOPSERW")?.ToString().Trim() ?? "NULL"}'," +
                                        $"'{reader.GetValue("WERPSB")?.ToString().Trim() ?? "NULL"}'," +
                                        $"'{reader.GetValue("DATAOP")?.ToString().Trim() ?? "NULL"}'," +
                                        $"'{reader.GetValue("GODZOP")?.ToString().Trim() ?? "NULL"}'," +
                                        $"'{reader.GetValue("NR_SLUZBOP")?.ToString().Trim() ?? "NULL"}'" +
                                        ");";


                    insertKasy.Add(insert);
                    bileterki++;
                }
            }
            Dispatcher.Invoke(() => textBox.Text += $"\nBileterki: {bileterki}\n");           
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

                var encoding = Encoding.GetEncoding(1250);
                var reader = table.OpenReader(encoding);

                while (reader.Read())
                {
                    string statusRaw = reader.GetValue("STATUS")?.ToString().Trim();
                    string status = (statusRaw != "7" && statusRaw != "8") ? "8" : (statusRaw ?? "NULL");

                    string insert = "INSERT INTO IMP_Pracownicy (NR_SLUZB, IMIE, NAZWISKO, STATUS, ODDZIAL, DATA_PRZYJ, DATA_ZWOL, TRYBROZWUM, ADRES, KODP, MIEJSC, SYM, GMINA, NAZKR, KOD2, KIER, TEL, TELKOM, ROKROZL, PLANDNIURL, DNIURLDO, DATAURL1, LBDNIROB1, DATAURL2, LBDNIROB2, DATAURL3, LBDNIROB3, NRREJAUT, DATPRZYAUT, NRINWBIL, LOGOPB, DATPRZYBIL, NRKARPAM, DATPRZKP, NAZPRZ, NRPRZ, DATAREJP, GODZREJP, DATAPOCZP, DATAZAKP, NRPRF, LOGOPRF, NRPZAD, NRPKARDR, NRKPPZAD, NAZORZ, NRORZ, DATAREJO, GODZREJO, DATAPOCZO, DATAZAKO, NRORF, LOGOORF, NRZAD, NRKARDR, NRKPOZAD, NRSPLAC, WYSTNOTY, WYSTFAKT, DRUKDOK, ETAT, NUMERPRJ, DATAWYDPRJ, DATAWAZPRJ, PRAWOJ, DWAZLEK, DWAZPSYCH, DATAPKRT, DATAKKRT, NUMERKRT, WYDKRT, DATA_UR, NR_DOWOD, NR_PASZ, PESEL, DATABLEK, DATABPSYCH, NRZASKDPO, DATKDPO, NRZASKDPR, DATKDPR, NRTELKOM, DTELKOM, DATAOP, GODZOP, NR_SLUZBOP) VALUES (" +
                        $"'{reader.GetValue("NR_SLUZB")?.ToString().Trim() ?? "NULL"}'," +
                        $"'{reader.GetValue("IMIE")?.ToString().Trim() ?? "NULL"}'," +
                        $"'{reader.GetValue("NAZWISKO")?.ToString().Trim() ?? "NULL"}'," +
                        $"'{status}'," +
                        $"'{reader.GetValue("ODDZIAL")?.ToString().Trim() ?? "NULL"}'," +
                        $"'{reader.GetValue("DATA_PRZYJ")?.ToString().Trim() ?? "NULL"}'," +
                        $"'{reader.GetValue("DATA_ZWOL")?.ToString().Trim() ?? "NULL"}'," +
                        $"'{reader.GetValue("TRYBROZWUM")?.ToString().Trim() ?? "NULL"}'," +
                        $"'{reader.GetValue("ADRES")?.ToString().Trim() ?? "NULL"}'," +
                        $"'{reader.GetValue("KODP")?.ToString().Trim() ?? "NULL"}'," +
                        $"'{reader.GetValue("MIEJSC")?.ToString().Trim() ?? "NULL"}'," +
                        $"'{reader.GetValue("SYM")?.ToString().Trim() ?? "NULL"}'," +
                        $"'{reader.GetValue("GMINA")?.ToString().Trim() ?? "NULL"}'," +
                        $"'{reader.GetValue("NAZKR")?.ToString().Trim() ?? "NULL"}'," +
                        $"'{reader.GetValue("KOD2")?.ToString().Trim() ?? "NULL"}'," +
                        $"'{reader.GetValue("KIER")?.ToString().Trim() ?? "NULL"}'," +
                        $"'{reader.GetValue("TEL")?.ToString().Trim() ?? "NULL"}'," +
                        $"'{reader.GetValue("TELKOM")?.ToString().Trim() ?? "NULL"}'," +
                        $"'{reader.GetValue("ROKROZL")?.ToString().Trim() ?? "NULL"}'," +
                        $"'{reader.GetValue("PLANDNIURL")?.ToString().Trim() ?? "NULL"}'," +
                        $"'{reader.GetValue("DNIURLDO")?.ToString().Trim() ?? "NULL"}'," +
                        $"'{reader.GetValue("DATAURL1")?.ToString().Trim() ?? "NULL"}'," +
                        $"'{reader.GetValue("LBDNIROB1")?.ToString().Trim() ?? "NULL"}'," +
                        $"'{reader.GetValue("DATAURL2")?.ToString().Trim() ?? "NULL"}'," +
                        $"'{reader.GetValue("LBDNIROB2")?.ToString().Trim() ?? "NULL"}'," +
                        $"'{reader.GetValue("DATAURL3")?.ToString().Trim() ?? "NULL"}'," +
                        $"'{reader.GetValue("LBDNIROB3")?.ToString().Trim() ?? "NULL"}'," +
                        $"'{reader.GetValue("NRREJAUT")?.ToString().Trim() ?? "NULL"}'," +
                        $"'{reader.GetValue("DATPRZYAUT")?.ToString().Trim() ?? "NULL"}'," +
                        $"'{reader.GetValue("NRINWBIL")?.ToString().Trim() ?? "NULL"}'," +
                        $"'{reader.GetValue("LOGOPB")?.ToString().Trim() ?? "NULL"}'," +
                        $"'{reader.GetValue("DATPRZYBIL")?.ToString().Trim() ?? "NULL"}'," +
                        $"'{reader.GetValue("NRKARPAM")?.ToString().Trim() ?? "NULL"}'," +
                        $"'{reader.GetValue("DATPRZKP")?.ToString().Trim() ?? "NULL"}'," +
                        $"'{reader.GetValue("NAZPRZ")?.ToString().Trim() ?? "NULL"}'," +
                        $"'{reader.GetValue("NRPRZ")?.ToString().Trim() ?? "NULL"}'," +
                        $"'{reader.GetValue("DATAREJP")?.ToString().Trim() ?? "NULL"}'," +
                        $"'{reader.GetValue("GODZREJP")?.ToString().Trim() ?? "NULL"}'," +
                        $"'{reader.GetValue("DATAPOCZP")?.ToString().Trim() ?? "NULL"}'," +
                        $"'{reader.GetValue("DATAZAKP")?.ToString().Trim() ?? "NULL"}'," +
                        $"'{reader.GetValue("NRPRF")?.ToString().Trim() ?? "NULL"}'," +
                        $"'{reader.GetValue("LOGOPRF")?.ToString().Trim() ?? "NULL"}'," +
                        $"'{reader.GetValue("NRPZAD")?.ToString().Trim() ?? "NULL"}'," +
                        $"'{reader.GetValue("NRPKARDR")?.ToString().Trim() ?? "NULL"}'," +
                        $"'{reader.GetValue("NRKPPZAD")?.ToString().Trim() ?? "NULL"}'," +
                        $"'{reader.GetValue("NAZORZ")?.ToString().Trim() ?? "NULL"}'," +
                        $"'{reader.GetValue("NRORZ")?.ToString().Trim() ?? "NULL"}'," +
                        $"'{reader.GetValue("DATAREJO")?.ToString().Trim() ?? "NULL"}'," +
                        $"'{reader.GetValue("GODZREJO")?.ToString().Trim() ?? "NULL"}'," +
                        $"'{reader.GetValue("DATAPOCZO")?.ToString().Trim() ?? "NULL"}'," +
                        $"'{reader.GetValue("DATAZAKO")?.ToString().Trim() ?? "NULL"}'," +
                        $"'{reader.GetValue("NRORF")?.ToString().Trim() ?? "NULL"}'," +
                        $"'{reader.GetValue("LOGOORF")?.ToString().Trim() ?? "NULL"}'," +
                        $"'{reader.GetValue("NRZAD")?.ToString().Trim() ?? "NULL"}'," +
                        $"'{reader.GetValue("NRKARDR")?.ToString().Trim() ?? "NULL"}'," +
                        $"'{reader.GetValue("NRKPOZAD")?.ToString().Trim() ?? "NULL"}'," +
                        $"'{reader.GetValue("NRSPLAC")?.ToString().Trim() ?? "NULL"}'," +
                        $"'{reader.GetValue("WYSTNOTY")?.ToString().Trim() ?? "NULL"}'," +
                        $"'{reader.GetValue("WYSTFAKT")?.ToString().Trim() ?? "NULL"}'," +
                        $"'{reader.GetValue("DRUKDOK")?.ToString().Trim() ?? "NULL"}'," +
                        $"'{reader.GetValue("ETAT")?.ToString().Trim() ?? "NULL"}'," +
                        $"'{reader.GetValue("NUMERPRJ")?.ToString().Trim() ?? "NULL"}'," +
                        $"'{reader.GetValue("DATAWYDPRJ")?.ToString().Trim() ?? "NULL"}'," +
                        $"'{reader.GetValue("DATAWAZPRJ")?.ToString().Trim() ?? "NULL"}'," +
                        $"'{reader.GetValue("PRAWOJ")?.ToString().Trim() ?? "NULL"}'," +
                        $"'{reader.GetValue("DWAZLEK")?.ToString().Trim() ?? "NULL"}'," +
                        $"'{reader.GetValue("DWAZPSYCH")?.ToString().Trim() ?? "NULL"}'," +
                        $"'{reader.GetValue("DATAPKRT")?.ToString().Trim() ?? "NULL"}'," +
                        $"'{reader.GetValue("DATAKKRT")?.ToString().Trim() ?? "NULL"}'," +
                        $"'{reader.GetValue("NUMERKRT")?.ToString().Trim() ?? "NULL"}'," +
                        $"'{reader.GetValue("WYDKRT")?.ToString().Trim() ?? "NULL"}'," +
                        $"'{reader.GetValue("DATA_UR")?.ToString().Trim() ?? "NULL"}'," +
                        $"'{reader.GetValue("NR_DOWOD")?.ToString().Trim() ?? "NULL"}'," +
                        $"'{reader.GetValue("NR_PASZ")?.ToString().Trim() ?? "NULL"}'," +
                        $"'{reader.GetValue("PESEL")?.ToString().Trim() ?? "NULL"}'," +
                        $"'{reader.GetValue("DATABLEK")?.ToString().Trim() ?? "NULL"}'," +
                        $"'{reader.GetValue("DATABPSYCH")?.ToString().Trim() ?? "NULL"}'," +
                        $"'{reader.GetValue("NRZASKDPO")?.ToString().Trim() ?? "NULL"}'," +
                        $"'{reader.GetValue("DATKDPO")?.ToString().Trim() ?? "NULL"}'," +
                        $"'{reader.GetValue("NRZASKDPR")?.ToString().Trim() ?? "NULL"}'," +
                        $"'{reader.GetValue("DATKDPR")?.ToString().Trim() ?? "NULL"}'," +
                        $"'{reader.GetValue("NRTELKOM")?.ToString().Trim() ?? "NULL"}'," +
                        $"'{reader.GetValue("DTELKOM")?.ToString().Trim() ?? "NULL"}'," +
                        $"'{reader.GetValue("DATAOP")?.ToString().Trim() ?? "NULL"}'," +
                        $"'{reader.GetValue("GODZOP")?.ToString().Trim() ?? "NULL"}'," +
                        $"'{reader.GetValue("NR_SLUZBOP")?.ToString().Trim() ?? "NULL"}');";

                    insertPracownicy.Add(insert);
                    pracownicy++;
                }
            }
            Dispatcher.Invoke(() => textBox.Text += $"\nPracownicy: {pracownicy}\n");
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

                var encoding = Encoding.GetEncoding(1250);
                var reader = table.OpenReader(encoding);

                while (reader.Read())
                {
                    string insert = "INSERT INTO IMP_KartyPamieci (NRKP, DATAPROD, POJEMN, TYP, BILETERKA, TYPBIL, ODDZ, DATAREJ, GODZREJ, NRPRACREJ, PINKARTY, NRSLUZBP, DATAZMPIN, NRKIER, IMIE, NAZWISKO, STATUS, STKP1, STKP2, MAXLPAS, NRWSIECI, KASA, LDNIREJ, DATAZMST, DATAWST, NROPST, DATAZAPST, GODZZAPST, BLOKADA, DATAWBLOK, NROPWB, DATAZBLOK, NROPZB, LZZBAKP, NAZWAZBA, DATAZZBA, GODZZZBA, NROPZZBA, NRPZAD, LZAD, PDZAD, ODZAD, NRKARDR, DATAZZAD, GODZZZAD, NROPZZAD, LZZADKP, DATAZIN, GODZZIN, NROPZIN, LZINKP, NAZPRZ, NRPRZ, DATAREJP, GODZREJP, DATAPOCZP, DATAZAKP, NRPRF, LOGOPRF, NRPZADKK, NRPKARDR, NRPRACPZ, IMIEPRPZ, NAZWPRPZ, STATPRPZ, NAZORZ, NRORZ, DATAREJO, GODZREJO, DATAPOCZO, DATAZAKO, NRORF, LOGOORF, NROZAD, NROKARDR, NRPRACOZ, IMIEPROZ, NAZWPROZ, STATPROZ, STKP3, LDNIBLOK, STKP4, INFOK1, INFOK2, MINDWOD, MAXPRZ, DATAWYCOF, DATALIKW, ST205_1, DATAOP, GODZOP, NR_SLUZBOP) VALUES (" +
                                        $"'{reader.GetValue("NRKP")?.ToString().Trim() ?? "NULL"}'," +
                                        $"'{reader.GetValue("DATAPROD")?.ToString().Trim() ?? "NULL"}'," +
                                        $"'{reader.GetValue("POJEMN")?.ToString().Trim() ?? "NULL"}'," +
                                        $"'{reader.GetValue("TYP")?.ToString().Trim() ?? "NULL"}'," +
                                        $"'{reader.GetValue("BILETERKA")?.ToString().Trim() ?? "NULL"}'," +
                                        $"'{reader.GetValue("TYPBIL")?.ToString().Trim() ?? "NULL"}'," +
                                        $"'{reader.GetValue("ODDZ")?.ToString().Trim() ?? "NULL"}'," +
                                        $"'{reader.GetValue("DATAREJ")?.ToString().Trim() ?? "NULL"}'," +
                                        $"'{reader.GetValue("GODZREJ")?.ToString().Trim() ?? "NULL"}'," +
                                        $"'{reader.GetValue("NRPRACREJ")?.ToString().Trim() ?? "NULL"}'," +
                                        $"'{reader.GetValue("PINKARTY")?.ToString().Trim() ?? "NULL"}'," +
                                        $"'{reader.GetValue("NRSLUZBP")?.ToString().Trim() ?? "NULL"}'," +
                                        $"'{reader.GetValue("DATAZMPIN")?.ToString().Trim() ?? "NULL"}'," +
                                        $"'{reader.GetValue("NRKIER")?.ToString().Trim() ?? "NULL"}'," +
                                        $"'{reader.GetValue("IMIE")?.ToString().Trim() ?? "NULL"}'," +
                                        $"'{reader.GetValue("NAZWISKO")?.ToString().Trim() ?? "NULL"}'," +
                                        $"'{reader.GetValue("STATUS")?.ToString().Trim() ?? "NULL"}'," +
                                        $"'{reader.GetValue("STKP1")?.ToString().Trim() ?? "NULL"}'," +
                                        $"'{reader.GetValue("STKP2")?.ToString().Trim() ?? "NULL"}'," +
                                        $"'{reader.GetValue("MAXLPAS")?.ToString().Trim() ?? "NULL"}'," +
                                        $"'{reader.GetValue("NRWSIECI")?.ToString().Trim() ?? "NULL"}'," +
                                        $"'{reader.GetValue("KASA")?.ToString().Trim() ?? "NULL"}'," +
                                        $"'{reader.GetValue("LDNIREJ")?.ToString().Trim() ?? "NULL"}'," +
                                        $"'{reader.GetValue("DATAZMST")?.ToString().Trim() ?? "NULL"}'," +
                                        $"'{reader.GetValue("DATAWST")?.ToString().Trim() ?? "NULL"}'," +
                                        $"'{reader.GetValue("NROPST")?.ToString().Trim() ?? "NULL"}'," +
                                        $"'{reader.GetValue("DATAZAPST")?.ToString().Trim() ?? "NULL"}'," +
                                        $"'{reader.GetValue("GODZZAPST")?.ToString().Trim() ?? "NULL"}'," +
                                        $"'{reader.GetValue("BLOKADA")?.ToString().Trim() ?? "NULL"}'," +
                                        $"'{reader.GetValue("DATAWBLOK")?.ToString().Trim() ?? "NULL"}'," +
                                        $"'{reader.GetValue("NROPWB")?.ToString().Trim() ?? "NULL"}'," +
                                        $"'{reader.GetValue("DATAZBLOK")?.ToString().Trim() ?? "NULL"}'," +
                                        $"'{reader.GetValue("NROPZB")?.ToString().Trim() ?? "NULL"}'," +
                                        $"'{reader.GetValue("LZZBAKP")?.ToString().Trim() ?? "NULL"}'," +
                                        $"'{reader.GetValue("NAZWAZBA")?.ToString().Trim() ?? "NULL"}'," +
                                        $"'{reader.GetValue("DATAZZBA")?.ToString().Trim() ?? "NULL"}'," +
                                        $"'{reader.GetValue("GODZZZBA")?.ToString().Trim() ?? "NULL"}'," +
                                        $"'{reader.GetValue("NROPZZBA")?.ToString().Trim() ?? "NULL"}'," +
                                        $"'{reader.GetValue("NRPZAD")?.ToString().Trim() ?? "NULL"}'," +
                                        $"'{reader.GetValue("LZAD")?.ToString().Trim() ?? "NULL"}'," +
                                        $"'{reader.GetValue("PDZAD")?.ToString().Trim() ?? "NULL"}'," +
                                        $"'{reader.GetValue("ODZAD")?.ToString().Trim() ?? "NULL"}'," +
                                        $"'{reader.GetValue("NRKARDR")?.ToString().Trim() ?? "NULL"}'," +
                                        $"'{reader.GetValue("DATAZZAD")?.ToString().Trim() ?? "NULL"}'," +
                                        $"'{reader.GetValue("GODZZZAD")?.ToString().Trim() ?? "NULL"}'," +
                                        $"'{reader.GetValue("NROPZZAD")?.ToString().Trim() ?? "NULL"}'," +
                                        $"'{reader.GetValue("LZZADKP")?.ToString().Trim() ?? "NULL"}'," +
                                        $"'{reader.GetValue("DATAZIN")?.ToString().Trim() ?? "NULL"}'," +
                                        $"'{reader.GetValue("GODZZIN")?.ToString().Trim() ?? "NULL"}'," +
                                        $"'{reader.GetValue("NROPZIN")?.ToString().Trim() ?? "NULL"}'," +
                                        $"'{reader.GetValue("LZINKP")?.ToString().Trim() ?? "NULL"}'," +
                                        $"'{reader.GetValue("NAZPRZ")?.ToString().Trim() ?? "NULL"}'," +
                                        $"'{reader.GetValue("NRPRZ")?.ToString().Trim() ?? "NULL"}'," +
                                        $"'{reader.GetValue("DATAREJP")?.ToString().Trim() ?? "NULL"}'," +
                                        $"'{reader.GetValue("GODZREJP")?.ToString().Trim() ?? "NULL"}'," +
                                        $"'{reader.GetValue("DATAPOCZP")?.ToString().Trim() ?? "NULL"}'," +
                                        $"'{reader.GetValue("DATAZAKP")?.ToString().Trim() ?? "NULL"}'," +
                                        $"'{reader.GetValue("NRPRF")?.ToString().Trim() ?? "NULL"}'," +
                                        $"'{reader.GetValue("LOGOPRF")?.ToString().Trim() ?? "NULL"}'," +
                                        $"'{reader.GetValue("NRPZADKK")?.ToString().Trim() ?? "NULL"}'," +
                                        $"'{reader.GetValue("NRPKARDR")?.ToString().Trim() ?? "NULL"}'," +
                                        $"'{reader.GetValue("NRPRACPZ")?.ToString().Trim() ?? "NULL"}'," +
                                        $"'{reader.GetValue("IMIEPRPZ")?.ToString().Trim() ?? "NULL"}'," +
                                        $"'{reader.GetValue("NAZWPRPZ")?.ToString().Trim() ?? "NULL"}'," +
                                        $"'{reader.GetValue("STATPRPZ")?.ToString().Trim() ?? "NULL"}'," +
                                        $"'{reader.GetValue("NAZORZ")?.ToString().Trim() ?? "NULL"}'," +
                                        $"'{reader.GetValue("NRORZ")?.ToString().Trim() ?? "NULL"}'," +
                                        $"'{reader.GetValue("DATAREJO")?.ToString().Trim() ?? "NULL"}'," +
                                        $"'{reader.GetValue("GODZREJO")?.ToString().Trim() ?? "NULL"}'," +
                                        $"'{reader.GetValue("DATAPOCZO")?.ToString().Trim() ?? "NULL"}'," +
                                        $"'{reader.GetValue("DATAZAKO")?.ToString().Trim() ?? "NULL"}'," +
                                        $"'{reader.GetValue("NRORF")?.ToString().Trim() ?? "NULL"}'," +
                                        $"'{reader.GetValue("LOGOORF")?.ToString().Trim() ?? "NULL"}'," +
                                        $"'{reader.GetValue("NROZAD")?.ToString().Trim() ?? "NULL"}'," +
                                        $"'{reader.GetValue("NROKARDR")?.ToString().Trim() ?? "NULL"}'," +
                                        $"'{reader.GetValue("NRPRACOZ")?.ToString().Trim() ?? "NULL"}'," +
                                        $"'{reader.GetValue("IMIEPROZ")?.ToString().Trim() ?? "NULL"}'," +
                                        $"'{reader.GetValue("NAZWPROZ")?.ToString().Trim() ?? "NULL"}'," +
                                        $"'{reader.GetValue("STATPROZ")?.ToString().Trim() ?? "NULL"}'," +
                                        $"'{reader.GetValue("STKP3")?.ToString().Trim() ?? "NULL"}'," +
                                        $"'{reader.GetValue("LDNIBLOK")?.ToString().Trim() ?? "NULL"}'," +
                                        $"'{reader.GetValue("STKP4")?.ToString().Trim() ?? "NULL"}'," +
                                        $"'{reader.GetValue("INFOK1")?.ToString().Trim() ?? "NULL"}'," +
                                        $"'{reader.GetValue("INFOK2")?.ToString().Trim() ?? "NULL"}'," +
                                        $"'{reader.GetValue("MINDWOD")?.ToString().Trim() ?? "NULL"}'," +
                                        $"'{reader.GetValue("MAXPRZ")?.ToString().Trim() ?? "NULL"}'," +
                                        $"'{reader.GetValue("DATAWYCOF")?.ToString().Trim() ?? "NULL"}'," +
                                        $"'{reader.GetValue("DATALIKW")?.ToString().Trim() ?? "NULL"}'," +
                                        $"'{reader.GetValue("ST205_1")?.ToString().Trim() ?? "NULL"}'," +
                                        $"'{reader.GetValue("DATAOP")?.ToString().Trim() ?? "NULL"}'," +
                                        $"'{reader.GetValue("GODZOP")?.ToString().Trim() ?? "NULL"}'," +
                                        $"'{reader.GetValue("NR_SLUZBOP")?.ToString().Trim() ?? "NULL"}'," +
                                        ");";

                    insertKarty.Add(insert);
                    karty++;
                }
            }
            Dispatcher.Invoke(() => textBox.Text += $"\nKarty Pamięci: {karty}\n");
        }

        private void SaveSQLBtn_Click(object sender, RoutedEventArgs e)
        {
            if (insertKarty.Count > 0 || insertPracownicy.Count > 0 || insertKarty.Count > 0 )
            {
                SaveAll();
            }
            else
            {
                MessageBox.Show("Nie ma nic do zapisania.");
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
            }
        }
    }   
}