using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ADODB;
using Hydra;
using System.Globalization;
using System.Reflection;
using cdn_api;
using Microsoft.VisualBasic;

[assembly: CallbackAssemblyDescription(Name: "ONES0221-CSVExport2",
                                           Description: "CSVExport2",
                                           Author: "OneSupport",
                                           Version: "1.0",
                                           SystemVersion: "2020.2.0",
                                           Created: "20-12-2022")]
namespace ZZnaCSV
{
    public class ToCSV
    {
        public string Indeks { get; set; }
        public string sztuk { get; set; }
        public string JM { get; set; }
        
    }


    [SubscribeProcedure(Procedures.ZaNLista, "CSVExport2")]
    public class ZZnaCSV : Callback
    {   
        private ClaWindow exportButton;
        private ClaWindow exportOpen;
        private ClaWindow pathButton;
        private ClaWindow pathOpen;
        private ClaWindow insertButton;
        private int XLWersjaAPI = 20202;
        private int XLTrybWsadowy = 1;
        private int XLWinieta = -1;
        private int XLUtworzWlasnaSesje = 1;
        private int sesja = -1;
        private string XLSQLServer = "192.168.4.15";   //"192.168.137.149";
        private string XLDatabase = "rpxl_2020_2";// "Domar2";
        private string XLUser = "Sabisup"; //"ADMIN";
        private string XLPass = "#0ut2510gro0ne.34";//"Dupa2003!";
        private string path = "";
        private string connection_string;
        private int listID;
        public class Foo
{
    public int Id { get; set; }
    public string Name { get; set; }
}
        public override void Init()
        {
            AddSubscription(true, 0, Events.JustAfterWindowOpening, new TakeEventDelegate(OnOpenwindow));
            AddSubscription(false, 0, Events.ResizeWindow, new TakeEventDelegate(Resize));
        }
        public override void Cleanup()
        {

        }
        public bool OnOpenwindow(Procedures ProcId, int ControlId, Events Event)
        {
            ClaWindow Parent = GetWindow();
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Rectangle WndBnd = GetWindow().Bounds;
            exportButton = GetWindow().AllChildren["?TabZam"];
            pathButton = GetWindow().AllChildren["?TabZam"];
            GetWindow().OnBeforeSized += new TakeEventDelegate(Resize);
            exportOpen = exportButton.Children.Add(ControlTypes.button);
            pathOpen = pathButton.Children.Add(ControlTypes.button);
            insertButton = GetWindow().AllChildren["?Close"];
            
            exportOpen.Bounds = new Rectangle((insertButton.Bounds.X - 20), insertButton.Bounds.Y - 3, 19, 17);
            exportOpen.TextRaw = "Do CSV2";
            exportOpen.ToolTipRaw = "Wyeksportuj zaznaczone towary do pliku CSV2";
            exportOpen.Visible = true;
            pathOpen.Bounds = new Rectangle((exportOpen.Bounds.X), exportOpen.Bounds.Y - 12, 12, 10);
            pathOpen.TextRaw = "..";
            pathOpen.ToolTipRaw = "Wybierz ścieżkę do eksportu CSV2";
            pathOpen.Visible = true;

            exportOpen.OnAfterAccepted += new TakeEventDelegate(OpenForm);

            pathOpen.OnAfterAccepted += new TakeEventDelegate(changePath);
            listID = GetWindow().AllChildren["?ListaZam"].Id;

            //popFiltr = GetWindow().AllChildren["?bUkryjFiltrDodatkowy"];
            //naduszenie przycisku
            return true;
        }
        private bool changePath(Procedures ProcedureId, int ControlId, Events Event)
        {
            using (var fbd = new FolderBrowserDialog())
            {
                DialogResult result = fbd.ShowDialog();

                if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(fbd.SelectedPath))
                {
                    try
                    {
                        path = fbd.SelectedPath;
                        File.WriteAllText(@"C:\OneSupport\Extradach\CSV\path2.txt", fbd.SelectedPath);
                        MessageBox.Show("Pliki będą zapisywane do ścieżki "+System.IO.File.ReadAllText(fbd.SelectedPath + @"\path2.txt"));
                    }
                    catch (Exception)
                    {

                        throw;
                    }

                }
            }
            return true;
        }
        private bool OpenForm(Procedures ProcedureId, int ControlId, Events Event)
        {
            Cursor.Current = Cursors.WaitCursor;
            List<ZamElemList> listaZamowien = new List<ZamElemList>();
            listaZamowien.Clear();
            _Recordset recordset = Runtime.WindowController.GetQueueMarked((int)ProcedureId, listID, GetCallbackThread());
            connection_string = GetConnectionString();
            try
            {
                path = System.IO.File.ReadAllText(@"C:\OneSupport\Extradach\CSV\path2.txt");
                var records = new List<ToCSV>();
                System.IO.Directory.CreateDirectory(path);
                if (recordset != null)
                {
                    int count = recordset.RecordCount;

                    if (count != 0)
                    {
                        recordset.MoveFirst();

                        while (recordset.EOF == false)
                        {

                            ADODB.Fields fields = recordset.Fields;
                            listaZamowien.Add(new ZamElemList()
                            {

                                GIDTyp = Convert.ToInt16(fields[0].Value),
                                GIDFirma = Convert.ToInt32(fields[1].Value),
                                GIDNumer = Convert.ToInt32(fields[2].Value),
                                GIDLp = Convert.ToInt16(fields[3].Value),

                            });

                            recordset.MoveNext();
                        }
                        
                        string msg = "select distinct count(DISTINCT ZaN_KntNumer) from cdn.ZamNag where ";
                        string text = "";
                        string subquery = "(select AtK_ID from CDN.AtrybutyKlasy where AtK_Nazwa = 'Parametry x#y#z')";
                        for (int i = 0; i < count; i++)
                        {
                            //MessageBox.Show(listaZamowien[i].GIDNumer.ToString());
                            if (i==0)
                            {
                                msg += "ZaN_KntNumer=(select ZaN_KntNumer from cdn.ZamNag where ZaN_GIDNumer =" + listaZamowien[i].GIDNumer+")";
                            }
                            else
                            {
                                msg += " or ZaN_KntNumer=(select ZaN_KntNumer from cdn.ZamNag where ZaN_GIDNumer =" + listaZamowien[i].GIDNumer+")";
                            }

                        } 
                        string Knt_amnt = ConnectToSQL(connection_string, msg);
                        if (Knt_amnt =="1")
                        {
                            var format = ConnectToSQL(connection_string, "select Atr_Wartosc from cdn.Atrybuty where Atr_ObiNumer = (select ZaN_KntNumer from cdn.ZamNag where ZaN_GIDNumer=" + listaZamowien[0].GIDNumer + ") and Atr_AtkId = (select AtK_ID from CDN.AtrybutyKlasy where AtK_Nazwa = 'format_eksportu')");
                            if (format.Length<2)
                            {
                                format = "txt";
                            }
                            //TODO: Kontrola formatu przy pomocy atybutu kontrahenta (domyslnie tylko CSV)
                            //MessageBox.Show(format);
                            if(!Directory.Exists(path))
                            {
                                throw new ArgumentException("Wybrana ścieżka nie istnieje! Zaktualizuj ścieżkę do eksportu plików obecna to: "+path);
                            }
                            for (int i = 0; i < count; i++)
                            {
                                
                                msg = "select CONCAT(ZaE_TwrNumer,' ',ZaE_TwrTyp,' ',ZaE_TwrLp,' ',ZaE_TwrFirma) from cdn.ZamElem where ZaE_GIDTyp = " + listaZamowien[i].GIDTyp + "and ZaE_GIDNumer = " + listaZamowien[i].GIDNumer +" and ZaE_GIDFirma=" + listaZamowien[i].GIDFirma;
                                string[] twr = new string[4];
                                int gidlp = 0;
                                foreach (var item in ConnectListToSQL(connection_string, msg))
                                {
                                    gidlp++;
                                    string dlugosc,sztuk;
                                    dlugosc = sztuk = "";
                                    twr = item.Split(' ');
                                    msg = "if exists (select 1 from cdn.Atrybuty where Atr_ObiTyp =" + twr[1] + " AND Atr_ObiNumer = " + twr[0] + " AND Atr_ObiLp =" + twr[2] + " AND Atr_ObiFirma =" + twr[3] + " AND Atr_AtkId = " + subquery + ") begin "
                                    + "select Atr_Wartosc from cdn.Atrybuty where Atr_ObiTyp =" + twr[1] + " AND Atr_ObiNumer = " + twr[0] + " AND Atr_ObiLp =" + twr[2] + " AND Atr_ObiSubLp =" + twr[3] + " AND Atr_AtkId = " + subquery + " end else begin select 'brak' end";

                                    if (ConnectToSQL(connection_string, msg).ToString() == "brak")
                                    {
                                        MessageBox.Show("Element o GIDNumer " + twr[0] + " Typ " + twr[1] + " nie posiada wymaganego parametru.", "KOMUNIKAT", MessageBoxButtons.OK, MessageBoxIcon.Question);
                                    }
                                    else
                                    {
                                        msg = "select TwK_Kod from cdn.TwrKody where TwK_TypKodu = 3 and TwK_TwrNumer = " + twr[0];
                                        msg = ConnectToSQL(connection_string, msg);

                                        string jm= ConnectToSQL(connection_string, "SELECT ZaE_JmZ FROM cdn.ZamElem where ZaE_GIDTyp = " + listaZamowien[i].GIDTyp + "and ZaE_GIDNumer = " + listaZamowien[i].GIDNumer + " and ZaE_GIDFirma=" + listaZamowien[i].GIDFirma); 
                                        foreach (var lista in ConnectListToSQL(connection_string, "select Concat(Ilosc,' ', Dlugosc) from cdn.One_ZamMacierz where GIDTyp=" + listaZamowien[i].GIDTyp + "and GIDNumer=" + listaZamowien[i].GIDNumer + " and GIDLp=" + gidlp))
                                        {
                                            string[] twr_szcz = new string[2];
                                            twr_szcz = lista.Split(' ');
                                            sztuk = twr_szcz[0].Replace(",", ".").Split('.')[0];
                                            dlugosc = twr_szcz[1].Replace(",",".");
                                            
                                            //SQL QUERY - Wszystki wartości z tabeli o danym GID, wrzucić do listy, pewnie foreach. jak pusty to pusta linijka.
                                            if (msg != "" && sztuk != "" && dlugosc != "")
                                            {
                                                records.Add(new ToCSV() { Indeks = msg,  sztuk = sztuk,  JM = jm });
                                            }
                                            else if (msg != "")
                                            {
                                                records.Add(new ToCSV() { Indeks = msg, sztuk = "1", JM = "m2"});
                                                break;
                                            }
                                            else break;
                                        }
                                    }
                                }
                                
                                CreateCSV(records, path + "\\" + listaZamowien[i].GIDNumer +"."+format);
                            }
                            //string csv = String.Join(",", records.Select(x => x.ToString()).ToArray());
                            MessageBox.Show("Eksport zakonczony!", "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        }
                        else
                        {
                            if (Int32.TryParse(Knt_amnt, out int res)==true)
                            {
                            MessageBox.Show($"Wybrano różnych kontrahentów - ilość: {res}");
                            }
                            else
                            {
                                MessageBox.Show($"Błąd połączenia SQL: {Knt_amnt}");

                            }



                        }

                    }
                    else
                    {
                        MessageBox.Show("Nie wybrano żadnego zamówienia!", "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        return true;
                    }
                }
                else
                {
                    MessageBox.Show("Nie wybrano żadnego zamówienia!", "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Błąd pobierania zaznaczonych dokumentów --> " + ex.Message, "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return true;
            }

            //MessageBox.Show("CosDziala, a dane to " + listaZamowien[0].GIDTyp + "  " + listaZamowien[0].GIDFirma + "  " + listaZamowien[0].GIDNumer + "  " + listaZamowien[0].GIDLp, "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            return true;
        }
        private static void CreateHeader<T>(List<T> list, StreamWriter sw)
        {
            PropertyInfo[] properties = typeof(T).GetProperties();
            for (int i = 0; i < properties.Length - 1; i++)
            {
                sw.Write(properties[i].Name + ",");
            }
            var lastProp = properties[properties.Length - 1].Name;
            sw.Write(lastProp + sw.NewLine);
        }
        private static void CreateRows<T>(List<T> list, StreamWriter sw)
        {
            foreach (var item in list)
            {
                PropertyInfo[] properties = typeof(T).GetProperties();
                for (int i = 0; i < properties.Length - 1; i++)
                {
                    var prop = properties[i];
                    sw.Write(prop.GetValue(item) + ";");
                }
                var lastProp = properties[properties.Length - 1];
                sw.Write(lastProp.GetValue(item) + sw.NewLine);
            }
        }
        public static void CreateCSV<T>(List<T> list, string filePath)
        {
            using (StreamWriter sw = new StreamWriter(filePath))
            {
                //CreateHeader(list, sw);
                CreateRows(list, sw);
            }
        }
        public int XLLogin()
        {
            return cdn_api.cdn_api.XLLogin(new XLLoginInfo_20202()
            {
                Wersja = 20202,
                ProgramID = "OONES0221-CSVExport2",
                //PlikLog = (Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\\api.log"),
                //PlikLog = (@"C:\OneSupport\ONE-S01_GEN_AWF\api.log"),
                Baza = XLDatabase,
                OpeIdent = XLUser,
                OpeHaslo = XLPass,
                TrybWsadowy = XLTrybWsadowy,
                //TrybNaprawy = 0,
                UtworzWlasnaSesje = XLUtworzWlasnaSesje,
                Winieta = XLWinieta,
            }, ref sesja);
        }
        public void Logout()
        {
            cdn_api.cdn_api.XLLogout(sesja);
        }
        public string GetConnectionString()
        {
            XLLogin();
            XLPolaczenieInfo_20202 polaczenieInfo = new XLPolaczenieInfo_20202 { Wersja = 20202 };
            cdn_api.cdn_api.XLPolaczenie(polaczenieInfo);

            //connection_string = @"Provider=MSOLEDBSQL.1;Password=43gQdfs31so3;Persist Security Info=True;User ID=onesupport_dod;Initial Catalog="";Data Source=SRV1;Initial File Name="";Server SPN="";Authentication="";Access Token=""";
            string connection_string = "Server=" + polaczenieInfo.Serwer + ";Database=" + polaczenieInfo.Baza + ";User Id=" + XLUser + ";Password=" + XLPass;
            //MessageBox.Show(connection_string);
             
            Logout();
            return connection_string;
        }

        public string ConnectToSQL(string connection_string, string query)
        {
            string str = "";
            try
            {
                using (SqlConnection connection = new SqlConnection(connection_string))
                {
                    SqlDataAdapter adapter = new SqlDataAdapter();
                    adapter.SelectCommand = new SqlCommand(query, connection);
                    DataSet dataSet = new DataSet();
                    adapter.Fill(dataSet, "Query");
                    if (dataSet.Tables.Count == 0)
                    {
                        return "";
                    }
                    foreach (DataRow row in (InternalDataCollectionBase)dataSet.Tables[0].Rows)
                    {
                        str = row[0].ToString();
                    }
                }
            }
            catch (Exception ex)
            {
                str = str + "Bład SQL:" + ex.ToString() + "\n";
            }
            return str;
        }
        public List<string> ConnectListToSQL(string connection_string, string query)
        {
            List<String> columnData = new List<String>();

            using (SqlConnection connection = new SqlConnection(connection_string))
            {
                connection.Open();
                
                using (SqlCommand command = new SqlCommand(query, connection))
                {
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            columnData.Add(reader.GetString(0));
                        }
                    }
                }
            }
            return columnData;
        }
        private bool Resize(Procedures ProcedureId, int ControlId, Events Event)
        {
            insertButton = GetWindow().AllChildren["?Close"];
            exportOpen.Bounds = new Rectangle(insertButton.Bounds.X-20, insertButton.Bounds.Y- 23, 19, 17);
            pathOpen.Bounds = new Rectangle((exportOpen.Bounds.X), exportOpen.Bounds.Y - 12, 12, 10);
            return true;
        }

    }
}
