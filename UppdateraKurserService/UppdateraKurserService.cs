using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using System.Timers;
using System.Net.Http;
//using MySql.Data.MySqlClient;
using Newtonsoft.Json;
using System.Configuration;
using System.Data.SqlClient;

/* Den större delan av aktiekurser använder jag finnhub.io för att uppdatera då jag får göra mer frekvent API anrop än vad man är tillåten mot Alpha Vantage.
 * Däremot så har inte finnhub alla svenska aktier, de som de saknar hämtar jag istället med Alpha Vantage
 * Dock så måste jag ändra symbolen från .ST till .STO i tabellen för att detta ska funka.
 * 
 * Lägg till dessa i App.config under SpecialAkiter. Om flera avskilj med _
 * 
 * 
 * 
 * 
 * 
 * 
 */




namespace UppdateraKurserService
{
    static class GlobalVars
    {
        public static string conString = "";
        public static string SQLconString = "";
        public static string quote = "\"";
        public static string TablesToUpdate = "";
        public static int Delay = 0;
        public static string SpecialAktier = "";
    }



    public partial class UppdateraKurserService : ServiceBase
    {

        System.Timers.Timer timer = null;
        string runAtStart = "";
        string runAtHour = "";

        public UppdateraKurserService()
        {
            InitializeComponent();
            //    <add key="TablesToUpdate"  value="AF_KF_ISK_IPS_TJP" />
            //<add key="SpecialAktier"  value="PARA.STO_CLS-B.STO_NENT-B.STO" />
            //timer = new System.Timers.Timer(2700000); // Intervallet i millisekunder mellan körningarna av ElapsedEventHandler
            //timer = new System.Timers.Timer(300000); // Intervallet i millisekunder mellan körningarna av ElapsedEventHandler
            timer = new System.Timers.Timer(50000); // Intervallet i millisekunder mellan körningarna av ElapsedEventHandler
            timer.Elapsed += new ElapsedEventHandler(OnTimedEvent); // Sätter ElapsedEventHandler till subrutinen OnTimedEvent
        }

        protected override void OnStart(string[] args)
        {
            //Kickar igång debuggern vid uppstart
            //System.Diagnostics.Debugger.Launch();

            GlobalVars.conString = ConfigurationManager.AppSettings["MySqlConnectionString"];
            GlobalVars.SQLconString = ConfigurationManager.AppSettings["SqlConnectionString"];
            GlobalVars.TablesToUpdate = ConfigurationManager.AppSettings["TablesToUpdate"];
            GlobalVars.Delay = Int32.Parse(ConfigurationManager.AppSettings["Delay"]);
            runAtStart = ConfigurationManager.AppSettings["runAtStart"];
            runAtHour = ConfigurationManager.AppSettings["runAtHour"];

            Logger("INFO", "Starting Uppdatera Kurser service");

            timer.Start();
        }

        protected override void OnStop()
        {
            Logger("INFO", "Stopping Uppdatera Kurser service");
        }

        private void OnTimedEvent(object source, ElapsedEventArgs e)
        {
            timer.Stop();
            char[] delim = { ',' };
           
            foreach (object AtHour in runAtHour.Split(delim))
            {
       
                if (DateTime.Now.Hour.ToString() == AtHour.ToString().Trim())
                {
                    Logger("DEBUG", "Nu kör vi !! " + DateTime.Now.Hour.ToString());

                    if (GlobalVars.TablesToUpdate.Contains("_"))
                    {
                        string[] tables = GlobalVars.TablesToUpdate.Split('_');

                        foreach (string table in tables)
                        {
                            GetStockPrices(table);
                            GetStockPriceSpecial(table);
                        }

                    }
                    else
                    {
                        GetStockPrices(GlobalVars.TablesToUpdate);
                        GetStockPriceSpecial(GlobalVars.TablesToUpdate);
                    }
                    GetCryptoPrices();
                    UpdateTotal();
                }
            }
            timer.Start();
        }

        public static void Logger(string type, string message)
        {
            try
            {
                using (SqlConnection connection = new SqlConnection(GlobalVars.SQLconString))
                {
                    connection.Open();

                    string CommandText = "INSERT INTO money.log (type,message) values(@type,@message)";
                    SqlCommand command = new SqlCommand(CommandText, connection);

                    command.Parameters.AddWithValue("@type", type);
                    command.Parameters.AddWithValue("@message", message);
                    command.ExecuteNonQuery();
                }
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        public class FinnhubData
        {
            public decimal C { get; set; }
            public decimal H { get; set; }
            public decimal L { get; set; }
            public decimal O { get; set; }
            public decimal PC { get; set; }
        }


        public static void GetCryptoPrices()
        {
            //  Welcome to Alpha Vantage! Here is your API key: APA3UI90FWXJA9IM
            //  string ApiURL = "https://www.alphavantage.co/query?function=CURRENCY_EXCHANGE_RATE&from_currency=BTC&to_currency=SEK&apikey=APA3UI90FWXJA9IM";
            // NOTERA: För att använda detta så måste stockholms börsens aktier markeras med STO och inte ST
            
            string sqlcmnd = "SELECT * FROM money.crypto";

            DataTable dt = new DataTable();
            var client = new System.Net.WebClient();
            var apiKey = "APA3UI90FWXJA9IM";

            //Logger("INFO", "Updatera priser för Krypto");

            try
            {
                using (SqlConnection connection = new SqlConnection(GlobalVars.SQLconString))
                {
                    connection.Open();

                    using (SqlCommand myCommand = new SqlCommand(sqlcmnd, connection))
                    {

                        using (SqlDataAdapter mysqlDa = new SqlDataAdapter(myCommand))
                            mysqlDa.Fill(dt);

                        foreach (DataRow row in dt.Rows)
                        {
                            string symbol = row[10].ToString();
                            string responseBody = "";

                            try
                            {
                                string url = $"https://www.alphavantage.co/query?function=CURRENCY_EXCHANGE_RATE&from_currency={symbol}&to_currency=SEK&apikey={apiKey}";
                                responseBody = client.DownloadString(url);

                                int position = responseBody.IndexOf("Swedish Krona");
                                string substring = responseBody.Substring(position + 45, 30);
                                int endposition = substring.IndexOf("\",");
                                string sekprice = substring.Substring(0, endposition);

                                //Logger("INFO0", "Position : " + position + " EndPosition: " + endposition);

                                decimal CurrentOpenPrice = decimal.Parse(sekprice);

                                System.Threading.Thread.Sleep(GlobalVars.Delay);

                                UpdateInvestment("crypto", symbol, CurrentOpenPrice, 1);
                                //Logger("INFO1", "Updating Crypto : " + symbol + " " + CurrentOpenPrice);

                            }
                            catch (Exception ex)
                            {
                                Logger("ERROR1", ex.Message);
                            }

                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Logger("ERROR2", ex.Message);
            }
        }

        public static void GetStockPrices(string table)
        {
            // key: bo5suuvrh5rbvm1sl1t0   https://finnhub.io/dashboard
            // https://finnhub.io/api/v1/quote?symbol=AAPL&token=bo5suuvrh5rbvm1sl1t0

            string sqlcmnd = "SELECT * FROM money." + table + ";";
            string apiKey = "bo5suuvrh5rbvm1sl1t0";
            decimal rate = 1;

            DataTable dt = new DataTable();
            var client = new System.Net.WebClient();

            //Logger("INFO", "Uppdatera priser för " + table);

            try
            {
                using (SqlConnection connection = new SqlConnection(GlobalVars.SQLconString))
                {
                    connection.Open();

                    using (SqlCommand myCommand = new SqlCommand(sqlcmnd, connection))
                    {
                        using (SqlDataAdapter mysqlDa = new SqlDataAdapter(myCommand))
                            mysqlDa.Fill(dt);

                        foreach (DataRow row in dt.Rows)
                        {
                            string symbol = row[10].ToString();
                            string valuta = row[9].ToString();

                            // Lite fullösning, men här behöver vi inte kolla upp aktier som har en symbol som slutar på .STO.
                            // Detta tas hand om av GetStockPriceSpecial
                            // Vi passar på att även hoppa över kontanter.

                            if (!symbol.Contains(".STO") && !symbol.Contains("Kontanter"))
                            { 
                                try
                                {
                                    string url = $"https://finnhub.io/api/v1/quote?symbol={symbol}&token={apiKey}";
                                
                                    string responseBody = client.DownloadString(url);
                                    FinnhubData StockData = JsonConvert.DeserializeObject<FinnhubData>(responseBody);
                                    decimal CurrentOpenPrice = StockData.C;

                                    if (string.Compare(valuta, "SEK") != 0)
                                    {
                                        rate = ConvertExchangeRates(valuta);
                                    }

                                    System.Threading.Thread.Sleep(GlobalVars.Delay);
                                    UpdateInvestment(table, symbol, CurrentOpenPrice,rate);
                                    rate = 1;
                                    //Logger("INFO", table + "." + symbol + " " + CurrentOpenPrice);
                                }
                                catch (Exception ex)
                                {
                                    Logger("ERROR3", symbol + " " + ex.Message);
                                }

                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Logger("ERROR4", ex.Message);
            }

        }

        public static void GetStockPriceSpecial(string table)
        {
            //  Welcome to Alpha Vantage! Here is your API key: APA3UI90FWXJA9IM
            //  string ApiURL = "https://www.alphavantage.co/query?function=TIME_SERIES_INTRADAY&symbol=MSFT&interval=1min&apikey=APA3UI90FWXJA9IM";
            // NOTERA: För att använda detta så måste stockholms börsens aktier markeras med STO och inte ST

            GlobalVars.SpecialAktier = ConfigurationManager.AppSettings["SpecialAktier"];
            string sqlcmnd = "";
            decimal rate = 1;

            //Logger("INFO", "Uppdatera (special) priser för " + table);

            if (GlobalVars.SpecialAktier.Contains("_"))
            {
                sqlcmnd = "SELECT * FROM money." + table + " WHERE Symbol = ";
                string[] aktier = GlobalVars.SpecialAktier.Split('_');
                string aktietemp = GlobalVars.quote + aktier[0] + GlobalVars.quote;

                for (int i=1; i < aktier.Length;i++)
                {
                    aktietemp = aktietemp + " OR Symbol = " + GlobalVars.quote + aktier[i] + GlobalVars.quote;
                }

                sqlcmnd = sqlcmnd + aktietemp + ";";
            }
            else
            {
                sqlcmnd = "SELECT * FROM money." + table + " WHERE Symbol = ' + GlobalVars.quote + GlobalVars.SpecialAktier + GlobalVars.quote + ';";
            }


            DataTable dt = new DataTable();
            var client = new System.Net.WebClient();
            var apiKey = "APA3UI90FWXJA9IM";


            try
            {
                using (SqlConnection connection = new SqlConnection(GlobalVars.SQLconString))
                {
                    connection.Open();

                    using (SqlCommand myCommand = new SqlCommand(sqlcmnd, connection))
                    {

                        using (SqlDataAdapter mysqlDa = new SqlDataAdapter(myCommand))
                            mysqlDa.Fill(dt);

                        foreach (DataRow row in dt.Rows)
                        {
                            string symbol = row[10].ToString();
                            string valuta = row[9].ToString();
                            string responseBody = "";

                            try
                            {
                                string url = $"https://www.alphavantage.co/query?function=TIME_SERIES_INTRADAY&symbol={symbol}&interval=1min&apikey={apiKey}&datatype=csv";

                                responseBody = client.DownloadString(url);
                                string[] tmpString = responseBody.Split(new string[] {System.Environment.NewLine}, StringSplitOptions.None);
                                tmpString = tmpString[1].Split(',');
                                decimal CurrentOpenPrice = decimal.Parse(tmpString[1]);
                                System.Threading.Thread.Sleep(GlobalVars.Delay);

                                if (string.Compare(valuta, "SEK") != 0)
                                {
                                    rate = ConvertExchangeRates(valuta);
                                }


                                UpdateInvestment(table, symbol, CurrentOpenPrice,rate);
                                rate = 1;
                                //Logger("INFO", "SPECIAL: " + table + "." + symbol + " " + CurrentOpenPrice);

                            }
                            catch (Exception ex)
                            {
                                Logger("ERROR5", table + "." + symbol + " " + responseBody + " " + ex.Message);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Logger("ERROR6", ex.Message);
            }
        }

        public static void UpdateInvestment(string table, string symbol, decimal Kurs, decimal rate)
        {
            //Uppdaterar Kurs och SEKKurs 

            decimal SEKKurs = Kurs * rate;
            string sqlcmnd = "UPDATE money." + table + " SET Kurs =  " + Kurs + ", SEKKURS = " + SEKKurs + " WHERE Symbol = '" +  symbol + "';";
            //Logger("Info7", sqlcmnd);

            try
            {
                using (SqlConnection connection = new SqlConnection(GlobalVars.SQLconString))
                {
                    connection.Open();
                    SqlCommand command = new SqlCommand(sqlcmnd, connection);
                    command.ExecuteNonQuery();
                }
            }
            catch (Exception ex)
            {
                Logger("ERROR7", ex.Message);
            }
        }

        public static void UpdateTotal()
        {
            string sqlcmnd = "EXEC money.update_total;";
            Logger("INFO8", sqlcmnd);

            try
            {
                using (SqlConnection connection = new SqlConnection(GlobalVars.SQLconString))
                {
                    connection.Open();
                    SqlCommand command = new SqlCommand(sqlcmnd, connection);
                    command.ExecuteNonQuery();
                }
            }
            catch (Exception ex)
            {
                Logger("ERROR8", ex.Message);
            }

        }

        public static decimal ConvertExchangeRates(string Valuta)
        {
            //https://api.exchangeratesapi.io/latest?base=USD

            string url = $"https://api.exchangeratesapi.io/latest?base={Valuta}";
            var client = new System.Net.WebClient();

            try
            {
                string responseBody = client.DownloadString(url);
                int position = responseBody.IndexOf("SEK");
                string substring = responseBody.Substring(position + 5, 20);
                int endposition = substring.IndexOf(",");
                string rate = substring.Substring(0, endposition - 1);
                return decimal.Parse(rate);

            }
            catch (Exception ex)
            {
                Logger("ERROR9", "Problem med att hämta Exchange rates " + ex.Message);
                return 0;
            }

        }


    }
}
