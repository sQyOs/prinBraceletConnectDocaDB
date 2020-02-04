using FirebirdSql.Data.FirebirdClient;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Common;
using System.Drawing;
using System.Drawing.Text;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Configuration;
using System.Runtime.InteropServices;
using System.Net.NetworkInformation;
using System.Threading;
using System.Drawing.Printing;
using System.Windows.Forms;

namespace prinBraceletConnectDocaDB
{
    class Program
    {
        private static string barCode;
        private static string strFio;
        private static string strFirst;
        private static string strIb;
        private static string strDepart;
        private static string idHsp;
        private static string strPhone;
        private static string strDateIn;
        private static string strFam;
        private static string strNam;
        private static string strOts;
        private static string strDbir;
        private static string printerForBracelet;
        private static string printerForLab;
        private static string docaAddr;
        private static string strSex;
        private static PrivateFontCollection pfc = new PrivateFontCollection();
        private static Image logoD = Properties.Resources.logo_kkb2;
        private static StreamWriter streamForLog = new StreamWriter("log.txt", append: true);
        private static DateTime dateTimeLastBar = DateTime.Now.AddMinutes(-1);
        private static string lastBarCode;
        public static DateTime dateTime = new DateTime(1970, 01, 01, 10, 0, 0);

        static void Main(string[] args)
        {
            streamForLog.AutoFlush = true;
            Console.SetError(streamForLog);
            try
            {
                byte[] fontdata = Properties.Resources.Verdana_Bold;
                IntPtr data = Marshal.AllocCoTaskMem(fontdata.Length);
                Marshal.Copy(fontdata, 0, data, fontdata.Length);
                pfc.AddMemoryFont(data, fontdata.Length);
            }
            catch (Exception e)
            {
                wrLine(e.Message);
            }
            int delayRepeat = Convert.ToInt32(ConfigurationManager.AppSettings.Get("delayRepeat"));
            printerForBracelet = ConfigurationManager.AppSettings.Get("printerForBracelet");
            printerForLab = ConfigurationManager.AppSettings.Get("printerForLab");
            docaAddr = ConfigurationManager.AppSettings.Get("docaAddr");
            bool labEnable = false;
            bool braceletEnable = false;
            if (ConfigurationManager.AppSettings.Get("labEnable").Equals("1"))
            {
                labEnable = true;
            }
            if (ConfigurationManager.AppSettings.Get("braceletEnable").Equals("1"))
            {
                braceletEnable = true;
            }
            bool isAvailable = false;


            FbConnectionStringBuilder cs = new FbConnectionStringBuilder();

            cs.DataSource = docaAddr;
            cs.Database = "/opt/firebird/database/docaplus.gdb";
            cs.UserID = "SYSDBA";
            cs.Password = "masterkey";
            cs.Charset = "UTF8";
            cs.Pooling = false;

            FbConnection connection = new FbConnection(cs.ToString());

            int checkCount = 0;

        restartApp:
            if (!isAvailable)
            {
                isAvailable = availableCheck();
                checkCount++;
                if (checkCount > 3)
                {
                    Thread.Sleep(1000);
                }
            }

            if (isAvailable)
            {
            zapros:
                wrLine("barcode:");
                barCode = Console.ReadLine();
                wrLine(barCode);
                if (barCode.Length == 20)
                {
                    checkCount = 0;
                anotherTry:
                    isAvailable = availableCheck();
                    string idPat = barCode.Substring(2, 9);
                    idHsp = barCode.Substring(11, 9);
                    if (isAvailable)
                    {
                        try //connect to db
                        {
                            string query = $@"select 
                                    pat_list.fam,
                                    pat_list.nam,
                                    pat_list.ots,
                                    depart.name,
                                    depart.phone,
                                    hosp.nom,
                                    dep_hsp.date_in,
                                    pat_list.id_pat,
                                    pat_list.d_bir,
                                    pat_list.sex
                                from dep_hsp
                                   inner join pat_list on(dep_hsp.id_pat = pat_list.id_pat)
                                   inner join depart on(dep_hsp.id_dep = depart.id)
                                   inner join hosp on(dep_hsp.id_pat = hosp.id_pat) and (dep_hsp.id_hsp = hosp.id_hsp)
                                where
                                   (
                                      (dep_hsp.id_hsp = {idHsp})
                                   )";
                            connection.Open();
                            wrLine(connection.State.ToString());

                            FbCommand fbCommand = new FbCommand(query, connection);
                            FbDataAdapter fbDataAdapter = new FbDataAdapter(fbCommand);
                            DataTable dataTable = new DataTable();
                            fbDataAdapter.Fill(dataTable);
                            connection.Close();
                            wrLine(connection.State.ToString());
                            foreach (DataRow dataRow in dataTable.Rows)
                            {
                                foreach (DataColumn column in dataTable.Columns)
                                {
                                    Console.WriteLine(column.ColumnName + " " + dataRow[column]);
                                }
                            }
                            if (dataTable.Rows.Count > 0)
                            {
                                strFam = spaceKiller(dataTable.Rows[dataTable.Rows.Count - 1].ItemArray[0].ToString());
                                strNam = spaceKiller(dataTable.Rows[dataTable.Rows.Count - 1].ItemArray[1].ToString());
                                strOts = spaceKiller(dataTable.Rows[dataTable.Rows.Count - 1].ItemArray[2].ToString());
                                strDepart = spaceKiller(dataTable.Rows[dataTable.Rows.Count - 1].ItemArray[3].ToString());
                                strPhone = spaceKiller(dataTable.Rows[dataTable.Rows.Count - 1].ItemArray[4].ToString());
                                strIb = spaceKiller(dataTable.Rows[dataTable.Rows.Count - 1].ItemArray[5].ToString());
                                DateTime dateIn = dateTime.AddSeconds(Convert.ToInt32(spaceKiller(dataTable.Rows[dataTable.Rows.Count - 1].ItemArray[6].ToString())));
                                DateTime dBir = dateTime.AddSeconds(Convert.ToInt32(spaceKiller(dataTable.Rows[dataTable.Rows.Count - 1].ItemArray[8].ToString())) + 3600);
                                strSex = dataTable.Rows[dataTable.Rows.Count - 1].ItemArray[9].ToString();

                                strDateIn = dateIn.ToString("dd.MM.yyyy") + " " + dateIn.ToString("HH:mm");
                                double age = Math.Truncate(DateTime.Now.Subtract(dBir).TotalDays / 365.25);
                                string ageSuffix = "лет";
                                if (!(age < 21 && age > 4))
                                {
                                    if (age % 10 == 1)
                                    {
                                        ageSuffix = "год";
                                    }
                                    else if (age % 10 > 1 && age % 10 < 5)
                                    {
                                        ageSuffix = "года";
                                    }
                                }
                                strDbir = $"{dBir.ToString("dd.MM.yyyy")} ({age} {ageSuffix})";
                            }
                            else
                            {
                                wrLine($"udefined pacient");
                                goto zapros;
                            }
                        }
                        catch (Exception e)
                        {
                            wrLine(e.Message);
                            connection.Close();
                        }
                        wrLine("start parse");
                        strFio = $"{strFam}{strNam.Substring(0, 1)}.{strOts.Substring(0, 1)}.";
                        wrLine(strFio);
                        if (strFio.Length > 3)
                        {
                            if (DateTime.Compare(DateTime.Now, dateTimeLastBar.AddSeconds(delayRepeat)) > 0 || lastBarCode != barCode)
                            {
                                lastBarCode = barCode;
                                dateTimeLastBar = DateTime.Now;
                                //опредление на неопознанного пациента
                                if (strFio.ToLower().IndexOf("неизвестный") >= 0)
                                {
                                    strFirst = $"Неизвестный, поступил: {strDateIn} ({strSex})";
                                    strFio = $"Неизвестный, {strDateIn} ({strSex})";
                                }
                                else
                                {
                                    strFirst = $"{strFio} {strDbir}";
                                }
                                wrLine($"idPat: {idPat} |idHsp: {idHsp} | {strFirst}, {strIb} {strDepart}, tel: {strPhone}");

                                wrLine("end parse");

                                wrLine("start generating & printing");
                                //generating and printing
                                int printOrPreview = Convert.ToInt32(ConfigurationManager.AppSettings.Get("printOrPreview"));
                                PrintDocument printBracelet = new PrintDocument();
                                printBracelet.DefaultPageSettings.Landscape = true;
                                printBracelet.DefaultPageSettings.PaperSize = new PaperSize("new size", 95, 700);
                                printBracelet.PrinterSettings.PrinterName = printerForBracelet;
                                printBracelet.DefaultPageSettings.PrinterResolution.Kind = PrinterResolutionKind.High;



                                printBracelet.PrintPage += PrintPageBracelet;

                                if (braceletEnable)
                                {
                                    if (printOrPreview == 2)
                                    {
                                        PrintPreviewDialog ppvw = new PrintPreviewDialog();
                                        ppvw.Document = printBracelet;
                                        ppvw.ShowDialog();
                                    }
                                    if (printOrPreview == 1)
                                    {
                                        printBracelet.Print();
                                    }
                                }
                                else
                                {
                                    wrLine("app.config disabled print bracelet");
                                }


                                PrintDocument printLab = new PrintDocument();
                                printLab.PrinterSettings.PrinterName = printerForLab;
                                printLab.PrintPage += PrintPageLab;
                                if (labEnable)
                                {
                                    if (printOrPreview == 2)
                                    {
                                        PrintPreviewDialog ppvw = new PrintPreviewDialog();
                                        ppvw.Document = printLab;
                                        ppvw.ShowDialog();
                                    }
                                    if (printOrPreview == 1)
                                    {
                                        printLab.Print();
                                    }
                                }
                                else
                                {
                                    wrLine("app.config disabled print lab list");
                                }
                                wrLine("end generating & printing");
                                strFirst = null;
                                strIb = null;
                                strDepart = null;
                                idHsp = null;
                                strFio = null;
                                strPhone = null;
                                strDateIn = null;
                            }
                            else
                            {
                                wrLine($"this barCode:\"{barCode}\" repeated {DateTime.Now.Subtract(dateTimeLastBar).ToString()} before");
                            }
                            goto zapros;
                        }
                    }
                    else
                    {
                        isAvailable = false;
                        wrLine("connection check...");
                        checkCount++;
                        if (checkCount < 5)
                        {
                            Thread.Sleep(500);
                            goto anotherTry;
                        }
                    }

                }
                else goto zapros;



                if (barCode != "exit")
                {
                    goto zapros;
                }
            }
            else goto restartApp;
            Console.ReadLine();

            connection.Close();
            Console.WriteLine(connection.State);

            Console.Read();
        }
        private static string spaceKiller(string s)
        {
            Regex regex = new Regex(@"\s+");
            return regex.Replace(s, " ");
        }
        private static bool availableCheck()
        {
            try
            {
                Ping pingSender = new Ping();
                string data = "ismyserverpingable";
                byte[] buffer = Encoding.ASCII.GetBytes(data);
                int timeout = 5;
                PingReply reply = pingSender.Send(docaAddr, timeout, buffer);
                if (reply.Status == IPStatus.Success)
                {
                    wrLine(docaAddr + " is available");
                    return true;
                }
                else
                {
                    wrLine(docaAddr + " no connection");
                    return false;
                }
            }
            catch (Exception e)
            {
                wrLine(e.Message);
                return false;
            }
        }
        private static void wrLine(string s)
        {
            Console.WriteLine(DateTime.Now + "| " + s);
            streamForLog.WriteLine(DateTime.Now + "| " + s);
        }
        private static void PrintPageBracelet(object sender, PrintPageEventArgs e)
        {

            //pfc.AddFontFile(@"Resources\IDAutomationSC128L SymbolEncoded.ttf");
            //pfc.AddFontFile(@"Resources\Verdana-Bold.ttf");
            //pfc.AddFontFile(Properties.Resources.Verdana_Bold)
            //Image logoD = Image.FromFile(@"Resources\logo_kkb2.png");

            e.Graphics.DrawImage(logoD, 0, 4, 28, 88);

            e.Graphics.DrawString(strFirst, new Font(pfc.Families[0], 11, FontStyle.Bold), Brushes.Black, 32, 5);
            e.Graphics.DrawString("ИБ:" + strIb + " " + strDepart + " Тел:" + strPhone, new Font(pfc.Families[0], 11, FontStyle.Bold), Brushes.Black, 32, 74);
            //generate and add barcode
            iTextSharp.text.pdf.Barcode128 code128 = new iTextSharp.text.pdf.Barcode128();
            code128.CodeType = iTextSharp.text.pdf.Barcode.CODE128;
            code128.ChecksumText = true;
            code128.GenerateChecksum = true;
            code128.Code = barCode;
            code128.BarHeight = 40;
            System.Drawing.Bitmap bm = new System.Drawing.Bitmap(code128.CreateDrawingImage(System.Drawing.Color.Black, System.Drawing.Color.White));
            //bm.SetResolution(500, 500);
            e.Graphics.DrawImage(bm, new PointF(40, 28));
        }
        private static void PrintPageLab(object sender, PrintPageEventArgs e)
        {
            //PrivateFontCollection pfc = new PrivateFontCollection();
            //pfc.AddFontFile(@"Resources\IDAutomationSC128L SymbolEncoded.ttf");
            //pfc.AddFontFile(@"Resources\Verdana-Bold.ttf");
            //generate and add barcode
            iTextSharp.text.pdf.Barcode128 code128 = new iTextSharp.text.pdf.Barcode128();
            code128.CodeType = iTextSharp.text.pdf.Barcode.CODE128;
            code128.ChecksumText = true;
            code128.GenerateChecksum = true;
            code128.Code = "00" + idHsp;
            code128.BarHeight = 35;
            System.Drawing.Bitmap bm = new System.Drawing.Bitmap(code128.CreateDrawingImage(System.Drawing.Color.Black, System.Drawing.Color.White));

            float rxStart = (float)(8 * 3.9 - 3 * 3.9);
            float ryStart = (float)(21 * 3.9 - 3 * 3.9);
            float rWidth = (float)(48.5 * 3.9);
            float rHeight = (float)(25 * 3.9);
            //float widthStrFio = e.Graphics.MeasureString(strFio, new Font(pfc.Families[1], 7, FontStyle.Bold)).Width;
            //float widthstrIb = e.Graphics.MeasureString(strIb, new Font(pfc.Families[1], 7, FontStyle.Bold)).Width;
            //float widthstrOtd = e.Graphics.MeasureString(strOtd, new Font(pfc.Families[1], 7, FontStyle.Bold)).Width;
            for (int i = 0; i < 4; i++)
            {
                for (int j = 0; j < 5; j++)
                {
                    //e.Graphics.DrawRectangle(Pens.Black, rxStart + rWidth * i, ryStart + rHeight * j, rWidth, rHeight);
                    StringFormat strFormat = new StringFormat();
                    strFormat.Alignment = StringAlignment.Center;
                    strFormat.LineAlignment = StringAlignment.Center;
                    e.Graphics.DrawImage(bm, new PointF(rxStart + rWidth * i + rWidth / 2 - bm.Width / 2, ryStart + rHeight * j + 5));
                    e.Graphics.DrawString("00" + idHsp + "\n" + strFio + " | " + strIb + " | " + strDepart, new Font(pfc.Families[0], 8, FontStyle.Bold), Brushes.Black, Rectangle.Round(new RectangleF(rxStart + rWidth * i, ryStart + rHeight * j + bm.Height, rWidth, rHeight - bm.Height)), strFormat);
                }
            }
        }
    }
}
