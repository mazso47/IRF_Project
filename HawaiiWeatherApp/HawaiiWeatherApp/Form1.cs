using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;

namespace HawaiiWeatherApp
{
    public partial class Form1 : Form
    {
        List<string> filenames = new List<string>();
        List<string> locations = new List<string>();
        List<weatherLinks> links = new List<weatherLinks>();

        Excel.Application xlApp; // A Microsoft Excel alkalmazás
        Excel.Workbook xlWB; // A létrehozott munkafüzet
        Excel.Worksheet xlSheet; // Munkalap a munkafüzeten belül

        public Form1()
        {
            InitializeComponent();
            fillLists();
            updateData();
        }

        private void fillLists()
        {
            filenames.AddRange
            (
                new List<string>
                {
                    "PHHI.xml", "PHSF.xml", "PHNL.xml", "PHTO.xml", "PHOG.xml", "PHKO.xml", "PHNG.xml", "PHBK.xml", "PHJH.xml", "PHNY.xml", "PHLI.xml", "PHJR.xml"
                }
            );

            locations.AddRange
            (
                new List<string>
                {
                    "Oahu", "Bradshaw Army Air Field", "Daniel K Inouye International Airport", "Hilo", "Kahului", "Kailua / Kona", "Kaneohe", "Kekaha", "Lahaina", "Lanai City", "Lihue", "Oahu"
                }
            );
        }

        private void getWeatherData()
        {

            //if (Directory.GetFileSystemEntries(Application.StartupPath.ToString() + "\\xmlFiles\\").Length == 0)
            //{
            //    updateData();
            //}

            XmlDocument xml = new XmlDocument();
            string selected = "";           
            for (int i = 0; i < links.Count; i++)
            {
                if (links[i].Location == textBox1.Text)
                {
                    selected = links[i].fileName;
                }
            }
            xml.Load("xmlFiles\\" + selected);
            label6.Text = xml.GetElementsByTagName("observation_time")[0].InnerText;
            label7.Text = xml.GetElementsByTagName("weather")[0].InnerText;
            label8.Text = xml.GetElementsByTagName("temp_c")[0].InnerText + "°C";
            label10.Text = xml.GetElementsByTagName("wind_mph")[0].InnerText;
            //https://stackoverflow.com/questions/897466/filter-list-object-without-using-foreach-loop-in-c2-0
        }

        
        private void updateData()
        {
            clearData();
            

            for (int i = 0; i < filenames.Count; i++)
            {
                weatherLinks l = new weatherLinks();
                l.Location = locations[i];
                l.fileName = filenames[i];
                links.Add(l);
            }
         
            WebClient webClient = new WebClient { UseDefaultCredentials = true };
            foreach (weatherLinks link in links)
            {
                webClient.Headers.Add("User-Agent: Other");
                string url = "https://w1.weather.gov/xml/current_obs/" + link.fileName;
                string localFilePath = Application.StartupPath.ToString() + "\\xmlFiles\\" + link.fileName;
                webClient.DownloadFile(url, localFilePath);
            }
        }

        private void clearData()
        {
            DirectoryInfo di = new DirectoryInfo(Application.StartupPath.ToString() + "\\xmlFiles");

            foreach (FileInfo file in di.GetFiles())
            {
                file.Delete();
            }
        }

        private void CreateExcel()
        {
            try
            {
                // Excel elindítása és az applikáció objektum betöltése
                xlApp = new Excel.Application();

                // Új munkafüzet
                xlWB = xlApp.Workbooks.Add(Missing.Value);

                // Új munkalap
                xlSheet = xlWB.ActiveSheet;

                // Tábla létrehozása
                CreateTable(); // Ennek megírása a következő feladatrészben következik

                // Control átadása a felhasználónak
                xlApp.Visible = true;
                xlApp.UserControl = true;
            }
            catch (Exception ex) // Hibakezelés a beépített hibaüzenettel
            {
                string errMsg = string.Format("Error: {0}\nLine: {1}", ex.Message, ex.Source);
                MessageBox.Show(errMsg, "Error");

                // Hiba esetén az Excel applikáció bezárása automatikusan
                xlWB.Close(false, Type.Missing, Type.Missing);
                xlApp.Quit();
                xlWB = null;
                xlApp = null;
            }
        }

        private void CreateTable()
        {
            string[] headers = new string[]
            {
                "Location",
                "Observation Time",
                "Weather",
                "Temperature (°C)",
                "Wind (MpH)"
            };
            for (int i = 0; i < headers.Length; i++)
            {
                xlSheet.Cells[1, i+1] = headers[i];
            }


            XmlDocument xml = new XmlDocument();

            object[,] values = new object[links.Count, headers.Length];

            int cnt = 0;
            foreach (weatherLinks link in links)
            {
                xml.Load(Application.StartupPath.ToString() + "\\xmlFiles\\" + link.fileName);
                values[cnt, 0] = xml.GetElementsByTagName("location")[0].InnerText;

                string obsTmp = xml.GetElementsByTagName("observation_time")[0].InnerText;
                values[cnt, 1] = obsTmp.Substring(obsTmp.Length - 25, 25);

                values[cnt, 2] = xml.GetElementsByTagName("weather")[0].InnerText;
                values[cnt, 3] = xml.GetElementsByTagName("temp_c")[0].InnerText;
                values[cnt, 4] = xml.GetElementsByTagName("wind_mph")[0].InnerText;
                cnt++;
            }

            xlSheet.get_Range(
            GetCell(2, 1),
            GetCell(1 + values.GetLength(0), values.GetLength(1))).Value2 = values;

            Excel.Range headerRange = xlSheet.get_Range(GetCell(1, 1), GetCell(1, headers.Length));
            headerRange.Font.Bold = true;
            headerRange.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            headerRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            headerRange.EntireColumn.AutoFit();
            headerRange.RowHeight = 40;
            headerRange.Interior.Color = Color.LightPink;
            headerRange.BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThick);

            Excel.Range fullRange = xlSheet.get_Range(GetCell(1, 1), GetCell(1 + values.GetLength(0), headers.Length));
            fullRange.RowHeight = 20;
            fullRange.BorderAround2(Excel.XlLineStyle.xlContinuous);
            fullRange.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;

            Excel.Range rightRange = xlSheet.get_Range(GetCell(1, 2), GetCell(1 + values.GetLength(0), headers.Length));
            rightRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

            Excel.Range leftRange = xlSheet.get_Range(GetCell(1, 1), GetCell(1 + values.GetLength(0), 1));
            leftRange.Font.Bold = true;

            Excel.Range tempRange = xlSheet.get_Range(GetCell(1, 4), GetCell(1 + values.GetLength(0), 4));
            tempRange.Font.Bold = true;

        }

        private string GetCell(int x, int y)
        {
            string ExcelCoordinate = "";
            int dividend = y;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                ExcelCoordinate = Convert.ToChar(65 + modulo).ToString() + ExcelCoordinate;
                dividend = (int)((dividend - modulo) / 26);
            }
            ExcelCoordinate += x.ToString();

            return ExcelCoordinate;
        }

        

        private void button1_Click(object sender, EventArgs e)
        {
            updateData();
           
        }

        private void button2_Click(object sender, EventArgs e)
        {
            getWeatherData();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            CreateExcel();
        }
    }
}
