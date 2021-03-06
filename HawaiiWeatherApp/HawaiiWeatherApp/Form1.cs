﻿using System;
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

        Excel.Application xlApp; 
        Excel.Workbook xlWB; 
        Excel.Worksheet xlSheet;

        public Form1()
        {
            InitializeComponent();
            setDecimalSeparator();
            checkDir();
            fillLists();
            updateData();
            setTexts();
            textBoxAutoComplete();
        }

        private void setDecimalSeparator()
        {
            System.Globalization.CultureInfo customCulture = (System.Globalization.CultureInfo)System.Threading.Thread.CurrentThread.CurrentCulture.Clone();
            customCulture.NumberFormat.NumberDecimalSeparator = ".";
            System.Threading.Thread.CurrentThread.CurrentCulture = customCulture;
        }

        private void checkDir()
        {
            if (Directory.Exists(Application.StartupPath.ToString() + "\\xmlFiles"))
            {
                return;
            }
            else
            {
               Directory.CreateDirectory(Application.StartupPath.ToString() + "\\xmlFiles");
            }
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
                    "Oahu", "Bradshaw Army Air Field", "Daniel K Inouye International Airport", "Hilo", "Kahului", "Kailua-Kona", "Kaneohe", "Kekaha", "Lahaina", "Lanai City", "Lihue", "Oahu"
                }
            );
        }

        private void updateData()
        {
            clearFiles();
            clearLinks();

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

        private void clearFiles()
        {
            DirectoryInfo di = new DirectoryInfo(Application.StartupPath.ToString() + "\\xmlFiles");

            foreach (FileInfo file in di.GetFiles())
            {
                file.Delete();
            }
        }

        private void clearLinks()
        {
            links.Clear();
        }

        private void getWeatherData()
        {

            XmlDocument xml = new XmlDocument();
            string selected = "";
            for (int i = 0; i < links.Count; i++)
            {
                if (links[i].Location == locationTextBox1.Text)
                {
                    selected = links[i].fileName;
                }
            }
            xml.Load("xmlFiles\\" + selected);
            string obsTimeTmp = xml.GetElementsByTagName("observation_time")[0]?.InnerText;
            obsTimeLabel.Text = obsTimeTmp.Substring(obsTimeTmp.Length - 25, 25);
            weatherLabel.Text = xml.GetElementsByTagName("weather")[0]?.InnerText;
            tempLabel.Text = xml.GetElementsByTagName("temp_c")[0]?.InnerText + " °C";
            windLabel.Text = xml.GetElementsByTagName("wind_mph")[0]?.InnerText;
            humidityLabel.Text = xml.GetElementsByTagName("relative_humidity")[0]?.InnerText + "%";
        }

        private void getRecommendation()
        {
            if (double.Parse(tempLabel.Text.Substring(0, 2).Replace(" ", "")) > 35 || double.Parse(windLabel.Text) > 40 || double.Parse(tempLabel.Text.Substring(0, 2).Replace(" ", "")) < 0)
            {
                outdoorsLabel.Text = Outdoors.Not_recommended.ToString();
            }
            else
            {
                outdoorsLabel.Text = Outdoors.Recommended.ToString();
            }
        }

        private void CreateExcel()
        {
            try
            {                
                xlApp = new Excel.Application();
                xlWB = xlApp.Workbooks.Add(Missing.Value);
                xlSheet = xlWB.ActiveSheet;
                
                CreateTable(); 
                
                xlApp.Visible = true;
                xlApp.UserControl = true;
            }
            catch (Exception ex) 
            {
                string errMsg = string.Format("Error: {0}\nLine: {1}", ex.Message, ex.Source);
                MessageBox.Show(errMsg, "Error");

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
                "Wind (MpH)",
                "Relative Humidity (%)"
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
                
                var location = xml.GetElementsByTagName("location")[0]?.InnerText.ToString();
                var obsTime = xml.GetElementsByTagName("observation_time")[0]?.InnerText.ToString();
                var weather = xml.GetElementsByTagName("weather")[0]?.InnerText.ToString();
                var temp = xml.GetElementsByTagName("temp_c")[0]?.InnerText.ToString();
                var wind = xml.GetElementsByTagName("wind_mph")[0]?.InnerText.ToString();
                var hum = xml.GetElementsByTagName("relative_humidity")[0]?.InnerText.ToString();

                values[cnt, 0] = location;
                values[cnt, 1] = obsTime.Substring(obsTime.Length - 25, 25);
                values[cnt, 2] = weather;
                values[cnt, 3] = temp;
                values[cnt, 4] = wind;
                values[cnt, 5] = hum;

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

            Excel.Range leftRange = xlSheet.get_Range(GetCell(2, 1), GetCell(1 + values.GetLength(0), 1));
            leftRange.Font.Bold = true;
            leftRange.Interior.Color = Color.Linen;

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

        private void setTexts()
        {
            obsNameTimeLabel.Text = Resource1.obsTime;
            weatherNameLabel.Text = Resource1.weather;
            tempNameLabel.Text = Resource1.temp;
            windNameLabel.Text = Resource1.wind;
            humidityNameLabel.Text = Resource1.hum;
            outdoorsNameLabel.Text = Resource1.outdoors;

            obsTimeLabel.Text = Resource1.emptyValue;
            weatherLabel.Text = Resource1.emptyValue;
            tempLabel.Text = Resource1.emptyValue;
            windLabel.Text = Resource1.emptyValue;
            humidityLabel.Text = Resource1.emptyValue;
            outdoorsLabel.Text = Resource1.emptyValue;
            exampleLabel.Text = Resource1.exampleLabel;
            locationLabel.Text = Resource1.locationLabel;

            weatherButton.Text = Resource1.weatherButton;
            updateButton.Text = Resource1.updateButton;
            excelButton.Text = Resource1.excelButton;
            exitButton.Text = Resource1.exitButton;
            csvButton.Text = Resource1.csvButton;
            this.Text = Resource1.appName;
        }

        private void textBoxAutoComplete()
        {
            locationTextBox1.AutoCompleteMode = AutoCompleteMode.Suggest;
            locationTextBox1.AutoCompleteSource = AutoCompleteSource.CustomSource;
            AutoCompleteStringCollection autoStrings = new AutoCompleteStringCollection();
            foreach (string location in locations)
            {
                autoStrings.Add(location);
            };
            locationTextBox1.AutoCompleteCustomSource = autoStrings;

        }

        private void weatherButton_Click(object sender, EventArgs e)
        {
            if (locationTextBox1.validateLocation(locationTextBox1.Text))
            {
                getWeatherData();
                getRecommendation();
            }
            else
            {
                MessageBox.Show("Invalid location!");
            }
        }

        private void updateButton_Click(object sender, EventArgs e)
        {
            updateData();
        }

        private void excelButton_Click(object sender, EventArgs e)
        {
            CreateExcel();
        }

        private void exitButton_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void csvButton_Click(object sender, EventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();

            sfd.InitialDirectory = Application.StartupPath;
            sfd.Filter = "Comma Seperated Values (*.csv)|*.csv";
            sfd.DefaultExt = "csv"; 
            sfd.AddExtension = true;

            if (sfd.ShowDialog() != DialogResult.OK) return;
            XmlDocument xml = new XmlDocument();
            using (StreamWriter sw = new StreamWriter(sfd.FileName, false, Encoding.UTF8))
            {
                sw.Write("Location");
                sw.Write(";");
                sw.Write("Observation Time");
                sw.Write(";");
                sw.Write("Weather");
                sw.Write(";");
                sw.Write("Temperature");
                sw.Write(";");
                sw.Write("Wind");
                sw.Write(";");
                sw.Write("Relative Humidity");
                sw.Write(";");
                sw.WriteLine();

                foreach (weatherLinks link in links)
                {
                    xml.Load(Application.StartupPath.ToString() + "\\xmlFiles\\" + link.fileName);

                    sw.Write(xml.GetElementsByTagName("location")[0]?.InnerText.ToString());
                    sw.Write(";");
                    sw.Write(xml.GetElementsByTagName("observation_time")[0]?.InnerText.ToString());
                    sw.Write(";");
                    sw.Write(xml.GetElementsByTagName("weather")[0]?.InnerText.ToString());
                    sw.Write(";");
                    sw.Write(xml.GetElementsByTagName("temp_c")[0]?.InnerText.ToString());
                    sw.Write(";");
                    sw.Write(xml.GetElementsByTagName("wind_mph")[0]?.InnerText.ToString());
                    sw.Write(";");
                    sw.Write(xml.GetElementsByTagName("relative_humidity")[0]?.InnerText.ToString());
                    sw.WriteLine();
                }
            }
        }
    }
}
