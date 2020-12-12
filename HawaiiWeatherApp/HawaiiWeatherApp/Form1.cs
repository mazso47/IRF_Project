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

      

        public Form1()
        {
            InitializeComponent();
            updateData();
        }


        private void getWeatherData()
        {
            XmlDocument xml = new XmlDocument();
            string selected = "";           
            for (int i = 0; i < links.Count; i++)
            {
                if (links[i].Location == textBox1.Text)
                {
                    selected = links[i].fileName;
                }
            }
            xml.Load("xmlFiles/" + selected);
            label6.Text = xml.GetElementsByTagName("observation_time")[0].InnerText;
            label7.Text = xml.GetElementsByTagName("weather")[0].InnerText;
            label8.Text = xml.GetElementsByTagName("temp_c")[0].InnerText + "°C";
            label9.Text = xml.GetElementsByTagName("relative_humidity")[0].InnerText;
            label10.Text = xml.GetElementsByTagName("wind_mph")[0].InnerText;
            //https://stackoverflow.com/questions/897466/filter-list-object-without-using-foreach-loop-in-c2-0
        }

        
        private void updateData()
        {
            clearData();
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

        private void button1_Click(object sender, EventArgs e)
        {
            updateData();
           
        }

        private void button2_Click(object sender, EventArgs e)
        {
            getWeatherData();
        }
    }
}
