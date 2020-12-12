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

namespace HawaiiWeatherApp
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            clearData();
            
        }

        private void getWeatherData()
        {
            XmlDocument xml = new XmlDocument();
            xml.Load("xmlFiles/oahu.xml");
            label1.Text = xml.GetElementsByTagName("temp_c")[0].InnerText;
        }

        private void updateData()
        {
            List<string> filenames = new List<string>();
            filenames.AddRange(new List<string>
            {
                "PHHI.xml", "PHSF.xml", "PHNL.xml", "PHTO.xml", "PHOG.xml", "PHKO.xml", "PHNG.xml", "PHBK.xml", "PHJH.xml", "PHNY.xml", "PHLI.xml", "PHJR.xml"
            });

            List<string> locations = new List<string>();
            locations.AddRange(new List<string>
            {
                "Oahu", "Bradshaw Army Air Field", "Daniel K Inouye International Airport", "Hilo", "Kahului", "Kailua / Kona", "Kaneohe", "Kekaha", "Lahaina", "Lanai City", "Lihue", "Oahu"
            });

            List<weatherLinks> links = new List<weatherLinks>();
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
    }
}
