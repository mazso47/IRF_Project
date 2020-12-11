using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
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
            getWeatherData();
        }

        private void getWeatherData()
        {
            XmlDocument xml = new XmlDocument();
            xml.Load("xmlFiles/oahu.xml");
            label1.Text = xml.GetElementsByTagName("temp_c")[0].InnerText;
        }
    }
}
