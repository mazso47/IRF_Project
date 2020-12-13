using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace HawaiiWeatherApp
{
    public class locationTextBox : TextBox
    {
        public locationTextBox()
        {
            this.ForeColor = Color.SaddleBrown;
            this.Font = new Font("Nirmala UI", 8);
        }
        public bool validateLocation(string location)
        {
            return Regex.IsMatch(
                location,
                @"^[A-Z][a-zA-Z\s\-]+$");
        }
    }
}
