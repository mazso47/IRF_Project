using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace HawaiiWeatherApp
{
    public class locationTextBox : TextBox
    {
        public bool validateLocation(string location)
        {
            return Regex.IsMatch(
                location,
                @"^[A-Z][a-zA-Z]+$");
                
        }
    }
}
