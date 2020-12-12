using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace HawaiiWeatherApp
{
    public class locationTextBox
    {
        public bool validateLocation(string location)
        {
            return Regex.IsMatch(
                location,
                @"^[A-Z][a-zA-Z]+$");
                
        }
    }
}
