using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace HawaiiWeatherApp.Classes
{
    class hawaiiButton : Button
    {
        public hawaiiButton()
        {
            this.Width = 141;
            this.Height = 25;
            this.Font = new Font("Nirmala UI", 8);
            this.BackColor = Color.PeachPuff;
            this.ForeColor = Color.SaddleBrown;
            this.FlatStyle = FlatStyle.Flat;
            this.FlatAppearance.BorderColor = Color.SaddleBrown;
            this.FlatAppearance.BorderSize = 2;
            this.TextAlign = ContentAlignment.MiddleCenter;
            this.Anchor = (AnchorStyles.Top | AnchorStyles.Right | AnchorStyles.Left);
        }
    }
}
