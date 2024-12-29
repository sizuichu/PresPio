using System;
using System.Windows.Forms;

namespace PresPio
    {
    public partial class Con_TextTools : UserControl
        {
        public Con_TextTools()
            {
            InitializeComponent();
            }

        private void elementHost1_VisibleChanged(object sender, EventArgs e)
            {
            MyRibbon RB = Globals.Ribbons.Ribbon1;
            RB.splitButton15.Enabled = true;
            }
        }
    }