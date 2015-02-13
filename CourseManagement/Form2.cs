using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CourseManagementProject
{
    public partial class Splash_Screen : Form
    {
        public Splash_Screen()
        {
            InitializeComponent();
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            progressBar1.Increment(1);
            if (progressBar1.Value == 1)
            {
                label1.Text = "Start up ...";
            }
            else if (progressBar1.Value == 25)
            {
                label1.Text = "Connecting to Database ...";
            }
            else if (progressBar1.Value == 50)
            {
                label1.Text = "Done ...";
            }
            else if (progressBar1.Value == 75)
            {
                label1.Text = "Loading Component ...";
            }
            else if (progressBar1.Value == 100)
            {
                timer1.Stop();
                label1.Text = "Done ...";
            }
            this.LostFocus += Form_LostFocus;
        }
        private void Form_LostFocus(object sender, EventArgs e)
        {
            if (!this.ContainsFocus && !this.ContainsFocus)
            {
                this.Hide();
            }
        }
    }
}
