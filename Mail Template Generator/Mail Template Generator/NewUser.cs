using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Mail_Template_Generator
{
    public partial class NewUser : Form
    {
        private Form1 Sender;

        public NewUser(Form1 form1)
        {
            InitializeComponent();
            Sender = form1;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Sender.add_To_CSV(textBox1.Text, textBox2.Text);
            this.Close();
        }
    }
}
