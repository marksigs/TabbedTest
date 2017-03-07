using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Configuration;
using MSC.IntegrationService.ConfigService;

namespace ConfigTestHarness
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            ConfigManager cm = new ConfigManager();
            MessageConfig msg = cm.GetMessageConfig("AIG");
            label1.Text = msg.Description;
            
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        
    }
}