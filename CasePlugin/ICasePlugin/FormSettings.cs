using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace ICasePlugin
{
    public partial class FormSettings : Form
    {
        public FormSettings()
        {
            InitializeComponent();
        }

        private void buttonSave_Click(object sender, EventArgs e)
        {
            classService.Password = textBoxPassword.Text;
            classService.URL = textBoxURL.Text;
            classService.Username = textBoxUsername.Text;

            classService.SaveSetting();

            this.DialogResult = DialogResult.OK;
            this.Hide();
        }

        ClassService classService;
        private void FormSettings_Load(object sender, EventArgs e)
        {

            classService = new ClassService();
            textBoxPassword.Text = classService.Password;
            textBoxURL.Text = classService.URL;
            textBoxUsername.Text = classService.Username;
        }

        private void buttonCancel_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.No;
            this.Hide();

        }

        private void buttonTest_Click(object sender, EventArgs e)
        {
            XCaseService.CaseService service = new XCaseService.CaseService();

            try
            {
                service.Url = textBoxURL.Text;
                if (service.Login(textBoxUsername.Text, textBoxPassword.Text))
                {
                    MessageBox.Show("The web service is valid. You must close and open Outlook to make it affected.");
                }
                else
                {
                    MessageBox.Show("The web service is invalid");
                }

            }
            catch
            {
                MessageBox.Show("Inavlid web service information");
            }


        }
    }
}
