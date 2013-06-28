using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;

namespace ICasePlugin
{
    public partial class Ribbon3
    {
        ClassService classService = new ClassService();
        private void GetCase()
        {
            try
            {
                XCaseService.CaseService service = new XCaseService.CaseService();
                service.Url = classService.URL;

                string[] caseNumbers = service.GetCaseNumber();

                RibbonDropDownItem item;
                comboBox1.Items.Clear();
                foreach (string caseNumber in caseNumbers)
                {
                    item = Factory.CreateRibbonDropDownItem();
                    item.Label = caseNumber.ToString();
                    comboBox1.Items.Add(item);
                }

                service.Dispose();
                service = null;
            }
            catch
            {
            }
        }

        private void Ribbon3_Load(object sender, RibbonUIEventArgs e)
        {
            this.GetCase();

        }

        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            using (FormSettings formSettings = new FormSettings())
            {
                formSettings.ShowDialog();
            }
        }

        private void comboBox1_TextChanged(object sender, RibbonControlEventArgs e)
        {
            if (comboBox1.Text != "")
            {
                Globals.ThisAddIn.useCaseNumber = true;
                if (comboBox1.Text.Contains(":"))
                {
                    Globals.ThisAddIn.intCaseNumber = comboBox1.Text.Substring(comboBox1.Text.IndexOf(":") + 2);
                }
                else
                {
                    Globals.ThisAddIn.intCaseNumber = comboBox1.Text;
                }
            }
            else
            {
                Globals.ThisAddIn.useCaseNumber = false;
                Globals.ThisAddIn.intCaseNumber = "";
            }
        }

    }
}
