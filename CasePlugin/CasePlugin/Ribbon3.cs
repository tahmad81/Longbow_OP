using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;

namespace CasePlugin
{
    public partial class Ribbon3
    {

        private void GetCase()
        {
            XCaseService.CaseService service = new XCaseService.CaseService();
            int[] caseNumbers = service.GetCaseNumber();

            RibbonDropDownItem item;
            comboBox1.Items.Clear();
            foreach (int caseNumber in caseNumbers)
            {
                item = Factory.CreateRibbonDropDownItem();
                item.Label = caseNumber.ToString();
                comboBox1.Items.Add(item);
            }

            service.Dispose();
            service = null;
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
                Globals.ThisAddIn.intCaseNumber = Convert.ToInt32(comboBox1.Text);
            }
            else
            {
                Globals.ThisAddIn.useCaseNumber = false;
                Globals.ThisAddIn.intCaseNumber = 0;
            }
        }
    }
}
