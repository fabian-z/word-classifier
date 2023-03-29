using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace word_classifier
{
    public partial class Ribbon
    {
        private void Ribbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void tlpHelp_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.ToggleTaskPane();
        }

        private void tlpWhite_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.Classify(ThisAddIn.Classification.White);
        }

        private void tlpGreen_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.Classify(ThisAddIn.Classification.Green);
        }

        private void tlpAmber_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.Classify(ThisAddIn.Classification.Amber);
        }

        private void tlpRed_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.Classify(ThisAddIn.Classification.Red);
        }
    }
}
