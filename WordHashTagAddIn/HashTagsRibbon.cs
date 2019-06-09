using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;

namespace WordHashTagAddIn
{
    public partial class HashTagsRibbon
    {

        private void HashTagsRibbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void ShowHashTagsBtn_Click(object sender, RibbonControlEventArgs e)
        {
            if (Panes.HashTags != null)
                Panes.HashTags.Visible = !Panes.HashTags.Visible;
        }
    }
}
