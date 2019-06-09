using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Forms.Integration;

namespace WordHashTagAddIn
{
    public partial class HashTagsForm : UserControl
    {
        public HashTagsForm(IHashTagsViewModel vm)
        {
            
            InitializeComponent();
            this.Controls.Add(new ElementHost(){Child = new HashTagsControl(vm), Dock = DockStyle.Fill});
        }
    }
}
