using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace BasedOnText
{
    public partial class ForceExit : Form
    {        
        public ForceExit()
        {
            InitializeComponent();
        }

        public void ShowText(string str)
        {
            waningMessageLabel.Text = str + "\n파일이 존재합니다. 오버로드 하시겠습니까?";
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            Common.over = false;
            this.Hide();
        }

        private void btnForceExit_Click(object sender, EventArgs e)
        {
            Common.over = true;
            this.Hide();
        }
    }
}
