using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace KIWI
{
    public partial class FormUserSimulate : Form
    {
        private FormUserInfo mFormUserInfo;

        public FormUserSimulate()
        {
            InitializeComponent();
        }

        public FormUserSimulate(FormUserInfo formUserInfo)
        {
            InitializeComponent();

            //더블 버퍼
            this.SetStyle(ControlStyles.DoubleBuffer, true);
            this.SetStyle(ControlStyles.AllPaintingInWmPaint, true);
            this.SetStyle(ControlStyles.UserPaint, true);

            mFormUserInfo = formUserInfo;
        }

    }
}
