using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace QuanLiBanGas.View
{
    public partial class TakePass : Form
    {
        public TakePass()
        {
            InitializeComponent();
        }

        private void btXN_Click(object sender, EventArgs e)
        {
            string pass = " ";
            string query = "exec U_TP '"+txtUser.Text+"' , '"+txtPhone.Text+"' ";
            object res = Model.Model.Instance.GetScalar(query);
            if (res!=null)
            {
                MessageBox.Show("Mật khẩu của bạn là: " + res.ToString(), "Thông báo");
            }
            else
            {
                MessageBox.Show("Tài khoản và số điện thoại không đúng!!!", "Thông báo");
            }
        }

        private void btThoat_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Bạn có muốn thoát không?", "Thông báo", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                this.Close();
            }
        }
    }
}
