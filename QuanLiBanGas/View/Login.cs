using QuanLiBanGas.View;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace QuanLiBanGas
{
    public partial class Login : Form
    {
        public const int WM_NCLBUTTONDOWN = 0xA1;
        public const int HT_CAPTION = 0x2;

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        public static extern int SendMessage(IntPtr hWnd, int Msg, int wParam, int lParam);
        [System.Runtime.InteropServices.DllImport("user32.dll")]
        public static extern bool ReleaseCapture();

        public Login()
        {
            InitializeComponent();
        }
        private bool login(string user, string pass)
        {
            string query = "exec U_Login '" + txtUser.Text + "' , '" + txtPass.Text + "'";
            DataTable result = Model.Model.Instance.GetTable(query);
            return result.Rows.Count > 0;
        }
        private void Login_Load(object sender, EventArgs e)
        {
            panel1.BackColor = Color.FromArgb(30, 0, 0,0);
        }

        private void btThoat_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Bạn có muốn thoát không?","Thông báo",MessageBoxButtons.YesNo)==DialogResult.Yes)
            {
                this.Close();
            }
        }

        private void bt(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                ReleaseCapture();
                SendMessage(Handle, WM_NCLBUTTONDOWN, HT_CAPTION, 0);
            }
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked==true)
            {
                txtPass.UseSystemPasswordChar = false;
            }
            else
            {
                txtPass.UseSystemPasswordChar = true;
            }
        }

        public string username;
        private void btLogin_Click(object sender, EventArgs e)
        {
            string user = txtUser.Text;
            string pass = txtPass.Text;
            if (login(user, pass))
            {
                username = txtUser.Text;
                BinhGas binhGas = new BinhGas();
                binhGas.userName = txtUser.Text;
                binhGas.ShowDialog();
                this.Show();
                txtUser.Text="";
                txtPass.Text = "";
               
            }
            else
            {
                MessageBox.Show("Tài khoản hoặc mật khẩu không chính xác", "Thông báo");
                txtUser.Focus();
                txtPass.Text = "";
            }
        }

        private void TP_Click(object sender, EventArgs e)
        {
            TakePass take = new TakePass();
            this.Hide();
            take.ShowDialog();
            this.Show();
        }
    }
}
