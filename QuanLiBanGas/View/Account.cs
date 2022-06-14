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
    public partial class Account : Form
    {
        public string userName;
        private int pos = -1;
        public Account()
        {     
            InitializeComponent();
        }

        
        private void Account_Load(object sender, EventArgs e)
        {
            lbUser.Text = "Người dùng: " +userName;
            if (chUser(userName)==false)
            {
                cbHien.Hide();
                btAdd.Enabled = false;
                btDelete.Enabled = false;
                btRevise.Enabled = false;
            }
            else
            {
                cbHien.Show();
                btAdd.Enabled = true;
                btDelete.Enabled = true;
                btRevise.Enabled = true;
            }
            this.dgvAccount.RowTemplate.Height = 50;
            dgvAccount.DataSource = Model.Model.Instance.GetTable("select  * from DangNhap");
           
        }

        private bool chUser(string user)
        {
            bool res = true;
            if (userName != "admin" && userName != "Admin")
            {
                res = false;
            }
            return res;
        }
        private void bìnhGaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            BinhGas binh = new BinhGas();
            binh.userName = userName;
            this.Hide();
            binh.ShowDialog();
            this.Close();
        }

        private void nhânViênToolStripMenuItem_Click(object sender, EventArgs e)
        {
            NhanVien nhanVien = new NhanVien();
            nhanVien.userName = userName;
            this.Hide();
            nhanVien.ShowDialog();
            this.Close();
        }

        private void hóaĐơnBánToolStripMenuItem_Click(object sender, EventArgs e)
        {
            HoaDonBan hoaDonBan = new HoaDonBan();
            hoaDonBan.userName = userName;
            this.Hide();
            hoaDonBan.ShowDialog();
            this.Close();
        }

        private void hóaĐơnNhậpToolStripMenuItem_Click(object sender, EventArgs e)
        {
            HoaDonNhap hoaDonNhap = new HoaDonNhap();
            hoaDonNhap.userName = userName;
            this.Hide();
            hoaDonNhap.ShowDialog();
            this.Close();
        }

        private void kháchHàngToolStripMenuItem_Click(object sender, EventArgs e)
        {
            KhachHang khachHang = new KhachHang();
            khachHang.userName=userName;
            this.Hide();
            khachHang.ShowDialog();
            this.Close();
        }

        private void nhàCungCấpToolStripMenuItem_Click(object sender, EventArgs e)
        {
            NhaCungCap nhaCungCap = new NhaCungCap();
            nhaCungCap.userName = userName;
            this.Hide();
            nhaCungCap.ShowDialog();
            this.Close();
        }

        private void mntThongKe_Click(object sender, EventArgs e)
        {
            ThongKe thongKe = new ThongKe();
            thongKe.userName = userName;
            this.Hide();
            thongKe.ShowDialog();
            this.Close();
        }

        private void CheckNull()
        {
            if (txtUser.Text=="")
            {
                MessageBox.Show("Hãy nhập vào tài khoản");
                txtUser.Focus();
            }
            if (txtPass.Text == "")
            {
                MessageBox.Show("Hãy nhập vào mật khẩu");
                txtPass.Focus();
            }
            if (txtPhone.Text=="")
            {
                MessageBox.Show("Hãy nhập vào số điện thoại");
                txtPhone.Focus();
            }
        }
        private void btAdd_Click(object sender, EventArgs e)
        {
            try
            {
                CheckNull();
                if (txtUser.Text != "" && txtPass.Text != "")
                {
                    int res = Model.Model.Instance.GetResIUD("insert into DangNhap values('" + txtUser.Text + "','" + txtPass.Text + "','" + txtPhone.Text + "')");
                    dgvAccount.DataSource = Model.Model.Instance.GetTable("select * from DangNhap");
                    if (res > 0)
                    {
                        MessageBox.Show("Thành công", "Thông báo");
                    }
                    else
                    {
                        MessageBox.Show("Thất bại", "Thông báo");
                    }
                }
            }
            catch (Exception)
            {
                MessageBox.Show("Tên tài khoản đã tồn tại", "Thông báo");
                txtUser.Focus();
            }
            
            
        }

        private void btRevise_Click(object sender, EventArgs e)
        {
            if (pos == -1)
            {
                MessageBox.Show("Hãy chọn hóa đơn muốn sửa");
                return;
            }
            else
            {
                CheckNull();
                if (txtUser.Text != "" && txtPass.Text != "")
                {
                    int res = Model.Model.Instance.GetResIUD("update DangNhap set  MK=N'" + txtPass.Text + "', SDT='" + txtPhone.Text + "' where TK=N'" + txtUser.Text + "'  ");
                    if (res > 0)
                    {
                        MessageBox.Show("Thành công", "Thông báo");
                        dgvAccount.DataSource = Model.Model.Instance.GetTable("select  * from DangNhap");
                    }
                    else
                    {
                        MessageBox.Show("Thất bại", "Thông báo");
                    }
                }
             }
        }

        private void btDelete_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Bạn có muốn Xóa không?", "Thông báo", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                
                string query = "Delete from DangNhap where TK = @ma";
                int result = Model.Model.Instance.GetResIUD(query, new object[] { txtUser.Text });
                if (result > 0)
                {
                    MessageBox.Show("Thành công", "Thông báo", MessageBoxButtons.OK);
                    dgvAccount.DataSource = Model.Model.Instance.GetTable("select  * from DangNhap");
                }
                else
                {
                    MessageBox.Show("Thất bại", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);

                }
            }

        }
       
        private void dgvAccount_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (chUser(userName)==true)
            {
                pos = e.RowIndex;
                try
                {
                    if (pos == -1)
                    {
                        return;
                    }
                    DataRow row = Model.Model.Instance.GetTable("select  * from DangNhap").Rows[pos];
                    txtUser.Text = row["TK"].ToString();
                    txtPass.Text = row["MK"].ToString();
                    txtPhone.Text = row["SDT"].ToString();
                }
                catch (Exception)
                {
                    MessageBox.Show("Vui lòng chọn lại", "Thông báo");
                }
            }
            else
            {
                MessageBox.Show("Không có quyền", "Thông báo");
            }
            
        }

        private void dgvAccount_CellFormatting_1(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (cbHien.Checked == false)
            {

                if ((e.ColumnIndex == 1 || e.ColumnIndex == 2) && e.Value != null)
                {
                    e.Value = new String('*', e.Value.ToString().Length);
                }
            }
            
        }

        private void cbHien_CheckedChanged(object sender, EventArgs e)
        {
            if (cbHien.Checked==true)
            {
                dgvAccount.DataSource = Model.Model.Instance.GetTable("select  * from DangNhap");
            }
            else
            {
                dgvAccount.DataSource = Model.Model.Instance.GetTable("select  * from DangNhap");
            }
        }

        private void txtPhone_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar >= '0' && e.KeyChar <= '9') || (Convert.ToInt32(e.KeyChar) == 8) || (Convert.ToInt32(e.KeyChar) == 13))
            {
                e.Handled = false;
            }
            else e.Handled = true;
        }
    }
}
