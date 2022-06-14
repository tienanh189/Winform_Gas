using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace QuanLiBanGas
{
    public partial class KhachHang : Form
    {
        public string userName;
        private int pos = -1;   //Chỉ số hàng bảng KH
        private int index = -1; //CHỉ số cbTK
        public KhachHang()
        {
            InitializeComponent();
        }
        private void KhachHang_Load(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Maximized;
            btRs.Visible = false;
            this.dgvKH.RowTemplate.Height = 50;       
            dgvKH.DataSource = Model.Model.Instance.GetTable("select  * from KhachHang");
        }

        private void ResetTXT()
        {
            txtMaKH.Text = "";
            txtTenKH.Text = "";
            txtSDT.Text = "";
            txtDiaChi.Text = "";
        }

        private void mntNguoiDung_Click(object sender, EventArgs e)
        {
            Account account = new Account();
            account.userName = userName;
            this.Hide();
            account.ShowDialog();
            this.Close();
        }

        private void bìnhGaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            BinhGas binh = new BinhGas();
            binh.userName = userName;
            this.Hide();
            binh.ShowDialog();
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

        private void nhânViênToolStripMenuItem_Click(object sender, EventArgs e)
        {
            NhanVien nhanVien = new NhanVien();
            nhanVien.userName = userName;
            this.Hide();
            nhanVien.ShowDialog();
            this.Close();
        }

        private void CheckNull()
        {
            if (txtMaKH.Text == "")
            {
                MessageBox.Show("Hãy nhập MaKH");
                txtMaKH.Focus();
            }
            if (txtTenKH.Text == "")
            {
                MessageBox.Show("Hãy nhập TenKH");
                txtTenKH.Focus();
            }
            if (txtDiaChi.Text == "")
            {
                MessageBox.Show("Hãy nhập địa chỉ");
                txtDiaChi.Focus();
            }
            if (txtSDT.Text == "")
            {
                MessageBox.Show("Hãy nhập SDT");
                txtSDT.Focus();
            }
        }
        private void btAdd_Click(object sender, EventArgs e)
        {
            try
            {
                CheckNull();
                if (txtMaKH.Text!=""&&txtTenKH.Text!=""&&txtDiaChi.Text!=""&&txtSDT.Text!="")
                {
                    if (MessageBox.Show("Bạn có muốn thêm không?", "Thông báo", MessageBoxButtons.YesNo) == DialogResult.Yes)
                    {
                        int res = Model.Model.Instance.GetResIUD("Insert into KhachHang values('" + txtMaKH.Text + "', N'" + txtTenKH.Text + "', N'" + txtDiaChi.Text + "', '" + txtSDT.Text + "')");
                        dgvKH.DataSource = Model.Model.Instance.GetTable("select  * from KhachHang");
                        if (res > 0)
                        {
                            MessageBox.Show("Thành công", "Thông báo");
                            ResetTXT();
                        }
                        else
                        {
                            MessageBox.Show("Thất bại", "Thông báo");
                        }
                    }
                }               
            }
            catch (Exception)
            {
                MessageBox.Show("MaKH đã tồn tại", "Thông báo");
            }
            
        }

        private void btRevise_Click(object sender, EventArgs e)
        {
            try
            {
                if (pos == -1)
                {
                    MessageBox.Show("Hãy chọn khách hàng muốn sửa");
                    return;
                }
                else
                {
                    CheckNull();
                    if (txtMaKH.Text != "" && txtTenKH.Text != "" && txtDiaChi.Text != "" && txtSDT.Text != "")
                    {
                        if (MessageBox.Show("Bạn có muốn sửa không?", "Thông báo", MessageBoxButtons.YesNo) == DialogResult.Yes)
                        {
                            int res = Model.Model.Instance.GetResIUD("update KhachHang set   TenKH= N'" + txtTenKH.Text + "', DiaChi=  N'" + txtDiaChi.Text + "' , SDTKH = '" + txtSDT.Text + "' where MaKH= '" + txtMaKH.Text + "' ");
                            if (res > 0)
                            {
                                MessageBox.Show("Thành công", "Thông báo");
                                dgvKH.DataSource = Model.Model.Instance.GetTable("select  * from KhachHang");
                                ResetTXT();
                            }
                            else
                            {
                                MessageBox.Show("Thất bại", "Thông báo");
                            }
                        }
                    }
                }
            }
            catch (Exception)
            {
                MessageBox.Show("Hãy kiểm tra lại các thông tin của khách hàng", "Thông báo");
            }
        }

        private void btDelete_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Bạn có muốn xóa không?", "Thông báo", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                if (pos == -1)
                {
                    MessageBox.Show("Hãy chọn khách hàng muốn xóa");
                    return;
                }
                string query = "Delete from KhachHang where MaKH = @ma";
                int result = Model.Model.Instance.GetResIUD(query, new object[] { txtMaKH.Text });
                if (result > 0)
                {
                    MessageBox.Show("Thành công", "Thông báo", MessageBoxButtons.OK);
                    dgvKH.DataSource = Model.Model.Instance.GetTable("select  * from KhachHang");
                    ResetTXT();
                }
                else
                {
                    MessageBox.Show("Thất bại", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);

                }
            }
        }
        
        private void dgvKH_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            pos = e.RowIndex;
            try
            {
                if (pos == -1)
                {
                    return;
                }
                DataRow row = Model.Model.Instance.GetTable("select  * from KhachHang").Rows[pos];
                txtMaKH.Text = row["MaKH"].ToString();
                txtTenKH.Text = row["TenKH"].ToString();
                txtDiaChi.Text = row["DiaChi"].ToString();
                txtSDT.Text = row["SDTKH"].ToString();
            }
            catch (Exception)
            {
                MessageBox.Show("Vui lòng chọn lại", "Thông báo");
            }
            
           
        }

        private void cbHTTK_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cbHTTK.SelectedIndex == 0 )
            {
                index = 0;
            }
            if (cbHTTK.SelectedIndex == 1)
            {
                index = 1;
            }
        }

        private void btSearch_Click(object sender, EventArgs e)
        {
            btRs.Visible = true;
            if (index == 0)
            {
                dgvKH.DataSource = Model.Model.Instance.GetTable("Select * from KhachHang where TenKH like '%" + txtTTTK.Text + "%' ");
            }
            if (index == 1)
            {
                dgvKH.DataSource = Model.Model.Instance.GetTable("Select * from KhachHang where DiaChi like '%" + txtTTTK.Text + "%' ");
            }
        }

        private void txtSDT_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar >= '0' && e.KeyChar <= '9') || (Convert.ToInt32(e.KeyChar) == 8) || (Convert.ToInt32(e.KeyChar) == 13))
            {
                e.Handled = false;
            }
            else e.Handled = true;
        }

        private void btExcel_Click(object sender, EventArgs e)
        {
            Excel.Application exApp = new Excel.Application();
            Excel.Workbook exBook = exApp.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);
            Excel.Worksheet exSheet = (Excel.Worksheet)exBook.Worksheets[1];

            Excel.Range header = (Excel.Range)exSheet.Cells[1, 1];
            Excel.Range title = (Excel.Range)exSheet.Cells[3, 4];
            exSheet.get_Range("D3:G3").HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            exSheet.get_Range("D3:G3").Merge(true);
            exSheet.get_Range("A1:F1").Merge(true);
            header.Font.Size = 24;
            header.Font.Bold = true;
            header.Font.Color = Color.Red;
            header.Value = "CỬA HÀNG BÁN GAS AN DƯƠNG";
            title.Font.Size = 18;
            title.Font.Bold = true;
            title.Font.Color = Color.Blue;
            title.Value = "DANH SÁCH KHÁCH HÀNG";


            //In dữ liệu
            exSheet.get_Range("D5:G5").Font.Bold = true;
            exSheet.get_Range("D5:G5").HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            exSheet.get_Range("D5", "G5").ColumnWidth = 25;
            exSheet.get_Range("E5:F5").ColumnWidth = 35;
            exSheet.get_Range("D5").Value = "Mã khách hàng";
            exSheet.get_Range("E5").Value = "Tên khách hàng";
            exSheet.get_Range("F5").Value = "Địa chỉ";
            exSheet.get_Range("G5").Value = "Số điện thoại";

            //dgvHDB.Columns[2].DefaultCellStyle.Format = "dd/MM/yyyy";
            for (int i = 0; i < dgvKH.Rows.Count; i++)
            {
                for (int j = 0; j < dgvKH.Columns.Count; j++)
                {
                    if (dgvKH.Rows[i].Cells[j].Value != null)
                    {
                        exApp.Cells[i + 7, j + 4] = dgvKH.Rows[i].Cells[j].Value.ToString();
                    }
                }
            }

            exSheet.Name = "Sheet1";
            exBook.Activate(); //Kích hoạt file Excel
            //Thiết lập các thuộc tính của SaveFileDialog
            SaveFileDialog dlgSave = new SaveFileDialog();
            dlgSave.Filter = "Excel Document(*.xlsx)|*.xlsx |Word Document(*.doc) | *.doc | All files(*.*) | *.* ";
            dlgSave.FilterIndex = 1;
            dlgSave.AddExtension = true;
            dlgSave.DefaultExt = ".xls";
            if (dlgSave.ShowDialog() == DialogResult.OK)
            {
                exBook.SaveAs(dlgSave.FileName.ToString());//Lưu file Excel
                exApp.Quit();//Thoát khỏi ứng dụng
            }
            else MessageBox.Show("Không có danh sách để in");

        }

        private void btRs_Click(object sender, EventArgs e)
        {
            dgvKH.DataSource = Model.Model.Instance.GetTable("select  * from KhachHang");
            btRs.Visible = false;
        }
    }
}
