using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel=Microsoft.Office.Interop.Excel;

namespace QuanLiBanGas
{
    public partial class NhanVien : Form
    {
        public string userName;
        private int index = -1;     //Chỉ số cbDM
        private int pos = -1;       //Chỉ số hàng bảng NV
        private int posDM = -1;     //CHỉ số hàng bảng DM
        private int indexTK = -1;   //Chỉ số cbTK
        public NhanVien()
        {
            InitializeComponent();
        }
        private void NhanVien_Load(object sender, EventArgs e)
        {

            this.WindowState = FormWindowState.Maximized;
            btRs.Visible = false;
            this.dgvNV.RowTemplate.Height = 50;
            this.dgvDM.RowTemplate.Height = 45;
            dgvNV.DataSource = Model.Model.Instance.GetTable("select  * from NhanVien");
            dgvNV.Columns[1].Width = 150;
        }

        private void ResetTXT()
        {
            txtMaNV.Text = "";
            txtTenNV.Text = "";
            txtSex.Text = "";
            txtSDT.Text = "";
            txtDiaChi.Text = "";
            txtMaCa.Text = "";
            txtMaCV.Text = "";
           
        }
        private void ResetTXTDM()
        {
            txtTenDM.Text = "";
            txtMaDM.Text = "";
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

        private void kháchHàngToolStripMenuItem_Click(object sender, EventArgs e)
        {
            KhachHang khachHang = new KhachHang();
            khachHang.userName = userName;
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
        
        private void cbDM_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cbDM.SelectedIndex == 0)
            {
                dgvDM.DataSource = Model.Model.Instance.GetTable("Select * from CaLam");
                index = 0;
                lbMa.Text = "Mã ca";
                lbTen.Text = "Tên ca";
            }
            if (cbDM.SelectedIndex == 1)
            {
                index = 1;
                dgvDM.DataSource = Model.Model.Instance.GetTable("Select * from CongViec");
                lbMa.Text = "Mã CV";
                lbTen.Text = "Tên CV";
            }

        }

        private void dgvNV_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            pos = e.RowIndex;
            try
            {
                if (pos == -1)
                {
                    return;
                }
                DataRow row = Model.Model.Instance.GetTable("select  * from NhanVien").Rows[pos];
                txtMaNV.Text = row["MaNV"].ToString();
                txtTenNV.Text = row["TenNV"].ToString();
                dtpNS.Text = row["NgaySinh"].ToString();
                txtSex.Text = row["GioiTinh"].ToString();
                txtDiaChi.Text = row["DiaChi"].ToString();
                txtSDT.Text = row["DienThoai"].ToString();
                txtMaCa.Text = row["MaCa"].ToString();
                txtMaCV.Text = row["MaCV"].ToString();
            }
            catch (Exception)
            {
                MessageBox.Show("Vui lòng chọn lai", "Thông báo");
            }
        }

        private void btAdd_Click(object sender, EventArgs e)
        {
            try
            {
                CheckNull();
                if (txtMaNV.Text!=""&&txtTenNV.Text!=""&&txtSex.Text!=""&&txtDiaChi.Text!=""&&txtSDT.Text!=""
                    &&txtMaCV.Text!=""&&txtMaCa.Text!="")
                {
                    if (MessageBox.Show("Bạn có muốn thêm không?", "Thông báo", MessageBoxButtons.YesNo) == DialogResult.Yes)
                    {
                        int res = Model.Model.Instance.GetResIUD("insert into NhanVien values('" + txtMaNV.Text + "',N'" + txtTenNV.Text + "',N'" + txtSex.Text + "','" + dtpNS.Text + "',N'" + txtDiaChi.Text + "','" + txtSDT.Text + "','" + txtMaCa.Text + "','" + txtMaCV.Text + "')");
                        dgvNV.DataSource = Model.Model.Instance.GetTable("select * from NhanVien");
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
                MessageBox.Show("Mã nhân viên đã tồn tại hoặc không tồn tại MaCa hoặc MaCV");
            }
            
        }

        private void CheckNull()
        {
            if (txtMaNV.Text == "")
            {
                MessageBox.Show("Hãy nhập MaNV", "Thông báo");
                txtMaNV.Focus();
            }
            if (txtTenNV.Text == "")
            {
                MessageBox.Show("Hãy nhập TenNV", "Thông báo");
                txtTenNV.Focus();
            }
            if (txtSex.Text == "")
            {
                MessageBox.Show("Hãy nhập Giới tính", "Thông báo");
                txtSex.Focus();
            }
            if (txtDiaChi.Text == "")
            {
                MessageBox.Show("Hãy nhập địa chỉ", "Thông báo");
                txtDiaChi.Focus();
            }
            if (txtSDT.Text == "")
            {
                MessageBox.Show("Hãy nhập SDT", "Thông báo");
                txtSDT.Focus();
            }
            if (txtMaCa.Text == "")
            {
                MessageBox.Show("Hãy nhập MaCa", "Thông báo");
                txtMaCa.Focus();
            }
            if (txtMaCV.Text == "")
            {
                MessageBox.Show("Hãy nhập MaCV", "Thông báo");
                txtMaCV.Focus();
            }
        }

        private void btRevise_Click(object sender, EventArgs e)
        {
            try
            {
                if (pos == -1)
                {
                    MessageBox.Show("Hãy chọn nhân viên muốn sửa");
                    return;
                }
                else
                {
                    CheckNull();
                    if (txtMaNV.Text != "" && txtTenNV.Text != "" && txtSex.Text != "" && txtDiaChi.Text != "" && txtSDT.Text != ""
                       && txtMaCV.Text != "" && txtMaCa.Text != "")
                    {
                        if (MessageBox.Show("Bạn có muốn sửa không?", "Thông báo", MessageBoxButtons.YesNo) == DialogResult.Yes)
                        {
                            int res = Model.Model.Instance.GetResIUD("update NhanVien set TenNV = N'" + txtTenNV.Text + "', GioiTinh=N'" + txtSex.Text + "',NgaySinh='" + dtpNS.Text + "', DiaChi=N'" + txtDiaChi.Text + "',DienTHoai='" + txtSDT.Text + "',MaCa='" + txtMaCa.Text + "',MaCV='" + txtMaCV.Text + "' where MaNV='" + txtMaNV.Text + "'");
                            if (res > 0)
                            {
                                MessageBox.Show("Thành công", "Thông báo");
                                dgvNV.DataSource = Model.Model.Instance.GetTable("select  * from NhanVien");
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
                MessageBox.Show("Không tồn tại MaCa hoặc MaCV");
            }
            
        }

        private void btDelete_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Bạn có muốn xóa không?", "Thông báo", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                if (pos == -1)
                {
                    MessageBox.Show("Hãy chọn nhân viên muốn xóa");
                    return;
                }
                string query = "Delete from NhanVien where MaNV = @ma";
                int result = Model.Model.Instance.GetResIUD(query, new object[] { txtMaNV.Text });
                if (result > 0)
                {
                    MessageBox.Show("Thành công", "Thông báo", MessageBoxButtons.OK);
                    dgvNV.DataSource = Model.Model.Instance.GetTable("select  * from NhanVien");
                    ResetTXT();
                }
                else
                {
                    MessageBox.Show("Thất bại", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);

                }
            }
        }
        private void CheckDM()
        {
            if (txtMaDM.Text == "")
            {
                MessageBox.Show("Hãy nhập vào mã");
                txtMaDM.Focus();
            }
            if (txtTenDM.Text == "")
            {
                MessageBox.Show("Hãy nhập vào tên");
                txtMaDM.Focus();
            }
        }
        private void btAddDM_Click(object sender, EventArgs e)
        {
            int res = 0;
            try
            {
                CheckDM();
                if (txtMaDM.Text != "" && txtTenDM.Text != "")
                {
                    if (MessageBox.Show("Bạn có muốn thêm không?", "Thông báo", MessageBoxButtons.YesNo) == DialogResult.Yes)
                    {
                        if (index == 0)
                        {
                            res = Model.Model.Instance.GetResIUD("insert into CaLam values( '" + txtMaDM.Text + "' , N'" + txtTenDM.Text + "' )");
                            dgvDM.DataSource = Model.Model.Instance.GetTable("select  * from CaLam");
                        }
                        if (index == 1)
                        {
                            res = Model.Model.Instance.GetResIUD("Insert into CongViec values( '" + txtMaDM.Text + "' , N'" + txtTenDM.Text + "' )");
                            dgvDM.DataSource = Model.Model.Instance.GetTable("select  * from CongViec");
                        }
                        if (res > 0)
                        {
                            MessageBox.Show("Thành công", "Thông báo");
                            ResetTXTDM();
                        }
                        else
                        {
                            MessageBox.Show("Thất bại", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                }
            }
            catch (Exception)
            {
                MessageBox.Show("Mã đã tồn tại", "Thông báo");
            }
            
                
           
        }

        private void btReviseDM_Click(object sender, EventArgs e)
        {
            int res = 0;

            CheckDM();
            if (txtMaDM.Text != "" && txtTenDM.Text != "")
            {
                if (MessageBox.Show("Bạn có muốn sửa không?", "Thông báo", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    if (index == 0)
                    {
                        res = Model.Model.Instance.GetResIUD("update CaLam set TenCA = N'" + txtTenDM.Text + "' where MaCa = '" + txtMaDM.Text + "'");
                        dgvDM.DataSource = Model.Model.Instance.GetTable("select  * from Calam");
                    }
                    if (index == 1)
                    {
                        res = Model.Model.Instance.GetResIUD("update CongViec set TenCV = N'" + txtTenDM.Text + "' where MaCV = '" + txtMaDM.Text + "'");
                        dgvDM.DataSource = Model.Model.Instance.GetTable("select  * from COngViec");
                    }
                    if (res > 0)
                    {
                        MessageBox.Show("Thành công", "Thông báo");
                        ResetTXTDM();
                    }
                    else
                    {
                        MessageBox.Show("Thất bại", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
           
        }

        private void btDeleteDM_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Bạn có muốn xóa không?", "Thông báo", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                if (posDM == -1)
                {
                    MessageBox.Show("Hãy chọn danh mục muốn xóa");
                    return;
                }
                string query = null;
                if (index == 0)
                {
                    query = "Delete from CaLam where MaCa = @ma";
                }
                if (index == 1)
                {
                    query = "Delete from CongViec where MaCV = @ma";
                }
               

                int result = Model.Model.Instance.GetResIUD(query, new object[] { txtMaDM.Text });
                if (result > 0)
                {
                    MessageBox.Show("Thành công", "Thông báo", MessageBoxButtons.OK);
                    if (index == 0) dgvDM.DataSource = Model.Model.Instance.GetTable("select  * from Calam ");
                    if (index == 1) dgvDM.DataSource = Model.Model.Instance.GetTable("select  * from CongViec");
                    ResetTXTDM();
                }
                else
                {
                    MessageBox.Show("Thất bại", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);

                }
            }
        }

        private void cbHTTK_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cbHTTK.SelectedIndex == 0)
            {
                indexTK = 0;
            }
            if (cbHTTK.SelectedIndex == 1)
            {
                indexTK = 1;
            }
            if (cbHTTK.SelectedIndex == 2)
            {
                indexTK = 2;
            }
            if (cbHTTK.SelectedIndex == 3)
            {
                indexTK = 3;
            }
        }

        private void btSearch_Click(object sender, EventArgs e)
        {
            btRs.Visible = true;
            if (indexTK == 0)
            {
                dgvNV.DataSource = Model.Model.Instance.GetTable("select * from NhanVien where MaCa = '"+txtTTTK.Text+"' ");
            }
            if (indexTK == 1)
            {
                dgvNV.DataSource = Model.Model.Instance.GetTable("select * from NhanVien where MaCV = '" + txtTTTK.Text + "' ");
            }
            if (indexTK == 2)
            {
                dgvNV.DataSource = Model.Model.Instance.GetTable("select * from NhanVien where TenNV like '%" + txtTTTK.Text + "%' ");
            }
            if (indexTK == 3)
            {
                dgvNV.DataSource = Model.Model.Instance.GetTable("select * from NhanVien where DiaCHi like '%" + txtTTTK.Text + "%' ");
            }
        }

        private void dgvDM_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            posDM = e.RowIndex;
            try
            {
                if (posDM == -1)
                {
                    return;
                }
                if (index == 0)
                {
                    DataRow row = Model.Model.Instance.GetTable("select * from Calam ").Rows[posDM];

                    txtMaDM.Text = row["MaCa"].ToString();
                    txtTenDM.Text = row["TenCA"].ToString();
                }
                if (index == 1)
                {
                    DataRow row = Model.Model.Instance.GetTable("select * from CongViec ").Rows[posDM];

                    txtMaDM.Text = row["MaCV"].ToString();
                    txtTenDM.Text = row["TenCV"].ToString();
                }
            }
            catch (Exception)
            {
                MessageBox.Show("Vui lòng chọn lai", "Thông báo");
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
            Excel.Range title = (Excel.Range)exSheet.Cells[3, 2];
            exSheet.get_Range("B3:I3").HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            exSheet.get_Range("B3:I3").Merge(true);
            exSheet.get_Range("A1:F1").Merge(true);
            header.Font.Size = 24;
            header.Font.Bold = true;
            header.Font.Color = Color.Red;
            header.Value = "CỬA HÀNG BÁN GAS AN DƯƠNG";
            title.Font.Size = 18;
            title.Font.Bold = true;
            title.Font.Color = Color.Blue;
            title.Value = "DANH SÁCH NHÂN VIÊN";


            //In dữ liệu
            exSheet.get_Range("B5:I5").Font.Bold = true;
            exSheet.get_Range("B5:i5").HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            exSheet.get_Range("B5", "G5:I5").ColumnWidth = 20;
            exSheet.get_Range("C5", "F5").ColumnWidth = 40;
            exSheet.get_Range("D5", "E5").ColumnWidth = 20;
            exSheet.get_Range("B5").Value = "Mã nhân viên";
            exSheet.get_Range("C5").Value = "Tên nhân viên";
            exSheet.get_Range("D5").Value = "Giới tính";
            exSheet.get_Range("E5").Value = "Ngày sinh";
            exSheet.get_Range("F5").Value = "Địa chỉ";
            exSheet.get_Range("G5").Value = "Điện thoại";
            exSheet.get_Range("H5").Value = "Mã ca";
            exSheet.get_Range("I5").Value = "Mã công việc";

           
            for (int i = 0; i < dgvNV.Rows.Count; i++)
            {
                for (int j = 0; j < dgvNV.Columns.Count; j++)
                {
                    if (dgvNV.Rows[i].Cells[j].Value != null)
                    {                    
                        if (j==3)
                        {
                            string[] str = dgvNV.Rows[i].Cells[j].Value.ToString().Split(' ');
                            exApp.Cells[i + 7, j + 2] = str[0];              
                        }
                        else
                        {
                            exApp.Cells[i + 7, j + 2] = dgvNV.Rows[i].Cells[j].Value.ToString();
                        }
                       
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
            dgvNV.DataSource = Model.Model.Instance.GetTable("select  * from NhanVien");
            btRs.Visible = false;
        }
    }
}
