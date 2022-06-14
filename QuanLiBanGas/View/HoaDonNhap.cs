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
    public partial class HoaDonNhap : Form
    {
        public string userName;
        private int index = -1;  //Chỉ sốcbTK
        private int pos = -1;   //Chỉ số hàng bảng HDN
        private int posCT = -1; //CHỉ số hàng bảng CTHDN

        public HoaDonNhap()
        {
            InitializeComponent();
        }

        private void HoaDonNhap_Load(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Maximized;
            btRs.Visible = false;
            this.dgvHDN.RowTemplate.Height = 50;
            this.dgvCTHDN.RowTemplate.Height = 50;
            dgvHDN.DataSource = getHDN();
            dgvCTHDN.DataSource =getCTHDN();
        }
        private DataTable getHDN()
        {
            return Model.Model.Instance.GetTable("select  * from HoaDonNhap");
        }
        private DataTable getCTHDN()
        {
            return Model.Model.Instance.GetTable("select* from ChiTietHDN");
        }
        private void ResetTXT()
        {
            txtMaHDN.Text = "";
            txtMaNV.Text = "";
            txtMaNCC.Text = "";
            txtTongTien.Text = "";
        }
        private void ResetTXTCT()
        {
            txtMaHDNCT.Text = "";
            txtMaBinh.Text = "";
            txtDGN.Text = "";
            txtSL.Text = "";
            txtGiamGia.Text = "";
            txtThanhTien.Text = "";
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
            BinhGas binh  = new BinhGas();
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

        private void nhânViênToolStripMenuItem_Click(object sender, EventArgs e)
        {
            NhanVien nhanVien = new NhanVien();
            nhanVien.userName = userName;
            this.Hide();
            nhanVien.ShowDialog();
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

        private void cbHTTK_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cbHTTK.SelectedIndex == 0)
            {
                index = 0;
            }
            if (cbHTTK.SelectedIndex == 1)
            {
                index = 1;
            }
            if (cbHTTK.SelectedIndex == 2)
            {
                index = 2;
            }
            if (cbHTTK.SelectedIndex == 3)
            {
                index = 3;
            }
        }

        private void btSearch_Click(object sender, EventArgs e)
        {
            btRs.Visible = true;
            if (index == 0)
            {
                dgvHDN.DataSource = Model.Model.Instance.GetTable("Select * from HoaDOnNhap where SoHDN = @shd", new object[] {txtTTTK.Text});
            }
            if (index == 1)
            {
                dgvHDN.DataSource = Model.Model.Instance.GetTable("Select HoaDOnNhap.SoHDN,MaNV,NgayNhap,MaNCC,TongTien from HoaDonNhap join ChiTietHDN on HoaDOnNhap.SoHDN = CHiTietHDN.SoHDN where MaBinh = @ma ", new object[] { txtTTTK.Text });
            }
            if (index == 2)
            {
                dgvHDN.DataSource = Model.Model.Instance.GetTable("Select * from HoaDOnNhap where MaNCC = @ncc", new object[] { txtTTTK.Text });
            }
            if (index == 3)
            {
                dgvHDN.DataSource = Model.Model.Instance.GetTable("Select * from HoaDOnNhap where NgayNhap = @ngay", new object[] { txtTTTK.Text });
            }
        }

        private void CheckNull()
        {
            if (txtMaHDN.Text == "")
            {
                MessageBox.Show("Hãy nhập vào MAHDN");
                txtMaHDN.Focus();
            }
            if (txtMaNV.Text == "")
            {
                MessageBox.Show("Hãy nhập vào MaNV");
                txtMaNV.Focus();
            }
            if (txtMaNCC.Text == "")
            {
                MessageBox.Show("Hãy nhập vào MaNCC");
                txtMaNCC.Focus();
            }
        }
       
        private void btThem_Click(object sender, EventArgs e)
        {
            try
            {
                CheckNull();
                if (txtMaHDN.Text!=""&&txtMaNV.Text!=""&&txtMaNCC.Text!="")
                {
                    if (MessageBox.Show("Bạn có muốn thêm không?", "Thông báo", MessageBoxButtons.YesNo) == DialogResult.Yes)
                    {
                        if (txtTongTien.Text == "")
                        {
                            txtTongTien.Text = "0";
                        }
                        int res = Model.Model.Instance.GetResIUD("insert into HoaDonNhap values('" + txtMaHDN.Text + "','" + txtMaNV.Text + "','" + dtpNN.Text + "','" + txtMaNCC.Text + "','" + txtTongTien.Text + "')");
                        dgvHDN.DataSource = Model.Model.Instance.GetTable("select  * from HoaDonNhap");
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
                MessageBox.Show("Hãy kiểm tra lại thông tin hóa đơn nhập ", "Thông báo");
            }
            

        }

        private void btSua_Click(object sender, EventArgs e)
        {
            try
            {         
                if (pos == -1)
                {
                    MessageBox.Show("Hãy chọn hóa đơn muốn sửa");
                    return;
                }
                else
                {
                    CheckNull();
                    if (txtMaHDN.Text != "" && txtMaNV.Text != "" && txtMaNCC.Text != "")
                    {
                        if (MessageBox.Show("Bạn có muốn sửa không?", "Thông báo", MessageBoxButtons.YesNo) == DialogResult.Yes)
                        {
                            int res = Model.Model.Instance.GetResIUD("update HoaDonNhap set  MaNv= '" + txtMaNV.Text + "' , NgayNhap= '" + dtpNN.Text + "', MaNCC =  '" + txtMaNCC.Text + "' , TongTien = '" + txtTongTien.Text + "' where SoHDN = '" + txtMaHDN.Text + "' ");
                            if (res > 0)
                            {
                                MessageBox.Show("Thành công", "Thông báo");
                                dgvHDN.DataSource = Model.Model.Instance.GetTable("select  * from HoaDonNhap");
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
                MessageBox.Show("Hãy kiểm tra lại thông tin hóa đơn nhập", "Thông báo");
            }
           

        }

        private void btXoa_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Bạn có chắc chắn muốn xóa không?", "Thông báo", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                if (pos == -1)
                {
                    MessageBox.Show("Hãy chọn hóa đơn muốn xóa");
                    return;
                }
                string query = "Delete from HoaDonNhap where SoHDN = @ma";
                int result = Model.Model.Instance.GetResIUD(query, new object[] { txtMaHDN.Text });
                if (result > 0)
                {
                    MessageBox.Show("Thành công", "Thông báo", MessageBoxButtons.OK);
                    dgvHDN.DataSource = Model.Model.Instance.GetTable("select  * from HoaDonNhap");
                    ResetTXT();
                }
                else
                {
                    MessageBox.Show("Thất bại", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);

                }
            }

        }

        private void CheckNullCT()
        {
            if (txtMaHDNCT.Text == "")
            {
                MessageBox.Show("Hãy nhập vào MAHDN");
                txtMaHDNCT.Focus();
            }
            if (txtMaBinh.Text == "")
            {
                MessageBox.Show("Hãy nhập vào MaBinh");
                txtMaBinh.Focus();
            }
            if (txtSL.Text == "")
            {
                MessageBox.Show("Hãy nhập vào SL");
                txtSL.Focus();
            }
            if (txtDGN.Text == "")
            {
                MessageBox.Show("Hãy nhập vào SL");
                txtDGN.Focus();
            }
        }
        private void btAddCT_Click(object sender, EventArgs e)
        {
            try
            {
                CheckNullCT();
                if (txtMaHDNCT.Text != "" && txtMaBinh.Text != "" && txtSL.Text != "" && txtDGN.Text != "")
                {
                    if (MessageBox.Show("Bạn có muốn thêm không?", "Thông báo", MessageBoxButtons.YesNo) == DialogResult.Yes)
                    {
                        int res = Model.Model.Instance.GetResIUD("insert into ChiTietHDN values('" + txtMaHDNCT.Text + "','" + txtMaBinh.Text + "','" + txtSL.Text + "','" + txtDGN.Text + "','" + txtGiamGia.Text + "','" + txtThanhTien.Text + "')");
                        dgvCTHDN.DataSource = Model.Model.Instance.GetTable("select * from ChiTietHDN");
                        if (res > 0)
                        {
                            MessageBox.Show("Thành công", "Thông báo");
                            dgvHDN.DataSource = Model.Model.Instance.GetTable("select  * from HoaDonNhap");
                            ResetTXTCT();
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
                MessageBox.Show("Không thể thêm chi tiết HDN này hãy kiểm tra lại MaHDN,MaBinh,SoLuong ", "Thông báo");
            }

        }

        private void btReviseCT_Click(object sender, EventArgs e)
        {
            try
            {
                if (posCT == -1)
                {
                    MessageBox.Show("Hãy chọn hóa đơn muốn sửa");
                    return;
                }
                else
                {
                    CheckNullCT();
                    if (txtMaHDNCT.Text != "" && txtMaBinh.Text != "" && txtSL.Text != "" && txtDGN.Text != "")
                    {
                        if (MessageBox.Show("Bạn có muốn sửa không?", "Thông báo", MessageBoxButtons.YesNo) == DialogResult.Yes)
                        {
                            int res = Model.Model.Instance.GetResIUD("update ChiTietHDN set  SLNhap='" + txtSL.Text + "', DonGia='" + txtDGN.Text + "',GiamGia='" + txtGiamGia.Text + "',ThanhTien='" + txtThanhTien.Text + "' where  SoHDN='" + txtMaHDNCT.Text + "' and MaBinh='" + txtMaBinh.Text + "'");
                            if (res > 0)
                            {
                                MessageBox.Show("Thành công", "Thông báo");
                                dgvCTHDN.DataSource = Model.Model.Instance.GetTable("select  * from ChiTietHDN");
                                dgvHDN.DataSource = Model.Model.Instance.GetTable("select  * from HoaDonNhap");
                                ResetTXTCT();
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
                MessageBox.Show("Không thể sửa chi tiết HDN này hãy kiểm tra lại MaHDN,MaBinh,SoLuong", "Thông báo");
            }


        }

        private void btDeleteCT_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Bạn có chắc chắn muốn xóa không?", "Thông báo", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                if (posCT == -1)
                {
                    MessageBox.Show("Hãy chọn hóa đơn muốn xóa");
                    return;
                }
                string query = "Delete from ChiTietHDN where  SoHDN='" + txtMaHDNCT.Text + "' and MaBinh='" + txtMaBinh.Text + "'";
                int result = Model.Model.Instance.GetResIUD(query);
                if (result > 0)
                {
                    MessageBox.Show("Thành công", "Thông báo", MessageBoxButtons.OK);
                    dgvCTHDN.DataSource = Model.Model.Instance.GetTable("select  * from ChiTietHDN");
                    dgvHDN.DataSource = Model.Model.Instance.GetTable("select  * from HoaDonNhap");
                    ResetTXTCT();
                }
                else
                {
                    MessageBox.Show("Thất bại", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);

                }
            }

        }
        
        private void dgvHDN_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            pos = e.RowIndex;
            try
            {
                if (pos == -1)
                {
                    return;
                }
                DataRow row = Model.Model.Instance.GetTable("select  * from HoaDonNhap").Rows[pos];
                txtMaHDN.Text = row["SoHDN"].ToString();
                txtMaNV.Text = row["MaNV"].ToString();
                dtpNN.Text = row["NgayNhap"].ToString();
                txtMaNCC.Text = row["MaNCC"].ToString();
                txtTongTien.Text = row["TongTien"].ToString();

            }
            catch (Exception)
            {
                MessageBox.Show("Vui lòng chọn lại", "Thông báo");
            }
           
        }
       
        private void dgvCTHDN_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            posCT = e.RowIndex;
            try
            {
                if (posCT == -1)
                {
                    return;
                }
                DataRow row = Model.Model.Instance.GetTable("select  * from ChiTietHDN").Rows[posCT];
                txtMaHDNCT.Text = row["SoHDN"].ToString();
                txtMaBinh.Text = row["MaBinh"].ToString();
                txtSL.Text = row["SLNhap"].ToString();
                txtDGN.Text = row["DonGia"].ToString();
                txtGiamGia.Text = row["GiamGia"].ToString();
                txtThanhTien.Text = row["ThanhTien"].ToString();
            }
            catch (Exception)
            {
                MessageBox.Show("Vui lòng chọn lại", "Thông báo");
            }
           

        }

        private void txtSL_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar >= '0' && e.KeyChar <= '9') || (Convert.ToInt32(e.KeyChar) == 8) || (Convert.ToInt32(e.KeyChar) == 13))
            {
                e.Handled = false;
            }
            else e.Handled = true;
        }

        private void txtGiamGia_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar >= '0' && e.KeyChar <= '9') || (Convert.ToInt32(e.KeyChar) == 8) || (Convert.ToInt32(e.KeyChar) == 13))
            {
                e.Handled = false;
            }
            else e.Handled = true;
        }

        private void txtĐGN_KeyPress(object sender, KeyPressEventArgs e)
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
            exSheet.get_Range("D3:H3").HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            exSheet.get_Range("D3:H3").Merge(true);
            exSheet.get_Range("A1:F1").Merge(true);
            header.Font.Size = 24;
            header.Font.Bold = true;
            header.Font.Color = Color.Red;
            header.Value = "CỬA HÀNG BÁN GAS AN DƯƠNG";
            title.Font.Size = 18;
            title.Font.Bold = true;
            title.Font.Color = Color.Blue;
            title.Value = "DANH SÁCH HÓA ĐƠN NHẬP";


            //In dữ liệu
            exSheet.get_Range("D5:H5").Font.Bold = true;
            exSheet.get_Range("D5:H5").HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            exSheet.get_Range("D5:H5").ColumnWidth = 20;
            exSheet.get_Range("D5").Value = "Số hóa đơn nhập";
            exSheet.get_Range("E5").Value = "Mã nhân viên";
            exSheet.get_Range("F5").Value = "Ngày nhập";
            exSheet.get_Range("G5").Value = "Mã nhà cung cấp";
            exSheet.get_Range("H5").Value = "Tổng tiền";

           
            for (int i = 0; i < dgvHDN.Rows.Count; i++)
            {
                for (int j = 0; j < dgvHDN.Columns.Count; j++)
                {
                    if (dgvHDN.Rows[i].Cells[j].Value != null)
                    {
                        if (j == 2)
                        {
                            string[] str = dgvHDN.Rows[i].Cells[j].Value.ToString().Split(' ');
                            exApp.Cells[i + 7, j + 4] = str[0];
                        }
                        else
                        {
                            exApp.Cells[i + 7, j + 4] = dgvHDN.Rows[i].Cells[j].Value.ToString();
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
            dgvHDN.DataSource = getHDN();
            btRs.Visible = false;
        }
    }

}
