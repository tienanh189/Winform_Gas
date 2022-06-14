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
    public partial class HoaDonBan : Form
    {
        public string userName;
        private int posDM = -1; //Chỉ số hàng bảng CTHDB
        private int pos = -1;   //Chỉ số hàng bảng HDB
        private int index = -1; //Chỉ số cbTK
        public HoaDonBan()
        {
            InitializeComponent();
        }

        private void HoaDonBan_Load(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Maximized;
            btRs.Visible = false;
            this.dgvHDB.RowTemplate.Height = 50;
            this.dgvCTHDB.RowTemplate.Height = 45;
            dgvHDB.DataSource = Model.Model.Instance.GetTable("select  * from HoaDonBan");      
            dgvCTHDB.DataSource = Model.Model.Instance.GetTable("select* from ChiTietHDB");
          
        }

        private void ResetTXT()
        {
            txtMaHDB.Text = "";
            txtMaNV.Text = "";
            txtKH.Text = "";
            txtTongTien.Text = "";
        }
        private void ResetTXTCT()
        {
            txtMaHDBCT.Text = "";
            txtMaBinh.Text = "";
            txtSLB.Text = "";
            txtGiamGia.Text = "";
            txtThanhTien.Text = "";
        }
        private void bìnhGaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            BinhGas binhGas = new BinhGas();
            binhGas.userName = userName;
            this.Hide();
            binhGas.ShowDialog();
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

        private void mntNguoiDung_Click(object sender, EventArgs e)
        {
            Account account = new Account();
            account.userName = userName;
            this.Hide();
            account.ShowDialog();
            this.Close();
        }

        private void CheckNull()
        {
            if (txtMaHDB.Text=="")
            {
                MessageBox.Show("Hãy nhập vào MAHDB");
                txtMaHDB.Focus();
            }
            if (txtMaNV.Text == "")
            {
                MessageBox.Show("Hãy nhập vào MaNV");
                txtMaNV.Focus();
            }
            if (txtKH.Text == "")
            {
                MessageBox.Show("Hãy nhập vào MaKH");
                txtKH.Focus();
            }
        }
        private void btThem_Click(object sender, EventArgs e)
        {
            try
            {
                CheckNull();
                if (txtMaHDB.Text!=""&&txtMaNV.Text!=""&&txtKH.Text!="")
                {
                    if (txtTongTien.Text=="")
                    {
                        txtTongTien.Text = "0";
                    }
                    if (MessageBox.Show("Bạn có muốn thêm không?", "Thông báo", MessageBoxButtons.YesNo) == DialogResult.Yes)
                    {
                        int res = Model.Model.Instance.GetResIUD("Insert into HoaDonBan values('" + txtMaHDB.Text + "', '" + txtMaNV.Text + "', '" + dtpNB.Text + "', '" + txtKH.Text + "', '" + txtTongTien.Text + "')");
                        dgvHDB.DataSource = Model.Model.Instance.GetTable("select  * from HoaDonBan");
                        if (res > 0)
                        {
                            MessageBox.Show("Thành công", "Thông báo");
                            ResetTXT();
                        }
                        else
                        {
                            MessageBox.Show("Thất bại", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }  
                    
                }              
            }
            catch (Exception )
            {
                MessageBox.Show("Hãy kiểm tra lại các thông tin của hóa đơn", "Thông báo");
            }
            
        }

        private void btSua_Click(object sender, EventArgs e)
        {
            try
            {
                CheckNull();
                if (pos == -1)
                {
                    MessageBox.Show("Hãy chọn hóa đơn muốn sửa");
                    return;
                }
                else
                {
                    if (txtMaHDB.Text != "" && txtMaNV.Text != "" && txtKH.Text != "")
                    {
                        if (MessageBox.Show("Bạn có muốn sửa không?", "Thông báo", MessageBoxButtons.YesNo) == DialogResult.Yes)
                        {
                            int res = Model.Model.Instance.GetResIUD("update HoaDonban set  MaNv= '" + txtMaNV.Text + "' , NgayBan= '" + dtpNB.Text + "', MaKH =  '" + txtKH.Text + "' , TongTien = '" + txtTongTien.Text + "' where SoHDB = '" + txtMaHDB.Text + "' ");
                            if (res > 0)
                            {
                                MessageBox.Show("Thành công", "Thông báo");
                                dgvHDB.DataSource = Model.Model.Instance.GetTable("select  * from HoaDonBan");
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
                MessageBox.Show("Hãy kiểm tra lại các thông tin của hóa đơn ", "Thông báo");
            }
                
        }

        private void btXoa_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Bạn có chắc chắn muốn xóa không?", "Thông báo",MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                if (pos == -1)
                {
                    MessageBox.Show("Hãy chọn hóa đơn muốn xóa");
                    return;
                }
                string query = "Delete from HoaDonBan where SoHDB = @ma";
                int result = Model.Model.Instance.GetResIUD(query, new object[] { txtMaHDB.Text });
                if (result > 0)
                {
                    MessageBox.Show("Thành công", "Thông báo", MessageBoxButtons.OK);
                    dgvHDB.DataSource = Model.Model.Instance.GetTable("select  * from HoaDonBan");
                    ResetTXT();
                }
                else
                {
                    MessageBox.Show("Thất bại", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);

                }
            }
        }
        
        private void dgvHDB_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            pos = e.RowIndex;
            try
            {
                if (pos == -1)
                {
                    return;
                }
                DataRow row = Model.Model.Instance.GetTable("select  * from HoaDonBan").Rows[pos];
                txtMaHDB.Text = row["SoHDB"].ToString();
                txtMaNV.Text = row["MaNV"].ToString();
                dtpNB.Text = row["NgayBan"].ToString();
                txtKH.Text = row["MaKH"].ToString();
                txtTongTien.Text = row["TongTien"].ToString();
            }
            catch (Exception)
            {
                MessageBox.Show("Vui lòng chọn lại", "Thông báo");
            }       
        }

        private void CheckNullCT()
        {
            if (txtMaHDBCT.Text == "")
            {
                MessageBox.Show("Hãy nhập vào MAHDB");
                txtMaHDBCT.Focus();
            }
            if (txtMaBinh.Text == "")
            {
                MessageBox.Show("Hãy nhập vào MaBinh");
                txtMaBinh.Focus();
            }
            if (txtSLB.Text == "")
            {
                MessageBox.Show("Hãy nhập vào SLBan");
                txtSLB.Focus();
            }
        }
        private void btThemCT_Click(object sender, EventArgs e)
        {
            try
            {
                CheckNullCT();
                if (txtMaHDBCT.Text!=""&&txtMaBinh.Text!=""&&txtSLB.Text!="")
                {
                    if (MessageBox.Show("Bạn có muốn thêm không?", "Thông báo", MessageBoxButtons.YesNo) == DialogResult.Yes)
                    {
                        int res = Model.Model.Instance.GetResIUD("insert into ChiTietHDB values('" + txtMaHDBCT.Text + "','" + txtMaBinh.Text + "','" + txtSLB.Text + "','" + txtGiamGia.Text + "','" + txtThanhTien.Text + "')");
                        dgvCTHDB.DataSource = Model.Model.Instance.GetTable("select * from ChiTietHDB");
                        if (res > 0)
                        {
                            MessageBox.Show("Thành công", "Thông báo");
                            dgvHDB.DataSource = Model.Model.Instance.GetTable("select  * from HoaDonBan");
                            ResetTXTCT();
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
                MessageBox.Show("Không thể thêm chi tiết HDB này hãy kiểm tra lại MaHDB,MaBinh,SoLuong ", "Thông báo");
            }
        }

        private void btSuaCT_Click(object sender, EventArgs e)
        {
            try
            {
                if (posDM == -1)
                {
                    MessageBox.Show("Hãy chọn hóa đơn muốn sửa");
                    return;
                }
                else
                {
                    CheckNullCT();
                    if (txtMaHDBCT.Text != "" && txtMaBinh.Text != "" && txtSLB.Text != "")
                    {
                        if (MessageBox.Show("Bạn có muốn sửa không?", "Thông báo", MessageBoxButtons.YesNo) == DialogResult.Yes)
                        {
                            int res = Model.Model.Instance.GetResIUD("update ChiTietHDB set   SLBan='" + txtSLB.Text + "', GiamGia='" + txtGiamGia.Text + "',ThanhTien='" + txtThanhTien.Text + "' where SoHDB='" + txtMaHDBCT.Text + "' and MaBinh='" + txtMaBinh.Text + "' ");
                            if (res > 0)
                            {
                                MessageBox.Show("Thành công", "Thông báo");
                                dgvCTHDB.DataSource = Model.Model.Instance.GetTable("select  * from ChiTietHDB");
                                dgvHDB.DataSource = Model.Model.Instance.GetTable("select  * from HoaDonBan");
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
                MessageBox.Show("Không thể sửa chi tiết HDB này hãy kiểm tra lại MaHDB,MaBinh,SoLuong ", "Thông báo");
            }
           

        }

        private void btXoaCT_Click(object sender, EventArgs e)
        {

            if (MessageBox.Show("Bạn có muốn Xóa không?", "Thông báo", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                if (posDM == -1)
                {
                    MessageBox.Show("Hãy chọn hàng hóa muốn xóa");
                    return;
                }
                string query = "Delete from ChiTietHDB where SoHDB = '"+txtMaHDBCT.Text+"' and MaBinh = '"+txtMaBinh.Text+"'";
                int result = Model.Model.Instance.GetResIUD(query);
                if (result > 0)
                {
                    MessageBox.Show("Thành công", "Thông báo", MessageBoxButtons.OK);
                    dgvCTHDB.DataSource = Model.Model.Instance.GetTable("select  * from ChiTietHDB");
                    dgvHDB.DataSource = Model.Model.Instance.GetTable("select  * from HoaDonBan");
                    ResetTXTCT();
                }
                else
                {
                    MessageBox.Show("Thất bại", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

        }
        
        private void dgvCTHDB_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            posDM = e.RowIndex;
            try
            {
                if (posDM == -1)
                {
                    return;
                }
                DataRow row = Model.Model.Instance.GetTable("select  * from ChiTietHDB").Rows[posDM];
                txtMaHDBCT.Text = row["SoHDB"].ToString();
                txtMaBinh.Text = row["MaBinh"].ToString();
                txtSLB.Text = row["SLBan"].ToString();
                txtGiamGia.Text = row["GiamGia"].ToString();
                txtThanhTien.Text = row["ThanhTien"].ToString();
            }
            catch (Exception)
            {
                MessageBox.Show("Vui lòng chọn lại", "Thông báo");
            }           
        }
        private void txtSLB_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar >= '0' && e.KeyChar <= '9')||(Convert.ToInt32(e.KeyChar) == 8)||(Convert.ToInt32(e.KeyChar) == 13))
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

        private void btSearch_Click(object sender, EventArgs e)
        {
            btRs.Visible = true;
            if (index == 0)
            {
                dgvHDB.DataSource = Model.Model.Instance.GetTable("Select * from HoaDOnBan where SoHDB = @shd", new object[] { txtTTTK.Text });
            }
            if (index == 1)
            {
                dgvHDB.DataSource = Model.Model.Instance.GetTable("Select HoaDOnBan.SoHDB,MaNV,NgayBan,MaKH,TongTien from HoaDonBan join ChiTietHDB on HoaDOnBan.SoHDB = CHiTietHDB.SoHDB where MaBinh = @ma ", new object[] { txtTTTK.Text });
            }
            if (index == 2)
            {
                dgvHDB.DataSource = Model.Model.Instance.GetTable("Select * from HoaDOnBan where MaKH = @nkh", new object[] { txtTTTK.Text });
            }
            if (index == 3)
            {
                dgvHDB.DataSource = Model.Model.Instance.GetTable("Select * from HoaDOnBan where NgayBan = @ngay", new object[] { txtTTTK.Text });
            }
        }

        private void cbHTTK_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cbHTTK.SelectedIndex==0)
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

        private void btExel_Click(object sender, EventArgs e)
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
            title.Value = "DANH SÁCH HÓA ĐƠN BÁN";


            //In dữ liệu
            exSheet.get_Range("D5:H5").Font.Bold = true;
            exSheet.get_Range("D5:H5").HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            exSheet.get_Range("D5:H5").ColumnWidth = 20;
            exSheet.get_Range("D5").Value = "Số hóa đơn bán";
            exSheet.get_Range("E5").Value = "Mã nhân viên";
            exSheet.get_Range("F5").Value = "Ngày bán";
            exSheet.get_Range("G5").Value = "Mã khách hàng";
            exSheet.get_Range("H5").Value = "Tổng tiền";

           
            for (int i = 0; i < dgvHDB.Rows.Count; i++)
            {
                for (int j = 0; j < dgvHDB.Columns.Count; j++)
                {
                    if (dgvHDB.Rows[i].Cells[j].Value != null)
                    {
                        if (j == 2)
                        {
                            string[] str = dgvHDB.Rows[i].Cells[j].Value.ToString().Split(' ');
                            exApp.Cells[i + 7, j + 4] = str[0];
                        }
                        else
                        {
                            exApp.Cells[i + 7, j + 4] = dgvHDB.Rows[i].Cells[j].Value.ToString();
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
            dgvHDB.DataSource = Model.Model.Instance.GetTable("select  * from HoaDonBan");
            btRs.Visible = false;
        }
    }
}
