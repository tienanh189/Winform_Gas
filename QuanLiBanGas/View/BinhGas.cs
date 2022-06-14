using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

namespace QuanLiBanGas
{
    public partial class BinhGas : Form
    {
        public string userName;
        private byte[] imageRe = null;  //Mảng lưu dãy bit của ảnh
        private string imgLo = " ";     //Địa chỉ ảnh
        private int indexDM = -1;        //Chỉ số danh mục
        private int posDM = -1;         //Vị trí  hàng bảng DM
        private int pos = -1;           //Vị trí hàng bảng DMGa
        private int index = -1;          //Chỉ số combobox tìm kiếm
        public BinhGas()
        {
            InitializeComponent();
        }

        private void ResetTXT()
        {
            txtMaBinh.Text = "";
            txtTenBinh.Text = "";
            txtMaMau.Text = "";
            txtMaKL.Text = "";
            txtMaLoai.Text = "";
            txtNuocSX.Text = "";
            txtSL.Text = "";
            txtTGBH.Text = "";
            txtGhiChu.Text = "";
            txtDGN.Text = "";
            txtDGB.Text = "";
            pictrureBoxGas.Image = null;
        }
        private void ResetTXTDM()
        {
            txtTen.Text = "";
            txtMa.Text = "";
        }
        private void BinhGas_Load(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Maximized;
            btRs.Visible = false;
            this.dtgvBG.RowTemplate.Height = 60;
            this.dtgvCT.RowTemplate.Height = 45;         
            dtgvBG.DataSource = Model.Model.Instance.GetTable("select  * from DMBinhGa");
            for (int i = 0; i < dtgvBG.Columns.Count; i++)
            {
                if (dtgvBG.Columns[i] is DataGridViewImageColumn)
                    ((DataGridViewImageColumn)dtgvBG.Columns[i]).ImageLayout = DataGridViewImageCellLayout.Zoom;
            }
        }

        private void CheckNull()
        {
            if (txtMaBinh.Text == "")
            {
                MessageBox.Show("Hãy nhập mã bình", "Thông báo");
                txtMaBinh.Focus();
            }
            if (txtTenBinh.Text == "")
            {
                MessageBox.Show("Hãy nhập tên bình", "Thông báo");
                txtTenBinh.Focus();
            }
            if (txtMaLoai.Text == "")
            {
                MessageBox.Show("Hãy nhập MaLoai", "Thông báo");
                txtMaLoai.Focus();
            }
            if (txtMaMau.Text == "")
            {
                MessageBox.Show("Hãy nhập MaMau", "Thông báo");
                txtMaMau.Focus();
            }
            if (txtMaKL.Text == "")
            {
                MessageBox.Show("Hãy nhập MaKL", "Thông báo");
                txtMaKL.Focus();
            }
            if (txtNuocSX.Text == "")
            {
                MessageBox.Show("Hãy nhập MaNSX", "Thông báo");
                txtNuocSX.Focus();
            }
        }

        private void btThem_Click(object sender, EventArgs e)
        {
            try
            {
                CheckNull();
                byte[] image = null;
                if (imgLo != null)
                {   
                    FileStream fileStream = new FileStream(imgLo, FileMode.Open, FileAccess.Read);
                    BinaryReader binaryReader = new BinaryReader(fileStream);
                    image = binaryReader.ReadBytes((int)fileStream.Length);
                }

                if (image == null)
                {
                    image = imageRe;

                }
                if (txtMaBinh.Text != "" && txtMaLoai.Text != "" && txtMaMau.Text != "" && txtMaKL.Text != "" && txtNuocSX.Text != "")
                {
                    if (MessageBox.Show("Bạn có muốn thêm không?","Thông báo",MessageBoxButtons.YesNo)==DialogResult.Yes)
                    {
                        string query = " insert into DMBinhga values ( '" + txtMaBinh.Text + "' , N'" + txtTenBinh.Text + "', '" + txtMaLoai.Text + "', '" + txtMaMau.Text + "' ,'" + txtMaKL.Text + "', '" + txtNuocSX.Text + "' , '" + txtSL.Text + "' , '" + txtDGN.Text + "' , '" + txtDGB.Text + "' , '" + txtTGBH.Text + "' , @image , N'" + txtGhiChu.Text + "' )";
                        int res = Model.Model.Instance.GetResIUD(query, new object[] { image });
                        if (res > 0)
                        {
                            MessageBox.Show("Thêm thành công", "Thông báo");
                            dtgvBG.DataSource = Model.Model.Instance.GetTable("select  * from DMBinhGa");
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
                MessageBox.Show("Hãy kiểm tra lại các thông tin của bình ga", "Thông báo");  
            }
     
        }

        private void btSua_Click_1(object sender, EventArgs e)
        {
            if (pos == -1)
            {
                MessageBox.Show("Hãy chọn binh ga muốn sửa","Thông báo");
                return;
            }
            else
            {
                byte[] image = null;
                try
                {
                    CheckNull();
                    if (imgLo != null)
                    {
                        FileStream fileStream = new FileStream(imgLo, FileMode.Open, FileAccess.Read);
                        BinaryReader binaryReader = new BinaryReader(fileStream);
                        image = binaryReader.ReadBytes((int)fileStream.Length);
                    }

                    if (image == null)
                    {
                        image = imageRe;

                    }

                    string query = "update DMBinhGa set TenBinh = '" + txtTenBinh.Text + "', MaLoai = '" +
                        txtMaLoai.Text + "', MaMau = '" +
                        txtMaMau.Text + "', MaKL = '" +
                        txtMaKL.Text + "', MaNuocSX='" +
                        txtNuocSX.Text + "', SoLuong = '" +
                        txtSL.Text + "', DonGiaNhap = '" +
                        txtDGN.Text + "', DonGiaBan = '" +
                        txtDGB.Text + "', TGBaoHanh =  '" +
                        txtTGBH.Text + "' , Anh = @image , GhiChu = N'" +
                        txtGhiChu.Text + "' where  Mabinh = '" + txtMaBinh.Text + "' ";
                    if (txtMaBinh.Text != ""&& txtMaLoai.Text != ""&& txtMaMau.Text != ""&& txtMaKL.Text!= ""&&txtNuocSX.Text!="")
                    {
                        if (MessageBox.Show("Bạn có muốn sửa không?", "Thông báo", MessageBoxButtons.YesNo) == DialogResult.Yes)
                        {
                            int result = Model.Model.Instance.GetResIUD(query, new object[] { image });
                            if (result > 0)
                            {
                                MessageBox.Show("Thành công", "Thông báo", MessageBoxButtons.OK);
                                dtgvBG.DataSource = Model.Model.Instance.GetTable("select  * from DMBinhGa");
                                ResetTXT();
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
                    MessageBox.Show("Hãy kiểm tra lại các thông tin của bình ga ", "Thông báo ");
                }
                
            }
        }

        private void btXoa_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Bạn có chắc chắn muốn xóa không?", "Thông báo", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                if (pos == -1)
                {
                    MessageBox.Show("Hãy chọn bình ga muốn xóa");
                    return;
                }
                string query = "Delete from DMBinhGa where MaBinh = @ma";
                int result = Model.Model.Instance.GetResIUD(query,new object[] {txtMaBinh.Text});
                if (result > 0)
                {
                    MessageBox.Show("Thành công", "Thông báo", MessageBoxButtons.OK);
                    dtgvBG.DataSource = Model.Model.Instance.GetTable("select  * from DMBinhGa");
                    ResetTXT();
                }
                else
                {
                    MessageBox.Show("Thất bại", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);

                }
            }
        }
    
        private void cbDM_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cbDM.SelectedIndex == 0)
            {
                dtgvCT.DataSource = Model.Model.Instance.GetTable("select  * from LoaiGa");
                lbMa.Text = "Ma Loai";
                lbTen.Text = "Ten Loai";
                indexDM = 0;
            }
            if (cbDM.SelectedIndex == 1)
            {
                dtgvCT.DataSource = Model.Model.Instance.GetTable("select  * from KhoiLuong");
                lbMa.Text = "Ma KL";
                lbTen.Text = "Trong Luong";
                indexDM = 1;
            }
            if (cbDM.SelectedIndex == 2)
            {
                dtgvCT.DataSource = Model.Model.Instance.GetTable("select  * from Mau");
                lbMa.Text = "Ma mau";
                lbTen.Text = "Ten mau";
                indexDM = 2;
            }
            if (cbDM.SelectedIndex == 3)
            {
                dtgvCT.DataSource = Model.Model.Instance.GetTable("select  * from NuocSX");
                lbMa.Text = "Ma nuoc";
                lbTen.Text = "Ten nuoc";
                indexDM = 3;
            }
        }
       
        private void btSearch_Click(object sender, EventArgs e)
        {
            btRs.Visible = true;
            if (index == 0)
            {
                dtgvBG.DataSource = Model.Model.Instance.GetTable("select * from DMbinhGa join KhoiLuong on DmBinhga.MaKL = KhoiLuong.MaKL where KhoiLuong.MaKL = @tl", new object[] { txtTTTK.Text});
            }
            if (index == 1)
            {
                dtgvBG.DataSource = Model.Model.Instance.GetTable("select * from DMbinhGa join LoaiGa on DmBinhga.MaLoai = LoaiGa.MaLoai where LoaiGa.MaLoai = @tl", new object[] { txtTTTK.Text });
            }
            if (index == 2)
            {
                dtgvBG.DataSource = Model.Model.Instance.GetTable("select * from DMbinhGa where TGBaoHanh = @tl", new object[] { txtTTTK.Text });
            }
        }

        private void CheckDM()
        {
            if (txtMa.Text=="")
            {
                MessageBox.Show("Hãy nhập vào mã");
                txtMa.Focus();
            }
            if (txtTen.Text=="")
            {
                if (indexDM==1)
                {
                    MessageBox.Show("Hãy nhập vào trọng lượng");
                }
                else
                {
                    MessageBox.Show("Hãy nhập vào tên");
                }
                
                txtMa.Focus();
            }
        }
        private void btAddDM_Click(object sender, EventArgs e)
        {
            int res = 0;
            try
            {
                CheckDM();
                if (txtMa.Text != "" && txtTen.Text != "")
                {
                    if (MessageBox.Show("Bạn có muốn thêm không?", "Thông báo", MessageBoxButtons.YesNo) == DialogResult.Yes)
                    {
                        if (indexDM == 0)
                        {
                            res = Model.Model.Instance.GetResIUD("insert into LoaiGa values( '" + txtMa.Text + "' , '" + txtTen.Text + "' )");
                            dtgvCT.DataSource = Model.Model.Instance.GetTable("select  * from LoaiGa");
                        }
                        if (indexDM == 1)
                        {
                            res = Model.Model.Instance.GetResIUD("Insert into KhoiLuong values( '" + txtMa.Text + "' , '" + txtTen.Text + "' )");
                            dtgvCT.DataSource = Model.Model.Instance.GetTable("select  * from KhoiLuong");
                        }
                        if (indexDM == 2)
                        {
                            res = Model.Model.Instance.GetResIUD("Insert into Mau values( '" + txtMa.Text + "' , '" + txtTen.Text + "' )");
                            dtgvCT.DataSource = Model.Model.Instance.GetTable("select  * from Mau");
                        }
                        if (indexDM == 3)
                        {
                            res = Model.Model.Instance.GetResIUD("Insert into NuocSX values( '" + txtMa.Text + "' , '" + txtTen.Text + "')");
                            dtgvCT.DataSource = Model.Model.Instance.GetTable("select  * from NuocSX");
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
                MessageBox.Show("Mã đã tồn tại","Thông báo");
            }
   
        }

        private void btSuaDM_Click(object sender, EventArgs e)
        {
            CheckDM();
            int res = 0;
            if (MessageBox.Show("Bạn có muốn sửa không?", "Thông báo", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                if (indexDM == 0)
                {
                    res = Model.Model.Instance.GetResIUD("update LoaiGa set TenLoai = '" + txtTen.Text + "' where MaLoai = '" + txtMa.Text + "'");
                    dtgvCT.DataSource = Model.Model.Instance.GetTable("select  * from LoaiGa");
                }
                if (indexDM == 1)
                {
                    res = Model.Model.Instance.GetResIUD("update KhoiLuong set TrongLuong = '" + txtTen.Text + "' where MaKL = '" + txtMa.Text + "'");
                    dtgvCT.DataSource = Model.Model.Instance.GetTable("select  * from KhoiLuong");
                }
                if (indexDM == 2)
                {
                    res = Model.Model.Instance.GetResIUD("update Mau set TenMau = '" + txtTen.Text + "' where MaMau = '" + txtMa.Text + "'");
                    dtgvCT.DataSource = Model.Model.Instance.GetTable("select  * from Mau");
                }
                if (indexDM == 3)
                {
                    res = Model.Model.Instance.GetResIUD("update NuocSX set TenNuocSX = '" + txtTen.Text + "' where MaNuocSX = '" + txtMa.Text + "'");
                    dtgvCT.DataSource = Model.Model.Instance.GetTable("select  * from NuocSX");
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

        private void btXoaDM_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Bạn có chắc chắn muốn xóa không?", "Thông báo", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                if (posDM == -1)
                {
                    MessageBox.Show("Hãy chọn danh mục muốn xóa", "Thông báo");
                    return;
                }
                string query = null;
                if (indexDM == 0)
                {
                   query = "Delete from LoaiGa where MaLoai = @ma";
                }
                if (indexDM == 1)
                {
                    query = "Delete from KhoiLuong where MaKL = @ma";
                }
                if (indexDM == 2)
                {
                    query = "Delete from Mau where MaMau = @ma";
                }
                if (indexDM == 3)
                {
                    query = "Delete from NuocSX where MaNuocSX = @ma";
                }

                int result = Model.Model.Instance.GetResIUD(query, new object[] { txtMa.Text });
                if (result > 0)
                {
                    MessageBox.Show("Thành công", "Thông báo", MessageBoxButtons.OK);
                    if (indexDM == 0) dtgvCT.DataSource = Model.Model.Instance.GetTable("select  * from LoaiGa");
                    if (indexDM == 1) dtgvCT.DataSource = Model.Model.Instance.GetTable("select  * from KhoiLuong");
                    if (indexDM == 2) dtgvCT.DataSource = Model.Model.Instance.GetTable("select  * from Mau");
                    if (indexDM == 3) dtgvCT.DataSource = Model.Model.Instance.GetTable("select  * from NuocSX");
                    ResetTXTDM();
                }
                else
                {
                    MessageBox.Show("Thất bại", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);

                }
            }
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
            KhachHang khach = new KhachHang();
            khach.userName = userName;
            this.Hide();
            khach.ShowDialog();
            this.Close();
        }

        private void nhânViênToolStripMenuItem_Click(object sender, EventArgs e)
        {
            NhanVien nhan = new NhanVien();
            nhan.userName = userName;
            this.Hide();
            nhan.ShowDialog();
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

        private void cbHTTK_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cbHTTK.SelectedIndex ==0)
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
        }
        
        private void btImage_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Filter = "png file(*.png)|*.png|jpg flie(*.jpg)|*.jpg|All files(*.*)| *.*";
            if (dialog.ShowDialog()==DialogResult.OK)
            {
                imgLo = dialog.FileName;
                pictrureBoxGas.ImageLocation = imgLo;
            }
        }
        
        private void dtgvBG_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            imgLo = null;
            pos = e.RowIndex;
            try
            {         
                if (pos == -1)
                {
                    return;
                }

                DataRow row = Model.Model.Instance.GetTable("select * from DMBinhGa ").Rows[pos];

                txtMaBinh.Text = row["MaBinh"].ToString();
                txtTenBinh.Text = row["TenBinh"].ToString();
                txtMaLoai.Text = row["MaLoai"].ToString();
                txtMaMau.Text = row["MaMau"].ToString();
                txtMaKL.Text = row["MaKL"].ToString();
                txtNuocSX.Text = row["MaNuocSX"].ToString();
                txtSL.Text = row["SoLuong"].ToString();
                txtDGN.Text = row["DonGiaNhap"].ToString();
                txtDGB.Text = row["DonGIaBan"].ToString();
                txtTGBH.Text = row["TGBaoHanh"].ToString();
                imageRe = ((byte[])row["Anh"]);
                if (imageRe == null) { pictrureBoxGas.Image = null; }
                else
                {
                    MemoryStream memoryStream = new MemoryStream(imageRe);
                    pictrureBoxGas.Image = Image.FromStream(memoryStream);
                }
                txtGhiChu.Text = row["GhiChu"].ToString();
            }
            catch (Exception)
            {
                MessageBox.Show("Vui lòng chọn lại","Thông báo");
            }
     
        }
        
        private void dtgvCT_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            posDM = e.RowIndex;
            try
            {       
                if (posDM == -1)
                {
                    return;
                }
                if (indexDM == 0)
                {
                    DataRow row = Model.Model.Instance.GetTable("select * from Loaiga ").Rows[posDM];

                    txtMa.Text = row["MaLoai"].ToString();
                    txtTen.Text = row["TenLoai"].ToString();
                }
                if (indexDM == 1)
                {
                    DataRow row = Model.Model.Instance.GetTable("select * from KhoiLuong ").Rows[posDM];

                    txtMa.Text = row["MaKL"].ToString();
                    txtTen.Text = row["TrongLuong"].ToString();
                }
                if (indexDM == 2)
                {
                    DataRow row = Model.Model.Instance.GetTable("select * from Mau ").Rows[posDM];

                    txtMa.Text = row["MaMau"].ToString();
                    txtTen.Text = row["TenMau"].ToString();
                }
                if (indexDM == 3)
                {
                    DataRow row = Model.Model.Instance.GetTable("select * from NuocsX ").Rows[posDM];

                    txtMa.Text = row["MaNuocSX"].ToString();
                    txtTen.Text = row["TenNUOCSX"].ToString();
                }
            }
            catch (Exception)
            {
                MessageBox.Show("Vui lòng chọn lại", "Thông báo");
            }
           
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
            title.Value = "DANH SÁCH BÌNH GAS";


            //In dữ liệu
            exSheet.get_Range("A5:K5").Font.Bold = true;
            exSheet.get_Range("A5:K5").HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            exSheet.get_Range("A5:K5").ColumnWidth = 20;
            exSheet.get_Range("A5").Value = "Mã bình";
            exSheet.get_Range("B5").Value = "Tên bình";
            exSheet.get_Range("C5").Value = "Mã loại";
            exSheet.get_Range("D5").Value = "Mã màu";
            exSheet.get_Range("E5").Value = "Mã khối lượng";
            exSheet.get_Range("F5").Value = "Mã nước sản xuất";
            exSheet.get_Range("G5").Value = "Số lượng";
            exSheet.get_Range("H5").Value = "Đơn giá nhập";
            exSheet.get_Range("I5").Value = "Đơn giá bán";
            exSheet.get_Range("J5").Value = "Thời gian bảo hành";
            exSheet.get_Range("K5").Value = "Ghi chú";

            for (int i = 0; i < dtgvBG.Rows.Count; i++)
            {
                for (int j = 0; j < dtgvBG.Columns.Count; j++)
                {
                    if (j == 10) continue;
                    if (dtgvBG.Rows[i].Cells[j].Value != null)
                    {
                        exApp.Cells[i + 7, j + 1] = dtgvBG.Rows[i].Cells[j].Value.ToString();
                        if (j == 11)
                        {
                            exApp.Cells[i + 7, j] = dtgvBG.Rows[i].Cells[j].Value.ToString();
                        }
                    }
                }
                exApp.Cells[i + 7, 12] = "";
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
            
        }

        private void txtTGBH_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar >= '0' && e.KeyChar <= '9') || (Convert.ToInt32(e.KeyChar) == 8) || (Convert.ToInt32(e.KeyChar) == 13))
            {
                e.Handled = false;
            }
            else e.Handled = true;
        }

        private void btRs_Click(object sender, EventArgs e)
        {
            dtgvBG.DataSource = Model.Model.Instance.GetTable("select  * from DMBinhGa");
            btRs.Visible = false;
        }
    }
}
