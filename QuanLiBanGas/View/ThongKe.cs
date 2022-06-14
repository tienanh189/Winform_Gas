using Microsoft.Reporting.WinForms;
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
    public partial class ThongKe : Form
    {
        public string userName;
        private int index = -1;//Chỉ sô cbHTBC
        public ThongKe()
        {
            InitializeComponent();
        }

        private void ThongKe_Load(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Maximized;
            this.reportViewerTK.RefreshReport();
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
            BinhGas binhGas = new BinhGas();
            binhGas.userName = userName;
            this.Hide();
            binhGas.ShowDialog();
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
  
        private void XoaTXT()
        {
            txtTTBC.Text = "";
            txtThang.Text = "";
            txtNam.Text = "";
        }
        private void cbBaoCao_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cbBaoCao.SelectedIndex == 0) {
                index = 0;
                XoaTXT();
                lbTT.Text = "Mã NV:";      
                txtTTBC.Enabled = true;
                txtThang.Enabled = false;
                txtNam.Enabled = false;
            }
            if (cbBaoCao.SelectedIndex == 1)
            {
                XoaTXT();
                lbTT.Text = "Tên NCC:";
                index = 1;
                txtTTBC.Enabled = true;
                txtThang.Enabled =true;
                txtNam.Enabled = false;
            }


            if (cbBaoCao.SelectedIndex == 2)
            {
                XoaTXT();
                lbTT.Text = "Thông tin: ";
                index = 2;
                txtTTBC.Enabled = false;
                txtThang.Enabled = true;
                txtNam.Enabled = false;
            }

            if (cbBaoCao.SelectedIndex == 3)
            {
                XoaTXT();
                lbTT.Text = "Thông tin: ";
                index = 3;
                txtTTBC.Enabled = false;
                txtThang.Enabled = false;
                txtNam.Enabled = true;
            }

            if (cbBaoCao.SelectedIndex == 4)
            {
                XoaTXT();
                lbTT.Text = "Thông tin: ";
                index = 4;
                txtTTBC.Enabled = false;
                txtThang.Enabled = true;
                txtNam.Enabled = true;
            }
            
        }

        private void btOutput_Click(object sender, EventArgs e)
        {
            this.reportViewerTK.Clear();
            this.reportViewerTK.LocalReport.DataSources.Clear();
            if (index==0)
            { 
                reportViewerTK.LocalReport.ReportEmbeddedResource = "QuanLiBanGas.View.ReportSP.rdlc";
                ReportDataSource rds = new ReportDataSource();
                rds.Name = "DataSetSP1";
                rds.Value = Model.Model.Instance.GetTable("exec BC1 '" + txtTTBC.Text + "'");
                reportViewerTK.LocalReport.DataSources.Add(rds);
               
            }
            if (index == 1)
            {
                
                reportViewerTK.LocalReport.ReportEmbeddedResource = "QuanLiBanGas.View.ReportNCC.rdlc";
                ReportDataSource rds = new ReportDataSource();
                rds.Name = "DataSetNCC";
                rds.Value = Model.Model.Instance.GetTable("exec BC2 '" + txtTTBC.Text + "','"+txtThang.Text+"'");
                reportViewerTK.LocalReport.DataSources.Add(rds);
                
            }
            if (index == 2)
            {  
                reportViewerTK.LocalReport.ReportEmbeddedResource = "QuanLiBanGas.View.ReportDTT.rdlc";
                ReportDataSource rds = new ReportDataSource();
                rds.Name = "DataSetDTT";
                rds.Value = Model.Model.Instance.GetTable("exec BCDT_theothang '" + txtThang.Text + "'");
                reportViewerTK.LocalReport.DataSources.Add(rds);
                
            }
            if (index == 3)
            {

                reportViewerTK.LocalReport.ReportEmbeddedResource = "QuanLiBanGas.View.ReportDTN.rdlc";
                ReportDataSource rds = new ReportDataSource();
                rds.Name = "DataSetDTN";
                rds.Value = Model.Model.Instance.GetTable("exec BCDT_theonam '" + txtNam.Text + "'");
                reportViewerTK.LocalReport.DataSources.Add(rds);
            }
            if (index == 4)
            {
               
                reportViewerTK.LocalReport.ReportEmbeddedResource = "QuanLiBanGas.View.ReportKH.rdlc";
                ReportDataSource rds = new ReportDataSource();
                rds.Name = "DataSetKH";
                rds.Value = Model.Model.Instance.GetTable("exec BC4 '" + txtThang.Text + "','" + txtNam.Text + "'");
                reportViewerTK.LocalReport.DataSources.Add(rds);
               
            }
            this.reportViewerTK.RefreshReport();
        }
    }
}
