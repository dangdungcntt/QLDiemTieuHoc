using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using QuanLyDiemTieuHoc.Controller;
namespace QuanLyDiemTieuHoc.View
{
    public partial class LoginForm : Form
    {
        private Controller_User cUser = new Controller_User();
        public LoginForm()
        {
            InitializeComponent();
        }

        private void BtnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void BtnMinimize_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
        }

        private void BtnDangNhap_Click(object sender, EventArgs e)
        {
            String TaiKhoan = txtTaiKhoan.Text;
            String MatKhau = txtMatKhau.Text;
            if (TaiKhoan == "" && MatKhau == "")
                lblThongBao.Text = "Tài khoản, mật khẩu không được bỏ trống.";
            else if (TaiKhoan == "")
                lblThongBao.Text = "Tài khoản không được bỏ trống.";
            else if (MatKhau == "")
                lblThongBao.Text = "Mật khẩu không được bỏ trống.";
            else if (cUser.CheckUserNameAndPassword(TaiKhoan, MatKhau))
            {
                this.Hide();
                MainForm mainForm = new MainForm(TaiKhoan);
                mainForm.ShowDialog();
                this.Close();
            }
            else
                lblThongBao.Text = "Tài khoản hoặc mật khẩu không chính xác";
        }

        private void TxtMatKhau_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                BtnDangNhap_Click(sender, e);
            }
        }
    }
}
