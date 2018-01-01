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
using QuanLyDiemTieuHoc.Model;
using QuanLyDiemTieuHoc.View;
using System.Windows.Forms.DataVisualization.Charting;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Threading;

namespace QuanLyDiemTieuHoc
{
    public partial class MainForm : Form
    {
        private Controller_Layout cLayout;
        private Controller_User cUser;
        private Database db = new Database();
        private Panel currenPanelActive;
        private Button currentButtonActive;
        private bool mouseDown;
        private Point lastLocation;
        private String Username;
        /* code cho chỉnh chu
         * viết nhớ cmt
         * đặt tên cho rõ ràng
         * những đoạn code khác nhiệm vụ nên cách nhau 1 dòng
         */
        public MainForm(String uname)
        {
            InitializeComponent();
            Username = uname;
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            // cấp phát đối tượng
            cLayout = new Controller_Layout();
            cUser = new Controller_User();
            //cLayout.SinhDiemNgauNhien();

            // lấy dữ liệu
            String TenHienThi = cUser.GetTenHienThi(Username); // lấy tên hiển thị
            DataTable dtHocKy = cLayout.GetDanhSachHocKy(); // lấy danh sách học kỳ
            DataTable dtKhoi = cLayout.GetDanhSachKhoi(); // lấy danh sách khối

            // Hiển thị lên giao diện
            SetComboBox(ref cbHocKy_NhapDiem, dtHocKy, "TenHocKi", "MaHocKi");
            SetComboBox(ref cbKhoi_NhapDiem, dtKhoi, "Khoi", "Khoi");
            SetComboBox(ref cbKhoi_BangDiemLop, dtKhoi, "Khoi", "Khoi");
            SetComboBox(ref cbHocKy_BangDiemLop, dtHocKy, "TenHocKi", "MaHocKi");
            SetComboBox(ref cbHocKy_DanhGia, dtHocKy, "TenHocKi", "MaHocKi");
            SetComboBox(ref cbKhoi_DanhGia, dtKhoi, "Khoi", "Khoi");
            // khác
            lbTenHienThi.Text = TenHienThi;
            txtTenHIenThi.Text = TenHienThi;
            lbUsername.Text = Username;
            currenPanelActive = pnGioiThieu;
            currentButtonActive = btnGioiThieu;
            UpdateMenu(ref pnGioiThieu, ref btnGioiThieu, ref tpGioiThieu);
        }

        // sự kiện click ở menu
        #region Sự kiện click ở menu
        private void UpdateMenu(ref Panel selectedPanel, ref Button selectedButton, ref TabPage selectedTabPage)
        {
            currenPanelActive.BackColor = Color.FromArgb(42, 63, 84);
            currentButtonActive.BackColor = Color.FromArgb(42, 63, 84);

            currenPanelActive = selectedPanel;
            currentButtonActive = selectedButton;

            selectedPanel.BackColor = Color.FromArgb(26, 187, 156);
            selectedButton.BackColor = Color.FromArgb(53, 73, 93);

            tabControl.SelectedTab = selectedTabPage;
        }

        private void BtnTraCuu_Click(object sender, EventArgs e)
        {
            UpdateMenu(ref pnTraCuu, ref btnBangDiemLop, ref tpBangDiemLop);
        }

        private void BtnNhapDiem_Click(object sender, EventArgs e)
        {
            UpdateMenu(ref pnNhapDiem, ref btnNhapDiem, ref tpNhapDiem);
        }

        private void BtnDanhGia_Click(object sender, EventArgs e)
        {
            UpdateMenu(ref pnDanhGia, ref btnDanhGia, ref tpDanhGia);
        }

        private void BtnTaiKhoan_Click(object sender, EventArgs e)
        {
            UpdateMenu(ref pnTaiKhoan, ref btnTaiKhoan, ref tpTaiKhoan);
        }

        private void BtnGioiThieu_Click(object sender, EventArgs e)
        {
            UpdateMenu(ref pnGioiThieu, ref btnGioiThieu, ref tpGioiThieu);
        }

        private void BtnDangXuat_Click(object sender, EventArgs e)
        {
            this.Hide();
            LoginForm x = new LoginForm();
            x.ShowDialog();
            this.Close();
        }
        #endregion

        // tùy chỉnh thanh header
        #region Thanh Header
        private void BtnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void BtnMinimize_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
        }

        // kéo thả header
        private void PanelMoveForm_MouseMove(object sender, MouseEventArgs e)
        {
            if (mouseDown)
            {
                this.Location = new Point(
                    (this.Location.X - lastLocation.X) + e.X, (this.Location.Y - lastLocation.Y) + e.Y);

                this.Update();
            }
        }

        private void PanelMoveForm_MouseUp(object sender, MouseEventArgs e)
        {
            mouseDown = false;
        }

        private void PanelMoveForm_MouseDown(object sender, MouseEventArgs e)
        {
            mouseDown = true;
            lastLocation = e.Location;
        }
        #endregion

        //tab giới thiệu
        #region Tab giới thiệu
        private void AvatarDung_Click(object sender, EventArgs e)
        {
            OpenBrowser("https://facebook.com/DangDungCNTT");
        }

        private void AvatarTai_Click(object sender, EventArgs e)
        {
            OpenBrowser("https://www.facebook.com/jos.tai.1");
        }

        private void AvatarTien_Click(object sender, EventArgs e)
        {
            OpenBrowser("https://www.facebook.com/SocolaDaiCa1997");
        }

        private void AvatarThanh_Click(object sender, EventArgs e)
        {
            OpenBrowser("https://www.facebook.com/nguyenthanhat1997");
        }

        private void AvatarQuan_Click(object sender, EventArgs e)
        {
            OpenBrowser("https://www.facebook.com/dinhquan.157");
        }

        private void Copyright_Click(object sender, EventArgs e)
        {
            OpenBrowser("https://tentstudy.xyz/");
        }
        #endregion

        // tab nhập điểm
        #region Tab Nhập điểm
        private void ShowMonHocVaLop_NhapDiem(object sender, EventArgs e)
        {
            // chỉ khi chọn khối và học kỳ thì mới đc chọn lớp và môn
            cbMonHoc_NhapDiem.Enabled = false;
            cbLop_NhapDiem.Enabled = false;

            if (cbHocKy_NhapDiem.SelectedIndex == -1 || cbHocKy_NhapDiem.Items.Count == 0)
                return;
            if (cbKhoi_NhapDiem.SelectedIndex == -1 || cbKhoi_NhapDiem.Items.Count == 0)
                return;

            cbMonHoc_NhapDiem.Enabled = true;
            cbLop_NhapDiem.Enabled = true;

            DataTable dtMonHoc = cLayout.GetDanhSachMonHoc(cbKhoi_NhapDiem.Text);
            SetComboBox(ref cbMonHoc_NhapDiem, dtMonHoc, "TenMon", "MaMon");

            DataTable dtLop = cLayout.GetDanhSachLop(cbKhoi_NhapDiem.Text, cbHocKy_NhapDiem.SelectedValue.ToString());
            SetComboBox(ref cbLop_NhapDiem, dtLop, "TenLop", "MaLop");
        }

        private void ShowBtnHienThi_NhapDiem(object sender, EventArgs e)
        {
            // chọn lớp và môn xong mới đc nhập điểm
            DisableButton(ref btnHienThi_NhapDiem);
            if (cbMonHoc_NhapDiem.SelectedIndex == -1 || cbMonHoc_NhapDiem.Items.Count == 0)
                return;
            if (cbLop_NhapDiem.SelectedIndex == -1 || cbLop_NhapDiem.Items.Count == 0)
                return;
            EnableButton(ref btnHienThi_NhapDiem);
        }

        private void BtnHienThi_NhapDiem_Click(object sender, EventArgs e)
        {
            //En - Disable controls
            cbHocKy_NhapDiem.Enabled = false;
            cbKhoi_NhapDiem.Enabled = false;
            cbLop_NhapDiem.Enabled = false;
            cbMonHoc_NhapDiem.Enabled = false;

            EnableButton(ref btnLuu_NhapDiem);
            EnableButton(ref btnNhapExcel_NhapDiem);
            EnableButton(ref btnXuatExcel_NhapDiem);
            DisableButton(ref btnHienThi_NhapDiem);

            dgvNhapDiem.Enabled = true;
            dgvNhapDiem.MultiSelect = false;
            //xóa hết các cột hiện tại
            dgvNhapDiem.Columns.Clear();

            String MaLop = cbLop_NhapDiem.SelectedValue.ToString();
            String MaMon = cbMonHoc_NhapDiem.SelectedValue.ToString();
            String HocKy = cbHocKy_NhapDiem.SelectedValue.ToString();

            dgvNhapDiem.DataSource = cLayout.GetDanhSachDiem(MaLop, MaMon, HocKy);

            //thêm cột thứ tự
            AddColumnToDGV(ref dgvNhapDiem, "STT", "STT");

            //không cho sửa 
            dgvNhapDiem.Columns[0].ReadOnly = true;
            dgvNhapDiem.Columns[1].ReadOnly = true;
            dgvNhapDiem.Columns[2].ReadOnly = true;
            dgvNhapDiem.Columns[3].ReadOnly = true;
            dgvNhapDiem.Columns[4].ReadOnly = true;

            // kích thước
            dgvNhapDiem.Columns[0].Width = 40;
            dgvNhapDiem.Columns[1].Width = 120;
            dgvNhapDiem.Columns[2].Width = 120;
            dgvNhapDiem.Columns[3].Width = 100;
            dgvNhapDiem.Columns[4].Width = 100;

            //căn lề giữa cho ngày sinh
            dgvNhapDiem.Columns[4].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            //định dạng ngày sinh
            dgvNhapDiem.Columns[4].DefaultCellStyle.Format = "dd/MM/yyyy";

        }

        private void BtnLuu_NhapDiem_Click(object sender, EventArgs e)
        {
            // lưu điểm
            int soHang = dgvNhapDiem.Rows.Count;
            // kiểm tra điểm không hợp lệ
            float Diem = -10;
            String value;
            for (int i = 0; i < soHang - 1; i++)
            {
                value = dgvNhapDiem[5, i].Value.ToString();
                if (value == "" || (float.TryParse(value, out Diem) && 0 <= Diem && Diem <= 10))
                {
                    // thỏa mãn thì thôi
                }
                else
                {
                    dgvNhapDiem.ClearSelection();
                    dgvNhapDiem.Focus();
                    dgvNhapDiem[5, i].Selected = true;
                    MessageBox.Show("Điểm không hợp lệ: " + value);
                    return;
                }
            }
            String MaMon = cbMonHoc_NhapDiem.SelectedValue.ToString();
            String HocKy = cbHocKy_NhapDiem.SelectedValue.ToString();
            for (int i = 0; i < soHang - 1; i++)
            {
                String MaHocSinh = dgvNhapDiem[1, i].Value.ToString();
                value = dgvNhapDiem[5, i].Value.ToString();
                cLayout.CapNhatDiem(MaHocSinh, MaMon, HocKy, value);
            }
            //En - Disable controls
            cbHocKy_NhapDiem.Enabled = true;
            cbKhoi_NhapDiem.Enabled = true;
            cbLop_NhapDiem.Enabled = true;
            cbMonHoc_NhapDiem.Enabled = true;

            EnableButton(ref btnHienThi_NhapDiem);
            DisableButton(ref btnLuu_NhapDiem);
            DisableButton(ref btnNhapExcel_NhapDiem);
            DisableButton(ref btnXuatExcel_NhapDiem);

            dgvNhapDiem.Enabled = false;
        }

        private void BtnNhapExcel_NhapDiem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Chú ý:\n- Các ô điểm trống sẽ được bỏ qua.\n- Điểm sẽ được đọc theo thứ tự từ trên xuống dưới");
            //tọa độ Sĩ Số
            int SSX = 6, SSY = 8;
            //mở excel
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "Excel Document(*.xlsx)|*.xlsx|Excel Document 2003(*.xls)|*.xls|All files(*.*)|*.*",
                RestoreDirectory = true
            };

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                Excel.Application exApp = new Excel.Application();
                Excel.Workbook exBook = exApp.Workbooks.Open(openFileDialog.FileName);
                Excel.Worksheet exSheet = (Excel.Worksheet)exBook.Worksheets[1];
                Excel.Range cell;

                int DongBatDau = 10;
                int DongHienTai = 0;
                int STT = 0;
                //Sĩ số
                cell = (Excel.Range)exSheet.Cells[SSX, SSY];
                int SiSo = int.Parse(cell.Value.ToString());

                for (int i = 1; i <= SiSo; i++)
                {
                    DongHienTai = DongBatDau + STT;
                    cell = (Excel.Range)exSheet.Cells[DongHienTai, 6];
                    if (cell.Value == null)
                    {
                        STT++; continue;
                    }
                    dgvNhapDiem[5, STT].Value = cell.Value;
                    STT++;
                }
                exBook.Close(false);
                exApp.Quit();
            }

        }

        private void BtnXuatExcel_NhapDiem_Click(object sender, EventArgs e)
        {
            if (dgvNhapDiem.Rows.Count > 0) //có dữ liệu để xuất
            {
                //lấy dữ liệu trên form
                String MaHocKy = cbHocKy_NhapDiem.SelectedValue.ToString();
                String TenLop = cbLop_NhapDiem.Text;
                String TenMon = cbMonHoc_NhapDiem.Text;
                int SiSo = dgvNhapDiem.RowCount;

                //tách chuỗi mã học kỳ lấy kỳ học và năm học
                String[] mang = MaHocKy.Split('_');
                int NamHoc1 = int.Parse(mang[0]);
                int NamHoc2 = int.Parse(mang[1]);
                int HocKy = int.Parse(mang[2]);

                //cấu hình cho việc ghi dữ liệu
                //tọa độ tiêu đề
                int TDX = 4, TDY = 1;
                //tọa độ Học Kỳ
                int HKX = 6, HKY = 3;
                //tọa độ Năm Học
                int NHX = 7, NHY = 3;
                //tọa độ Sĩ Số
                int SSX = 6, SSY = 8;
                //dòng dữ liệu đầu tiên
                int DongDauTien = 10;

                //các thuộc tính của dữ liệu nguồn
                int soCot = dgvNhapDiem.ColumnCount;
                int soHang = dgvNhapDiem.RowCount;
                int DongHienTai = 0;
                int STT = 0;

                //mở excel
                Excel.Application exApp = new Excel.Application();
                Excel.Workbook exBook = exApp.Workbooks.Open(Path.GetFullPath("ExcelTemplate/BangDiemMonHoc.xlsx"));
                Excel.Worksheet exSheet = (Excel.Worksheet)exBook.Worksheets[1];
                Excel.Range cell;

                //ghi thông tin bảng điểm
                //Tiêu đề bảng điểm
                cell = (Excel.Range)exSheet.Cells[TDX, TDY];
                cell.Value = String.Format("BẢNG ĐIỂM MÔN {0} LỚP {1}", TenMon.ToUpper(), TenLop);
                //Học kỳ
                cell = (Excel.Range)exSheet.Cells[HKX, HKY];
                cell.Value = HocKy == 1 ? "I" : "II";
                //Năm học
                cell = (Excel.Range)exSheet.Cells[NHX, NHY];
                cell.Value = NamHoc1 + " - " + NamHoc2;
                //Sĩ số
                cell = (Excel.Range)exSheet.Cells[SSX, SSY];
                cell.Value = SiSo;

                //ghi danh sách học sinh, điểm

                foreach (DataGridViewRow row in dgvNhapDiem.Rows)
                {
                    DongHienTai = DongDauTien + STT;

                    //ghi từ dgv ra
                    for (int j = 0; j < soCot; j++)
                    {
                        cell = (Excel.Range)exSheet.Cells[DongHienTai, j + 1];
                        cell.Value = row.Cells[j].Value;
                    }

                    //ghi đánh giá
                    cell = (Excel.Range)exSheet.Cells[DongHienTai, soCot + 1];
                    cell.Value = String.Format("=IF(F{0} = \"\", \"\", IF(F{0} >= 9, \"Giỏi\", IF(F{0} >= 7, \"Khá\", \"Trung Bình\")))", DongHienTai);

                    STT++;
                }

                exSheet.Name = "BANG DIEM " + TenMon.ToUpper();
                exBook.Activate();
                SaveFileDialog saveFileDialog = new SaveFileDialog
                {
                    Filter = "Excel Document(*.xlsx)|*.xlsx|Excel Document 2003(*.xls)|*.xls|All files(*.*)|*.*",
                    FilterIndex = 1,
                    AddExtension = true,
                    DefaultExt = ".xlsx",
                    RestoreDirectory = true,
                    FileName = String.Format("Bang-Diem-Mon-{0}-Lop-{1}-Ky-{2}-{3}-{4}", TenMon, TenLop, HocKy, NamHoc1, NamHoc2)
                };

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        exBook.SaveAs(saveFileDialog.FileName.ToString());
                    } catch (Exception)
                    {
                        MessageBox.Show("Không thể lưu file");
                    }
                }
                exBook.Close(false);
                exApp.Quit();
            }
            else
            {
                MessageBox.Show("Không có dữ liệu");
            }
        }
        #endregion

        // tab bảng điểm lớp
        #region Tab bảng điểm lớp
        private void Show_Lop_BangDiemLop(object sender, EventArgs e)
        {
            //chỉ khi chọn Học kỳ và Khối thì mới được chọn lớp
            cbLop_BangDiemLop.Enabled = false;

            if (cbHocKy_BangDiemLop.SelectedIndex == -1 || cbHocKy_BangDiemLop.Items.Count == 0)
                return;
            if (cbKhoi_BangDiemLop.SelectedIndex == -1 || cbKhoi_BangDiemLop.Items.Count == 0)
                return;

            cbLop_BangDiemLop.Enabled = true;

            DataTable dtLop = cLayout.GetDanhSachLop(cbKhoi_BangDiemLop.SelectedValue.ToString(), cbHocKy_BangDiemLop.SelectedValue.ToString());
            SetComboBox(ref cbLop_BangDiemLop, dtLop, "TenLop", "MaLop");
        }

        private void BtnHienThi_BangDiemLop_Click(object sender, EventArgs e)
        {
            if (cbHocKy_BangDiemLop.SelectedIndex == -1 || cbLop_BangDiemLop.SelectedIndex == -1 || cbKhoi_BangDiemLop.SelectedIndex == -1)
            {
                MessageBox.Show("Vui lòng chọn học kỳ, khối và lớp");
                return;
            }

            String maLop = cbLop_BangDiemLop.SelectedValue.ToString();
            String HocKy = cbHocKy_BangDiemLop.SelectedValue.ToString();
            int Khoi = int.Parse(cbKhoi_BangDiemLop.SelectedValue.ToString());

            dgvBangDiemLop.Enabled = true;
            //xóa các cột hiện tại của dgv
            dgvBangDiemLop.Columns.Clear();

            //lấy bảng điểm của lớp
            dgvBangDiemLop.DataSource = cLayout.GetBangDiemLop(Khoi, maLop, HocKy);

            //Chèn thêm cột số thứ tự vào dgv
            AddColumnToDGV(ref dgvBangDiemLop, "STT", "STT");

            //điều chỉnh kích thước các cột
            dgvBangDiemLop.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            dgvBangDiemLop.Columns[0].Width = 40;
            dgvBangDiemLop.Columns[1].Width = 110;
            dgvBangDiemLop.Columns[2].Width = 120;
            dgvBangDiemLop.Columns[3].Width = 60;
            dgvBangDiemLop.Columns[4].Width = 90;


            //căn lề giữa cho ngày sinh
            dgvBangDiemLop.Columns[4].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            //định dạng ngày sinh
            dgvBangDiemLop.Columns[4].DefaultCellStyle.Format = "dd/MM/yyyy";

            dgvBangDiemLop.ReadOnly = true;
        }

        private void BtnXuatExcel_BangDiemLop_Click(object sender, EventArgs e)
        {

            if (dgvBangDiemLop.Rows.Count > 0) //có dữ liệu để xuất
            {
                //lấy dữ liệu trên form
                int Khoi = int.Parse(cbKhoi_BangDiemLop.SelectedValue.ToString());
                String MaLop = cbLop_BangDiemLop.SelectedValue.ToString();

                //lấy dữ liệu database (Tên Lớp, Sĩ Số, Học kỳ, Năm Học)
                DataTable data = cLayout.GetThongTinLop(MaLop);
                String MaHocKy = data.Rows[0]["HocKi"].ToString();
                String TenLop = data.Rows[0]["TenLop"].ToString();
                int SiSo = int.Parse(data.Rows[0]["SiSo"].ToString());

                //tách chuỗi mã học kỳ lấy kỳ học và năm học
                String[] mang = MaHocKy.Split('_');
                int NamHoc1 = int.Parse(mang[0]);
                int NamHoc2 = int.Parse(mang[1]);
                int HocKy = int.Parse(mang[2]);

                //cấu hình cho việc ghi dữ liệu
                //tọa độ tiêu đề
                int TDX = 4, TDY = 1;
                //tọa độ Học Kỳ
                int HKX = 6, HKY = 3;
                //tọa độ Năm Học
                int NHX = 7, NHY = 3;
                //tọa độ Sĩ Số
                int SSX = 6, SSY = 11;
                //dòng dữ liệu đầu tiên
                int DongDauTien = 10;

                //các thuộc tính của dữ liệu nguồn
                int soCot = dgvBangDiemLop.ColumnCount;
                int soHang = dgvBangDiemLop.RowCount;
                int DongHienTai = 0;
                int STT = 0;

                //mở excel
                Excel.Application exApp = new Excel.Application();
                Excel.Workbook exBook = exApp.Workbooks.Open(Path.GetFullPath("ExcelTemplate/BangDiemLop" + Khoi + ".xlsx"));
                Excel.Worksheet exSheet = (Excel.Worksheet)exBook.Worksheets[1];
                Excel.Range cell;

                //ghi thông tin bảng điểm
                //Tiêu đề bảng điểm
                cell = (Excel.Range)exSheet.Cells[TDX, TDY];
                cell.Value = String.Format("BẢNG ĐIỂM LỚP {0}", TenLop);
                //Học kỳ
                cell = (Excel.Range)exSheet.Cells[HKX, HKY];
                cell.Value = HocKy == 1 ? "I" : "II";
                //Năm học
                cell = (Excel.Range)exSheet.Cells[NHX, NHY];
                cell.Value = NamHoc1 + " - " + NamHoc2;
                //Sĩ số
                cell = (Excel.Range)exSheet.Cells[SSX, SSY];
                cell.Value = SiSo;

                //ghi danh sách học sinh, điểm

                foreach (DataGridViewRow row in dgvBangDiemLop.Rows)
                {
                    DongHienTai = DongDauTien + STT;

                    //ghi từ dgv ra
                    for (int j = 0; j < soCot; j++)
                    {
                        cell = (Excel.Range)exSheet.Cells[DongHienTai, j + 1];
                        cell.Value = row.Cells[j].Value ?? "";
                    }

                    //ghi học lực
                    cell = (Excel.Range)exSheet.Cells[DongHienTai, soCot + 1];
                    cell.Value = String.Format("=IF(OR(F{0} = \"\", G{0} = \"\"), \"\", IF(AND(F{0} >= 9, G{0} >= 9), \"Giỏi\", IF(AND(F{0} >= 7, G{0} >= 7), \"Khá\", \"Trung Bình\")))", DongHienTai);

                    STT++;
                }

                //ghi các thông tin khác
                exSheet.Name = "BẢNG ĐIỂM LỚP " + TenLop;
                exBook.Activate();

                //Mở hộp thoại save file
                SaveFileDialog saveFileDialog = new SaveFileDialog
                {
                    Filter = "Excel Document(*.xlsx)|*.xlsx|Excel Document 2003(*.xls)|*.xls|All files(*.*)|*.*",
                    FilterIndex = 1,
                    AddExtension = true,
                    DefaultExt = ".xlsx",
                    RestoreDirectory = true,
                    FileName = String.Format("Bang-Diem-Lop-{0}-Ky-{1}-{2}-{3}", TenLop, HocKy, NamHoc1, NamHoc2)
                };
                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        exBook.SaveAs(saveFileDialog.FileName.ToString());
                    }
                    catch (Exception)
                    {
                        MessageBox.Show("Không thể lưu file");
                    }
                }
                exBook.Close(false);
                exApp.Quit(); //thoát excel
            }
            else
            {
                MessageBox.Show("Không có dữ liệu");
            }
        }
        #endregion

        // tab đánh giá
        #region Tab Đánh giá
        private void BtnHienThi_DanhGia_Click(object sender, EventArgs e)
        {
            if (cbKhoi_DanhGia.SelectedIndex == -1 || cbHocKy_DanhGia.SelectedIndex == -1)
            {
                MessageBox.Show("Vui lòng chọn học kỳ và khối"); return;
            }

            //lấy dữ liệu trên form
            String Khoi = cbKhoi_DanhGia.SelectedValue.ToString();
            String MaHocKy = cbHocKy_DanhGia.SelectedValue.ToString();

            //tách chuỗi mã học kỳ lấy kỳ học và năm học
            String[] mang = MaHocKy.Split('_');
            int NamHoc1 = int.Parse(mang[0]);
            int NamHoc2 = int.Parse(mang[1]);
            int HocKi = int.Parse(mang[2]);

            //lấy, xử lý dữ liệu từ database
            DataTable dtDanhGia = cLayout.GetDanhGia(Khoi, MaHocKy);
            int[] a = new int[5];
            int[] b = new int[5];
            int[] c = new int[5];
            string[] lop = new string[5];

            for (int i = 0; i < 5; i++)
            {
                a[i] = int.Parse(dtDanhGia.Rows[i][2].ToString());
                b[i] = int.Parse(dtDanhGia.Rows[i][3].ToString());
                c[i] = int.Parse(dtDanhGia.Rows[i][4].ToString());
                lop[i] = dtDanhGia.Rows[i][1].ToString();
            }


            //đặt tên cho biểu đồ
            chtDanhGia.Titles[0].Text = String.Format("Biểu đồ học lực khối {0}", Khoi);
            chtDanhGia.Titles[1].Text = String.Format("Học kỳ {0} năm học {1} - {2}", HocKi == 1 ? "I" : "II", NamHoc1, NamHoc2);

            //đổ dữ liệu cho biểu đồ
            chtDanhGia.Series["HSGioi"].Points.DataBindY(a);
            chtDanhGia.Series["HSKha"].Points.DataBindY(b);
            chtDanhGia.Series["HSTrungBinh"].Points.DataBindY(c);
            
            //xóa tiêu đề các cột
            chtDanhGia.ChartAreas[0].AxisX.CustomLabels.Clear();

            //đặt tiêu đề cho từng cột
            int d = 2;
            for (int i = 0; i < 5; i++)
            {
                chtDanhGia.ChartAreas[0].AxisX.CustomLabels.Add(new CustomLabel(0, d, lop[i], 0, LabelMarkStyle.None));
                d += 2;
            }

            //En - Disable Controls
            EnableButton(ref btnXuatPNGChart);
        }

        private void BtnXuatPNGChart_Click(object sender, EventArgs e)
        {
            if (chtDanhGia.Series["HSGioi"].Points.Count == 0) //kiểm tra biểu đồ rỗng
            {
                MessageBox.Show("Biểu đồ rỗng");
                return;
            }

            //lấy dữ liệu trên form
            String Khoi = cbKhoi_DanhGia.SelectedValue.ToString();
            String MaHocKy = cbHocKy_DanhGia.SelectedValue.ToString();
            
            //tách chuỗi mã học kỳ lấy kỳ học và năm học
            String[] mang = MaHocKy.Split('_');
            int NamHoc1 = int.Parse(mang[0]);
            int NamHoc2 = int.Parse(mang[1]);
            int HocKy = int.Parse(mang[2]);

            //lưu file
            SaveFileDialog saveFileDialog = new SaveFileDialog
            {
                Filter = "PNG|*.png|All files(*.*)|*.*",
                FilterIndex = 1,
                AddExtension = true,
                DefaultExt = ".png",
                RestoreDirectory = true,
                FileName = String.Format("Bieu-do-hoc-luc-khoi-{0}-Ky-{1}-{2}-{3}", Khoi, HocKy, NamHoc1, NamHoc2)
            };
            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                chtDanhGia.SaveImage(saveFileDialog.FileName.ToString(), ChartImageFormat.Png);
            }
        }
        #endregion

        // tab tài khoản
        #region Tab Tài Khoản
        private void TxtTenHIenThi_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter && txtTenHIenThi.Text != "")
            {
                BtnLuuTenNguoiDung_Click(sender, e);
                return;
            }
        }

        private void TxtNhapLaiMatKhauMoi_KeyUp(object sender, KeyEventArgs e)
        {

            if (e.KeyCode == Keys.Enter && txtNhapLaiMatKhauMoi.Text != "" && lblThongBaoThatBai.Text == "")
            {
                BtnDoiMatKhau_Click(sender, e);
                return;
            }
            lblThongBaoThatBai.Text = "";
            if (txtMatKhauMoi.Text.Equals(txtNhapLaiMatKhauMoi.Text) == false)
            {
                lblThongBaoThatBai.Text = "Hai mật khẩu không trùng khớp";
                DisableButton(ref btnDoiMatKhau);
            }
            else
            {
                EnableButton(ref btnDoiMatKhau);
            }
        }

        private void BtnLuuTenNguoiDung_Click(object sender, EventArgs e)
        {
            String tb = "";
            if (txtTenHIenThi.Text != "")
            {
                if (cUser.CapNhatTenHienThi(Username, txtTenHIenThi.Text))
                {
                    tb = "Cập nhật tên người dùng thành công";
                    MessageBox.Show(tb, "Thành công");
                }
                else
                {
                    tb = "Cập nhật tên người dùng thất bại";
                    MessageBox.Show(tb, "Thất bại", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void BtnDoiMatKhau_Click(object sender, EventArgs e)
        {
            String tb = "";
            if (txtMatKhauMoi.Text == txtNhapLaiMatKhauMoi.Text && txtNhapLaiMatKhauMoi.Text != "")
            {
                if (cUser.CapNhatMatKhau(Username, txtMatKhauHienTai.Text, txtMatKhauMoi.Text))
                {
                    tb = "Đổi mật khẩu thành công";
                    MessageBox.Show(tb, "Thành công");
                    BtnDangXuat_Click(sender, e);
                }
                else
                {
                    tb = "Đổi mật khẩu thất bại";
                    txtMatKhauHienTai.Focus();
                    txtMatKhauHienTai.SelectAll();
                    MessageBox.Show(tb, "Thất bại", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            DisableButton(ref btnDoiMatKhau);
        }
        #endregion

        //các hàm dùng chung
        #region Các hàm dùng chung
        private void OpenBrowser(String url)
        {
            System.Diagnostics.Process.Start(url);
        }

        private void AddColumnToDGV(ref CustomDGV dgv, String ColName, String ColText)
        {
            //Chèn thêm cột số thứ tự vào dgv
            DataGridViewTextBoxColumn colSTT = new DataGridViewTextBoxColumn
            {
                HeaderText = ColText,
                Name = ColName
            };
            dgv.Columns.Insert(0, colSTT);

            //đánh số thứ tự
            int i = 1;
            foreach (DataGridViewRow row in dgv.Rows)
            {
                row.Cells[0].Value = i;
                i++;
            }
        }

        private void SetComboBox(ref ComboBox cb, DataTable dt, String DisplayMember, String ValueMember)
        {
            cb.BindingContext = new BindingContext();
            cb.DataSource = dt;
            cb.DisplayMember = DisplayMember;
            cb.ValueMember = ValueMember;
            cb.SelectedIndex = -1;
        }

        //            

        private void EnableButton(ref Button button)
        {
            button.Enabled = true;
            button.BackColor = Color.White;
        }

        private void DisableButton(ref Button button)
        {
            button.Enabled = false;
            button.BackColor = Color.FromArgb(240, 240, 240);
        }
        #endregion
    }
}
