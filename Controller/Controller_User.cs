using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
namespace QuanLyDiemTieuHoc.Controller
{
    class Controller_User : Controller
    {
        String sql;
        public bool CheckUserNameAndPassword(String username, String password)
        {
            password = Md5(password);
            sql = String.Format(@"
                SELECT 
                    COUNT(*) 
                FROM users 
                WHERE TenTaiKhoan = '{0}' AND MatKhau = '{1}'
            ", username, password);
            int check = int.Parse(db.ExcuteScalar(sql).ToString());
            return check == 1;
        }
        public String GetTenHienThi(String tenTaiKhoan)
        {
            sql = String.Format(@"
                SELECT TenHienThi 
                FROM users 
                WHERE TenTaiKhoan = '{0}'
            ", tenTaiKhoan);
            return db.ExcuteScalar(sql).ToString();
        }
        public bool CapNhatTenHienThi(String TenTaiKhoan, String TenHienThi)
        {
            sql = String.Format(@"
                UPDATE users 
                SET TenHIenThi = N'{0}' WHERE TenTaiKhoan ='{1}'
            ", TenHienThi, TenTaiKhoan);
            return db.ExcuteNonQuery(sql) != 0;
        }
        public bool CapNhatMatKhau(String TenTaiKhoan, String MatKhauHienTai, String matKhau)
        {
            if (CheckUserNameAndPassword(TenTaiKhoan, MatKhauHienTai)) {
                sql = String.Format(@"
                    UPDATE users SET MatKhau = '{0}' 
                    WHERE TenTaiKhoan = '{1}'
                ", Md5(matKhau), TenTaiKhoan);
                return db.ExcuteNonQuery(sql) != 0;
            }
            return false;
        }
    }
}
