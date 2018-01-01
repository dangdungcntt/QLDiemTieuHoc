using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.SqlClient;
namespace QuanLyDiemTieuHoc.Controller
{
    class Controller_Layout : Controller
    {
        String sql;
        public DataTable GetDanhSachHocKy()
        {
            return db.Excute(@"SELECT * FROM HocKi");
        }
        public DataTable GetDanhSachKhoi()
        {
            return db.Excute(@"SELECT * FROM Khoi");
        }
        public DataTable GetDanhSachMonHoc(String khoi)
        {
            sql = String.Format(@"
                SELECT 
                    Khoi_MonHoc.MaMon, MonHoc.TenMon 
                FROM Khoi_MonHoc 
                JOIN MonHoc 
                ON Khoi_MonHoc.MaMon = MonHoc.MaMon 
                WHERE Khoi = {0}
            ", khoi.ToString());
            return db.Excute(sql);
        }
        public DataTable GetDanhSachLop(String khoi, String hocKi)
        {
            sql = String.Format(@"
                SELECT MaLop, TenLop 
                FROM Lop 
                WHERE Hocki = '{0}' AND Khoi = {1}
            ", hocKi, khoi);    
            return db.Excute(sql);
        }
        public DataTable GetDanhSachDiem(String maLop, String MaMon, String HocKi)
        {
            sql = String.Format(@"
                SELECT 
                    HocSinh.MaHocSinh AS 'Mã Học Sinh', HocSinh.Ho AS 'Họ đệm', HocSinh.Ten AS 'Tên',
                    HocSinh.NgaySinh AS 'Ngày Sinh', Diem AS 'Điểm'
                FROM (
                    SELECT 
                        MaHocSinh, MaMon, MaHocki FROM HocSinh_lop, MonHoc, HocKi 
                    WHERE HocSinh_Lop.MaLop = '{0}' AND MonHoc.MaMon = '{1}' AND Hocki.MaHocKi = '{2}'
                ) AS zz LEFT JOIN Diem 
                ON zz.MaHocSinh = Diem.MaHocSinh AND zz.MaMon = Diem.MaMon AND zz.MaHocKi = Diem.HocKi
                JOIN HocSinh 
                ON HocSinh.MaHocSinh = zz.MaHocSinh ORDER BY HocSinh.Ten ASC
            ", maLop, MaMon, HocKi);
            return db.Excute(sql);
        }
        public int CapNhatDiem(String MaHocSinh, String MaMon, String HocKi, String Diem)
        {
            int kq = 0;
            if (Diem == "") Diem = "NULL";
            sql = String.Format(@"
                SELECT 
                    COUNT(*) 
                FROM Diem 
                WHERE MaHocSinh='{0}' AND HocKi='{1}' AND MaMon='{2}'
            ", MaHocSinh, HocKi, MaMon);
            if (int.Parse(db.ExcuteScalar(sql).ToString()) != 0)
            {
                //Console.WriteLine("aa");
                sql = String.Format(@"
                    UPDATE Diem 
                    SET Diem = {0} 
                    WHERE MaHocSinh = '{1}' AND Hocki = '{2}' and MaMon = '{3}'
                ", Diem, MaHocSinh, HocKi, MaMon);
                kq += db.ExcuteNonQuery(sql);
            }
            else
            {
                sql = String.Format(@"
                    INSERT INTO Diem (MaHocSinh, HocKi, MaMon, Diem) 
                    VALUES ('{0}', '{1}', '{2}', {3})
                ", MaHocSinh, HocKi, MaMon, Diem);
                kq += db.ExcuteNonQuery(sql);
            }
            return kq;
        }
        public void SinhDiemNgauNhien()
        {
            sql = String.Format(@"
                SELECT 
                    MaHocSinh, Hocki, MaMon 
                FROM HocSinh_Lop 
                JOIN lop 
                ON HocSinh_Lop.MaLop = lop.MaLop
                JOIN Khoi_MonHoc 
                ON lop.Khoi = Khoi_MonHoc.khoi
            ");
            DataTable x = db.Excute(sql);
            Random rand = new Random();
            foreach (DataRow row in x.Rows)
            {
                string mahocsinh = row["MaHocSinh"].ToString();
                string hocki = row["Hocki"].ToString();
                string mamon = row["MaMon"].ToString();
                float diem = (float)(rand.Next(23, 41)) / 4;
                CapNhatDiem(mahocsinh, mamon, hocki, diem.ToString());
            }
        }
        public DataTable GetBangDiemLop(int khoi, String maLop, String HocKi)
        {
            string query;
            if (khoi == 1)
            {
                query = String.Format(@"
                    SELECT 
                        HocSinh.MaHocSinh AS 'Mã học sinh', HocSinh.Ho AS 'Họ đệm', HocSinh.Ten AS 'Tên',
                        HocSinh.NgaySinh AS 'Ngày sinh', Toan AS 'Toán' , TiengViet AS 'Tiếng Việt',
                        DaoDuc AS 'Đạo Đức', TuNhienXaHoi AS 'TNXH', AmNhac AS 'Âm Nhạc', MiThuat AS 'Mĩ Thuật'
                    FROM HocSinh JOIN (
	                    SELECT 
                            HocSinh.MaHocSinh,
		                    SUM(CASE WHEN Diem.MaMon = 'T' THEN ROUND(Diem.Diem,2) ELSE null END) AS Toan,
		                    SUM(CASE WHEN Diem.MaMon = 'TV' THEN ROUND(Diem.Diem,2) ELSE null END) AS TiengViet,
		                    SUM(CASE WHEN Diem.MaMon = 'DD' THEN ROUND(Diem.Diem,2) ELSE null END) AS DaoDuc,
		                    SUM(CASE WHEN Diem.MaMon = 'TNXH' THEN ROUND(Diem.Diem,2) ELSE null END) AS TuNhienXaHoi,
		                    SUM(CASE WHEN Diem.MaMon = 'AN' THEN ROUND(Diem.Diem,2) ELSE null END) AS AmNhac,
		                    SUM(CASE WHEN Diem.MaMon = 'MT' THEN ROUND(Diem.Diem,2) ELSE null END) AS MiThuat
	                    FROM HocSinh 
                        JOIN HocSinh_Lop 
		                ON HocSinh.MaHocSinh = HocSinh_Lop.MaHocSinh
		                JOIN Diem
                        ON Diem.MaHocSinh = HocSinh.MaHocSinh
	                    WHERE HocSinh_Lop.MaLop = '{0}' AND Diem.HocKi = '{1}'
	                    GROUP BY HocSinh.MaHocSinh
                    ) AS zz on HocSinh.MaHocSinh = zz.MaHocSinh ORDER BY HocSinh.Ten ASC
                ", maLop, HocKi);
                return db.Excute(query);
            }
            if (khoi == 2)
            {
                query = String.Format(@"
                    SELECT 
                        HocSinh.MaHocSinh AS 'Mã học sinh', HocSinh.Ho AS 'Họ đệm', HocSinh.Ten AS 'Tên',
                        HocSinh.NgaySinh AS 'Ngày sinh', Toan AS 'Toán' , TiengViet AS 'Tiếng Việt',
                        TiengAnh AS 'Tiếng Anh', DaoDuc AS 'Đạo Đức', TuNhienXaHoi AS 'TNXH',
                        AmNhac AS 'Âm Nhạc', MiThuat AS 'Mĩ Thuật', TinHoc AS 'Tin Học' 
                    FROM HocSinh JOIN (
	                    SELECT 
                            HocSinh.MaHocSinh,
		                    SUM(CASE WHEN Diem.MaMon = 'T' THEN ROUND(Diem.Diem,2) ELSE null END) AS Toan,
		                    SUM(CASE WHEN Diem.MaMon = 'TV' THEN ROUND(Diem.Diem,2) ELSE null END) AS TiengViet,
		                    SUM(CASE WHEN Diem.MaMon = 'TA' THEN ROUND(Diem.Diem,2) ELSE null END) AS TiengAnh,
		                    SUM(CASE WHEN Diem.MaMon = 'DD' THEN ROUND(Diem.Diem,2) ELSE null END) AS DaoDuc,
		                    SUM(CASE WHEN Diem.MaMon = 'TNXH' THEN ROUND(Diem.Diem,2) ELSE null END) AS TuNhienXaHoi,
		                    SUM(CASE WHEN Diem.MaMon = 'AN' THEN ROUND(Diem.Diem,2) ELSE null END) AS AmNhac,
		                    SUM(CASE WHEN Diem.MaMon = 'MT' THEN ROUND(Diem.Diem,2) ELSE null END) AS MiThuat,
		                    SUM(CASE WHEN Diem.MaMon = 'TH' THEN ROUND(Diem.Diem,2) ELSE null END) AS TinHoc
	                    FROM HocSinh 
                        JOIN HocSinh_Lop 
		                ON HocSinh.MaHocSinh = HocSinh_Lop.MaHocSinh
		                JOIN Diem
                        ON Diem.MaHocSinh = HocSinh.MaHocSinh
	                    WHERE HocSinh_Lop.MaLop = '{0}' AND Diem.HocKi = '{1}'
	                    GROUP BY HocSinh.MaHocSinh
                    ) AS zz ON HocSinh.MaHocSinh = zz.MaHocSinh ORDER BY HocSinh.Ten ASC
                ", maLop, HocKi);
                return db.Excute(query);
            }
            if (khoi == 3)
            {
                query = String.Format(@"
                    SELECT 
                        HocSinh.MaHocSinh AS 'Mã học sinh', HocSinh.Ho AS 'Họ đệm', HocSinh.Ten AS 'Tên',
                        HocSinh.NgaySinh AS 'Ngày sinh', Toan AS 'Toán' , TiengViet AS 'Tiếng Việt',
                        TiengAnh AS 'Tiếng Anh', DaoDuc AS 'Đạo Đức', KhoaHoc AS 'Khoa Học',
                        AmNhac AS 'Âm Nhạc', MiThuat AS 'Mĩ Thuật', TinHoc AS 'Tin Học' 
                    FROM HocSinh JOIN (
	                        SELECT 
                                HocSinh.MaHocSinh,
		                        SUM(CASE WHEN Diem.MaMon = 'T' THEN ROUND(Diem.Diem,2) ELSE null END) AS Toan,
		                        SUM(CASE WHEN Diem.MaMon = 'TV' THEN ROUND(Diem.Diem,2) ELSE null END) AS TiengViet,
		                        SUM(CASE WHEN Diem.MaMon = 'TA' THEN ROUND(Diem.Diem,2) ELSE null END) AS TiengAnh,
		                        SUM(CASE WHEN Diem.MaMon = 'DD' THEN ROUND(Diem.Diem,2) ELSE null END) AS DaoDuc,
		                        SUM(CASE WHEN Diem.MaMon = 'KH' THEN ROUND(Diem.Diem,2) ELSE null END) AS KhoaHoc,
		                        SUM(CASE WHEN Diem.MaMon = 'AN' THEN ROUND(Diem.Diem,2) ELSE null END) AS AmNhac,
		                        SUM(CASE WHEN Diem.MaMon = 'MT' THEN ROUND(Diem.Diem,2) ELSE null END) AS MiThuat,
		                        SUM(CASE WHEN Diem.MaMon = 'TH' THEN ROUND(Diem.Diem,2) ELSE null END) AS TinHoc
	                        FROM HocSinh 
                            JOIN HocSinh_Lop 
		                    ON HocSinh.MaHocSinh = HocSinh_Lop.MaHocSinh
		                    JOIN Diem
		                    ON Diem.MaHocSinh = HocSinh.MaHocSinh
	                        WHERE HocSinh_Lop.MaLop = '{0}' AND Diem.HocKi = '{1}'
	                        GROUP BY HocSinh.MaHocSinh
                    ) AS zz ON HocSinh.MaHocSinh = zz.MaHocSinh ORDER BY HocSinh.Ten ASC
                ", maLop, HocKi);
                return db.Excute(query);
            }
            query = string.Format(@"
                SELECT 
                    HocSinh.MaHocSinh AS 'Mã học sinh', HocSinh.Ho AS 'Họ đệm', HocSinh.Ten AS 'Tên',
                    HocSinh.NgaySinh AS 'Ngày sinh', Toan AS 'Toán' , TiengViet AS 'Tiếng Việt',
                    TiengAnh AS 'Tiếng Anh', DaoDuc AS 'Đạo Đức', KhoaHoc AS 'Khoa Học', AmNhac AS 'Âm Nhạc', 
                    KiThuat AS 'Kĩ Thuật', MiThuat AS 'Mĩ Thuật',
                    LichSuDiaLy AS 'LS-ĐL', TinHoc AS 'Tin Học'
                FROM HocSinh JOIN (
	                SELECT HocSinh.MaHocSinh,
		                SUM(CASE WHEN Diem.MaMon = 'T' THEN ROUND(Diem.Diem,2) ELSE null END) AS Toan,
		                SUM(CASE WHEN Diem.MaMon = 'TV' THEN ROUND(Diem.Diem,2) ELSE null END) AS TiengViet,
		                SUM(CASE WHEN Diem.MaMon = 'TA' THEN ROUND(Diem.Diem,2) ELSE null END) AS TiengAnh,
		                SUM(CASE WHEN Diem.MaMon = 'DD' THEN ROUND(Diem.Diem,2) ELSE null END) AS DaoDuc,
		                SUM(CASE WHEN Diem.MaMon = 'KH' THEN ROUND(Diem.Diem,2) ELSE null END) AS KhoaHoc,
		                SUM(CASE WHEN Diem.MaMon = 'AN' THEN ROUND(Diem.Diem,2) ELSE null END) AS AmNhac,
		                SUM(CASE WHEN Diem.MaMon = 'MT' THEN ROUND(Diem.Diem,2) ELSE null END) AS MiThuat,
                        SUM(CASE WHEN Diem.MaMon = 'KT' THEN ROUND(Diem.Diem,2) ELSE null END) AS KiThuat, 
                        SUM(CASE WHEN Diem.MaMon = 'TH' THEN ROUND(Diem.Diem,2) ELSE null END) AS TinHoc,
		                SUM(CASE WHEN Diem.MaMon = 'LSDL' THEN ROUND(Diem.Diem,2) ELSE null END) as LichSuDiaLy
	                FROM HocSinh
                    JOIN HocSinh_Lop 
		            ON HocSinh.MaHocSinh = HocSinh_Lop.MaHocSinh
		            JOIN Diem
		            ON Diem.MaHocSinh = HocSinh.MaHocSinh
	                WHERE HocSinh_Lop.MaLop = '{0}' AND Diem.HocKi = '{1}'
	                GROUP BY HocSinh.MaHocSinh
                ) AS zz ON HocSinh.MaHocSinh = zz.MaHocSinh ORDER BY HocSinh.Ten ASC
            ", maLop, HocKi);
            return db.Excute(query);
        }
        public DataTable GetDanhGia(String Khoi, String HocKi)
        {
            sql = String.Format(@"
                SELECT 
                    Lop.MaLop, Lop.TenLop,
                    SUM(CASE WHEN diem >= 8 THEN 1 ELSE 0 END) AS 'Gioi',
                    SUM(CASE WHEN 8 > diem and diem >= 6 THEN 1 ELSE 0 END) AS 'Kha',
                    SUM(CASE WHEN 6 > diem and diem >= 0 THEN 1 ELSE 0 END) AS 'TrungBinh',
                    COUNT(diem) AS 'siso'
                FROM Lop
                JOIN HocSinh_Lop 
                ON HocSinh_Lop.MaLop = lop.MaLop
                JOIN (
                    SELECT 
                        MaHocSinh, MIN(diem) AS 'diem' 
                    FROM diem
                    WHERE diem.MaMon in('T', 'TV') AND Diem.HocKi = '{0}'
                    GROUP BY diem.MaHocSinh
                ) AS diem 
                ON HocSinh_Lop.MaHocSinh = diem.MaHocSinh
                WHERE khoi = '{1}' AND lop.Hocki = '{0}'
                GROUP BY lop.MaLop, lop.TenLop
            ", HocKi, Khoi);
            return db.Excute(sql);
        }

        public DataTable GetThongTinLop(string maLop)
        {
            sql = String.Format(@"
                SELECT 
                    COUNT(*) AS SiSo, Lop.HocKi, Lop.TenLop
                FROM Lop 
                JOIN HocSinh_Lop 
                ON Lop.MaLop = HocSinh_Lop.MaLop
                WHERE Lop.MaLop = '{0}'
                GROUP BY Lop.HocKi, Lop.TenLop
            ", maLop);
            return db.Excute(sql);
        }
    }
}
