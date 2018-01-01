using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.SqlClient;
namespace QuanLyDiemTieuHoc.Model
{
    class Database
    {
        SqlConnection conn = null;
        String connectionString = @"Data Source=(local);Initial Catalog=QuanLyDiemTieuHoc;Integrated Security=True";

        public void CreateConnect()
        {
            conn = new SqlConnection(connectionString);
            conn.Open();
        }

        public void CloseConnect()
        {
            if (conn.State == ConnectionState.Open)
            {
                conn.Close();
            }
        }

        public DataTable Excute(String sql)
        {
            DataTable dat = new DataTable();
            CreateConnect();
            SqlDataAdapter dataAdapter = new SqlDataAdapter(sql, conn);
            dataAdapter.Fill(dat);
            CloseConnect();
            return dat;
        }
        // đổ dữ liệu vào dataset
        public DataSet Excute_DataSet(String sql)
        {
            DataSet das = new DataSet();
            CreateConnect();
            SqlDataAdapter dataAdapter = new SqlDataAdapter(sql, conn);
            dataAdapter.Fill(das);
            CloseConnect();
            return das;
        }

        //câu lệnh chuyên dùng để thực thi các truy vấn update ...
        public int ExcuteNonQuery(String sql)
        {
            CreateConnect();
            SqlCommand cmd = new SqlCommand(sql, conn);
            int resul = cmd.ExecuteNonQuery();
            CloseConnect();
            return resul;
        }
        // câu lệnh chuyền dùng để thức thi các truy vấn trả về 1 giá trị VD như Count
        public object ExcuteScalar(String sql)
        {
            CreateConnect();
            SqlCommand cmd = new SqlCommand(sql, conn);
            object obj = cmd.ExecuteScalar();
            CloseConnect();
            return obj;
        }
    }
}
