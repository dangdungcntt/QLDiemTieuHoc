using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using QuanLyDiemTieuHoc.Model;
using System.Security.Cryptography;
using System.Data;
using System.Data.SqlClient;
namespace QuanLyDiemTieuHoc.Controller
{
    class Controller
    {
        public Database db = new Database();
        public String Md5(String input)
        {
            StringBuilder hash = new StringBuilder();
            MD5CryptoServiceProvider md5provider = new MD5CryptoServiceProvider();
            byte[] bytes = md5provider.ComputeHash(new UTF8Encoding().GetBytes(input));

            for (int i = 0; i < bytes.Length; i++)
            {
                hash.Append(bytes[i].ToString("x2"));
            }
            return hash.ToString();
        }
    }
}
