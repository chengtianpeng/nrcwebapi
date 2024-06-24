using SqlSugar;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace NRCWebApi.Model
{
    /// <summary>
    /// 8.188
    /// </summary>
    [Table("ss_User")]
    public class ss_User
    {

        [SugarColumn(IsPrimaryKey = true)]
        public int UserId { get; set; }
        public string LoginName { get; set; }
        public string UserEName { get; set; }
        public string UserCName { get; set; }
        public string Password { get; set; }
        public int Gender { get; set; }
        public int Status { get; set; }
        public string Email { get; set; }
        public string Cellphone { get; set; }
        public string Telephone { get; set; }
        public string FinanceCode { get; set; }
        public int SortNumber { get; set; }
        public DateTime RegistrationTime { get; set; }
        public int Deleted { get; set; }
        public string Remark { get; set; }
        public string Esign { get; set; }

        public double PicSize { get; set; }
        public string PicType { get; set; }
        public string PicPath { get; set; }
        public bool IsAllowPic { get; set; }
        public string Barcode { get; set; }
        public DateTime LastPasswordModificationDate { get; set; }
        public DateTime PasswordExpirationDate { get; set; }

    }
}
