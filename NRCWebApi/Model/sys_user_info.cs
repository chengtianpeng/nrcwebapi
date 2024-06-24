using SqlSugar;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace NRCWebApi.Model
{
    /// <summary>
    /// OA
    /// </summary>
    [Table("sys_user_info")]
    public class sys_user_info
    {
        /// <summary>
        /// 用户ID
        /// </summary>
        [SugarColumn(IsPrimaryKey = true)]
        public long user_id { get; set; }
        /// <summary>
        /// 部门ID
        /// </summary>
        public long dept_id { get; set; }
        /// <summary>
        /// 用户账号
        /// </summary>
        public string user_name { get; set; }
        /// <summary>
        /// 用户昵称
        /// </summary>

        public string nick_name { get; set; }
        /// <summary>
        /// 用户邮箱
        /// </summary>
        public string email { get; set; }
        /// <summary>
        /// 手机号码
        /// </summary>
        public string phonenumber { get; set; }
      
    }
}
