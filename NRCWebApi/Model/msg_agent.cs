using SqlSugar;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace NRCWebApi.Model
{
    /// <summary>
    /// 8.99
    /// </summary>
    [Table("msg_agent")]
    public class msg_agent
    {
        /// <summary>
        /// 主键
        /// </summary>
        [SugarColumn(IsPrimaryKey = true)]
        public string msgid { get; set; }

        public string systemname { get; set; }
        public string msgfrom { get; set; }
        public string msgto { get; set; }
        public string msgtitle { get; set; }
        public string msgconten { get; set; }
        public string attachmentid { get; set; }
        public int msgstate { get; set; }
        public string msglog { get; set; }

        public DateTime createtime { get; set; }
        public DateTime finishtime { get; set; }
        public string r1 { get; set; }
        public string r2 { get; set; }
        public string r3 { get; set; }
        public string r4 { get; set; }
        public string r5 { get; set; }
        public string cc { get; set; }
        public string bc { get; set; }
    }
}
