using SqlSugar;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace NRCWebApi.Model
{
    /// <summary>
    /// 8.188
    /// </summary>
    [Table("IT_MMZH")]
    public class IT_MMZH
    {
        [SugarColumn(IsPrimaryKey = true)]
        public Guid IT_MMZHOID { get; set; }
        public string DLZH { get; set; }
        public string YWXTMC { get; set; }
        public string DHHM { get; set; }
        public string YGXM { get; set; }
        public string YGBH { get; set; }
        public string ZHFS { get; set; }
        public int CLZT { get; set; }
        public DateTime TJSJ { get; set; }
        public DateTime CLSJ { get; set; }
        public string XXNR { get; set; }
        public string WORKID { get; set; }
    }
}
