using SqlSugar;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace NRCWebApi.Model
{
    /// <summary>
    /// 8.188
    /// </summary>
    [Table("UCML_CONTACT")]
    public class UCML_CONTACT
    {
        [SugarColumn(IsPrimaryKey = true)]
        public Guid UCML_CONTACTOID { get; set; }
        public string PersonName { get; set; }
        public string WORK_PH_NUM { get; set; }
        public string CON_JOB_TITLE { get; set; }
        public string CON_EMAIL_ADDR { get; set; }
        public string HOME_PH_NUM { get; set; }
        public string MobilePhone { get; set; }
        public string HomeAddress { get; set; }
        public string QQNO { get; set; }
    }
}
