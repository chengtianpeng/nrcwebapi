using System.ComponentModel.DataAnnotations.Schema;
using System.ComponentModel.DataAnnotations;
using SqlSugar;

namespace NRCWebApi.Model
{
    /// <summary>
    /// 短信发送审核记录
    /// </summary>
    [Table("msg_review")]
    public class msg_review
    {
        #region Model

        /// <summary>
        /// 主键
        /// </summary>
        [SugarColumn(IsPrimaryKey = true)]
        public Guid Id { get; set; }

        /// <summary>
        /// 工作流ID
        /// </summary>
        public int WorkID { get; set; }

        /// <summary>
        /// 节点ID
        /// </summary>
        public string NodeID { get; set; }

        /// <summary>
        /// 业务系统（OA,PMS,CRS,HRIS,MMS）
        /// </summary>
        public string SystemName { get; set; }

        /// <summary>
        /// 业务类型(xqjh 需求计划,ssrr 三商入围,ssbg 三商变更)
        /// </summary>
        public string BusinessType { get; set; }

        /// <summary>
        /// 发送人
        /// </summary>
        public string MsgFrom { get; set; }

        /// <summary>
        /// 收件人
        /// </summary>
        public string MsgTo { get; set; }
        /// <summary>
        /// 短信接收人
        /// </summary>
        public string SMSTo { get; set; }
        /// <summary>
        /// 抄送人
        /// </summary>
        public string Cc { get; set; }

        /// <summary>
        /// 密送人
        /// </summary>
        public string Bc { get; set; }

        /// <summary>
        /// 标题
        /// </summary>
        public string MsgTitle { get; set; }

        /// <summary>
        /// 内容  
        /// </summary>
        public string MsgContent { get; set; }

        /// <summary>
        /// 附件ID
        /// </summary>
        public string AttachmentId { get; set; }

        /// <summary>
        /// 发送短信的状态，0 不发送短信 ，1 需要发送短信 2 发送短信成功  3发送失败
        /// </summary>
        public int SendSMSState { get; set; }

        /// <summary>
        /// 发送短信时间
        /// </summary>
        public DateTime? SendSMSTime { get; set; }

        /// <summary>
        /// 发送邮件的状态，0 不发送 ，1 需要发送 2 发送短信成功 3 发送失败
        /// </summary>
        public int SendEmailState { get; set; }

        /// <summary>
        /// 发送邮件时间
        /// </summary>
        public DateTime? SendEmailTime { get; set; }

        /// <summary>
        /// 创建时间
        /// </summary>
        public DateTime CreateTime { get; set; }


        #endregion
    }
}
