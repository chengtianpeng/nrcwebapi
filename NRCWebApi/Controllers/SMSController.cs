using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using Newtonsoft.Json;
using NRCWebApi.Common;
using NRCWebApi.Dto;
using NRCWebApi.Model;
using NRCWebApi.youdu;
using SqlSugar;
using System.DirectoryServices;
using System.Text.RegularExpressions;

// For more information on enabling Web API for empty projects, visit https://go.microsoft.com/fwlink/?LinkID=397860

namespace NRCWebApi.Controllers
{
    /// <summary>
    /// 密码重置接口
    /// </summary>
    [Route("api/[controller]/[action]")]
    [ApiController]
    public class SMSController : ControllerBase
    {

       
        private const int Buin = 14910893; // 请填写企业总机号
        private const string AppId = "ydD0A5A50A30904E06915C9B1D7C643515"; // 请填写AppId
        private const string Address = "172.16.35.66:7080"; // 请填写有度服务器地址
        private const string EncodingaesKey = "xIXJiNE2x3eSZrqBIzr9JskXDUMux5zDi29SaaGX7nk="; // 请填写企业应用EncodingaesKey
         

        private readonly IDictionary<ConnectionKey, ISqlSugarClient> _sqlSugarClients;
        private readonly ILogger<SMSController> _logger;

        /// <summary>
        /// 构造方法
        /// </summary>
        /// <param name="sqlSugarClients"></param>
        /// <param name="logger"></param>
        public SMSController(IDictionary<ConnectionKey, ISqlSugarClient> sqlSugarClients,
            ILogger<SMSController> logger)
        {
            _sqlSugarClients = sqlSugarClients;
            _logger = logger;
        }

        /// <summary>
        /// 测试
        /// </summary>
        /// <returns></returns>
        [HttpGet(Name = "Test")]
        public List<msg_review> Test()
        {
            var db = _sqlSugarClients[ConnectionKey.PMS];//其中Dbconnect就是连接数据库字符串的Key

            var data = db.Queryable<msg_review>().Where(u => u.SystemName == "PMS").ToList();
            _logger.LogInformation("LogInformation test");
            _logger.LogWarning("LogWarning test");
            _logger.LogError("LogError test");
          
            return data;
        }

        #region AD域控制

        /// <summary>
        /// 判断用户是否在AD中
        /// </summary>
        /// <param name="userid">用户AD账户</param>
        /// <returns></returns>
        [HttpGet(Name = "GetUserIsAdd")]
        public ReturnMsg IsUserinAd(string userid)
        {

            userid = userid.Replace(" ", "").ToLower();
            ReturnMsg returnMsg = new ReturnMsg();
            returnMsg.Message = CheckAccount(userid);
            returnMsg.AdSuccess = !returnMsg.Message.Contains("Null");

            if (returnMsg.AdSuccess)
            {

                var db = _sqlSugarClients[ConnectionKey.IT_8_188];
                UCML_CONTACT? first = db.Queryable<UCML_CONTACT>().Where(m => m.HomeAddress == userid).ToList().FirstOrDefault();
                if (first != null)
                {
                    returnMsg.UserName = userid;
                    returnMsg.BindSuccess = true;
                    returnMsg.Email = first.CON_EMAIL_ADDR;
                    returnMsg.PhoneNumber = first.MobilePhone;

                    if (!string.IsNullOrEmpty(returnMsg.Email))
                    {
                        Msgmethod msgmethod = new Msgmethod();
                        msgmethod.label = "Email: " + returnMsg.Email;
                        msgmethod.value = returnMsg.Email;
                        returnMsg.msgmethods.Add(msgmethod);
                    }

                    if (!string.IsNullOrEmpty(returnMsg.PhoneNumber))
                    {
                        Msgmethod msgmethod = new Msgmethod();
                        msgmethod.label = "PhoneNumber: " + returnMsg.PhoneNumber;
                        msgmethod.value = returnMsg.PhoneNumber;
                        returnMsg.msgmethods.Add(msgmethod);
                    }
                }
                else
                {
                    returnMsg.BindSuccess = false;
                    returnMsg.Message = "";
                }
            }


            return returnMsg;
        }

        /// <summary>
        /// 检查验证码是否正确
        /// </summary>
        /// <param name="vicode"></param>
        /// <param name="guid"></param>
        /// <param name="pwd"></param>
        /// <returns></returns>
        [HttpGet(Name = "GetIsVicCodeOk")]
        public ReturnMsg IsVicCodeOk(string vicode, string guid, string pwd)
        {
            ReturnMsg returnMsg = new ReturnMsg();
            string data = System.Text.RegularExpressions.Regex.Replace(vicode, @"[^0-9]+", "");
            returnMsg.Message = data;

            var db = _sqlSugarClients[ConnectionKey.IT_8_188];
            IT_MMZH? first = db.Queryable<IT_MMZH>().Where(m => m.IT_MMZHOID == Guid.Parse(guid) && m.XXNR.Contains(vicode) && m.CLZT == 2).ToList().FirstOrDefault();
            if (first != null && !string.IsNullOrEmpty(first.DHHM))
            {
                if (IsPasswordOk(pwd))
                {
                    string userac = first.DLZH;
                    Lock(userac);
                    string msg = UpdateADAccount(userac, pwd);

                    if (!msg.Contains("ERROR"))
                    {
                        returnMsg.AdSuccess = true;
                        returnMsg.Message = "Password changed successfully! 密码修改成功!";
                        Console.WriteLine(userac + ":" + pwd);


                        first.CLSJ = DateTime.Now;
                        first.CLZT = 3;
                        db.Updateable<IT_MMZH>(first).ExecuteCommand();

                    }
                    else
                    {
                        returnMsg.AdSuccess = false;
                        returnMsg.Message = msg;
                    }
                }
                else
                {
                    returnMsg.AdSuccess = false;
                    returnMsg.Message = "The password does not meet the complexity requirements ! 密码不满足复杂度要求 ";
                }
            }
            else
            {
                returnMsg.AdSuccess = false;
                returnMsg.Message = "The verification code is incorrect or expired ! 验证码错误或过期";
            }
            return returnMsg;
        }



        /// <summary>
        /// AD账户查询
        /// </summary>
        /// <param name="Domain"></param>
        /// <param name="UserAccount"></param>
        /// <param name="UserPassWord"></param>
        /// <returns></returns>
        private string CheckAccount(string acc)
        {
            string Domain = "tdnrc.com";
            string UserAccount = "administrator";
            string UserPassWord = "Stv7#1552770";
            string ReturnValue = string.Empty;
            string IsLock = "";
            try
            {
                DirectoryEntry entry = new DirectoryEntry("LDAP://" + Domain, UserAccount, UserPassWord, AuthenticationTypes.Secure);
                DirectorySearcher mySearcher = new DirectorySearcher(entry);
                mySearcher.Filter = ("(&(objectClass=user)(sAMAccountName=" + acc + "))");
                SearchResult searchResult = mySearcher.FindOne();
                ReturnValue = "Null";
                if (searchResult != null)
                {
                    DirectoryEntry userEntry = searchResult.GetDirectoryEntry();
                    if (userEntry != null)
                    {
                        IsLock = ((bool)userEntry.InvokeGet("IsAccountLocked")).ToString();
                    }
                    ReturnValue = searchResult.Path;
                }

            }
            catch (Exception ex)
            {
                ReturnValue = "Exception：" + ex.Message;
            }
            return ReturnValue + "IsAccountLocked:" + IsLock;
        }

        private string Lock(string acc)
        {
            string Domain = "tdnrc.com";
            string UserAccount = "administrator";
            string UserPassWord = "Stv7#1552770";

            string ReturnValue = string.Empty;
            try
            {
                //throw new InvalidPluginExecutionException(Domain + "|==|" + UserAccount + "|==|" + UserPassWord);
                DirectoryEntry entry = new DirectoryEntry("LDAP://" + Domain, UserAccount, UserPassWord, AuthenticationTypes.Secure);
                DirectorySearcher mySearcher = new DirectorySearcher(entry);
                mySearcher.Filter = ("(&(objectClass=user)(sAMAccountName=" + acc + "))");
                SearchResult searchResult = mySearcher.FindOne();
                DirectoryEntry userEntry = searchResult.GetDirectoryEntry();
                userEntry.InvokeSet("IsAccountLocked", new object[] { false });
                userEntry.CommitChanges();
                ReturnValue = "Account UnLocked";
            }
            catch (Exception ex)
            {
                ReturnValue = "ERROR：" + ex.Message;
            }
            return ReturnValue;
        }

        private string UpdateADAccount(string UserAccount, string UserNewPassWord)
        {

            string Domain = "10.162.128.74"; //tdnrc.com
            string adinUserAccount = "administrator";
            string UserPassWord = "Stv7#1552770";

            //反馈内容、
            string ReturnString = string.Empty;

            try
            {
                DirectoryEntry entry = new DirectoryEntry("LDAP://" + Domain, adinUserAccount, UserPassWord, AuthenticationTypes.Secure);
                DirectorySearcher mySearcher = new DirectorySearcher(entry);
                mySearcher.Filter = ("(&(objectClass=user)(sAMAccountName=" + UserAccount + "))");
                SearchResult searchResult = mySearcher.FindOne();
                DirectoryEntry userEntry = searchResult.GetDirectoryEntry();
                var data = userEntry.Invoke("SetPassword", new object[] { UserNewPassWord });
                ReturnString = UserNewPassWord;
                // System.IO.File.AppendAllText(AppDomain.CurrentDomain.BaseDirectory + "\\log.txt", UserAccount + ":    " + UserNewPassWord + Environment.NewLine);
            }
            catch (Exception ex)
            {
                ReturnString = "ERROR：" + ex.Message;
            }


            return ReturnString;
        }


        #endregion

        #region   有度密码重置


        [HttpGet(Name = "IsYouduVicCodeOk")]
        public async Task<ReturnMsg> IsYouduVicCodeOk(string vicode, string guid, string pwd)
        {
            ReturnMsg returnMsg = new ReturnMsg();
            string data = System.Text.RegularExpressions.Regex.Replace(vicode, @"[^0-9]+", "");
            returnMsg.Message = data;

            try
            {
                var db = _sqlSugarClients[ConnectionKey.IT_8_188];
                IT_MMZH? first = db.Queryable<IT_MMZH>().Where(m => m.IT_MMZHOID == Guid.Parse(guid) && m.XXNR.Contains(vicode) && m.CLZT == 2).ToList().FirstOrDefault();
                if (first != null && !string.IsNullOrEmpty(first.DHHM))
                {
                    if (IsPasswordOk(pwd))
                    {
                        string userac = first.DLZH;

                        AppClient m_appClient = new AppClient(Address, Buin, AppId, EncodingaesKey);
                        bool secc = await m_appClient.ChangePWD(new MessageSetauth(userac, pwd));

                        if (secc)
                        {
                            returnMsg.AdSuccess = true;
                            returnMsg.Message = "Password changed successfully! 密码修改成功!";
                            Console.WriteLine("修改成功|" + userac + ":" + pwd);


                            first.CLSJ = DateTime.Now;
                            first.CLZT = 3;
                            db.Updateable<IT_MMZH>(first).ExecuteCommand();
                        }
                        else
                        {
                            Console.WriteLine("修改失败|" + userac + ":" + pwd);
                            returnMsg.AdSuccess = false;
                            returnMsg.Message = "密码修改失败请联系IT部";
                        }
                    }
                    else
                    {
                        Console.WriteLine("复杂度验证不通过|" + guid + ":" + pwd);
                        returnMsg.AdSuccess = false;
                        returnMsg.Message = "The password does not meet the complexity requirements ! 密码不满足复杂度要求 ";
                    }
                }
                else
                {
                    returnMsg.AdSuccess = false;
                    returnMsg.Message = "The verification code is incorrect or expired ! 验证码错误或过期";
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("有度密码重置异常|" + ex.Message);

            }
            return returnMsg;
        }


        [HttpGet(Name = "IsUserinYoudu")]
        public async Task<ReturnMsg> IsUserinYoudu(string userloginid)
        {
            userloginid = userloginid.Replace(" ", "").ToLower();
            ReturnMsg returnMsg = new ReturnMsg();
            string empid = "";

            AppClient m_appClient = new AppClient(Address, Buin, AppId, EncodingaesKey);
            bool exist = await m_appClient.IsYDUserExist(userloginid);
            if (exist)
            {
                empid = userloginid;
                returnMsg.Message = empid;
                returnMsg.AdSuccess = true;


                var db = _sqlSugarClients[ConnectionKey.IT_8_188];
                UCML_CONTACT? first = db.Queryable<UCML_CONTACT>().Where(m => m.QQNO == empid).ToList().FirstOrDefault();
                if (first != null)
                {
                    returnMsg.UserName = empid;
                    returnMsg.BindSuccess = true;
                    returnMsg.Email = first.CON_EMAIL_ADDR;
                    returnMsg.PhoneNumber = first.MobilePhone;

                    if (!string.IsNullOrEmpty(returnMsg.Email))
                    {
                        Msgmethod msgmethod = new Msgmethod();
                        msgmethod.label = "Email: " + returnMsg.Email;
                        msgmethod.value = returnMsg.Email;
                        returnMsg.msgmethods.Add(msgmethod);
                    }

                    if (!string.IsNullOrEmpty(returnMsg.PhoneNumber))
                    {
                        Msgmethod msgmethod = new Msgmethod();
                        msgmethod.label = "PhoneNumber: " + returnMsg.PhoneNumber;
                        msgmethod.value = returnMsg.PhoneNumber;
                        returnMsg.msgmethods.Add(msgmethod);
                    }
                }
                else
                {
                    returnMsg.BindSuccess = false;
                    returnMsg.Message = "";
                }

            }
            else
            {
                returnMsg.AdSuccess = false;
                returnMsg.Message = "The user name does not exist 用户名不存在，请重新输入";
            }

            return returnMsg;
        }





        #endregion

        #region OA

        /// <summary>
        /// 判断用户是否在OA中
        /// </summary>
        /// <param name="userid">用户OA账户</param>
        /// <returns></returns>
        [HttpGet(Name = "IsUserinOA")]
        public ReturnMsg IsUserinOA(string userloginid)
        {
            userloginid = userloginid.Replace(" ", "").ToLower();
            ReturnMsg returnMsg = new ReturnMsg();
            string empid = "";
            string email = "";
            string phone = "";
            long userid = CheckOAAccount(userloginid, out empid, out email, out phone);
            if (userid > 0)
            {
                returnMsg.Message = empid;
                returnMsg.AdSuccess = true;

                returnMsg.UserName = empid;
                returnMsg.BindSuccess = true;
                returnMsg.Email = email;
                returnMsg.PhoneNumber = phone;

                if (!string.IsNullOrEmpty(returnMsg.Email))
                {
                    Msgmethod msgmethod = new Msgmethod();
                    msgmethod.label = "Email: " + returnMsg.Email;
                    msgmethod.value = returnMsg.Email;
                    returnMsg.msgmethods.Add(msgmethod);
                }

                if (!string.IsNullOrEmpty(returnMsg.PhoneNumber))
                {
                    Msgmethod msgmethod = new Msgmethod();
                    msgmethod.label = "PhoneNumber: " + returnMsg.PhoneNumber;
                    msgmethod.value = returnMsg.PhoneNumber;
                    returnMsg.msgmethods.Add(msgmethod);
                }

            }
            else
            {
                returnMsg.AdSuccess = false;
                returnMsg.Message = "The user name does not exist 用户名不存在，请重新输入";
            }


            return returnMsg;
        }


        private long CheckOAAccount(string acc, out string empid, out string email, out string phone)
        {
            long userid = -1;
            empid = "";
            email = "";
            phone = "";

            var db = _sqlSugarClients[ConnectionKey.OA];
            sys_user_info? sys_user_infoFirst = db.Queryable<sys_user_info>().Where(m => m.user_name == acc || m.email == (acc + "@tdnrc.com")).ToList().FirstOrDefault();

            if (sys_user_infoFirst != null && !string.IsNullOrEmpty(sys_user_infoFirst.user_name))
            {
                userid = sys_user_infoFirst.user_id;
                empid = sys_user_infoFirst.user_name;
                email = sys_user_infoFirst.email;
                phone = sys_user_infoFirst.phonenumber;
            }

            return userid;


        }

        /// <summary>
        /// 检查验证码是否正确
        /// </summary>
        /// <param name="vicode"></param>
        /// <param name="guid"></param>
        /// <param name="pwd"></param>
        /// <returns></returns>
        [HttpGet(Name = "IsOAVicCodeOk")]
        public ReturnMsg IsOAVicCodeOk(string vicode, string guid, string pwd)
        {
            ReturnMsg returnMsg = new ReturnMsg();
            string data = System.Text.RegularExpressions.Regex.Replace(vicode, @"[^0-9]+", "");
            returnMsg.Message = data;

            var db = _sqlSugarClients[ConnectionKey.IT_8_188];
            IT_MMZH? IT_MMZHFirst = db.Queryable<IT_MMZH>().Where(m => m.IT_MMZHOID == Guid.Parse(guid) && m.XXNR.Contains(vicode) && m.CLZT == 2).ToList().FirstOrDefault();
            if (IT_MMZHFirst != null && !string.IsNullOrEmpty(IT_MMZHFirst.DHHM))
            {
                if (IsPasswordOk(pwd))
                {
                    string userac = IT_MMZHFirst.DLZH;

                    WebRequestHelper wh = new WebRequestHelper();
                    ResetUserDto use = new ResetUserDto()
                    {
                        username = userac,
                        newPassword = pwd
                    };

                    string pdata = JsonConvert.SerializeObject(use);

                    string resBool = wh.HttpPost("http://oa.tdnrc.com/api/system/user/profile/updateUserPwd", pdata);

                    bool resB = bool.Parse(resBool);
                    //  int res = SqlHelperPMS.ExecuteNonQuery(sql, spup.ToArray());


                    if (resB)
                    {
                        returnMsg.AdSuccess = true;
                        returnMsg.Message = "Password changed successfully! 密码修改成功!";

                        IT_MMZHFirst.CLSJ = DateTime.Now;
                        IT_MMZHFirst.CLZT = 3;
                        db.Updateable(IT_MMZHFirst).ExecuteCommand();

                    }
                    else
                    {
                        Console.WriteLine("修改失败|" + userac + ":" + pwd);
                        returnMsg.AdSuccess = false;
                        returnMsg.Message = "密码修改失败请联系IT部";
                    }
                }
                else
                {
                    Console.WriteLine("复杂度验证不通过|" + guid + ":" + pwd);
                    returnMsg.AdSuccess = false;
                    returnMsg.Message = "The password does not meet the complexity requirements ! 密码不满足复杂度要求 ";
                }
            }
            else
            {
                returnMsg.AdSuccess = false;
                returnMsg.Message = "The verification code is incorrect or expired ! 验证码错误或过期";
            }

            return returnMsg;
        }


        #endregion

        #region PMS

        /// <summary>
        /// 判断用户是否在PMS中
        /// </summary>
        /// <param name="userid">用户PMS账户</param>
        /// <returns></returns>
        [HttpGet(Name = "IsUserinPMS")]
        public ReturnMsg IsUserinPMS(string userloginid)
        {
            userloginid = userloginid.Replace(" ", "").ToLower();
            ReturnMsg returnMsg = new ReturnMsg();
            string empid = "";
            string email = "";
            string phone = "";
            int userid = CheckPMSAccount(userloginid, out empid, out email, out phone);
            if (userid > 0)
            {
                returnMsg.Message = empid;
                returnMsg.AdSuccess = true;


                var db = _sqlSugarClients[ConnectionKey.IT_8_188];
                UCML_CONTACT? first = db.Queryable<UCML_CONTACT>().Where(m => m.QQNO == empid).ToList().FirstOrDefault();
                if (first != null)
                {
                    returnMsg.UserName = empid;
                    returnMsg.BindSuccess = true;
                    returnMsg.Email = first.CON_EMAIL_ADDR;
                    returnMsg.PhoneNumber = first.MobilePhone;

                    if (!string.IsNullOrEmpty(returnMsg.Email))
                    {
                        Msgmethod msgmethod = new Msgmethod();
                        msgmethod.label = "Email: " + returnMsg.Email;
                        msgmethod.value = returnMsg.Email;
                        returnMsg.msgmethods.Add(msgmethod);
                    }

                    if (!string.IsNullOrEmpty(returnMsg.PhoneNumber))
                    {
                        Msgmethod msgmethod = new Msgmethod();
                        msgmethod.label = "PhoneNumber: " + returnMsg.PhoneNumber;
                        msgmethod.value = returnMsg.PhoneNumber;
                        returnMsg.msgmethods.Add(msgmethod);
                    }
                }
                else
                {
                    returnMsg.BindSuccess = false;
                    returnMsg.Message = "";
                }
            }
            else
            {
                returnMsg.AdSuccess = false;
                returnMsg.Message = "The user name does not exist 用户名不存在，请重新输入";
            }

            return returnMsg;
        }


        private int CheckPMSAccount(string acc, out string empid, out string email, out string phone)
        {
            int userid = -1;
            empid = "";
            email = "";
            phone = "";

            string sql = "select userid,LoginName,Email,Cellphone from ss_User where LoginName=@account or lower(replace(UserEName,' ', ''))=@account";

            var db = _sqlSugarClients[ConnectionKey.PMS];
            List<SugarParameter> parameters = new List<SugarParameter>();
            parameters.Add(new SugarParameter("account", acc.Replace(" ", "").Trim().ToLower()));
            ss_User? first = db.Ado.SqlQuery<ss_User>(sql, parameters).FirstOrDefault();
            if (first != null && !string.IsNullOrEmpty(first.LoginName))
            {
                userid = first.UserId;
                empid = first.LoginName;
                email = first.Email;
                phone = first.Cellphone;
            }

            return userid;


        }



        /// <summary>
        /// 检查验证码是否正确
        /// </summary>
        /// <param name="vicode"></param>
        /// <param name="guid"></param>
        /// <param name="pwd"></param>
        /// <returns></returns>
        [HttpGet(Name = "IsPmsVicCodeOk")]
        public ReturnMsg IsPmsVicCodeOk(string vicode, string guid, string pwd)
        {
            ReturnMsg returnMsg = new ReturnMsg();
            string data = System.Text.RegularExpressions.Regex.Replace(vicode, @"[^0-9]+", "");
            returnMsg.Message = data;

            var db = _sqlSugarClients[ConnectionKey.IT_8_188];
            IT_MMZH? first = db.Queryable<IT_MMZH>().Where(m => m.IT_MMZHOID == Guid.Parse(guid) && m.XXNR.Contains(vicode) && m.CLZT == 2).ToList().FirstOrDefault();
            if (first != null && !string.IsNullOrEmpty(first.DHHM))
            {
                if (IsPasswordOk(pwd))
                {
                    string sql = "update ss_User set Password=@pwd,LastPasswordModificationDate=SYSDATETIME(),PasswordExpirationDate=DATEADD(M,3,SYSDATETIME()) where LoginName=@usid";
                    List<SugarParameter> spup = new List<SugarParameter>();
                    spup.Add(new SugarParameter("pwd", SecurityHelper.ToMD5String(pwd)));
                    spup.Add(new SugarParameter("usid", first.DLZH));

                    var dbPMS = _sqlSugarClients[ConnectionKey.PMS];
                    int res = dbPMS.Ado.ExecuteCommand(sql, spup);

                    if (res > 0)
                    {
                        returnMsg.AdSuccess = true;
                        returnMsg.Message = "Password changed successfully! 密码修改成功!";
                        Console.WriteLine("修改成功|" + first.DLZH + ":" + pwd);

                        first.CLSJ = DateTime.Now;
                        first.CLZT = 3;
                        db.Updateable<IT_MMZH>(first).ExecuteCommand();

                    }
                    else
                    {
                        Console.WriteLine("修改失败|" + first.DLZH + ":" + pwd);
                        returnMsg.AdSuccess = false;
                        returnMsg.Message = "密码修改失败请联系IT部";
                    }
                }
                else
                {
                    Console.WriteLine("复杂度验证不通过|" + guid + ":" + pwd);
                    returnMsg.AdSuccess = false;
                    returnMsg.Message = "The password does not meet the complexity requirements ! 密码不满足复杂度要求 ";
                }
            }
            else
            {
                returnMsg.AdSuccess = false;
                returnMsg.Message = "The verification code is incorrect or expired ! 验证码错误或过期";
            }
            return returnMsg;
        }



        #endregion

        #region CRS

        /// <summary>
        /// 判断用户是否在PMS中
        /// </summary>
        /// <param name="userid">用户PMS账户</param>
        /// <returns></returns>
        [HttpGet(Name = "IsUserinCRS")]
        public ReturnMsg IsUserinCRS(string userloginid)
        {
            userloginid = userloginid.Replace(" ", "").ToLower();
            ReturnMsg returnMsg = new ReturnMsg();
            string empid = "";
            string email = "";
            string phone = "";
            int userid = CheckCRSAccount(userloginid, out empid, out email, out phone);
            if (userid > 0)
            {
                returnMsg.Message = empid;
                returnMsg.AdSuccess = true;


                var db = _sqlSugarClients[ConnectionKey.IT_8_188];
                UCML_CONTACT? first = db.Queryable<UCML_CONTACT>().Where(m => m.QQNO == empid).ToList().FirstOrDefault();
                if (first != null)
                {
                    returnMsg.UserName = empid;
                    returnMsg.BindSuccess = true;
                    returnMsg.Email = first.CON_EMAIL_ADDR;
                    returnMsg.PhoneNumber = first.MobilePhone;

                    if (!string.IsNullOrEmpty(returnMsg.Email))
                    {
                        Msgmethod msgmethod = new Msgmethod();
                        msgmethod.label = "Email: " + returnMsg.Email;
                        msgmethod.value = returnMsg.Email;
                        returnMsg.msgmethods.Add(msgmethod);
                    }

                    if (!string.IsNullOrEmpty(returnMsg.PhoneNumber))
                    {
                        Msgmethod msgmethod = new Msgmethod();
                        msgmethod.label = "PhoneNumber: " + returnMsg.PhoneNumber;
                        msgmethod.value = returnMsg.PhoneNumber;
                        returnMsg.msgmethods.Add(msgmethod);
                    }
                }
                else
                {
                    returnMsg.BindSuccess = false;
                    returnMsg.Message = "";
                }
            }
            else
            {
                returnMsg.AdSuccess = false;
                returnMsg.Message = "The user name does not exist 用户名不存在，请重新输入";
            }

            return returnMsg;
        }


        private int CheckCRSAccount(string acc, out string empid, out string email, out string phone)
        {

            int userid = -1;
            empid = "";
            email = "";
            phone = "";

            string sql = "select userid,LoginName,Email,Cellphone from ss_User where LoginName=@account or lower(replace(UserEName,' ', ''))=@account";

            var db = _sqlSugarClients[ConnectionKey.CRS];
            List<SugarParameter> parameters = new List<SugarParameter>();
            parameters.Add(new SugarParameter("account", acc.Replace(" ", "").Trim().ToLower()));
            ss_User? first = db.Ado.SqlQuery<ss_User>(sql, parameters).FirstOrDefault();
            if (first != null && !string.IsNullOrEmpty(first.LoginName))
            {
                userid = first.UserId;
                empid = first.LoginName;
                email = first.Email;
                phone = first.Cellphone;
            }

            return userid;

        }

        /// <summary>
        /// 检查验证码是否正确
        /// </summary>
        /// <param name="vicode"></param>
        /// <param name="guid"></param>
        /// <param name="pwd"></param>
        /// <returns></returns>
        [HttpGet(Name = "IsCrsVicCodeOk")]
        public ReturnMsg IsCrsVicCodeOk(string vicode, string guid, string pwd)
        {
            ReturnMsg returnMsg = new ReturnMsg();
            string data = System.Text.RegularExpressions.Regex.Replace(vicode, @"[^0-9]+", "");
            returnMsg.Message = data;

            var db = _sqlSugarClients[ConnectionKey.IT_8_188];
            IT_MMZH? first = db.Queryable<IT_MMZH>().Where(m => m.IT_MMZHOID == Guid.Parse(guid) && m.XXNR.Contains(vicode) && m.CLZT == 2).ToList().FirstOrDefault();
            if (first != null && !string.IsNullOrEmpty(first.DHHM))
            {
                if (IsPasswordOk(pwd))
                {
                    string sql = "update ss_User set Password=@pwd,LastPasswordModificationDate=SYSDATETIME(),PasswordExpirationDate=DATEADD(M,3,SYSDATETIME()) where LoginName=@usid";
                    List<SugarParameter> spup = new List<SugarParameter>();
                    spup.Add(new SugarParameter("pwd", SecurityHelper.ToMD5String(pwd)));
                    spup.Add(new SugarParameter("usid", first.DLZH));

                    var dbPMS = _sqlSugarClients[ConnectionKey.CRS];
                    int res = dbPMS.Ado.ExecuteCommand(sql, spup);

                    if (res > 0)
                    {
                        returnMsg.AdSuccess = true;
                        returnMsg.Message = "Password changed successfully! 密码修改成功!";
                        Console.WriteLine("修改成功|" + first.DLZH + ":" + pwd);

                        first.CLSJ = DateTime.Now;
                        first.CLZT = 3;
                        db.Updateable<IT_MMZH>(first).ExecuteCommand();

                    }
                    else
                    {
                        Console.WriteLine("修改失败|" + first.DLZH + ":" + pwd);
                        returnMsg.AdSuccess = false;
                        returnMsg.Message = "密码修改失败请联系IT部";
                    }
                }
                else
                {
                    Console.WriteLine("复杂度验证不通过|" + guid + ":" + pwd);
                    returnMsg.AdSuccess = false;
                    returnMsg.Message = "The password does not meet the complexity requirements ! 密码不满足复杂度要求 ";
                }
            }
            else
            {
                returnMsg.AdSuccess = false;
                returnMsg.Message = "The verification code is incorrect or expired ! 验证码错误或过期";
            }
            return returnMsg;
        }


        #endregion

        #region 公共方法


        /// <summary>
        /// 验证码
        /// </summary>
        /// <param name="msgtype"></param>
        /// <param name="username"></param>
        /// <param name="systemtype"></param>
        /// <returns></returns>
        [HttpGet(Name = "GetVaverInfo")]
        public VaverifyInfoDto GetVaverInfo(string msgtype, string username, string systemtype)
        {
            VaverifyInfoDto vaverifyInfo = new VaverifyInfoDto();
            vaverifyInfo.VaverifyId = Guid.NewGuid().ToString();
            vaverifyInfo.AdSuccess = false;

            string dhhm = "";
            List<SqlParameter> sqlParams = new List<SqlParameter>();

            if (systemtype == "AD")
            {
                var db = _sqlSugarClients[ConnectionKey.IT_8_188];
                UCML_CONTACT? first = db.Queryable<UCML_CONTACT>().Where(m => m.HomeAddress == username).ToList().FirstOrDefault();
                if (first != null)
                {
                    dhhm = first.CON_EMAIL_ADDR;
                    if (msgtype == "TELMSG")
                    {
                        dhhm = first.MobilePhone;
                    }
                }
            }

            else if (systemtype == "OA")
            {
                var db = _sqlSugarClients[ConnectionKey.OA];
                sys_user_info? first = db.Queryable<sys_user_info>().Where(m => m.user_name == username.Trim() || m.email == (username.Trim() + "@tdnrc.com")).ToList().FirstOrDefault();

                if (first != null && !string.IsNullOrEmpty(first.user_name))
                {
                    dhhm = first.email;
                    if (msgtype == "TELMSG")
                    {
                        dhhm = first.phonenumber;
                    }
                }
            }

            else
            {
                var db = _sqlSugarClients[ConnectionKey.IT_8_188];
                UCML_CONTACT? first = db.Queryable<UCML_CONTACT>().Where(m => m.QQNO == username).ToList().FirstOrDefault();
                if (first != null)
                {
                    dhhm = first.CON_EMAIL_ADDR;
                    if (msgtype == "TELMSG")
                    {
                        dhhm = first.MobilePhone;
                    }
                }
            }

            if (!string.IsNullOrEmpty(dhhm))
            {
                Random rd = new Random();
                string info = "The verification code is:" + rd.Next(1000, 9999).ToString();


                // SendRecMSG(dhhm, info, username, msgtype, systemtype, vaverifyInfo.VaverifyId);
                #region 发送短信和邮件 ，存储数据库中
                string zt = "0";
                if (msgtype == "EMAIL")
                {

                    msg_agent msgagent = new msg_agent()
                    {
                        msgid = Guid.NewGuid().ToString(),
                        systemname = "ITSM",
                        msgfrom = "itsm@tdnrc.com",
                        msgto = dhhm,
                        msgtitle = "Password Reset",
                        msgconten = info,
                        attachmentid = "0",
                        msgstate = 1,
                        createtime = DateTime.Now

                    };
                    var db = _sqlSugarClients[ConnectionKey.PMS];
                    db.Insertable(msgagent).ExecuteCommand();
                    zt = "2";
                }

                IT_MMZH iT_MMZH = new IT_MMZH()
                {
                    IT_MMZHOID = Guid.Parse(vaverifyInfo.VaverifyId),
                    DLZH = username,
                    YWXTMC = systemtype,
                    DHHM = dhhm,
                    ZHFS = msgtype,
                    CLZT = int.Parse(zt),
                    TJSJ = DateTime.Now,
                    XXNR = info
                };
                var dbIT = _sqlSugarClients[ConnectionKey.IT_8_188];
                dbIT.Insertable(iT_MMZH).ExecuteCommand();


                #endregion

                vaverifyInfo.AdSuccess = true;
                Console.WriteLine(username + "|" + msgtype + "|" + systemtype + "|" + "获取验证码成功");
            }

            return vaverifyInfo;
        }

        /// <summary>
        /// 验证密码强度
        /// </summary>
        /// <param name="strpassword"></param>
        /// <returns></returns>
        private bool IsPasswordOk(string strpassword)
        {
            string pattern = @"^(?=.*[a-z])(?=.*[A-Z])(?=.*\d)(?=.*[!@#$%^&*()+=])[a-zA-Z\d!@#$%^&*()+=]{8,18}$";
            var regex = new Regex(pattern);
            return regex.IsMatch(strpassword);
        }

        #endregion

    }
}
