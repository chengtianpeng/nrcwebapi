using AutoMapper;
using Microsoft.AspNetCore.DataProtection.KeyManagement;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Caching.Memory;
using NPOI.SS.Formula;
using NRCWebApi.Common;
using NRCWebApi.Dto;
using NRCWebApi.Model;
using SqlSugar;

namespace NRCWebApi.Controllers
{
    /// <summary>
    /// IT操作类
    /// </summary>
    [Route("api/[controller]/[Action]")]
    [ApiController]
    public class ITController : ControllerBase
    {
        private readonly IDictionary<ConnectionKey, ISqlSugarClient> _sqlSugarClients;
        private readonly ILogger<HRController> _logger;
        private readonly IMemoryCache _cache;
        private readonly IMapper _mapper;

        private readonly string corpid = "ww80c2aa9931887ece"; //企业ID
        private readonly string corpsecret = "IOrp-xmtT1eTTidIhacdVPMy2gU8odl8mm-br_dB9NZBuUQoZ-zH6h-HWTn7NwxG"; //通讯录应用密码
        private readonly string gettokenUrl = "https://api.exmail.qq.com/cgi-bin/gettoken";  //获取token
        private readonly string gettokenkeyCache = "emailtoken";
        private readonly string userupdateUrl = "https://api.exmail.qq.com/cgi-bin/user/update";  //修改邮箱密码
        private readonly string permissionPass = "moss@55.cn";  //校验的密码

        /// <summary>
        /// 构造方法
        /// </summary>
        /// <param name="sqlSugarClients"></param>
        /// <param name="logger"></param>
        /// <param name="cache"></param>
        public ITController(IDictionary<ConnectionKey, ISqlSugarClient> sqlSugarClients,
            ILogger<HRController> logger,
            IMemoryCache cache,
            IMapper mapper)
        {
            _sqlSugarClients = sqlSugarClients;
            _logger = logger;
            _cache = cache;
            _mapper = mapper;
        }

        /// <summary>
        /// 修改Email密码(绑定了微信或手机无法修改密码)
        /// </summary>
        /// <param name="request"></param>
        /// <returns></returns>
        [HttpGet(Name = "UpdateEmailPassword")]
        public async Task<IActionResult> UpdateEmailPassword([FromQuery] UpdateEmailPasswordRequestDto request)
        {
            if (request == null)
            {
                return BadRequest("请求参数不能为空");
            }

            if (request.PermissionPass != permissionPass)
            {
                return BadRequest("校验的密码不正确，请咨询IT部");
            }

            string cachetoken;
            //获取cache
            if (!_cache.TryGetValue(gettokenkeyCache, out cachetoken))
            {
                //没获取到，则分配一个值
                Dictionary<string, string> tokkenPara = new Dictionary<string, string>();
                tokkenPara.Add("corpid", corpid);
                tokkenPara.Add(key: "corpsecret", corpsecret);
                UpdateEmailPasswordResponseDto updateEmailPasswordResponseDto = await HttpHelper.DoGet<UpdateEmailPasswordResponseDto>(gettokenUrl, tokkenPara);
                if (updateEmailPasswordResponseDto != null && updateEmailPasswordResponseDto.errcode == 0 && !string.IsNullOrEmpty(updateEmailPasswordResponseDto.access_token))
                {
                    cachetoken = updateEmailPasswordResponseDto.access_token;

                    //设置cache策略
                    var cacheEntryOptions = new MemoryCacheEntryOptions()
                    //设置3s的滑动过期时间（3s内如果有请求到，则自动再延长3s）
                        .SetSlidingExpiration(TimeSpan.FromSeconds(7000));
                    //保存到cache
                    _cache.Set(gettokenkeyCache, cachetoken, cacheEntryOptions);

                }

            }

            Dictionary<string, string> updatePara = new Dictionary<string, string>();
            updatePara.Add("userid", request.EmailAddriss);
            updatePara.Add("password", request.Password);
            string reuqestUrl = $"{userupdateUrl}?access_token={cachetoken}";
            EmailBaseResponseDto emailBaseResponseDto = await HttpHelper.DoPost<EmailBaseResponseDto>(reuqestUrl, updatePara);

            return Ok(emailBaseResponseDto);
        }


        /// <summary>
        /// 手动发送邮箱
        /// </summary>
        /// <param name="request"></param>
        /// <returns></returns>
        [HttpGet(Name = "HandSendEmail")]
        public IActionResult HandSendEmail([FromQuery] HandSendEmailRequestDto request)
        {
            if (request == null)
            {
                return BadRequest("请求参数不能为空");
            }
            if (request.PermissionPass != permissionPass)
            {
                return BadRequest("校验的密码不正确，请咨询IT部");
            }

            msg_review model = _mapper.Map<msg_review>(request);
            model.Id = Guid.NewGuid();
            model.WorkID = 0;
            model.NodeID = "0";
            model.SystemName = request.MsgFrom.ToLower().Contains("pms") ? "PMS" : "IT";
            model.BusinessType = "Hand";
            model.SendEmailState = 1;
            model.CreateTime = DateTime.Now;

            var db = _sqlSugarClients[ConnectionKey.PMS];//其中Dbconnect就是连接数据库字符串的Key
            var data = db.Insertable<msg_review>(model).ExecuteCommand();
            return Ok(data + "执行成功");
        }
    }
}
