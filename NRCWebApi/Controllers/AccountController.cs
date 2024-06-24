using AutoMapper;
using Microsoft.AspNetCore.Authentication;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.IdentityModel.Tokens;
using NPOI.OpenXmlFormats.Spreadsheet;
using NPOI.SS.Formula.Functions;
using NRCWebApi.Common;
using NRCWebApi.Dto.appsettings;
using NRCWebApi.Dto.test;
using NRCWebApi.Filter;
using NRCWebApi.Model;
using SqlSugar;
using StackExchange.Redis;
using System.IdentityModel.Tokens.Jwt;
using System.Security.Claims;
using System.Text;

namespace NRCWebApi.Controllers
{
    /// <summary>
    ///Account操作控制器
    /// </summary>
    [Route("api/[controller]/[Action]")]
    [ApiController]
    public class AccountController : ControllerBase
    {
        private readonly IDictionary<ConnectionKey, ISqlSugarClient> _sqlSugarClients;
        private readonly IMapper _mapper;
        private readonly ILogger<NpoiController> _logger;
        private readonly IConnectionMultiplexer _connectionMultiplexer;
        private readonly IDatabase _redis;
        private readonly IConfiguration _configuration;
        private readonly IHttpContextAccessor _httpContextAccessor;

        /// <summary>
        /// 构造方法
        /// </summary>
        /// <param name="sqlSugarClients"></param>
        /// /// <param name="mapper"></param>
        /// <param name="logger"></param>
        /// <param name="connectionMultiplexer"></param>
        /// <param name="configuration"></param>
        /// <param name="httpContextAccessor"></param>
        public AccountController(IDictionary<ConnectionKey, ISqlSugarClient> sqlSugarClients,
            IMapper mapper,
            ILogger<NpoiController> logger,
            IConnectionMultiplexer connectionMultiplexer,
            IConfiguration configuration,
            IHttpContextAccessor httpContextAccessor
          )
        {
            _sqlSugarClients = sqlSugarClients;
            _mapper = mapper;
            _logger = logger;
            _connectionMultiplexer = connectionMultiplexer;
            _redis = connectionMultiplexer.GetDatabase(0);
            _configuration = configuration;
            _httpContextAccessor = httpContextAccessor;
        }

        /// <summary>
        /// 登录
        /// </summary>
        /// <returns></returns>
        [HttpPost(Name = "Login")]
        [PreventDuplicateRequests(30)]
        [AllowAnonymous]
        public IActionResult Login([FromBody] TestRequestDto quest)
        {
            // 生成 JWT
            var tokenHandler = new JwtSecurityTokenHandler();
            //从配置中获取数据
            var jwtTokenOptions = _configuration.GetSection("JwtOptions").Get<JwtOptions>() ?? new JwtOptions();
            //密钥对象
            SymmetricSecurityKey key = new SymmetricSecurityKey(Encoding.UTF8.GetBytes(jwtTokenOptions.SecurityKey));

            //数据中查询

            var db = _sqlSugarClients[ConnectionKey.PMS];
            ss_User? first = db.Queryable<ss_User>().Where(m => m.LoginName == quest.name).First();
            if (first != null && !string.IsNullOrEmpty(first.LoginName))
            {
                //身份信息对象
                var tokenDescriptor = new SecurityTokenDescriptor
                {
                    Subject = new ClaimsIdentity(new Claim[]
                    {
                        new Claim(ClaimTypes.NameIdentifier, first.UserId.ToString()),
                        new Claim("UserName", first.UserCName),
                        new Claim("Email", first.Email)
                    }),
                    Expires = DateTime.Now.AddMilliseconds(jwtTokenOptions.ExpireSeconds), //过期时间
                    Issuer = jwtTokenOptions.Isuser,
                    Audience = jwtTokenOptions.Audience,
                    NotBefore = DateTime.Now, //立即生效  DateTime.Now.AddMilliseconds(30),//30s后有效
                    SigningCredentials = new SigningCredentials(key, SecurityAlgorithms.HmacSha256)
                };
                var token = tokenHandler.CreateToken(tokenDescriptor);
                var jwt = tokenHandler.WriteToken(token);

                return Ok(new { jwt });
            }
            else
            {
                return Ok("用户名或者密码失败");
            }
        }



        /// <summary>
        /// 登录后测试用户数据
        /// </summary>
        /// <returns></returns>
        [HttpGet(Name = "LoginAfter_GetData")]
        [Authorize]

        public async Task<IActionResult> LoginAfter_GetData()
        {
            //获取登录用户的AuthenticateResult
            var auth = await _httpContextAccessor.HttpContext?.AuthenticateAsync();
            if (auth != null && auth.Succeeded)
            {
                //在声明集合中获取ClaimTypes.NameIdentifier 的值就是用户ID
                var userCli = auth.Principal.Claims.FirstOrDefault(c => c.Type == ClaimTypes.NameIdentifier);
                if (userCli != null && !string.IsNullOrEmpty(userCli.Value))
                {
                    return Ok(userCli);
                }

            }

            return Ok("Authenticated!");
        }

        /// <summary>
        /// 登录后测试用户数据1
        /// </summary>
        /// <returns></returns>
        [HttpGet(Name = "LoginAfter_GetData1")]
        [Authorize]

        public IActionResult LoginAfter_GetData1()
        {
            string test = DateTime.Now.ToString("yyyy:MM:dd HH:mm:ss");
            return Ok(test);
        }

        /// <summary>
        /// 登录前测试AllowAnonymous1
        /// </summary>
        /// <returns></returns>
        [HttpGet(Name = "LoginBefore_GetData1")]
        [AllowAnonymous]

        public IActionResult LoginBefore_GetData1()
        {
            string test = DateTime.Now.ToString("yyyy:MM:dd HH:mm:ss");
            return Ok(test);
        }

        /// <summary>
        /// 登录前测试无 AllowAnonymous Authorize
        /// </summary>
        /// <returns></returns>
        [HttpGet(Name = "LoginBefore_GetData2")]

        public IActionResult LoginBefore_GetData2()
        {
            string test = DateTime.Now.ToString("yyyy:MM:dd HH:mm:ss");
            return Ok(test);
        }
    }
}
