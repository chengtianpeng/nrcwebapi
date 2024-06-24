using AutoMapper;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.StaticFiles;
using Microsoft.Extensions.Hosting.Internal;
using Microsoft.Net.Http.Headers;
using NPOI.HPSF;
using NPOI.OpenXmlFormats.Shared;
using NPOI.OpenXmlFormats.Spreadsheet;
using NPOI.OpenXmlFormats.Wordprocessing;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NRCWebApi.Common;
using NRCWebApi.Common.Redis;
using NRCWebApi.Dto;
using NRCWebApi.Dto.test;
using NRCWebApi.Filter;
using NRCWebApi.Model;
using SqlSugar;
using SqlSugar.Extensions;
using StackExchange.Redis;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Net;
using System.Net.Http.Headers;
using System.Web;

namespace NRCWebApi.Controllers
{
    /// <summary>
    /// Npoi操作控制器
    /// </summary>
    [Route("api/[controller]/[Action]")]
    [ApiController]
    public class NpoiController : ControllerBase
    {
        private readonly IDictionary<ConnectionKey, ISqlSugarClient> _sqlSugarClients;
        private readonly IMapper _mapper;
        private readonly ILogger<NpoiController> _logger;
        private readonly IConnectionMultiplexer _connectionMultiplexer;
        private readonly IDatabase _redis;

        /// <summary>
        /// 构造方法
        /// </summary>
        /// <param name="sqlSugarClients"></param>
        /// /// <param name="mapper"></param>
        /// <param name="logger"></param>
        /// <param name="connectionMultiplexer"></param>
        public NpoiController(IDictionary<ConnectionKey, ISqlSugarClient> sqlSugarClients,
            IMapper mapper,
            ILogger<NpoiController> logger,
            IConnectionMultiplexer connectionMultiplexer
          )
        {
            _sqlSugarClients = sqlSugarClients;
            _mapper = mapper;
            _logger = logger;
            _connectionMultiplexer = connectionMultiplexer;
            _redis = connectionMultiplexer.GetDatabase(0);
        }

        /// <summary>
        /// 测试导出-自动生成xlsx文件并导出xlsx文件流
        /// </summary>
        /// <returns></returns>
        [HttpGet(Name = "NopiAutoExport")]
        public FileContentResult NopiAutoExport()
        {
            var db = _sqlSugarClients[ConnectionKey.OA];//其中Dbconnect就是连接数据库字符串的Key
            var data = db.Queryable<sys_user_info>().ToList();

            List<NPOIRequestDto> dicColumns = new List<NPOIRequestDto>();
            dicColumns.Add(new NPOIRequestDto() { DataName = "user_name", TitleName = "用户账号", Width = 50 });
            dicColumns.Add(new NPOIRequestDto() { DataName = "nick_name", TitleName = "用户昵称", Width = 30 });
            dicColumns.Add(new NPOIRequestDto() { DataName = "email", TitleName = "用户邮箱" });
            dicColumns.Add(new NPOIRequestDto() { DataName = "phonenumber", TitleName = "手机号码" });

            byte[] buffer = ExcelUtil.ExportExcel<sys_user_info>(data, dicColumns, "数据列表");
            return File(buffer, "application/octet-stream", $"数据列表{DateTime.Now:yyyyMMddHHmmssfff}.xlsx");

            //using (MemoryStreamHelper ms = new MemoryStreamHelper())
            //{
            //    //加上设置大小下载下来的.xlsx文件打开时才不会报“Excel 已完成文件级验证和修复。此工作簿的某些部分可能已被修复或丢弃”
            //    // Response.Headers.Add("Content-Length", ms.Length.ToString());

            //    ms.AllowClose = false;
            //    workbook.Write(ms);
            //    workbook.Close();
            //    ms.Flush();
            //    ms.Position = 0;
            //    ms.AllowClose = true;

            //    size = ms.Length;
            //    buffer = ms.GetBuffer();

            //}


            //加上设置大小下载下来的.xlsx文件打开时才不会报“Excel 已完成文件级验证和修复。此工作簿的某些部分可能已被修复或丢弃”
            //  Response.Headers.Add("Content-Length", size.ToString());



            ////找到文件路径 find file path
            // string filePath = Path.Combine(Path.GetFullPath("Excel"), "Dem1o.xlsx");
            ////创建文件流 Create file stream
            //FileStream fs = new FileStream(filePath, FileMode.Create);
            //workbook.Write(fs);

            //  return Content(path);

            //  FileStream fsaa = System.IO.File.OpenRead(filePath);

            //  MemoryStream msss = new MemoryStream(buffer);

            //加上设置大小下载下来的.xlsx文件打开时才不会报“Excel 已完成文件级验证和修复。此工作簿的某些部分可能已被修复或丢弃”
            // Response.Headers.Add("Content-Length", "19213".ToString());

            // byte[] aaaa=  StreamToBytes(msss);
            //  return File(buffer, "application/octet-stream", "发送邮件的数据列表test.xlsx");

            //var dataaa = new FileStreamResult(msss, "application/octet-stream");
            //dataaa.FileDownloadName = "abc.xlsx";
            //return Task.FromResult<FileResult>(dataaa);




            //HttpResponseMessage result = new HttpResponseMessage(HttpStatusCode.OK);
            //result.Content = new StreamContent(fs);
            ////a text file is actually an octet-stream (pdf, etc)
            ////result.Content.Headers.ContentType = new MediaTypeHeaderValue("application/octet-stream");

            //result.Content.Headers.ContentType = new MediaTypeHeaderValue("application/octet-stream");
            ////we used attachment to force download
            //result.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment");
            //result.Content.Headers.ContentDisposition.FileName = "file.xlsx";
            //return result;


            // return File(buffer, "application/octet-stream", "发送邮件的数据列表test.xlsx");

            // string filePath = Path.Combine(Path.GetFullPath("Excel"), "Demo.xlsx");
            //string fileDirectoryPath = filePath.Substring(0, filePath.LastIndexOf("\\") + 1);
            //if (!Directory.Exists(fileDirectoryPath))
            //{
            //    Directory.CreateDirectory(fileDirectoryPath);
            //}
            //if (System.IO.File.Exists(filePath))
            //{
            //    System.IO.File.Delete(filePath);
            //}
            //System.IO.File.WriteAllBytes(filePath, buffer);

        }


        /// <summary>
        /// 测试导出-xlsx模板文件并导出xlsx文件流
        /// </summary>
        /// <returns></returns>
        [HttpGet(Name = "NopiTempExport")]
        public FileContentResult NopiTempExport()
        {
            var db = _sqlSugarClients[ConnectionKey.OA];//其中Dbconnect就是连接数据库字符串的Key
            var data = db.Queryable<sys_user_info>().ToList();

            List<string> dicColumns = new List<string>();
            dicColumns.Add("user_name");
            dicColumns.Add("nick_name");
            dicColumns.Add("email");
            dicColumns.Add("phonenumber");

            //找到文件路径 find file path
            string filePath = Path.Combine(Path.GetFullPath("Excel"), "NopiTempExport.xlsx");

            byte[] buffer = ExcelUtil.ExportExcel<sys_user_info>(filePath, data, dicColumns);
            return File(buffer, "application/octet-stream", $"数据Temp列表{DateTime.Now:yyyyMMddHHmmssfff}.xlsx");
        }


        /// <summary>
        /// 测试导入-文件导入成文件流，转DataTable导入
        /// </summary>
        /// <returns></returns>
        [HttpPost(Name = "NopiStreamDataTableImport")]
        public IActionResult NopiStreamDataTableImport(IFormFile excelFile)
        {
            //得到上传的文件
            var postFile = Request.Form.Files[0];
            //得到扩展名
            string extName = Path.GetExtension(postFile.FileName);
            //得到文件流
            MemoryStream ms = new MemoryStream();
            postFile.CopyTo(ms);
            ms.Position = 0;

            //转DataTable
            DataTable dt = ExcelUtil.ExcelToTable(ms, 0, 0, extName);

            return Ok(dt);
        }


        /// <summary>
        /// 测试导入-文件导入成文件流，转List导入
        /// </summary>
        /// <returns></returns>
        [HttpPost(Name = "NopiStreamListImport")]
        public IActionResult NopiStreamListImport(IFormFile excelFile)
        {
            //得到上传的文件
            var postFile = Request.Form.Files[0];
            //得到扩展名
            string extName = Path.GetExtension(postFile.FileName);
            //得到文件流
            MemoryStream ms = new MemoryStream();
            postFile.CopyTo(ms);
            ms.Position = 0;

            //导入的xlsx对应的数据库字段名称
            Dictionary<string, string> dict = new Dictionary<string, string>();
            dict.Add("用户账号", "user_name");
            dict.Add("用户昵称", "nick_name");
            dict.Add("用户邮箱", "email");
            dict.Add("手机号码", "phonenumber");

            //转List
            List<sys_user_info> list = ExcelUtil.ExcelToList<sys_user_info>(ms, 0, 0, extName, dict);

            return Ok(list);
        }



        /// <summary>
        /// 测试_AutoMapper To mdel
        /// </summary>
        /// <returns></returns>
        [HttpGet(Name = "Test_autoMapperToModel")]
        public IActionResult Test_autoMapperToModel()
        {
            var db = _sqlSugarClients[ConnectionKey.OA];//其中Dbconnect就是连接数据库字符串的Key
            var data = db.Queryable<sys_user_info>().ToList().FirstOrDefault();
            sys_user_infoDTO model = _mapper.Map<sys_user_infoDTO>(data);

            return Ok(model);
        }

        /// <summary>
        /// 测试_AutoMapper To List
        /// </summary>
        /// <returns></returns>
        [HttpGet(Name = "Test_autoMapperToList")]
        public IActionResult Test_autoMapperToList()
        {
            var db = _sqlSugarClients[ConnectionKey.OA];//其中Dbconnect就是连接数据库字符串的Key
            var data = db.Queryable<sys_user_info>().ToList();
            List<sys_user_infoDTO> models = _mapper.Map<List<sys_user_infoDTO>>(data);

            return Ok(models);
        }


        /// <summary>
        /// 测试Redis_1
        /// </summary>
        /// <returns></returns>
        [HttpGet(Name = "Test_Redis")]
        public IActionResult Test_Redis([FromQuery] string name)
        {

            //建立Redis 连接
            // ConnectionMultiplexer redis = ConnectionMultiplexer.Connect("172.16.8.188:6379,password=123456");

            _redis.StringGetSet("name", name);
            _redis.StringSet("nameT", name, TimeSpan.FromMinutes(1), When.Exists);
            string names = _redis.StringGet("name").ObjToString();
            return Ok(names);
        }


        /// <summary>
        /// 测试Get幂等性
        /// </summary>
        /// <returns></returns>
        [HttpGet(Name = "Test_Duplicate1")]
        [PreventDuplicateRequests(30)]
        public IActionResult Test_Duplicate1([FromQuery] TestRequestDto quest)
        {
            string test = quest.name + DateTime.Now.ToString("yyyy:MM:dd HH:mm:ss");
            return Ok(test);
        }


        /// <summary>
        /// 测试Post防重
        /// </summary>
        /// <returns></returns>
        [HttpPost(Name = "Test_Duplicate2")]
        [PreventDuplicateRequests(30)]
        public IActionResult Test_Duplicate2([FromBody] TestRequestDto quest)
        {
            string test = quest.name + DateTime.Now.ToString("yyyy:MM:dd HH:mm:ss");
            return Ok(test);
        }


    }
}
