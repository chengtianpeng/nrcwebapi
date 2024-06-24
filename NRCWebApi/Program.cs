
using Autofac.Core;
using Microsoft.OpenApi.Models;
using NRCWebApi.Common;
using SqlSugar;
using NLog;
using NLog.Web;
using NLog.Extensions.Logging;
using AutoMapper;
using Newtonsoft.Json;
using Newtonsoft.Json.Serialization;
using StackExchange.Redis;
using NRCWebApi.Dto.appsettings;
using NRCWebApi.Common.Redis;
using NRCWebApi.Filter;
using Microsoft.AspNetCore.Authentication.JwtBearer;
using Microsoft.IdentityModel.Tokens;
using System.Text;

var builder = WebApplication.CreateBuilder(args);

// Add services to the container.

builder.Services.AddControllers().AddNewtonsoftJson(opt =>
{
    //忽略循环引用
    opt.SerializerSettings.ReferenceLoopHandling = ReferenceLoopHandling.Ignore;

    //不改变字段大小
    //opt.SerializerSettings.ContractResolver = new DefaultContractResolver();
}); ;

// Learn more about configuring Swagger/OpenAPI at https://aka.ms/aspnetcore/swashbuckle
builder.Services.AddEndpointsApiExplorer();

//添加Memory缓存
builder.Services.AddMemoryCache();
//添加分布式缓存
builder.Services.AddDistributedMemoryCache();

//添加redis缓存服务
string configuration = builder.Configuration.GetValue<string>("RedisServer:Connection") ?? ""; // builder.Configuration.GetSection("RedisServer:Configuration").Value ?? "";
                                                                                               //第一种方式
builder.Services.AddSingleton<IConnectionMultiplexer>(sp =>
{
    return ConnectionMultiplexer.Connect(configuration);
});

//第二种方式 是 第一种的简写
//var redis = ConnectionMultiplexer.Connect(configuration);
//builder.Services.AddSingleton<IConnectionMultiplexer>(redis);

//第三种方式 是 配置到一个类
//获取 appsettings.json 中RedisServer的 配置
//RedisServer redisServer = builder.Configuration.GetSection("RedisServer").Get<RedisServer>() ?? new RedisServer();
////注册redis缓存服务
//builder.Services.AddSingleton(new RedisHelper(redisServer));


//注入 幂等性 和 接口防重组件
builder.Services.AddScoped<PreventDuplicateRequestsActionFilter>();



// 配置映射对象 Auto Mapper Configurations
var mapperConfig = new MapperConfiguration(mc =>
{
    mc.AddProfile(new MappingProfile());
});
IMapper mapper = mapperConfig.CreateMapper();
builder.Services.AddSingleton(mapper);



//配置Nlog日志
//配置从配置文件的`NLog` 节点读取配置
//var nlogConfig = builder.Configuration.GetSection("NLog");
//NLog.LogManager.Configuration = new NLogLoggingConfiguration(nlogConfig);
//清空其他日志Providers
builder.Logging.ClearProviders();
//设置日志等级
builder.Logging.SetMinimumLevel(Microsoft.Extensions.Logging.LogLevel.Trace);
//该配置用来指定使用ASP.NET Core 默认的日志过滤器
var nlogOptions = new NLogAspNetCoreOptions() { RemoveLoggerFactoryFilter = false };
builder.Host.UseNLog(nlogOptions); //启用NLog



// 配置 JWT 身份验证
JwtOptions jwtDto = builder.Configuration.GetSection("JwtOptions").Get<JwtOptions>() ?? new JwtOptions();
builder.Services.AddAuthentication(JwtBearerDefaults.AuthenticationScheme)//声明使用 JWT 进行身份验证。
    .AddJwtBearer(options =>
    {
        //TokenValidationParameters包含一组参数，用于验证 JWT 的有效性。
        options.TokenValidationParameters = new TokenValidationParameters
        {
            ValidateIssuer = true,
            ValidateAudience = true,
            ValidateLifetime = true,
            ValidateIssuerSigningKey = true,
            //ValidIssuer 和 ValidAudience 是我们信任的 JWT 签发者和受众者。
            ValidIssuer = jwtDto.Isuser,
            ValidAudience = jwtDto.Audience,
            //IssuerSigningKey 是我们用于加密和验证 JWT 的密钥。
            IssuerSigningKey = new SymmetricSecurityKey(Encoding.UTF8.GetBytes(jwtDto.SecurityKey))
        };
    });





//swagger
builder.Services.AddSwaggerGen(option =>
{
    foreach (var version in typeof(NRCWebApi.Common.ApiVersions).GetEnumNames())
    {
        option.SwaggerDoc(version, new OpenApiInfo()
        {
            Title = "Swagger的Api文档",
            Version = version,
            Description = "通用版本的CoreApi版本" + version
        });
    }


    option.AddSecurityDefinition("Bearer", new OpenApiSecurityScheme
    {
        Description = "Authorization header. \r\nExample:Bearer 12345ABCDE",
        Name = "Authorization",
        In = ParameterLocation.Header,
        Type = SecuritySchemeType.ApiKey,
        Scheme = "Bearer"
    });
    option.AddSecurityRequirement(new OpenApiSecurityRequirement()
        {{
            new OpenApiSecurityScheme
            {
                Reference = new OpenApiReference
                {
                    Type = ReferenceType.SecurityScheme,
                    Id = "Bearer"
                }, Scheme = "oauth2", Name = "Bearer", In = ParameterLocation.Header }, new List<string>()
            }
        });




    // xml文档绝对路径 
    var file = Path.Combine(AppContext.BaseDirectory, $"{AppDomain.CurrentDomain.FriendlyName}.xml");

    // true : 显示控制器层注释
    option.IncludeXmlComments(file, true);
    // 对action的名称进行排序，如果有多个，就可以看见效果了。
    option.OrderActionsBy(o => o.RelativePath);
});




//配置Cors
builder.Services.AddCors(option =>
{
    option.AddPolicy("any", builder =>
    {
        builder.SetIsOriginAllowed(_ => true).AllowAnyMethod().AllowAnyHeader().AllowCredentials();
    });
});

//配置数据库连接
//这三个方法是ASP.NET Core中的依赖注入容器中的注册方法，它们的区别如下：

//1. AddSingleton：将服务注册为单例模式，即在整个应用程序生命周期中只创建一个实例，每次请求都返回同一个实例。

//2. AddScoped：将服务注册为作用域模式，即在每个请求范围内创建一个实例，同一请求内多次请求返回同一个实例，不同请求返回不同实例。

//3. AddTransient：将服务注册为瞬态模式，即每次请求都创建一个新的实例，每次请求返回不同的实例。

//总结：

//- AddSingleton：适用于需要在整个应用程序中共享的服务，如数据库上下文、配置服务等。

//- AddScoped：适用于需要在同一请求内共享的服务，如请求级别的缓存、请求级别的数据库上下文等。

//- AddTransient：适用于每次请求都需要创建新实例的服务，如日志服务、随机数生成器等。
// 注册sqlsugar

//依赖注入 sqlsugar数据库连接
builder.Services.AddSqlsugarSetup(builder);


var app = builder.Build();

// Configure the HTTP request pipeline.
//if (app.Environment.IsDevelopment())
//{
//    app.UseSwagger();
//    app.UseSwaggerUI();
//}

app.UseSwagger();
app.UseSwaggerUI(option =>
{
    foreach (string version in typeof(NRCWebApi.Common.ApiVersions).GetEnumNames())
    {
        option.SwaggerEndpoint($"/swagger/{version}/swagger.json", $"Steven文档【{version}】版本");
        option.DefaultModelExpandDepth(-1);
    }
});


//认证：是启用身份验证中间件，它会验证请求中的身份信息，并将身份信息存储在HttpContext.User属性中
app.UseAuthentication();//注意，一定得先启动这个
//授权：是启用授权中间件，它会检查HttpContext.User中的身份信息是否有访问当前请求所需的权限。
app.UseAuthorization();

//配置Cors
app.UseCors("any");

app.MapControllers();

app.Run();
