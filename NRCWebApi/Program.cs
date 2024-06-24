
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
    //����ѭ������
    opt.SerializerSettings.ReferenceLoopHandling = ReferenceLoopHandling.Ignore;

    //���ı��ֶδ�С
    //opt.SerializerSettings.ContractResolver = new DefaultContractResolver();
}); ;

// Learn more about configuring Swagger/OpenAPI at https://aka.ms/aspnetcore/swashbuckle
builder.Services.AddEndpointsApiExplorer();

//���Memory����
builder.Services.AddMemoryCache();
//��ӷֲ�ʽ����
builder.Services.AddDistributedMemoryCache();

//���redis�������
string configuration = builder.Configuration.GetValue<string>("RedisServer:Connection") ?? ""; // builder.Configuration.GetSection("RedisServer:Configuration").Value ?? "";
                                                                                               //��һ�ַ�ʽ
builder.Services.AddSingleton<IConnectionMultiplexer>(sp =>
{
    return ConnectionMultiplexer.Connect(configuration);
});

//�ڶ��ַ�ʽ �� ��һ�ֵļ�д
//var redis = ConnectionMultiplexer.Connect(configuration);
//builder.Services.AddSingleton<IConnectionMultiplexer>(redis);

//�����ַ�ʽ �� ���õ�һ����
//��ȡ appsettings.json ��RedisServer�� ����
//RedisServer redisServer = builder.Configuration.GetSection("RedisServer").Get<RedisServer>() ?? new RedisServer();
////ע��redis�������
//builder.Services.AddSingleton(new RedisHelper(redisServer));


//ע�� �ݵ��� �� �ӿڷ������
builder.Services.AddScoped<PreventDuplicateRequestsActionFilter>();



// ����ӳ����� Auto Mapper Configurations
var mapperConfig = new MapperConfiguration(mc =>
{
    mc.AddProfile(new MappingProfile());
});
IMapper mapper = mapperConfig.CreateMapper();
builder.Services.AddSingleton(mapper);



//����Nlog��־
//���ô������ļ���`NLog` �ڵ��ȡ����
//var nlogConfig = builder.Configuration.GetSection("NLog");
//NLog.LogManager.Configuration = new NLogLoggingConfiguration(nlogConfig);
//���������־Providers
builder.Logging.ClearProviders();
//������־�ȼ�
builder.Logging.SetMinimumLevel(Microsoft.Extensions.Logging.LogLevel.Trace);
//����������ָ��ʹ��ASP.NET Core Ĭ�ϵ���־������
var nlogOptions = new NLogAspNetCoreOptions() { RemoveLoggerFactoryFilter = false };
builder.Host.UseNLog(nlogOptions); //����NLog



// ���� JWT �����֤
JwtOptions jwtDto = builder.Configuration.GetSection("JwtOptions").Get<JwtOptions>() ?? new JwtOptions();
builder.Services.AddAuthentication(JwtBearerDefaults.AuthenticationScheme)//����ʹ�� JWT ���������֤��
    .AddJwtBearer(options =>
    {
        //TokenValidationParameters����һ�������������֤ JWT ����Ч�ԡ�
        options.TokenValidationParameters = new TokenValidationParameters
        {
            ValidateIssuer = true,
            ValidateAudience = true,
            ValidateLifetime = true,
            ValidateIssuerSigningKey = true,
            //ValidIssuer �� ValidAudience ���������ε� JWT ǩ���ߺ������ߡ�
            ValidIssuer = jwtDto.Isuser,
            ValidAudience = jwtDto.Audience,
            //IssuerSigningKey ���������ڼ��ܺ���֤ JWT ����Կ��
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
            Title = "Swagger��Api�ĵ�",
            Version = version,
            Description = "ͨ�ð汾��CoreApi�汾" + version
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




    // xml�ĵ�����·�� 
    var file = Path.Combine(AppContext.BaseDirectory, $"{AppDomain.CurrentDomain.FriendlyName}.xml");

    // true : ��ʾ��������ע��
    option.IncludeXmlComments(file, true);
    // ��action�����ƽ�����������ж�����Ϳ��Կ���Ч���ˡ�
    option.OrderActionsBy(o => o.RelativePath);
});




//����Cors
builder.Services.AddCors(option =>
{
    option.AddPolicy("any", builder =>
    {
        builder.SetIsOriginAllowed(_ => true).AllowAnyMethod().AllowAnyHeader().AllowCredentials();
    });
});

//�������ݿ�����
//������������ASP.NET Core�е�����ע�������е�ע�᷽�������ǵ��������£�

//1. AddSingleton��������ע��Ϊ����ģʽ����������Ӧ�ó�������������ֻ����һ��ʵ����ÿ�����󶼷���ͬһ��ʵ����

//2. AddScoped��������ע��Ϊ������ģʽ������ÿ������Χ�ڴ���һ��ʵ����ͬһ�����ڶ�����󷵻�ͬһ��ʵ������ͬ���󷵻ز�ͬʵ����

//3. AddTransient��������ע��Ϊ˲̬ģʽ����ÿ�����󶼴���һ���µ�ʵ����ÿ�����󷵻ز�ͬ��ʵ����

//�ܽ᣺

//- AddSingleton����������Ҫ������Ӧ�ó����й���ķ��������ݿ������ġ����÷���ȡ�

//- AddScoped����������Ҫ��ͬһ�����ڹ���ķ��������󼶱�Ļ��桢���󼶱�����ݿ������ĵȡ�

//- AddTransient��������ÿ��������Ҫ������ʵ���ķ�������־����������������ȡ�
// ע��sqlsugar

//����ע�� sqlsugar���ݿ�����
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
        option.SwaggerEndpoint($"/swagger/{version}/swagger.json", $"Steven�ĵ���{version}���汾");
        option.DefaultModelExpandDepth(-1);
    }
});


//��֤�������������֤�м����������֤�����е������Ϣ�����������Ϣ�洢��HttpContext.User������
app.UseAuthentication();//ע�⣬һ�������������
//��Ȩ����������Ȩ�м����������HttpContext.User�е������Ϣ�Ƿ��з��ʵ�ǰ���������Ȩ�ޡ�
app.UseAuthorization();

//����Cors
app.UseCors("any");

app.MapControllers();

app.Run();
