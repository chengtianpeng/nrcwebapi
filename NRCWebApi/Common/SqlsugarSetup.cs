using SqlSugar;
using SqlSugar.Extensions;

namespace NRCWebApi.Common
{
    public static class SqlsugarSetup
    {
        public static void AddSqlsugarSetup(this IServiceCollection services, WebApplicationBuilder builder)
        {
            builder.Services.AddSingleton<IDictionary<ConnectionKey, ISqlSugarClient>>(context =>
            {
                var sqlSugarClients = new Dictionary<ConnectionKey, ISqlSugarClient>();

                //sql server
                var connectionStringsSqlServer = builder.Configuration.GetSection("ConnectionStringsSqlServer").GetChildren();
                foreach (var connectionString in connectionStringsSqlServer)
                {

                    SqlSugarClient db = new SqlSugarClient(new ConnectionConfig()
                    {
                        ConnectionString = connectionString.Value,
                        DbType = DbType.SqlServer,
                        IsAutoCloseConnection = true,
                    });

                    ConnectionKey key = (ConnectionKey)Enum.Parse(typeof(ConnectionKey), connectionString.Key);
                    sqlSugarClients.Add(key, db);
                }

                //mysql
                var connectionStringsMysql = builder.Configuration.GetSection("ConnectionStringsMysql").GetChildren();
                foreach (var connectionString in connectionStringsMysql)
                {

                    SqlSugarClient db = new SqlSugarClient(new ConnectionConfig()
                    {
                        ConnectionString = connectionString.Value,
                        DbType = DbType.MySql,
                        IsAutoCloseConnection = true,
                    });

                    ConnectionKey key = (ConnectionKey)Enum.Parse(typeof(ConnectionKey), connectionString.Key);
                    sqlSugarClients.Add(key, db);
                }

                return sqlSugarClients;
            });
        }
    }
}
