using System.ComponentModel;

namespace NRCWebApi.Common
{
    public class GlobalEnum
    {
    }

    public enum ConnectionKey
    {

        None = 0,
        PMS = 1,
        CRS = 2,
        OA = 3,
        HR = 4,
        MMS = 5,
        IT_8_188 = 6

    }

    /// <summary>
    /// 接口版本
    /// </summary>
    public enum ApiVersions
    {
        [Description("V1")]
        V1 = 0,

        [Description("V2")]
        V2 = 1,
    }


}
