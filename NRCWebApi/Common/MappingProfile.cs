using AutoMapper;
using NRCWebApi.Dto;
using NRCWebApi.Model;

namespace NRCWebApi.Common
{
    /// <summary>
    /// mapping
    /// </summary>
    public class MappingProfile : Profile
    {
        /// <summary>
        /// 构造方法
        /// </summary>
        public MappingProfile()
        {
            //添加你需要映射的对象:左边是原始对象，右边是结果对象
            CreateMap<sys_user_info, sys_user_infoDTO>();
            CreateMap<HandSendEmailRequestDto, msg_review>();
        }
    }
}
