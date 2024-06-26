using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Console_事件总线
{
    /// <summary>
    /// 鱼竿（被观察者） 发布方
    /// </summary>
    public class FishingRod
    {
        public delegate void FishingHandler(FishType type); //声明委托
        public event FishingHandler? FishingEvent; //声明事件


        public void ThrowHook(FishingMan fishingMan)
        {
            Console.WriteLine("{0}开始钓鱼！",fishingMan.Name);

            //用随机数模拟鱼咬钩，若随机数为偶数，则为鱼咬钩
            if (new Random().Next() % 2 == 0)
            {
                var type = (FishType)(new Random().Next(0, 5));
                Console.WriteLine("铃铛：叮叮叮，鱼儿咬钩了");

                if (FishingEvent != null)
                {
                    FishingEvent(type);
                }
            }

        }
    }
}
