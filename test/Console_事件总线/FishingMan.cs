using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Console_事件总线
{
    /// <summary>
    /// 垂钓者（观察者） 订阅方
    /// </summary>
    public class FishingMan
    {
        public string Name { get; set; }
        public int FishCount { get; set; }

        public FishingMan(string name)
        {
            Name = name;
        }

        /// <summary>
        /// 钓鱼者使用的鱼竿
        /// </summary>
        public FishingRod? FishingRod { get; set; }


        public void Fishing()
        {
            this.FishingRod?.ThrowHook(this);
        }

        public void Update(FishType type)
        {
            FishCount++;
            Console.WriteLine("{0}:钓到一条[{2}]，已经钓到{1}条鱼了！", Name, FishCount, type);
        }
    }
}
