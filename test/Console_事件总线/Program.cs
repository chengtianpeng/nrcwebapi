// See https://aka.ms/new-console-template for more information
using Console_事件总线;


Console.WriteLine("----------第一部分 https://www.cnblogs.com/bert-hubo/p/16229827.html  https://www.cnblogs.com/mq0036/p/16977665.html---------------------------");
//Task.Run 和 Async await的执行顺序
TaskRun.TestFun();
Console.WriteLine("----------第一部分总结：碰到await,必须执行完里面的方法，才能执行下一行的代码。 碰到Task.Run，new TaskFactory().StartNew . 下一行代码和里面的方法，同时执行。  ---------------------------");



Console.WriteLine("----------第二部分---------------------------");
//1. 初始化鱼竿
FishingRod fishingRod = new FishingRod();
//2. 声明钓鱼者
FishingMan fishingMan = new FishingMan("天鹏");
//3. 分配鱼竿
fishingMan.FishingRod = fishingRod;
//4. 注册观察者  订阅方
fishingRod.FishingEvent += fishingMan.Update;
//5. 循环钓鱼
while (fishingMan.FishCount < 5)
{
    fishingMan.Fishing();
    Console.WriteLine("-------------------------------------");

    Thread.Sleep(2000);
}


Console.WriteLine("Hello, World!");








/// <summary>
/// Async await的执行顺序 1
/// </summary>
public class TaskRun
{

    public static async void TestFun()
    {
        Console.WriteLine($"{DateTime.Now.ToString("HH:mm:ss.fff")} Test start!");
        await MainTask();
        Console.WriteLine($"{DateTime.Now.ToString("HH:mm:ss.fff")} Test End !");
    }

    private static async Task MainTask()
    {
        Console.WriteLine($"{DateTime.Now.ToString("HH:mm:ss.fff")} MainTask Start !");
        Thread.Sleep(2000);
        //await or not
        await Task.Run(() =>
        {
            Console.WriteLine($"{DateTime.Now.ToString("HH:mm:ss.fff")} Subtask Start !");
            Thread.Sleep(2000);
            Console.WriteLine($"{DateTime.Now.ToString("HH:mm:ss.fff")} Subtask End !");
        });
        Console.WriteLine($"{DateTime.Now.ToString("HH:mm:ss.fff")} MainTask End !");
    }
}


/// <summary>
/// Async await的执行顺序 2
/// </summary>
public class TaskRun2
{

    public static async void TestFun()
    {
        Console.WriteLine($"{DateTime.Now.ToString("HH:mm:ss.fff")} Test start!");
        MainTask();
        Console.WriteLine($"{DateTime.Now.ToString("HH:mm:ss.fff")} Test End !");
    }

    private static async Task MainTask()
    {
        Console.WriteLine($"{DateTime.Now.ToString("HH:mm:ss.fff")} MainTask Start !");
        Thread.Sleep(2000);
        //await or not
        Task.Run(() =>
        {
            Console.WriteLine($"{DateTime.Now.ToString("HH:mm:ss.fff")} Subtask Start !");
            for (int i = 0; i < 10; i++)
            {
                Console.WriteLine($"{DateTime.Now.ToString("HH:mm:ss.fff")} Subtask Start{i} !");
            }
            Thread.Sleep(2000);
            Console.WriteLine($"{DateTime.Now.ToString("HH:mm:ss.fff")} Subtask End !");
        });
        Console.WriteLine($"{DateTime.Now.ToString("HH:mm:ss.fff")} MainTask End !");
        for (int i = 0; i < 10; i++)
        {
            Console.WriteLine($"{DateTime.Now.ToString("HH:mm:ss.fff")} MainTask End{i} !");
        }
        Thread.Sleep(1000);
    }
}

