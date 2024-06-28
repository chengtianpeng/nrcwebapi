using System.Runtime.InteropServices;

namespace MyRabbitSend
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Hello, World!");
            test1();
        }


        /// <summary>
        /// 学习资料：https://blog.csdn.net/qq_58159494/article/details/136263157
        /// </summary>
        public static void test1() {

            MyRabbitMQ.RabbitConection con = new MyRabbitMQ.RabbitConection();
            // con.SendMsg();
            con.ReceiveMsg();
        }
    }
}