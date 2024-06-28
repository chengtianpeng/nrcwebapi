using RabbitMQ.Client;
using RabbitMQ.Client.Events;
using System;
using System.Text;
using static System.Net.WebRequestMethods;

namespace MyRabbitMQ
{
    public class RabbitConection
    {
        public ConnectionFactory conect()
        {
            var factory = new ConnectionFactory
            {
                HostName = "172.16.8.89", // RabbitMQ 服务器的主机名或 IP 地址,我这里为本地
                Port = 5672, // RabbitMQ 服务器的端口号  
                UserName = "admin", // 用于身份验证的用户名  
                Password = "123456", // 用于身份验证的密码  
                VirtualHost = "/"  // 虚拟主机名称
            };

            return factory;
        }


        /// <summary>
        /// 发送消息
        /// </summary>
        public void SendMsg()
        {
            using (var connection = conect().CreateConnection())//建立与 RabbitMQ 服务器的连接
            {
                using (var channel = connection.CreateModel()) //在已经建立的连接上创建一个新的信道（channel）
                {
                    //声明队列
                    channel.QueueDeclare(queue: "NewRabbitMQTest",//队列的名称。如果此参数为空字符串，服务器将生成一个唯一的队列名称
                                         durable: true,//true为开启队列持久化
                                         exclusive: false,//队列是否只能由创建它的连接使用。当连接关闭时，队列将被自动删除。
                                         autoDelete: false,//当队列中的所有消息都被消费者消费后，是否应该自动删除队列。
                                         arguments: null//用于声明队列时的其他属性设置
                                         );

                    channel.CreateBasicProperties().DeliveryMode = 2;//设置消息持久化

                    string message = "你好test";
                    var body = Encoding.UTF8.GetBytes(message);
                    channel.BasicPublish(exchange: string.Empty,//交换机名称。
                              routingKey: "NewRabbitMQTest",//路由键。它决定了消息应该被发送到哪个队列。
                              basicProperties: null,//消息的属性，如消息的持久性、优先级、内容类型等。
                              body: body //消息的主体，通常是一个字节数组。
                              );


                    Console.WriteLine("Send OK");
                }
            }
        }


        /// <summary>
        /// 接收消息
        /// EventingBasicConsumer（T） 是 RabbitMQ 中的一个消费者类，用于接收来自 RabbitMQ 服务器中T队列的消息。
        ///Received事件会在队列中有新消息到达时被触发。
        /// </summary>
        public void ReceiveMsg()
        {
            using (var connection = conect().CreateConnection())//建立与 RabbitMQ 服务器的连接
            {
                using (var channel = connection.CreateModel()) //在已经建立的连接上创建一个新的信道（channel）
                {
                    //声明队列
                    channel.QueueDeclare(queue: "NewRabbitMQTest",//队列的名称。如果此参数为空字符串，服务器将生成一个唯一的队列名称
                                                   durable: true,//true为开启队列持久化
                                                   exclusive: false,//队列是否只能由创建它的连接使用。当连接关闭时，队列将被自动删除。
                                                   autoDelete: false,//当队列中的所有消息都被消费者消费后，是否应该自动删除队列。
                                                   arguments: null//用于声明队列时的其他属性设置
                                                   );
                    //限流:在消费者中设置://告诉broker同一时间只处理一个消息
                    channel.BasicQos(
                        prefetchSize: 0,
                        prefetchCount: 1,
                        global: false
                        );


                    var consumer = new EventingBasicConsumer(channel); //接收来自 RabbitMQ 服务器的消息


                    //Received在队列中有新消息到达时被触发
                    consumer.Received += (model, ea) =>//第一个参数通常为创建的EventingBasicConsumer实例，第二个参数包含了与接收到的消息相关的信息
                    {
                        var body = ea.Body.ToArray();
                        var message = Encoding.UTF8.GetString(body);

                        //false 只是确认签收当前的消息，设置为true的时候则代表签收该消费者所有未签收的消息
                        channel.BasicAck(deliveryTag: ea.DeliveryTag, multiple: false); //ea.DeliveryTag队列编号
                        Console.WriteLine($" [x] Received {message}");
                    };


                    channel.BasicConsume(queue: "NewRabbitMQTest",
                                         autoAck: false,//是否自动确认消息。如果设置为 true，则每当消息被消费者接收时，RabbitMQ 会自动认为该消息已被成功处理，并将其从队列中移除。如果设置为 false，则消费者需要显式地发送一个确认（通过 BasicAck 方法）来告诉 RabbitMQ 该消息已被成功处理。
                                         consumer: consumer);

                    Console.ReadLine();
                }
            }

        }
    }
}