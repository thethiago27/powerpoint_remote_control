using System;
using System.Threading.Tasks;
using StackExchange.Redis;

using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace powerpoint_remote_control
{
    internal class Program
    {
        private static async Task Main(string[] args)
        {
            var redisConnection = Environment.GetEnvironmentVariable("REDIS_CONNECTION_STRING") ?? "localhost:6379";
            var channel = Environment.GetEnvironmentVariable("REDIS_CHANNEL") ?? "powerpoint-control";

            Console.WriteLine($"Connecting to Redis at {redisConnection}...");

            var redis = await ConnectionMultiplexer.ConnectAsync(redisConnection);
            var subscriber = redis.GetSubscriber();

            var app = new PowerPoint.Application();

            Console.WriteLine($"Listening on channel '{channel}'. Press Ctrl+C to exit.");

            var exitSignal = new TaskCompletionSource<bool>();

            Console.CancelKeyPress += (s, e) =>
            {
                e.Cancel = true;
                exitSignal.TrySetResult(true);
            };

            await subscriber.SubscribeAsync(RedisChannel.Literal(channel), (ch, message) =>
            {
                var command = (string)message;
                Console.WriteLine($"[{DateTime.Now:HH:mm:ss.fff}] Command received: {command}");

                try
                {
                    var presentation = app.ActivePresentation;

                    switch (command)
                    {
                        case "next":
                            presentation.SlideShowWindow.View.Next();
                            break;
                        case "previous":
                            presentation.SlideShowWindow.View.Previous();
                            break;
                        case "first":
                            presentation.SlideShowWindow.View.First();
                            break;
                        case "last":
                            presentation.SlideShowWindow.View.Last();
                            break;
                        default:
                            Console.WriteLine($"Unknown command: {command}");
                            break;
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing command '{command}': {ex.Message}");
                }
            });

            await exitSignal.Task;

            await subscriber.UnsubscribeAllAsync();
            redis.Dispose();

            Console.WriteLine("Disconnected. Goodbye!");
        }
    }
}
