using System;
using System.IO;
using System.Threading.Tasks;
using InTheHand.Net;
using InTheHand.Net.Bluetooth;
using InTheHand.Net.Sockets;

using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace powerpoint_remote_control
{
    internal class Program
    {
        // Serial Port Profile UUID — universally recognized by mobile Bluetooth clients
        private static readonly Guid ServiceUuid = new Guid("00001101-0000-1000-8000-00805f9b34fb");

        private static async Task Main(string[] args)
        {
            var app = new PowerPoint.Application();

            using var listener = new BluetoothListener(ServiceUuid);
            listener.ServiceName = "PowerPoint Remote Control";
            listener.Start();

            Console.WriteLine("Bluetooth server started. Pair your device and connect.");
            Console.WriteLine($"Service: {listener.ServiceName}");
            Console.WriteLine($"UUID:    {ServiceUuid}");
            Console.WriteLine("Press Ctrl+C to exit.\n");

            var exitSignal = new TaskCompletionSource<bool>();
            Console.CancelKeyPress += (s, e) =>
            {
                e.Cancel = true;
                exitSignal.TrySetResult(true);
            };

            // Accept clients concurrently
            _ = Task.Run(async () =>
            {
                while (!exitSignal.Task.IsCompleted)
                {
                    var client = await Task.Run(() => listener.AcceptBluetoothClient());
                    _ = HandleClientAsync(client, app);
                }
            });

            await exitSignal.Task;
            Console.WriteLine("Shutting down. Goodbye!");
        }

        private static async Task HandleClientAsync(BluetoothClient client, PowerPoint.Application app)
        {
            using (client)
            using (var stream = client.GetStream())
            using (var reader = new StreamReader(stream))
            {
                Console.WriteLine($"[{DateTime.Now:HH:mm:ss}] Connected: {client.RemoteMachineName}");

                try
                {
                    string line;
                    while ((line = await reader.ReadLineAsync()) != null)
                    {
                        ProcessCommand(line.Trim(), app);
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Connection error: {ex.Message}");
                }

                Console.WriteLine($"[{DateTime.Now:HH:mm:ss}] Disconnected: {client.RemoteMachineName}");
            }
        }

        private static void ProcessCommand(string command, PowerPoint.Application app)
        {
            Console.WriteLine($"[{DateTime.Now:HH:mm:ss.fff}] Command: {command}");

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
                Console.WriteLine($"Error processing '{command}': {ex.Message}");
            }
        }
    }
}
