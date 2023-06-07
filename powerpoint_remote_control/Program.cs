using System;
using Amazon;
using Amazon.SQS;
using Amazon.SQS.Model;
using System.Threading.Tasks;

using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace powerpoint_remote_control
{
  internal class Program
  {
      private static async Task Main(string[] args)
        {
            var accessKey = Environment.GetEnvironmentVariable("AWS_ACCESS_KEY_ID");
            var secretKey = Environment.GetEnvironmentVariable("AWS_SECRET_ACCESS_KEY");
            var region = RegionEndpoint.USEast1;
            
            var sqsClient = new AmazonSQSClient(accessKey, secretKey, region);
            
            var queueUrl = Environment.GetEnvironmentVariable("SQS_QUEUE_URL");
            
            var receiveMessageRequest = new ReceiveMessageRequest(queueUrl);
            
            while (true)
            {
                
                var app = new PowerPoint.Application();

                var receiveRequest = new ReceiveMessageRequest
                {
                    QueueUrl = queueUrl,
                    MaxNumberOfMessages = 1,
                    WaitTimeSeconds = 20
                };
                
                var receiveResponse = await sqsClient.ReceiveMessageAsync(receiveRequest);

                foreach (var message in receiveResponse.Messages)
                {
                    var body = message.Body;
                    
                    Console.WriteLine(body);
                    
                    var presentation = app.ActivePresentation;
                    
                    switch (body)
                    {
                        case "next":
                            presentation.SlideShowWindow.View.Next();
                            break;
                        case "previous":
                            presentation.SlideShowWindow.View.Previous();
                            break;
                        default:
                            Console.WriteLine("Unknown command");
                            break;
                    }
                    
                    var deleteRequest = new DeleteMessageRequest
                    {
                        QueueUrl = queueUrl,
                        ReceiptHandle = message.ReceiptHandle
                    };

                    await sqsClient.DeleteMessageAsync(deleteRequest);
                }
            }
        }
    }
}