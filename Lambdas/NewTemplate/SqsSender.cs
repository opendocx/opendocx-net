using Amazon.SQS;
using Amazon.SQS.Model;
using System;
using System.Text.Json;
using System.Threading.Tasks;

public class SqsSender
{
  private readonly IAmazonSQS _sqsClient;
  private string? _sqsUrl;

  public SqsSender(IAmazonSQS sqsClient)
  {
    _sqsClient = sqsClient;
  }

  public async Task SendMessageAsync(string messageBody, string workspace, string jobID, string? errMessage = null)
  {
    var queueUrl = await this.GetQueueUrl();
    var sendMessageRequest = new SendMessageRequest
    {
      QueueUrl = queueUrl,
      MessageBody = messageBody,
      MessageAttributes =
            {
                ["workspace"] = new MessageAttributeValue
                {
                    DataType = "String",
                    StringValue = workspace
                },
                ["jobID"] = new MessageAttributeValue
                {
                    DataType = "String",
                    StringValue = jobID
                },
                ["error"] = new MessageAttributeValue
                {
                    DataType = "String",
                    StringValue = errMessage ?? string.Empty
                }
            }
    };
    await _sqsClient.SendMessageAsync(sendMessageRequest);
  }
    
  private async Task<string> GetQueueUrl()
  {
    if (_sqsUrl == null)
    {
      // Get the ARN from environment variable
      var sqsQueueArn = Environment.GetEnvironmentVariable("SQS_QUEUE_ARN");
      if (sqsQueueArn != null)
      {
        // Use the SDK to get the URL
        var getQueueUrlResponse = await _sqsClient.GetQueueUrlAsync(new Amazon.SQS.Model.GetQueueUrlRequest
        {
          QueueName = sqsQueueArn.Split(':').Last() // Extract queue name from ARN
        });
        _sqsUrl = getQueueUrlResponse.QueueUrl;
      }
      else
      {
        throw new Exception($"Environment variable SQS_QUEUE_ARN not set!");
      }
    }
    return _sqsUrl;
  }
}