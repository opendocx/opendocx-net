using Amazon.Lambda.Core;
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

  public async Task SendMessageAsync(ILambdaLogger logger, string messageBody, string workspace, string jobID, string? errMessage = null)
  {
    var queueUrl = await this.GetQueueUrl(logger);
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
                }
            }
    };
    if (errMessage != null)
    {
      sendMessageRequest.MessageAttributes.Add("error", new MessageAttributeValue
      {
        DataType = "String",
        StringValue = errMessage
      });
    }
    await _sqsClient.SendMessageAsync(sendMessageRequest);
  }
    
  private async Task<string> GetQueueUrl(ILambdaLogger logger)
  {
    if (_sqsUrl == null)
    {
      // Get the ARN from environment variable
      var sqsQueueArn = Environment.GetEnvironmentVariable("SQS_QUEUE_ARN");
      if (sqsQueueArn != null)
      {
        logger.Log("SQS Queue ARN = " + sqsQueueArn);
        var queueName = sqsQueueArn.Split(':').Last(); // Extract queue name from ARN
        logger.Log("requesting queue URL for queue " + queueName);
        // Use the SDK to get the URL
        var getQueueUrlResponse = await _sqsClient.GetQueueUrlAsync(new GetQueueUrlRequest
        {
          QueueName = queueName
        });
        logger.Log("Got response");
        _sqsUrl = getQueueUrlResponse.QueueUrl;
        logger.Log("Queue URL is " + _sqsUrl);
      }
      else
      {
        throw new Exception($"Environment variable SQS_QUEUE_ARN not set!");
      }
    }
    return _sqsUrl;
  }
}