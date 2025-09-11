using Amazon.Lambda.Core;
using Amazon.Lambda.S3Events;
using Amazon.SQS;
using Amazon.S3;
using Amazon.S3.Util;
using Amazon.S3.Model;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.Json.Serialization;
using OpenDocx;
using DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing;

// Assembly attribute to enable the Lambda function's JSON input to be converted into a .NET class.
[assembly: LambdaSerializer(typeof(Amazon.Lambda.Serialization.SystemTextJson.DefaultLambdaJsonSerializer))]

namespace NewTemplate;

public class Functions
{
    IAmazonS3 S3Client { get; set; }
    SqsSender SQSSender { get; set; }

    /// <summary>
    /// Default constructor. This constructor is used by Lambda to construct the instance. When invoked in a Lambda environment
    /// the AWS credentials will come from the IAM role associated with the function and the AWS region will be set to the
    /// region the Lambda function is executed in.
    /// </summary>
    public Functions()
    {
        S3Client = new AmazonS3Client();
        SQSSender = new SqsSender(new AmazonSQSClient());
    }

    /// <summary>
    /// Constructs an instance with a preconfigured S3 client. This can be used for testing the outside of the Lambda environment.
    /// </summary>
    /// <param name="s3Client">The service client to access Amazon S3.</param>
    public Functions(IAmazonS3 s3Client, IAmazonSQS sqsClient)
    {
        this.S3Client = s3Client;
        this.SQSSender = new SqsSender(sqsClient);
    }

    /// <summary>
    /// This method is called for every Lambda invocation. This method takes in an S3 event object and can be used 
    /// to respond to S3 notifications.
    /// </summary>
    /// <param name="evnt">The event for the Lambda function handler to process.</param>
    /// <param name="context">The ILambdaContext that provides methods for logging and describing the Lambda environment.</param>
    /// <returns></returns>
    public async Task ExtractFields(S3Event evnt, ILambdaContext context)
    {
        var eventRecords = evnt.Records ?? new List<S3Event.S3EventNotificationRecord>();
        foreach (var record in eventRecords)
        {
            var s3Event = record.S3;
            if (s3Event == null)
            {
                continue;
            }

            string bucket = s3Event.Bucket.Name;
            string key = s3Event.Object.Key; // last portion should be "/template.docx", since that's what triggers this lambda
            string baseKey = key[..(key.LastIndexOf('/') + 1)];
            string jobId = LastFolder(baseKey);
            context.Logger.Log($"Template job {jobId} uploaded; starting...");
            byte[] docxBytes = await GetBytes(bucket, key, context);

            try
            {
                var options = new PrepareTemplateOptions()
                {
                    GenerateFlatPreview = true,
                    RemoveCustomProperties = true,
                    KeepPropertyNames = new List<string>() { "UpdateFields", "PlayMacros" },
                };

                var normalizeResult = FieldExtractor.NormalizeTemplate(docxBytes, options.RemoveCustomProperties, options.KeepPropertyNames);
                await PutBytes(bucket, baseKey + "normalized.obj.docx", normalizeResult.NormalizedTemplate);
                await PutString(bucket, baseKey + "fields.obj.json", normalizeResult.ExtractedFields);
            }
            catch (Exception e)
            {
                context.Logger.LogInformation("Failure encountered: " + e.Message);
                if (e.StackTrace != null)
                    context.Logger.LogInformation(e.StackTrace);
            }
        }
    }

    public async Task ReplaceFieldsForPreview(S3Event evnt, ILambdaContext context)
    {
        var eventRecords = evnt.Records ?? new List<S3Event.S3EventNotificationRecord>();
        foreach (var record in eventRecords)
        {
            var s3Event = record.S3;
            if (s3Event == null)
            {
                continue;
            }

            string bucket = s3Event.Bucket.Name;
            string key = s3Event.Object.Key; // last portion should be "/normalized.obj.docx", since that's what triggers this lambda
            string baseKey = key[..(key.LastIndexOf('/') + 1)];
            string jobId = LastFolder(baseKey);

            context.Logger.Log($"Preview processing for job {jobId} starting...");
            byte[] normalizedBytes = await GetBytes(bucket, key, context);

            try
            {
                var previewResult = TemplateTransformer.TransformTemplate(
                    normalizedBytes,
                    TemplateFormat.PreviewDocx,
                    null); // field map is ignored when output = TemplateFormat.PreviewDocx
                if (!previewResult.HasErrors)
                {
                    context.Logger.LogInformation("Preview generated");
                    await PutBytes(bucket, baseKey + "preview.obj.docx", previewResult.Bytes);
                }
                else
                {
                    context.Logger.LogInformation("Preview failed to generate:\n" + string.Join('\n', previewResult.Errors));
                }
            }
            catch (Exception e)
            {
                // some other random exception, typically an internal error
                context.Logger.LogInformation("Preview error: " + e.Message);
            }
        }
    }

    public async Task TransformTemplate(S3Event evnt, ILambdaContext context)
    {
        var options = new JsonSerializerOptions
        {
            Converters = { new JsonStringEnumConverter() },
            NumberHandling = JsonNumberHandling.AllowReadingFromString
        };
        var eventRecords = evnt.Records ?? new List<S3Event.S3EventNotificationRecord>();
        foreach (var record in eventRecords)
        {
            var s3Event = record.S3;
            if (s3Event == null)
            {
                continue;
            }
            string bucket = s3Event.Bucket.Name;
            string key = s3Event.Object.Key;
            var dirs = key.Split('/');
            string baseKey = key[..(key.LastIndexOf('/') + 1)];
            string workspace = GetWorkspace(baseKey);
            string jobId = LastFolder(baseKey);
            string fieldDictStr = await GetString(bucket, key, context);

            // we also need to re-retrieve the normalized template (stored in step 1)
            byte[] normalizedBytes = await GetBytes(bucket, baseKey + "normalized.obj.docx", context);
            var fieldDict = JsonSerializer.Deserialize<Dictionary<string, ParsedField>>(fieldDictStr, options);
            var compileResult = Templater.CompileTemplate(normalizedBytes, fieldDict);
            if (compileResult.HasErrors)
            {
                // send SQS message:
                await SQSSender.SendMessageAsync("OK", workspace, jobId, string.Join('\n', compileResult.Errors));
            }
            else
            {
                await PutBytes(bucket, baseKey + "oxpt.docx", compileResult.Bytes);
            }
            // clean up inter-lambda temp files
            await DeleteObject(bucket, key, context);
            // TODO: maybe we should check to make sure the ReplaceFieldsForPreview lambda already got this file
            // and began its conversion for the preview, before deleting it?  But I'm betting that has very
            // likely already happened by now, so unless we encounter issues, let's just go ahead and clean it up:
            await DeleteObject(bucket, baseKey + "normalized.obj.docx", context);
        }
    }

    private async Task<byte[]> GetBytes(string bucket, string key, ILambdaContext context)
    {
        try
        {
            // context.Logger.LogInformation(response.Headers.ContentType);
            using (var response = await S3Client.GetObjectAsync(bucket, key))
            {
                using (MemoryStream ms = new MemoryStream())
                {
                    response.ResponseStream.CopyTo(ms);
                    return ms.ToArray();
                }
            }
        }
        catch (Exception e)
        {
            context.Logger.LogError($"Error getting object {key} from bucket {bucket}. Make sure it exists and your bucket is in the same region as this function.");
            context.Logger.LogError(e.Message);
            context.Logger.LogError(e.StackTrace);
            throw;
        }
    }

    private async Task PutBytes(string bucket, string key, byte[] bytes)
    {
        using (var stream = new MemoryStream(bytes))
        {
            PutObjectResponse response = await S3Client.PutObjectAsync(new PutObjectRequest
            {
                BucketName = bucket,
                Key = key,
                InputStream = stream,
            });
        }
    }

    private async Task<string> GetString(string bucket, string key, ILambdaContext context)
    {
        try
        {
            // context.Logger.LogInformation(response.Headers.ContentType);
            using (var response = await this.S3Client.GetObjectAsync(bucket, key))
            {
                using (var reader = new StreamReader(response.ResponseStream))
                {
                    return await reader.ReadToEndAsync();
                }
            }
        }
        catch (Exception e)
        {
            context.Logger.LogError($"Error getting object {key} from bucket {bucket}. Make sure it exists and your bucket is in the same region as this function.");
            context.Logger.LogError(e.Message);
            context.Logger.LogError(e.StackTrace);
            throw;
        }
    }

    private async Task PutString(string bucket, string key, string str)
    {
        PutObjectResponse response = await S3Client.PutObjectAsync(new PutObjectRequest
        {
            BucketName = bucket,
            Key = key,
            ContentBody = str,
        });
    }

    private async Task DeleteObject(string bucket, string key, ILambdaContext context)
    {
        context.Logger.Log($"Cleaning up {key}");
        await S3Client.DeleteObjectAsync(new DeleteObjectRequest
        {
            BucketName = bucket,
            Key = key
        });
    }

    private static string LastFolder(string key)
    {
        var trimmed = key.TrimEnd('/');
        int pos = trimmed.LastIndexOf('/') + 1;
        return trimmed[pos..];
    }

    private static string GetWorkspace(string key)
    {
        var end = key.IndexOf('/');
        var folder1 = end == -1 ? key : key[..end];
        var start = folder1.IndexOf('-') + 1;
        return folder1[start..];
    }
}