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
using Peg.Base;

// Assembly attribute to enable the Lambda function's JSON input to be converted into a .NET class.
[assembly: LambdaSerializer(typeof(Amazon.Lambda.Serialization.SystemTextJson.DefaultLambdaJsonSerializer))]

namespace DocxTemplater;

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
    public async Task PrepareDocxTemplate(S3Event evnt, ILambdaContext context)
    {
        var eventRecords = evnt.Records ?? new List<S3Event.S3EventNotificationRecord>();
        foreach (var record in eventRecords)
        {
            var s3Entity = record.S3;
            if (s3Entity == null)
            {
                continue;
            }
            string key = s3Entity.Object.Key; // last portion should be "/template.docx", since that's what triggers this lambda
            if (record.EventName.StartsWith("s3:ObjectCreated:Copy")) // Skip processing for copy operations
            {
                context.Logger.LogLine($"Skipping PrepareDocxTemplate triggered by copy operation for {key}");
                continue;
            }
            if (!key.Contains("/templates/"))
            {
                context.Logger.LogLine($"Skipping PrepareDocxTemplate triggered on file NOT inside 'templates/': '{key}'");
                continue;
            }
            string baseKey = key[..(key.LastIndexOf('/') + 1)];
            string workspace = GetWorkspace(baseKey);
            string jobId = LastFolder(baseKey); // same as tvid (Template Version ID)
            context.Logger.Log($"Template job {jobId} uploaded; starting...");
            context.Logger.Log("Retrieving template from S3...");
            string bucket = s3Entity.Bucket.Name;
            byte[] docxBytes = await GetBytes(bucket, key, context);
            context.Logger.Log("Template retrieved successfully; normalizing (step 1A)...");
            byte[] normalizedBytes;
            try
            {
                var options = new PrepareTemplateOptions()
                {
                    GenerateFlatPreview = true,
                    RemoveCustomProperties = true,
                    KeepPropertyNames = new List<string>() { "UpdateFields", "PlayMacros" },
                };

                var normalizeResult = Normalizer.NormalizeTemplate(docxBytes, options.RemoveCustomProperties, options.KeepPropertyNames);
                normalizedBytes = normalizeResult.NormalizedTemplate;
                // await PutBytes(bucket, baseKey + "normalized.obj.docx", normalizedBytes);
                context.Logger.Log("Template normalized; building field dictionary (step 1B)...");
                var fieldDict = Templater.ParseFieldsToDict(normalizeResult.ExtractedFields);
                context.Logger.Log("Storing field dictionary in S3...");
                await PutString(bucket, baseKey + "fields.dict.json",
                    JsonSerializer.Serialize(fieldDict, OpenDocx.OpenDocx.DefaultJsonOptions));
                context.Logger.Log("Field dictionary stored (so other processes can use it); sending OK1 to SQS...");
                await SQSSender.SendMessageAsync(context.Logger, "OK1", workspace, jobId);

                context.Logger.Log("Transforming docx to oxpt (step 1C)...");
                var compileResult = Templater.CompileTemplate(normalizedBytes, fieldDict);
                if (compileResult.HasErrors)
                {
                    throw new Exception("Error while transforming DOCX:\n"
                        + string.Join('\n', compileResult.Errors));
                }
                context.Logger.Log("DOCX transformed; storing oxpt.docx in S3...");
                await PutBytes(bucket, baseKey + "oxpt.docx", compileResult.Bytes);
                context.Logger.Log("oxpt.docx stored; sending OK2 message to SQS...");
                await SQSSender.SendMessageAsync(context.Logger, "OK2", workspace, jobId);
                context.Logger.Log("Success");
            }
            catch (Exception e)
            {
                context.Logger.LogError(e.Message);
                if (e.StackTrace != null)
                    context.Logger.LogDebug(e.StackTrace);
                // send SQS message with error info:
                await SQSSender.SendMessageAsync(context.Logger, "Error: " + e.Message, workspace, jobId, e.Message);
                return;
            }
            // the following is OUTSIDE the above try/catch block, because we intentionally
            // do NOT want to report SQS whether it succeeds or fails.
            context.Logger.Log("Now replacing docx fields prior to markdown conversion (step 1D)...");
            try
            {
                var previewResult = TemplateTransformer.TransformTemplate(
                    normalizedBytes,
                    TemplateFormat.PreviewDocx,
                    null); // field map is ignored when output = TemplateFormat.PreviewDocx
                if (!previewResult.HasErrors)
                {
                    context.Logger.LogInformation("Preview generated; storing preview.obj.docx to S3...");
                    await PutBytes(bucket, baseKey + "preview.obj.docx", previewResult.Bytes);
                    // we do NOT report success via SQS, because the previews are not required for basic functionality
                }
                else
                {
                    context.Logger.LogInformation("Preview failed to generate:\n" + string.Join('\n', previewResult.Errors));
                    // we do NOT report errors via SQS, because the previews are not required for basic functionality
                }
            }
            catch (Exception e)
            {
                // some other random exception, typically an internal error
                context.Logger.LogInformation("Preview error: " + e.Message);
            }
        }
    }

    public async Task ImportDocxTemplate(S3Event evnt, ILambdaContext context)
    {
        var eventRecords = evnt.Records ?? new List<S3Event.S3EventNotificationRecord>();
        foreach (var record in eventRecords)
        {
            var s3Entity = record.S3;
            if (s3Entity == null)
            {
                continue;
            }
            string key = s3Entity.Object.Key; // last portion should be "/imported.docx", since that's what triggers this lambda
            if (record.EventName.StartsWith("s3:ObjectCreated:Copy")) // Skip processing for copy operations
            {
                context.Logger.LogLine($"Skipping ImportDocxTemplate triggered by copy operation for {key}");
                continue;
            }
            if (!key.Contains("/templates/"))
            {
                context.Logger.LogLine($"Skipping ImportDocxTemplate triggered on file NOT inside '/templates/': '{key}'");
                continue;
            }
            string baseKey = key[..(key.LastIndexOf('/') + 1)];
            string workspace = GetWorkspace(baseKey);
            string jobId = LastFolder(baseKey); // same as tvid (Template Version ID)
            context.Logger.Log($"Template Import job {jobId} uploaded; starting...");
            context.Logger.Log("Retrieving template from S3...");
            string bucket = s3Entity.Bucket.Name;
            byte[] docxBytes = await GetBytes(bucket, key, context);
            context.Logger.Log("Template retrieved successfully; beginning rename and normalizing (step 1A)...");
            // RenameObject uses COPY, so it should NOT trigger a redundant invocation of PrepareDocxTemplate
            var renameTask = RenameObject(bucket, key, baseKey + "template.docx", context);
            byte[] normalizedBytes;
            try
            {
                var options = new PrepareTemplateOptions()
                {
                    GenerateFlatPreview = true,
                    RemoveCustomProperties = true,
                    KeepPropertyNames = new List<string>() { "UpdateFields", "PlayMacros" },
                };

                var normalizeResult = Normalizer.NormalizeTemplate(docxBytes, options.RemoveCustomProperties, options.KeepPropertyNames);
                normalizedBytes = normalizeResult.NormalizedTemplate;
                // await PutBytes(bucket, baseKey + "normalized.obj.docx", normalizedBytes);
                context.Logger.Log("Template normalized; parsing fields and building logic tree...");
                var parseResult = Templater.ParseFields(normalizeResult.ExtractedFields, true);
                var fieldDict = parseResult.ParsedFields;
                context.Logger.Log("Parsing completed; storing field dictionary in S3...");
                await PutString(bucket, baseKey + "fields.dict.json",
                    JsonSerializer.Serialize(fieldDict, OpenDocx.OpenDocx.DefaultJsonOptions));
                // await SQSSender.SendMessageAsync(context.Logger, "OK1", workspace, jobId);
                context.Logger.Log("Field dictionary stored (so other processes can use it); storing logic.json...");
                await PutString(bucket, baseKey + "logic.json",
                    JsonSerializer.Serialize(parseResult.LogicTree, OpenDocx.OpenDocx.DefaultJsonOptions));

                context.Logger.Log("Logic tree stored; storing logic.js...");
                await PutString(bucket, baseKey + "logic.js", parseResult.LegacyLogicModule);
                context.Logger.Log("Logic module stored.");
                context.Logger.Log("Transforming docx to oxpt (step 1C)...");
                var compileResult = Templater.CompileTemplate(normalizedBytes, fieldDict);
                if (compileResult.HasErrors)
                {
                    throw new Exception("Error while transforming DOCX:\n"
                        + string.Join('\n', compileResult.Errors));
                }
                context.Logger.Log("DOCX transformed; storing oxpt.docx in S3...");
                await PutBytes(bucket, baseKey + "oxpt.docx", compileResult.Bytes);
                context.Logger.Log("oxpt.docx stored");
                // await SQSSender.SendMessageAsync(context.Logger, "OK2", workspace, jobId);
                // context.Logger.Log("Success");
            }
            catch (Exception e)
            {
                context.Logger.LogError(e.Message);
                if (e.StackTrace != null)
                    context.Logger.LogDebug(e.StackTrace);
                // send SQS message with error info:
                // await SQSSender.SendMessageAsync(context.Logger, "Error: " + e.Message, workspace, jobId, e.Message);
                return;
            }
            // the following is OUTSIDE the above try/catch block, because we intentionally
            // do NOT want to report SQS whether it succeeds or fails.
            context.Logger.Log("Now replacing docx fields prior to markdown conversion (step 1D)...");
            try
            {
                var previewResult = TemplateTransformer.TransformTemplate(
                    normalizedBytes,
                    TemplateFormat.PreviewDocx,
                    null); // field map is ignored when output = TemplateFormat.PreviewDocx
                if (!previewResult.HasErrors)
                {
                    context.Logger.LogInformation("Preview generated; storing preview.obj.docx to S3...");
                    await PutBytes(bucket, baseKey + "preview.obj.docx", previewResult.Bytes);
                    // we do NOT report success via SQS, because the previews are not required for basic functionality
                }
                else
                {
                    context.Logger.LogInformation("Preview failed to generate:\n" + string.Join('\n', previewResult.Errors));
                    // we do NOT report errors via SQS, because the previews are not required for basic functionality
                }
            }
            catch (Exception e)
            {
                // some other random exception, typically an internal error
                context.Logger.LogInformation("Preview error: " + e.Message);
            }
            await renameTask; // make sure rename has finished before we move on to the next event
        }
    }

    /// <summary>
    /// This method is called for every Lambda invocation.
    /// </summary>
    /// <param name="req">The request for the Lambda function handler to process.</param>
    /// <param name="context">The ILambdaContext that provides methods for logging and describing the Lambda environment.</param>
    /// <returns>The result of the assembly.</returns>
    public async Task<AssembleResponse> AssembleDocx(AssembleRequest req, ILambdaContext context)
    {
        if (string.IsNullOrEmpty(req.ID) && string.IsNullOrEmpty(req.OutputName))
        {
            return new AssembleResponse(req, false)
            {
                ErrorMessage = $"Invalid AssembleRequest: either OutputName or ID (or both) must be specified"
            };
        }
        try
        {
            // get the requested template
            byte[] templateBytes = await GetBytes(req.Bucket, req.TemplateKey, context);

            // get the requested composition sources, which should already have been completed
            var sourceTasks = req.Sources.Select(async src =>
            {
                var keepSections = req.Data.Contains($"oxpt://DocumentAssembler/insert/{src}?KeepSections=true");
                if (!keepSections)
                {
                    // sanity check -- make sure this approach works
                    if (!req.Data.Contains($"oxpt://DocumentAssembler/insert/{src}"))
                    {
                        throw new Exception("Unable to determine whether or not to keep Sections in insert operation");
                    }
                }
                var srcKey = req.DestinationKey + '/' + src;
                var sourceBytes = await GetBytes(req.Bucket, srcKey, context);
                return new IndirectSource(src, sourceBytes, keepSections);
            });
            var sources = await Task.WhenAll(sourceTasks);

            // assemble the document
            var result = await OpenDocx.Assembler.AssembleDocumentAsync(templateBytes, req.Data, sources);
            if (result.HasErrors)
            {
                return new AssembleResponse(req, false)
                {
                    ErrorMessage = $"Assembly error: {result.Error}"
                };
            }

            // Write result to S3
            string? resultKey = string.IsNullOrEmpty(req.OutputName) ? null : string.Join('/', req.DestinationKey, req.OutputName);
            string? interimResultKey = string.IsNullOrEmpty(req.ID) ? null : string.Join('/', req.DestinationKey, req.ID);
            string outKey = string.IsNullOrEmpty(interimResultKey) ? resultKey! : interimResultKey;
            await PutBytes(req.Bucket, outKey, result.Bytes);
            context.Logger.LogLine($"New object {outKey} stored");

            if (!string.IsNullOrEmpty(resultKey) && (outKey != resultKey))
            {
                // request included BOTH an ID (under which output has now already been stored)
                // AND a final output file name... in this case copy the output to the requested location
                await CopyObject(req.Bucket, interimResultKey!, resultKey, context);
            }

            return new AssembleResponse(req, true)
            {
                ResultKey = resultKey,
                InterimResultKey = interimResultKey,
            };
        }
        catch (Exception ex)
        {
            context.Logger.LogLine($"Exception: {ex.Message}");
            return new AssembleResponse(req, false)
            {
                ErrorMessage = $"Exception: {ex.Message}"
            };
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

    private async Task CopyObject(string bucket, string sourceKey, string destinationKey, ILambdaContext context)
    {
        try
        {
            var copyRequest = new CopyObjectRequest
            {
                SourceBucket = bucket,
                SourceKey = sourceKey,
                DestinationBucket = bucket,
                DestinationKey = destinationKey
            };
            await S3Client.CopyObjectAsync(copyRequest);
            context.Logger.LogLine($"Copied {sourceKey} to {destinationKey}");
        }
        catch (Exception e)
        {
            context.Logger.LogError($"Error copying object from {sourceKey} to {destinationKey} in bucket {bucket}");
            context.Logger.LogError(e.Message);
            throw;
        }
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

    private async Task RenameObject(string bucket, string sourceKey, string destinationKey, ILambdaContext context)
    {
        try
        {
            // Copy the object to the new location
            await CopyObject(bucket, sourceKey, destinationKey, context);
            
            // Delete the original object
            await DeleteObject(bucket, sourceKey, context);
            
            context.Logger.LogLine($"Renamed {sourceKey} to {destinationKey}");
        }
        catch (Exception e)
        {
            context.Logger.LogError($"Error renaming object from {sourceKey} to {destinationKey} in bucket {bucket}");
            context.Logger.LogError(e.Message);
            throw;
        }
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