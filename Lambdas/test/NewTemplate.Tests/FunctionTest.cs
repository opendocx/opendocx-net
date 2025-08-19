using Xunit;
using Amazon.Lambda;
using Amazon.Lambda.Core;
using Amazon.Lambda.TestUtilities;
using Amazon.Lambda.S3Events;

using Amazon;
using Amazon.S3;
using Amazon.S3.Model;
using Amazon.S3.Util;
using System.Collections.Generic;

namespace NewTemplate.Tests;

public class FunctionTest
{
    [Fact]
    public async Task TestS3EventLambdaFunction()
    {
        IAmazonS3 s3Client = new AmazonS3Client(RegionEndpoint.USWest2);

        var bucketName = "lambda-NewTemplate-".ToLower() + DateTime.Now.Ticks;
        var key = "text.txt";

        // Create a bucket an object to setup a test data.
        await s3Client.PutBucketAsync(bucketName);
        try
        {
            await s3Client.PutObjectAsync(new PutObjectRequest
            {
                BucketName = bucketName,
                Key = key,
                ContentBody = "sample data"
            });

            // Setup the S3 event object that S3 notifications would create with the fields used by the Lambda function.
            var s3Event = new S3Event
            {
                Records = new List<S3Event.S3EventNotificationRecord>
                {
                    new S3Event.S3EventNotificationRecord
                    {
                        S3 = new S3Event.S3Entity
                        {
                            Bucket = new S3Event.S3BucketEntity {Name = bucketName },
                            Object = new S3Event.S3ObjectEntity {Key = key }
                        }
                    }
                }
            };

            // Invoke the lambda function and confirm the content type was returned.
            var functions = new Functions(s3Client);
            var contentType = await functions.DocxUploadExtractFields(s3Event,new TestLambdaContext());

            Assert.Equal("text/plain", contentType);

        }
        finally
        {
            // Clean up the test data
            await AmazonS3Util.DeleteS3BucketWithObjectsAsync(s3Client, bucketName);
        }
    }
}