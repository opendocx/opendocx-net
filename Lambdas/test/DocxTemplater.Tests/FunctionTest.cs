using System.IO.Compression;
using System.Text;
using System.Text.Json;
using System.Xml.Linq;
using Amazon;
using Amazon.Lambda.S3Events;
using Amazon.Lambda.TestUtilities;
using Amazon.Runtime;
using Amazon.S3;
using Amazon.S3.Model;
using OpenDocx;
using Xunit;

namespace DocxTemplater.Tests;

public class FunctionTest
{
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public async Task ProvisionThenAssembleWithHostEvaluatedData(bool includeDetails)
    {
        const string prefix = "example/templates/greeting/";
        using var s3 = new MemoryS3();
        var source = Template();
        s3.Objects[prefix + "import.docx"] = source;
        var functions = new Functions(s3);
        var context = new TestLambdaContext();

        await functions.ImportDocxTemplate(new S3Event
        {
            Records = new List<S3Event.S3EventNotificationRecord>
            {
                new()
                {
                    EventName = "ObjectCreated:Put",
                    S3 = new S3Event.S3Entity
                    {
                        Bucket = new S3Event.S3BucketEntity { Name = "example-bucket" },
                        Object = new S3Event.S3ObjectEntity { Key = prefix + "import.docx" }
                    }
                }
            }
        }, context);

        Assert.False(s3.Objects.ContainsKey(prefix + "import.docx"));
        Assert.Equal(source, s3.Objects[prefix + "template.docx"]);
        Assert.True(s3.Objects.ContainsKey(prefix + "oxpt.docx"), ((TestLambdaLogger)context.Logger).Buffer.ToString());
        using var logic = JsonDocument.Parse(s3.Objects[prefix + "logic.json"]);
        Assert.NotEqual(0, logic.RootElement.GetArrayLength());
        var fields = JsonSerializer.Deserialize<Dictionary<string, ParsedField>>(
            s3.Objects[prefix + "fields.dict.json"], OpenDocx.OpenDocx.DefaultJsonOptions)!;
        var name = Assert.Single(fields.Values, f => f.Expression == "Customer.Name");
        var condition = Assert.Single(fields.Values, f => f.Type == FieldType.If);

        // The host evaluates expressions and builds XML using returned atom names.
        var data = new XElement("_odx",
            new XElement(name.Atom, "Ada Lovelace"),
            new XElement(condition.Atom + "b", includeDetails ? "true" : "false"));
        var response = await functions.AssembleDocx(new AssembleRequest
        {
            Bucket = "example-bucket",
            TemplateKey = prefix + "oxpt.docx",
            DestinationKey = "example/output",
            OutputName = "greeting.docx",
            Sources = new List<string>(),
            Data = data.ToString()
        }, context);

        Assert.True(response.Success, response.ErrorMessage);
        Assert.Equal("example/output/greeting.docx", response.ResultKey);
        using var zip = new ZipArchive(new MemoryStream(s3.Objects[response.ResultKey!]));
        using var xml = zip.GetEntry("word/document.xml")!.Open();
        XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        var text = string.Concat(XDocument.Load(xml).Descendants(w + "t").Select(t => t.Value));
        Assert.Equal("Hello Ada Lovelace." + (includeDetails ? "Details included." : "Summary only."), text);
        Assert.DoesNotContain("{[", text);
    }

    [Fact]
    public async Task MissingOutputIdentityIsRejectedBeforeStorageAccess()
    {
        using var s3 = new MemoryS3();
        var response = await new Functions(s3).AssembleDocx(new AssembleRequest
        {
            Bucket = "example-bucket",
            TemplateKey = "unused.docx",
            DestinationKey = "example/output",
            Sources = new List<string>(),
            Data = "<_odx/>"
        }, new TestLambdaContext());
        Assert.False(response.Success);
        Assert.Contains("either OutputName or ID", response.ErrorMessage);
        Assert.Equal(0, s3.Reads);
        Assert.Empty(s3.Objects);
    }

    private static byte[] Template()
    {
        using var bytes = new MemoryStream();
        using (var zip = new ZipArchive(bytes, ZipArchiveMode.Create, true))
        {
            void Add(string name, string content)
            {
                using var writer = new StreamWriter(zip.CreateEntry(name).Open(), new UTF8Encoding(false));
                writer.Write(content);
            }
            Add("[Content_Types].xml", """
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                </Types>
                """);
            Add("_rels/.rels", """
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
                """);
            Add("word/_rels/document.xml.rels", """
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rSettings" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
                </Relationships>
                """);
            Add("word/settings.xml", "<w:settings xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"/>");
            Add("word/document.xml", """
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
                  <w:p><w:r><w:t>Hello {[Customer.Name]}.</w:t></w:r></w:p>
                  <w:p><w:r><w:t>{[if IncludeDetails]}Details included.{[else]}Summary only.{[endif]}</w:t></w:r></w:p>
                  <w:sectPr/>
                </w:body></w:document>
                """);
        }
        return bytes.ToArray();
    }

    private sealed class MemoryS3 : AmazonS3Client
    {
        public Dictionary<string, byte[]> Objects { get; } = new();
        public int Reads { get; private set; }

        public MemoryS3() : base(new AnonymousAWSCredentials(), RegionEndpoint.USEast1) { }

        public override Task<GetObjectResponse> GetObjectAsync(string bucketName, string key,
            CancellationToken cancellationToken = default)
        {
            Reads++;
            return Task.FromResult(new GetObjectResponse { ResponseStream = new MemoryStream(Objects[key]) });
        }

        public override async Task<PutObjectResponse> PutObjectAsync(PutObjectRequest request,
            CancellationToken cancellationToken = default)
        {
            using var bytes = new MemoryStream();
            if (request.InputStream != null)
                await request.InputStream.CopyToAsync(bytes, cancellationToken);
            Objects[request.Key] = request.InputStream == null
                ? Encoding.UTF8.GetBytes(request.ContentBody) : bytes.ToArray();
            return new PutObjectResponse();
        }

        public override Task<CopyObjectResponse> CopyObjectAsync(CopyObjectRequest request,
            CancellationToken cancellationToken = default)
        {
            Objects[request.DestinationKey] = Objects[request.SourceKey].ToArray();
            return Task.FromResult(new CopyObjectResponse());
        }

        public override Task<DeleteObjectResponse> DeleteObjectAsync(DeleteObjectRequest request,
            CancellationToken cancellationToken = default)
        {
            Objects.Remove(request.Key);
            return Task.FromResult(new DeleteObjectResponse());
        }
    }
}
