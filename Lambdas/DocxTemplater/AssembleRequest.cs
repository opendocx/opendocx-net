namespace DocxTemplater;

public class AssembleRequest
{
    /// <summary>
    /// S3 bucket where templates and assembled documents and and will be stored
    /// </summary>
    public required string Bucket { get; set; }

    /// <summary>
    /// Key for OXPT DOCX template
    /// </summary>
    public required string TemplateKey { get; set; }

    /// <summary>
    /// XML data for assembly
    /// </summary>
    public required string Data { get; set; }

    /// <summary>
    /// Key (prefix) where sources are located and output will be stored
    /// </summary>
    public required string DestinationKey { get; set; }

    /// <summary>
    /// List of IDs for sources -- these keys are in the Data XML and are DOCX chunks already in DestinationKey
    /// </summary>
    public required List<string> Sources { get; set; }

    /// <summary>
    /// Anything that will be used as a source in the future, needs an ID -- intended to be a GUID
    /// </summary>
    public string? ID { get; set; }

    /// <summary>
    /// Anything that is a final document intended for delivery to a user, should have an OutputName
    /// </summary>
    public string? OutputName { get; set; }
}
