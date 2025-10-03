namespace DocxTemplater;

public class AssembleResponse
{
    public string DestinationKey { get; set; }
    public string? ID { get; set; }
    public string? OutputName { get; set; }
    public bool Success { get; set; }
    public string? ErrorMessage { get; set; }
    public string? ResultKey { get; set; }
    public string? InterimResultKey { get; set; }

    public AssembleResponse(AssembleRequest req, bool success)
  {
    Success = success;
    DestinationKey = req.DestinationKey;
    ID = req.ID;
    OutputName = req.OutputName;
  }
}
