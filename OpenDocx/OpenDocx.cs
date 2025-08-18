using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace OpenDocx
{
    public class IndirectSource
    {
        public string ID { get; set; }
        public byte[] Bytes { get; set; }
        public bool KeepSections { get; set; }

        public IndirectSource(string id, byte[] bytes, bool keepSections = false)
        {
            ID = id;
            Bytes = bytes;
            KeepSections = keepSections;
        }
    }

    public static class OpenDocx
    {
        public static readonly JsonSerializerOptions DefaultJsonOptions = new()
        {
            WriteIndented = false,
            DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingDefault,
            Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping,
            Converters =
                    {
                        new JsonStringEnumConverter()
                    },
        };
    }
}
