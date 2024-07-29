using System.IO;

namespace OpenDocx
{
    public class AssembleResult
    {
        public byte[] Bytes { get; }
        public string Error { get; }
        public bool HasErrors { get => string.IsNullOrEmpty(Error); }

        internal AssembleResult(string documentFilename, string error = null)
        {
            Bytes = File.ReadAllBytes(documentFilename);
            Error = error;
        }

        internal AssembleResult(byte[] document, string error = null)
        {
            Bytes = document;
            Error = error;
        }
    }
}
