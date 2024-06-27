using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenDocx
{
    public class FieldExprNamer
    {
        private Dictionary<string, string> atomStore;

        public FieldExprNamer()
        {
            atomStore = new Dictionary<string, string>();
        }

        public string GetFieldAtom(ParsedField fieldObj)
        {
            var str = fieldObj.Expression ?? throw new Exception("Unexpected: cannot atomize a null string");
            if (atomStore.TryGetValue(str, out var value))
                return value;
            value = (fieldObj.Type == FieldType.List ? "L" : "C") + fieldObj.Number.ToString();
            atomStore.Add(str, value);
            return value;
        }
    }
}
