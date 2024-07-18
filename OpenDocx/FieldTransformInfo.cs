using System;
using System.Collections.Generic;

namespace OpenDocx
{
    public interface IFieldTransformInfo
    {
        public string Content { get; }
    }

    public class FieldTransformInfo : IFieldTransformInfo
    {
        public string fieldType;
        public string atomizedExpr;

        private string Prefix {
            get {
                switch (fieldType) {
                    case "Content":
                        return string.Empty;
                    case "If":
                        return "if ";
                    case "EndIf":
                        return "endif";
                    case "Else":
                        return "else";
                    case "ElseIf":
                        return "elseif ";
                    case "List":
                        return "list ";
                    case "EndList":
                        return "endlist";
                }
                throw new FieldParseException("Unexpected fieldType '" + fieldType + "'");
            }
        }

        public string Content {
            get {
                return Prefix + atomizedExpr;
            }
        }
    }

    public class FieldTransformIndex : Dictionary<string, FieldTransformInfo>
    {

    }
}
