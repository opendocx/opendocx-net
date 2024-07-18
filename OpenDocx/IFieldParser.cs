/***************************************************************************

Copyright (c) Lowell Stewart 2018-2019.
Licensed under the Mozilla Public License. See LICENSE file in the project root for full license information.

Published at https://github.com/opendocx/opendocx
Developer: Lowell Stewart
Email: lowell@opendocx.com

***************************************************************************/

using System;
using System.Collections;
using System.Collections.Generic;
using System.Text.Json.Serialization;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace OpenDocx
{
    public enum FieldType
    {
        Content = 1,
        If,
        ElseIf,
        Else,
        EndIf,
        List,
        EndList,
        Insert
    }

    public class ParsedField : IFieldTransformInfo
    {
        [JsonPropertyName("fieldType")]
        public FieldType Type { get; set; }

        [JsonPropertyName("expr")]
        public string Expression { get; set; }

        [JsonPropertyName("atomizedExpr")]
        public string Atom { get; set; }

        [JsonPropertyName("parent")]
        public uint ParentNumber { get; set; }

        [JsonIgnore]
        public uint Number { get; set; }

        [JsonIgnore]
        public string Comment { get; set; }

        [JsonIgnore]
        internal List<ParsedField> ContentArray { get; set; } // only for temporary use when parsing Else and ElseIf fields

        [JsonIgnore]
        private string Prefix
        {
            get
            {
                switch (Type)
                {
                    case FieldType.Content:
                        return string.Empty;
                    case FieldType.If:
                        return "if ";
                    case FieldType.EndIf:
                        return "endif";
                    case FieldType.Else:
                        return "else";
                    case FieldType.ElseIf:
                        return "elseif ";
                    case FieldType.List:
                        return "list ";
                    case FieldType.EndList:
                        return "endlist";
                }
                throw new FieldParseException("Unexpected field type");
            }
        }

        [JsonIgnore]
        public string Text
        {
            get => Prefix + (
                string.IsNullOrWhiteSpace(Expression) ? string.Empty : Expression.Trim()
            );
        }

        [JsonIgnore]
        public string Content { get => Prefix + Atom; }
    }

    public class FieldLogicNode
    {
        [JsonPropertyName("type")]
        public FieldType Type { get; set; }

        [JsonPropertyName("expr")]
        public string Expression { get; set; }

        [JsonPropertyName("atom")]
        public string Atom { get; set; }

        [JsonPropertyName("id")]
        public uint FirstField { get; set; }

        [JsonPropertyName("idd")]
        public List<uint> OtherFields { get; set; }

        [JsonPropertyName("contentArray")]
        public List<FieldLogicNode> Content { get; set; }

        [JsonIgnore]
        internal LogicScope Scope { get; set; }

        internal FieldLogicNode(ParsedField field, List<FieldLogicNode> content = null, LogicScope scope = null)
        {
            Type = field.Type;
            Expression = field.Expression;
            Atom = field.Atom;
            FirstField = field.Number;
            Content = content;
            Scope = scope;
        }

        public void AddField(uint fieldNum)
        {
            if (OtherFields == null)
            {
                OtherFields = new List<uint>();
            }
            OtherFields.Add(fieldNum);
        }
    }

    public class FieldParseException : Exception
    {
        public FieldParseException() { }
        public FieldParseException(string message) : base(message) { }
        public FieldParseException(string message, Exception inner) : base(message, inner) { }
    }

}
