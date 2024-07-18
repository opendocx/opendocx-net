using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenDocx
{
    internal class LogicScope : IDictionary<string, FieldLogicNode>
    {
        private Dictionary<string, FieldLogicNode> This { get; set; }

        private LogicScope Parent { get; set; }

        public LogicScope(LogicScope parent = null)
        {
            This = new Dictionary<string, FieldLogicNode>();
            Parent = parent;
        }

        private LogicScope(Dictionary<string, FieldLogicNode> dictionary, LogicScope parent = null)
        {
            This = new Dictionary<string, FieldLogicNode>(dictionary);
            Parent = parent;
        }

        public LogicScope Copy()
        {
            return new LogicScope(This);
        }

        public FieldLogicNode this[string key] {
            set => ((IDictionary<string, FieldLogicNode>)This)[key] = value;
            get
            {
                if (Parent != null)
                {
                    if (This.ContainsKey(key))
                    {
                        return This[key];
                    }
                    return Parent[key];
                }
                return ((IDictionary<string, FieldLogicNode>)This)[key];
            }
        }

        public ICollection<string> Keys => (ICollection<string>)(Parent != null
            ? ((IDictionary<string, FieldLogicNode>)This).Keys.Concat(Parent.Keys)
            : ((IDictionary<string, FieldLogicNode>)This).Keys);

        public ICollection<FieldLogicNode> Values => (ICollection<FieldLogicNode>)(Parent != null
            ? ((IDictionary<string, FieldLogicNode>)This).Values.Concat(Parent.Values)
            : ((IDictionary<string, FieldLogicNode>)This).Values);

        public int Count => ((ICollection<KeyValuePair<string, FieldLogicNode>>)This).Count
            + (Parent != null ? Parent.Count : 0);

        public bool IsReadOnly => ((ICollection<KeyValuePair<string, FieldLogicNode>>)This).IsReadOnly;

        public void Add(string key, FieldLogicNode value)
        {
            ((IDictionary<string, FieldLogicNode>)This).Add(key, value);
        }

        public void Add(KeyValuePair<string, FieldLogicNode> item)
        {
            ((ICollection<KeyValuePair<string, FieldLogicNode>>)This).Add(item);
        }

        public void Clear()
        {
            ((ICollection<KeyValuePair<string, FieldLogicNode>>)This).Clear();
        }

        public bool Contains(KeyValuePair<string, FieldLogicNode> item)
        {
            return ((ICollection<KeyValuePair<string, FieldLogicNode>>)This).Contains(item)
                || (Parent != null && Parent.Contains(item));
        }

        public bool ContainsKey(string key)
        {
            return ((IDictionary<string, FieldLogicNode>)This).ContainsKey(key)
                || (Parent != null && Parent.ContainsKey(key));
        }

        public void CopyTo(KeyValuePair<string, FieldLogicNode>[] array, int arrayIndex)
        {
            ((ICollection<KeyValuePair<string, FieldLogicNode>>)This).CopyTo(array, arrayIndex);
        }

        public IEnumerator<KeyValuePair<string, FieldLogicNode>> GetEnumerator()
        {
            if (Parent != null)
            {
                return ((IEnumerable<KeyValuePair<string, FieldLogicNode>>)This).Union(Parent).GetEnumerator();
            }
            // else
            return ((IEnumerable<KeyValuePair<string, FieldLogicNode>>)This).GetEnumerator();
        }

        public bool Remove(string key)
        {
            return ((IDictionary<string, FieldLogicNode>)This).Remove(key);
        }

        public bool Remove(KeyValuePair<string, FieldLogicNode> item)
        {
            return ((ICollection<KeyValuePair<string, FieldLogicNode>>)This).Remove(item);
        }

        public bool TryGetValue(string key, [MaybeNullWhen(false)] out FieldLogicNode value)
        {
            return ((IDictionary<string, FieldLogicNode>)This).TryGetValue(key, out value)
                || (Parent != null && Parent.TryGetValue(key, out value));
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            // todo: how to coalesce this enumerator with the parent's?
            return ((IEnumerable)This).GetEnumerator();
        }
    }
}
