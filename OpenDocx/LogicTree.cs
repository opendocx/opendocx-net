using DocumentFormat.OpenXml.Office2010.Excel;
using OpenDocx;
using OpenXmlPowerTools;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenDocx
{
    public static class LogicTree
    {
        public static List<FieldLogicNode> BuildLogicTree(List<ParsedField> astBody)
        {
            // return a copy of astBody with all (or at least some) logically insignificant nodes pruned out:
            // remove EndIf and EndList nodes
            // remove Content nodes that are already defined in the same logical/list scope
            // always process down all if branches & lists
            // leave field ID metadata in place so we can still detect error locations
            var copy = ReduceContentArray(astBody);
            SimplifyContentArray2(copy);
            SimplifyContentArray3(copy);
            return copy;
        }

        private static List<FieldLogicNode> ReduceContentArray(
            List<ParsedField> astBody,
            List<FieldLogicNode> newBody = null,
            LogicScope scope = null,
            LogicScope parentScope = null)
        {
            // prune EndIf and EndList nodes (only important insofar as we need to match up nodes to fields --
            //   which will not be the case with a reduced logic tree)
            // prune redundant Content nodes that are already defined in the same logical & list scope
            // always process down all if branches & lists
            //    but check whether each if expression is the first (in its scope) to refer to the expression,
            //    and if so, indicate it on the node
            // future: compare logical & list scopes of each item, and eliminate logical branches
            //    and list iterations that are redundant?
            newBody ??= new();
            scope ??= new LogicScope();
            foreach (var obj in astBody)
            {
                var newObj = ReduceContentNode(obj, scope, parentScope);
                if (newObj != null)
                {
                    newBody.Add(newObj);
                }
            }
            return newBody;
        }

        private static FieldLogicNode ReduceContentNode(
            ParsedField astNode,
            LogicScope scope,
            LogicScope parentScope = null)
        {
            if (astNode.Type == FieldType.EndList)
            {
                // can we find the matching List field, to set its .endid = astNode.id before returning?
                return null;
            }
            if (astNode.Type == FieldType.EndIf)
            {
                // can we find the initial matching If field (not ElseIf or Else), to set its .endid = astNode.id before returning?
                return null;
            }
            if (astNode.Type == FieldType.Content)
            {
                if (astNode.Number == 0 && astNode.Expression == "_punc") // _punc nodes do not affect logic!
                {
                    return null;
                }
                if (scope.TryGetValue(astNode.Expression, out var existing) && existing != null)
                {
                    if (astNode.Number > 0)
                    {
                        existing.AddField(astNode.Number);
                    }
                    return null;
                }
                var copyOfNode = new FieldLogicNode(astNode);
                scope[astNode.Expression] = copyOfNode;
                return copyOfNode;
            }
            if (astNode.Type == FieldType.List)
            {
                if (scope.TryGetValue(astNode.Expression, out var existing) && existing != null)
                {
                    // this list has already been added to the parent scope; revisit it to add more content members if necessary
                    ReduceContentArray(astNode.ContentArray, existing.Content, existing.Scope);
                    if (astNode.Number > 0)
                    {
                        existing.AddField(astNode.Number);
                    }
                    return null;
                }
                else
                {
                    var newScope = new LogicScope(); // fresh new wholly separate scope for lists
                    var newContent = ReduceContentArray(astNode.ContentArray, null, newScope);
                    var copyOfNode = new FieldLogicNode(astNode, newContent, newScope);
                    scope[astNode.Expression] = copyOfNode; // set BEFORE recursion for consistent results?  (or is it intentionally after?)
                    return copyOfNode;
                }
            }
            if (astNode.Type == FieldType.If || astNode.Type == FieldType.ElseIf || astNode.Type == FieldType.Else)
            {
                // if's are always left in at this point (because of their importance to the logic;
                // a lot more work would be required to accurately optimize them out.)
                // HOWEVER, we can't add the expr to the parent scope by virtue of it having been referenced in a condition,
                // because it means something different for the same expression to be evaluated
                // in a Content node vs. an If/ElseIf node, and therefore an expression emitted/evaluated as part of a condition
                // still needs to be emitted/evaluated as part of a merge/content node.
                // AND VICE VERSA: an expression emitted as part of a content node STILL needs to be emitted as part of a condition,
                // too.
                var copyOfNode = new FieldLogicNode(astNode);

                var pscope = (parentScope != null) ? parentScope : scope;
                // this 'parentScope' thing is a bit tricky.  The parentScope argument is only supplied
                // when we're inside an If/ElseIf/Else block within the current scope.
                // If supplied, it INDIRECTLY refers to the actual scope -- basically, successive layers of "if" blocks
                // that each establish a new "mini" scope, that has the parent scope as its prototype.
                // This means, a reference to an identifier in a parent scope, will prevent that identifier from
                // subsequently appearing (redundantly) in a child; but a reference to an identifier in a child scope,
                // must NOT prevent that identifier from appearing subsequently in a parent scope.
                if (copyOfNode.Type == FieldType.If || copyOfNode.Type == FieldType.ElseIf)
                {
                    if (!pscope.ContainsKey("if$" + astNode.Expression))
                    {
                        pscope["if$" + astNode.Expression] = copyOfNode;
                    }
                }
                var childContext = new LogicScope(pscope);
                copyOfNode.Content = ReduceContentArray(astNode.ContentArray, null, childContext, pscope);
                return copyOfNode;
            }
            throw new FieldParseException("Unexpected ast node type");
        }

        private static void SimplifyContentArray2(List<FieldLogicNode> astBody)
        {
            // 2nd pass at simplifying logic
            // for now, just clean up scope helpers leftover from first pass
            foreach (FieldLogicNode obj in astBody)
            {
                SimplifyNode2(obj);
            }
        }

        private static void SimplifyNode2(FieldLogicNode astNode)
        {
            if (astNode.Scope != null)
            {
                astNode.Scope = null;
            }
            if (astNode.Content != null)
            {
                SimplifyContentArray2(astNode.Content);
            }
        }

        private static void SimplifyContentArray3(List<FieldLogicNode> body, LogicScope scope = null)
        {
            if (scope == null) { scope = new LogicScope(); }
            // 3rd pass at simplifying scopes
            var initialScope = scope.Copy(); // shallow-clone the scope to start with
            // first go through content fields
            var i = 0;
            while (i < body.Count)
            {
                var field = body[i];
                var nodeRemoved = false;
                if (field.Type == FieldType.Content)
                {
                    if (scope.TryGetValue(field.Expression, out var existing))
                    {
                        if (field.FirstField > 0)
                        {
                            existing.AddField(field.FirstField);
                        }
                        body.RemoveAt(i);
                        nodeRemoved = true;
                    }
                    else
                    {
                        scope[field.Expression] = field;
                    }
                }
                if (!nodeRemoved) i++;
            }
            // then recurse into ifs and lists
            foreach (var field in body)
            {
                if (field.Type == FieldType.List)
                {
                    if (!scope.ContainsKey(field.Expression))
                    {
                        scope[field.Expression] = field;
                    }
                    SimplifyContentArray3(field.Content, new LogicScope()); // new scope for lists
                }
                else if (field.Type == FieldType.If)
                {
                    // the content in an if block has everything in its parent scope
                    SimplifyContentArray3(field.Content, scope.Copy()); // copy the parent scope
                }
                else if (field.Type == FieldType.ElseIf || field.Type == FieldType.Else)
                {
                    // elseif and else fields are (in the logic tree) children of ifs,
                    // but they do NOT have access to the parent scope; reset to initial scope for if
                    SimplifyContentArray3(field.Content, initialScope.Copy());
                }
            }
            // note: although this will eliminate some redundant fields, it will not eliminate all of them.
            // This code was translated from the original JavaScript (in project Yatte) and some major
            // rework will be required to implement a new, more straight-forward, and more capable approach.
        }
    }
}
