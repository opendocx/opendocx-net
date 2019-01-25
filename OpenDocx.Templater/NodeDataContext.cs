﻿using System;
using System.Linq;
using System.Threading.Tasks;

namespace OpenDocx
{
    public class NodeDataContext : MetadataParser, IAsyncDataContext
    {
        private Func<object, Task<object>> _evaluateText;
        private Func<object, Task<object>> _evaluateBool;
        private Func<object, Task<object>> _evaluateList;
        private Func<object, Task<object>> _releaseContext;
        private string _contextId;

        public NodeDataContext(dynamic options)
        {
            _evaluateText = (Func<object, Task<object>>)options.evaluateText;
            _evaluateBool = (Func<object, Task<object>>)options.evaluateBool;
            _evaluateList = (Func<object, Task<object>>)options.evaluateList;
            _releaseContext = (Func<object, Task<object>>)options.releaseContext;
            _contextId = String.Empty;
        }

        public NodeDataContext(NodeDataContext parent, string contextId)
        {
            _evaluateText = parent._evaluateText;
            _evaluateBool = parent._evaluateBool;
            _evaluateList = parent._evaluateList;
            _releaseContext = parent._releaseContext;
            _contextId = contextId;
        }

        public async Task<string> EvaluateTextAsync(string selector, bool optional)
        {
            Console.WriteLine("DN: OpenDocx.NodeDataContext.EvaluateTextAsync called for '{0}' from context '{1}'...", selector, _contextId);
            try
            {
                var payload = new { contextId = _contextId, expr = selector };
                var result = (string)await _evaluateText(payload);
                Console.WriteLine("DN: OpenDocx.NodeDataContext.EvaluateTextAsync is returning '{0}'", result);
                return result;
            }
            catch (Exception e)
            {
                Console.WriteLine("DN: OpenDocx.NodeDataContext.EvaluateTextAsync error: '{0}'", e.Message);
                throw new EvaluationException("EvaluationException: " + e.Message, e);
            }
        }

        public async Task<bool> EvaluateBoolAsync(string selector, string match, string notMatch)
        {
            Console.WriteLine("DN: OpenDocx.NodeDataContext.EvaluateBoolAsync called for '{0}' from context '{1}'...", selector, _contextId);
            try
            {
                var payload = new { contextId = _contextId, expr = selector };
                var result = (bool)await _evaluateBool(payload);
                Console.WriteLine("DN: OpenDocx.NodeDataContext.EvaluateBoolAsync is returning '{0}'", result);
                return result;
            }
            catch (Exception e)
            {
                Console.WriteLine("DN: OpenDocx.NodeDataContext.EvaluateBoolAsync error: '{0}'", e.Message);
                throw new EvaluationException("EvaluationException: " + e.Message, e);
            }
        }

        public async Task<IAsyncDataContext[]> EvaluateListAsync(string selector, bool optional)
        {
            Console.WriteLine("DN: OpenDocx.NodeDataContext.EvaluateListAsync called for '{0}' from context '{1}'...", selector, _contextId);
            try
            {
                var payload = new { contextId = _contextId, expr = selector };
                dynamic result = await _evaluateList(payload);
                Console.WriteLine("NodeDataContext: await _evaluateList returned '{0}'", result.ToString());
                if (result.GetType().IsArray)
                {
                    var oary = (object[]) result;
                    Console.WriteLine("DN: OpenDocx.NodeDataContext.EvaluateListAsync returning contexts '{0}'", oary.Aggregate("", (acc, item) => acc + item.ToString() + ", "));
                    return oary.Select(contextId => new NodeDataContext(this, (string)contextId)).ToArray();
                }
                throw new EvaluationException("_evaluateList result is not an array");
            }
            catch (Exception e)
            {
                throw new EvaluationException("EvaluationException: " + e.Message, e);
            }
        }

        public async Task ReleaseAsync()
        {
            Console.WriteLine("DN: OpenDocx.NodeDataContext.ReleaseAsync called for context '{0}'...", _contextId);
            var actuallyDisposed = (bool)await _releaseContext(_contextId);
            Console.WriteLine("DN: OpenDocx.NodeDataContext.ReleaseAsync completed; {0}", actuallyDisposed ? "disposed" : "not disposed (still referenced)");
        }
    }

}

