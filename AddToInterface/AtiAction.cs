using EnvDTE;
using Microsoft.VisualStudio.Imaging.Interop;
using Microsoft.VisualStudio.Language.Intellisense;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace AddToInterface
{
    internal class AtiAction : ISuggestedAction
    {
        private readonly CodeInterface _interface;
        private readonly CodeFunction _srcFunction;

        public AtiAction(CodeInterface codeInterface, CodeFunction codeFunction)
        {
            _interface = codeInterface;
            _srcFunction = codeFunction;
        }

        public string DisplayText
        {
            get
            {
                return $"Add '{_srcFunction.Name}' to {_interface.Name} interface";
            }
        }

        public string IconAutomationText
        {
            get
            {
                return null;
            }
        }

        ImageMoniker ISuggestedAction.IconMoniker
        {
            get
            {
                return default(ImageMoniker);
            }
        }

        public string InputGestureText
        {
            get
            {
                return null;
            }
        }

        public bool HasActionSets
        {
            get
            {
                return false;
            }
        }

        public Task<IEnumerable<SuggestedActionSet>> GetActionSetsAsync(CancellationToken cancellationToken)
        {
            return null;
        }

        public bool HasPreview
        {
            get
            {
                return false;
            }
        }

        public Task<object> GetPreviewAsync(CancellationToken cancellationToken)
        {
            return Task.FromResult<object>(null);
        }

        public void Dispose()
        {
        }

        public void Invoke(CancellationToken cancellationToken)
        {
            if (cancellationToken.IsCancellationRequested)
            {
                return;
            }

            try
            {
                CodeFunction addedFunction = _interface.AddFunction(_srcFunction.Name, vsCMFunction.vsCMFunctionFunction, _srcFunction.Type, -1);

                var parameters = _srcFunction.Parameters.Cast<CodeParameter>().Reverse().ToList();

                foreach (var parameter in parameters)
                {
                    addedFunction.AddParameter(parameter.Name, parameter.Type);
                }
            }
            catch { }
        }

        public bool TryGetTelemetryId(out Guid telemetryId)
        {
            telemetryId = Guid.Empty;
            return false;
        }
    }
}
