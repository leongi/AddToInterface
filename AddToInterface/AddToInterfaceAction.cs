using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using Microsoft.VisualStudio.Imaging.Interop;
using Microsoft.VisualStudio.Language.Intellisense;
using Microsoft.VisualStudio.Text;
using EnvDTE;
using System.Linq;

namespace TestLightBulb
{
    internal class AddToInterfaceAction : ISuggestedAction
    {
        private readonly CodeInterface _interface;
        private readonly CodeFunction _srcFunction;

        public AddToInterfaceAction(CodeInterface codeInterface, CodeFunction codeFunction)
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
                return true;
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
