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
        private readonly CodeFunction _function;

        public AddToInterfaceAction(CodeInterface codeInterface, CodeFunction codeFunction)
        {
            _interface = codeInterface;
            _function = codeFunction;
        }

        public string DisplayText
        {
            get
            {
                return string.Format("Add '{0}' to {1} interface", _function.Name, _interface.Name);
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
            var textBlock = new TextBlock();
            textBlock.Padding = new Thickness(5);
            textBlock.Inlines.Add(new Run() { Text = string.Empty });
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

            CodeFunction func = _interface.AddFunction(_function.Name, vsCMFunction.vsCMFunctionFunction, _function.Type, -1);

            IEnumerable<CodeParameter> parameters = _function.Parameters.Cast<CodeParameter>().Reverse();

            foreach (var parameter in parameters)
            {
                func.AddParameter(parameter.Name, parameter.Type);
            }
        }

        public bool TryGetTelemetryId(out Guid telemetryId)
        {
            telemetryId = Guid.Empty;
            return false;
        }
    }
}
