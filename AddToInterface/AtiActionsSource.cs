using EnvDTE;
using Microsoft.VisualStudio.Language.Intellisense;
using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.Text.Editor;
using Microsoft.VisualStudio.Text.Operations;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace AddToInterface
{
    internal class AtiActionsSource : ISuggestedActionsSource
    {
        private readonly AtiActionsSourceProvider _factory;
        private readonly ITextBuffer _textBuffer;
        private readonly ITextView _textView;
        private readonly List<AtiAction> _suggestedActions = new List<AtiAction>();
        private readonly DTE _dte = (DTE)Microsoft.VisualStudio.Shell.Package.GetGlobalService(typeof(DTE));

        public AtiActionsSource(AtiActionsSourceProvider testSuggestedActionsSourceProvider, ITextView textView, ITextBuffer textBuffer)
        {
            _factory = testSuggestedActionsSourceProvider;
            _textBuffer = textBuffer;
            _textView = textView;
        }

        public Task<bool> HasSuggestedActionsAsync(ISuggestedActionCategorySet requestedActionCategories, SnapshotSpan range, CancellationToken cancellationToken)
        {
            return Task.Factory.StartNew(() =>
            {
                lock (_suggestedActions)
                {
                    return HasSuggestedActions();
                }
            });
        }

        public IEnumerable<SuggestedActionSet> GetSuggestedActions(ISuggestedActionCategorySet requestedActionCategories, SnapshotSpan range, CancellationToken cancellationToken)
        {
            TextExtent extent;
            bool isSignificant = TryGetWordUnderCaret(out extent) && extent.IsSignificant;

            if (isSignificant)
            {
                lock (_suggestedActions)
                {
                    if (_suggestedActions.Count > 0)
                    {
                        return new List<SuggestedActionSet> { new SuggestedActionSet(_suggestedActions) };
                    }
                }
            }

            return Enumerable.Empty<SuggestedActionSet>();
        }

        public event EventHandler<EventArgs> SuggestedActionsChanged;

        public void Dispose()
        {
        }

        public bool TryGetTelemetryId(out Guid telemetryId)
        {
            telemetryId = Guid.Empty;
            return false;
        }

        private bool HasSuggestedActions()
        {
            _suggestedActions.Clear();

            try
            {
                TextExtent extent;
                if (TryGetWordUnderCaret(out extent) && extent.IsSignificant)
                {
                    var activeDocument = _dte.ActiveDocument;
                    TextSelection textSelection = (TextSelection)activeDocument.Selection;
                    TextPoint point = (TextPoint)textSelection.ActivePoint;

                    CodeFunction srcFunction;
                    if (TryGetFunctionFromPoint(point, out srcFunction))
                    {
                        // Function validation
                        bool isFuncRegular = (srcFunction.FunctionKind == vsCMFunction.vsCMFunctionFunction);
                        bool isPublicAccess = (srcFunction.Access == vsCMAccess.vsCMAccessPublic);
                        bool isCursorOnDefinition = (point.Line == srcFunction.StartPoint.Line);

                        if (!isFuncRegular || !isPublicAccess || !isCursorOnDefinition)
                        {
                            return false;
                        }

                        // Code scan
                        CodeClass codeClass = srcFunction.Parent as CodeClass;
                        if (codeClass != null)
                        {
                            // Interfaces scan
                            var interfaces = GetAllImplementedInterfaces(codeClass);
                            bool foundInAny = interfaces.Any(i => IsFoundInInterface(srcFunction, i));

                            // Suggestions addition
                            if (interfaces.Count > 0 && !foundInAny)
                            {
                                foreach (var item in interfaces)
                                {
                                    _suggestedActions.Add(new AtiAction(item, srcFunction));
                                }

                                return true;
                            }
                        }
                    }
                }
            }
            catch { }

            return false;
        }

        private bool TryGetFunctionFromPoint(TextPoint point, out CodeFunction codeFunction)
        {
            try
            {
                FileCodeModel codeModel = _dte.ActiveDocument.ProjectItem.FileCodeModel;
                codeFunction = codeModel.CodeElementFromPoint(point, vsCMElement.vsCMElementFunction) as CodeFunction;
                return true;
            }
            catch
            {
                codeFunction = null;
                return false;
            }
        }

        private bool IsFoundInInterface(CodeFunction function, CodeInterface codeInterface)
        {
            foreach (CodeElement destFunc in codeInterface.Children)
            {
                if (destFunc.Kind == vsCMElement.vsCMElementFunction)
                {
                    if (function.Name == destFunc.Name && ContainSameParameters(function, (CodeFunction)destFunc))
                    {
                        return true;
                    }
                }
            }

            return false;
        }

        private bool TryGetWordUnderCaret(out TextExtent wordExtent)
        {
            ITextCaret caret = _textView.Caret;
            SnapshotPoint point;

            if (caret.Position.BufferPosition > 0)
            {
                point = caret.Position.BufferPosition - 1;
            }
            else
            {
                wordExtent = default(TextExtent);
                return false;
            }

            ITextStructureNavigator navigator = _factory.NavigatorService.GetTextStructureNavigator(_textBuffer);

            wordExtent = navigator.GetExtentOfWord(point);
            return true;

        }

        private bool ContainSameParameters(CodeFunction func1, CodeFunction func2)
        {
            List<CodeParameter> func1Params = func1.Parameters.Cast<CodeParameter>().ToList();
            List<CodeParameter> func2Params = func2.Parameters.Cast<CodeParameter>().ToList();

            if (func1Params.Count != func2Params.Count)
            {
                return false;
            }

            for (int i = 0; i < func1Params.Count; i++)
            {
                if (func1Params[i].Type.AsFullName != func2Params[i].Type.AsFullName)
                {
                    return false;
                }
            }

            return true;
        }

        private List<CodeInterface> GetAllImplementedInterfaces(CodeClass codeClass)
        {
            var list = new List<CodeInterface>();

            foreach (CodeElement item in codeClass.ImplementedInterfaces)
            {
                if (item.Kind == vsCMElement.vsCMElementInterface)
                {
                    FillAllBaseInterfaces(list, (CodeInterface)item);
                }
            }

            return list;
        }

        private void FillAllBaseInterfaces(List<CodeInterface> list, CodeInterface codeInterface)
        {
            if (IsPartOfTheSolution((CodeElement)codeInterface))
            {
                list.Add(codeInterface);
            }

            foreach (CodeElement item in codeInterface.Bases)
            {
                if (item.Kind == vsCMElement.vsCMElementInterface)
                {
                    FillAllBaseInterfaces(list, (CodeInterface)item);
                }
            }
        }

        private bool IsPartOfTheSolution(CodeElement codeElement)
        {
            try
            {
                var projectItem = codeElement.ProjectItem;
                return true;
            }
            catch
            {
                return false;
            }
        }
    }
}
