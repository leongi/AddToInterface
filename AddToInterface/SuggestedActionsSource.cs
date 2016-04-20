using Microsoft.VisualStudio.Language.Intellisense;
using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.Text.Editor;
using Microsoft.VisualStudio.Text.Operations;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TestLightBulb;
using System.Threading;
using EnvDTE;

namespace AddToInterface
{
    internal class SuggestedActionsSource : ISuggestedActionsSource
    {
        private readonly AddToInterfaceActionsSourceProvider _factory;
        private readonly ITextBuffer _textBuffer;
        private readonly ITextView _textView;
        private List<AddToInterfaceAction> _currentActions = new List<AddToInterfaceAction>();

        EnvDTE.DTE dte = (EnvDTE.DTE)Microsoft.VisualStudio.Shell.Package.GetGlobalService(typeof(EnvDTE.DTE));

        public SuggestedActionsSource(AddToInterfaceActionsSourceProvider testSuggestedActionsSourceProvider, ITextView textView, ITextBuffer textBuffer)
        {
            _factory = testSuggestedActionsSourceProvider;
            _textBuffer = textBuffer;
            _textView = textView;
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

        public Task<bool> HasSuggestedActionsAsync(ISuggestedActionCategorySet requestedActionCategories, SnapshotSpan range, CancellationToken cancellationToken)
        {
            return Task.Factory.StartNew(() =>
            {
                if (Monitor.TryEnter(_currentActions))
                {
                    try
                    {
                        TextExtent extent;
                        if (TryGetWordUnderCaret(out extent) && extent.IsSignificant)
                        {
                            var activeDocument = dte.ActiveDocument;
                            TextSelection textSelection = (TextSelection)activeDocument.Selection;
                            TextPoint point = (TextPoint)textSelection.ActivePoint;
                            FileCodeModel fileCodeModel = activeDocument.ProjectItem.FileCodeModel;


                            //! Can throw exception
                            CodeFunction srcFunction = fileCodeModel.CodeElementFromPoint(point, vsCMElement.vsCMElementFunction) as CodeFunction;


                            // Function validation
                            bool isFuncRegular = srcFunction.FunctionKind == vsCMFunction.vsCMFunctionFunction;
                            bool isPublicAccess = srcFunction.Access == vsCMAccess.vsCMAccessPublic;
                            bool isCursorOnDefinition = point.Line == srcFunction.StartPoint.Line;

                            if (!isFuncRegular || !isPublicAccess || !isCursorOnDefinition)
                            {
                                return false;
                            }


                            // Interfaces scan
                            CodeClass codeClass = srcFunction.Parent as CodeClass;
                            var interfaces = codeClass.ImplementedInterfaces.Cast<CodeInterface>().ToList();
                            bool isFoundInAny = interfaces.Any(i => IsFoundInInterface(srcFunction, i));


                            // Suggestions addition
                            if (interfaces.Count > 0 && !isFoundInAny)
                            {
                                _currentActions.Clear();

                                foreach (var item in interfaces)
                                {
                                    _currentActions.Add(new AddToInterfaceAction(item, srcFunction));
                                }

                                return true;
                            }
                        }
                    }
                    catch
                    { }
                    finally
                    {
                        Monitor.Exit(_currentActions);
                    }
                }
                return false;
            });
        }

        public IEnumerable<SuggestedActionSet> GetSuggestedActions(ISuggestedActionCategorySet requestedActionCategories, SnapshotSpan range, CancellationToken cancellationToken)
        {
            TextExtent extent;
            if (TryGetWordUnderCaret(out extent) && extent.IsSignificant)
            {
                return new List<SuggestedActionSet> { new SuggestedActionSet(_currentActions) };
            }
            else
            {
                return Enumerable.Empty<SuggestedActionSet>();
            }
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

        private bool IsFoundInInterface(CodeFunction function, CodeInterface codeInterface)
        {
            foreach (CodeFunction destFunc in codeInterface.Children)
            {
                if (destFunc != null)
                {
                    if (function.Name == destFunc.Name)
                    {
                        if (function.Parameters.Count == 0 && destFunc.Parameters.Count == 0)
                        {
                            return true;
                        }

                        int count = 0;
                        if (function.Parameters.Count == destFunc.Parameters.Count)
                        {
                            List<CodeParameter> srcParams = function.Parameters.Cast<CodeParameter>().ToList();
                            List<CodeParameter> destParams = destFunc.Parameters.Cast<CodeParameter>().ToList();

                            for (int i = 0; i < srcParams.Count; i++)
                            {
                                if (destParams[i].Type.AsFullName == srcParams[i].Type.AsFullName)
                                {
                                    count++;
                                }
                                else
                                {
                                    break;
                                }
                            }

                            if (srcParams.Count == count)
                            {
                                return true;
                            }
                        }
                    }
                }
            }

            return false;
        }
    }
}
