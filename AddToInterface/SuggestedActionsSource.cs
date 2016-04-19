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
        private readonly AddToInterfaceActionsSourceProvider m_factory;
        private readonly ITextBuffer m_textBuffer;
        private readonly ITextView m_textView;
        private List<AddToInterfaceAction> _currentActions = new List<AddToInterfaceAction>();

        EnvDTE.DTE dte = (EnvDTE.DTE)Microsoft.VisualStudio.Shell.Package.GetGlobalService(typeof(EnvDTE.DTE));

        public SuggestedActionsSource(AddToInterfaceActionsSourceProvider testSuggestedActionsSourceProvider, ITextView textView, ITextBuffer textBuffer)
        {
            m_factory = testSuggestedActionsSourceProvider;
            m_textBuffer = textBuffer;
            m_textView = textView;
        }

        private bool TryGetWordUnderCaret(out TextExtent wordExtent)
        {
            ITextCaret caret = m_textView.Caret;
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

            ITextStructureNavigator navigator = m_factory.NavigatorService.GetTextStructureNavigator(m_textBuffer);

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
                            TextPoint pnt = (TextPoint)textSelection.ActivePoint;
                            int currentLine = pnt.Line;
                            FileCodeModel fcm = activeDocument.ProjectItem.FileCodeModel;
                            CodeFunction srcFunc;
                            try
                            {
                                srcFunc = fcm.CodeElementFromPoint(pnt, vsCMElement.vsCMElementFunction) as CodeFunction;

                                if (srcFunc.Access != vsCMAccess.vsCMAccessPublic)
                                {
                                    return false;
                                }

                                int startLine = srcFunc.StartPoint.Line;
                                if (currentLine == startLine)
                                {
                                    CodeClass cl = srcFunc.Parent as CodeClass;
                                    var interfaces = cl.ImplementedInterfaces;

                                    if (interfaces.Count > 0)
                                    {
                                        List<CodeInterface> typedInterfaces = interfaces.Cast<CodeInterface>().ToList();

                                        bool found = false;
                                        foreach (var item in typedInterfaces)
                                        {
                                            if (IsFoundInInterface(srcFunc, item))
                                            {
                                                found = true;
                                                break;
                                            }
                                        }

                                        if (found)
                                        {
                                            return false;
                                        }
                                        else
                                        {
                                            _currentActions.Clear();

                                            foreach (var item in typedInterfaces)
                                            {
                                                _currentActions.Add(new AddToInterfaceAction(item, srcFunc));
                                            }

                                            return true;
                                        }
                                    }
                                }
                            }
                            catch
                            {

                            }
                        }
                    }
                    finally
                    {
                        Monitor.Exit(_currentActions);
                    }
                }
                return false;
            });
        }

        private bool IsFoundInInterface(CodeFunction srcFunc, CodeInterface codeInterface)
        {
            foreach (CodeFunction destFunc in codeInterface.Children)
            {
                if (destFunc != null)
                {
                    if (srcFunc.Name == destFunc.Name)
                    {
                        if (srcFunc.Parameters.Count == 0 && destFunc.Parameters.Count == 0)
                        {
                            return true;
                        }

                        int count = 0;
                        if (srcFunc.Parameters.Count == destFunc.Parameters.Count)
                        {
                            List<CodeParameter> srcParams = srcFunc.Parameters.Cast<CodeParameter>().ToList();
                            List<CodeParameter> destParams = destFunc.Parameters.Cast<CodeParameter>().ToList();

                            for (int i = 0; i < srcParams.Count; i++)
                            {
                                if (destParams[i].Type.AsFullName == srcParams[i].Type.AsFullName)
                                {
                                    count++;
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
            // This is a sample provider and doesn't participate in LightBulb telemetry
            telemetryId = Guid.Empty;
            return false;
        }
    }
}
