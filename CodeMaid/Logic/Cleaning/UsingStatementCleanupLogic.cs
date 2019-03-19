// --------------------------------------------------------------------------------------------------------------------
// <copyright file="SwssNotificationRequestHandler.cs" company="Microsoft">
//   Copyright (c) Microsoft Corporation. All rights reserved.
// </copyright>
// --------------------------------------------------------------------------------------------------------------------

namespace SteveCadwallader.CodeMaid.Logic.Cleaning
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using EnvDTE;
    using SteveCadwallader.CodeMaid.Helpers;
    using SteveCadwallader.CodeMaid.Model.CodeItems;
    using SteveCadwallader.CodeMaid.Properties;

    /// <summary>
    /// A class for encapsulating using statement cleanup logic.
    /// </summary>
    internal class UsingStatementCleanupLogic
    {
        #region Fields

        private readonly CodeMaidPackage _package;
        private readonly CommandHelper _commandHelper;

        private readonly CachedSettingSet<string> _usingStatementsToReinsertWhenRemoved =
            new CachedSettingSet<string>(() => Settings.Default.Cleaning_UsingStatementsToReinsertWhenRemovedExpression,
                                         expression =>
                                         expression.Split(new[] { "||" }, StringSplitOptions.RemoveEmptyEntries)
                                                   .Select(x => x.Trim())
                                                   .Where(y => !string.IsNullOrEmpty(y))
                                                   .ToList());

        #endregion Fields

        #region Constructors

        /// <summary>
        /// The singleton instance of the <see cref="UsingStatementCleanupLogic" /> class.
        /// </summary>
        private static UsingStatementCleanupLogic _instance;

        /// <summary>
        /// Gets an instance of the <see cref="UsingStatementCleanupLogic" /> class.
        /// </summary>
        /// <param name="package">The hosting package.</param>
        /// <returns>An instance of the <see cref="UsingStatementCleanupLogic" /> class.</returns>
        internal static UsingStatementCleanupLogic GetInstance(CodeMaidPackage package)
        {
            return _instance ?? (_instance = new UsingStatementCleanupLogic(package));
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="UsingStatementCleanupLogic" /> class.
        /// </summary>
        /// <param name="package">The hosting package.</param>
        private UsingStatementCleanupLogic(CodeMaidPackage package)
        {
            _package = package;

            _commandHelper = CommandHelper.GetInstance(_package);
        }

        #endregion Constructors

        #region Methods

        /// <summary>
        /// Run the visual studio built-in remove and sort using statements command.
        /// </summary>
        /// <remarks>
        /// Before VS2017 these were two separate commands.  Starting in VS2017 they were merged into one.
        /// </remarks>
        /// <param name="textDocument">The text document to update.</param>
        public void RemoveAndSortUsingStatements(TextDocument textDocument)
        {
            if (!Settings.Default.Cleaning_RunVisualStudioRemoveAndSortUsingStatements) return;
            if (_package.IsAutoSaveContext && Settings.Default.Cleaning_SkipRemoveAndSortUsingStatementsDuringAutoCleanupOnSave) return;

            // Capture all existing using statements that should be re-inserted if removed.
            const string patternFormat = @"^[ \t]*{0}[ \t]*\r?\n";

            var points = (from usingStatement in _usingStatementsToReinsertWhenRemoved.Value
                          from editPoint in TextDocumentHelper.FindMatches(textDocument, string.Format(patternFormat, usingStatement))
                          select new { editPoint, text = editPoint.GetLine() }).Reverse().ToList();

            // Shift every captured point one character to the right so they will auto-advance
            // during new insertions at the start of the line.
            foreach (var point in points)
            {
                point.editPoint.CharRight();
            }

            if (_package.IDEVersion >= 15)
            {
                _commandHelper.ExecuteCommand(textDocument, "EditorContextMenus.CodeWindow.RemoveAndSort");
            }
            else
            {
                _commandHelper.ExecuteCommand(textDocument, "Edit.RemoveUnusedUsings");
                _commandHelper.ExecuteCommand(textDocument, "Edit.SortUsings");
            }

            // Check each using statement point and re-insert it if removed.
            foreach (var point in points)
            {
                string text = point.editPoint.GetLine();
                if (text != point.text)
                {
                    point.editPoint.StartOfLine();
                    point.editPoint.Insert(point.text);
                    point.editPoint.Insert(Environment.NewLine);
                }
            }
        }

        /// <summary>
        /// Sorts all using statements in ascending order, with System using statements on top.
        /// </summary>
        /// <param name="usingStatementsItems">List of using Statement codeItems</param>
        /// <param name="namespaceItems">List of namespace codeItems</param>
        internal void MoveUsingStatementsWithinNamespace(List<CodeItemUsingStatement> usingStatementsItems, List<CodeItemNamespace> namespaceItems)
        {
            if (namespaceItems.Count != 1)
            {
                //We return back as is, if multiple namespaces are found.
                return;
            }

            CodeItemNamespace theOnlyNamespace = namespaceItems.First();

            EditPoint namespaceInsertCursor = theOnlyNamespace.StartPoint;

            // Setting the start point where we will start inserting using statements.
            namespaceInsertCursor.LineDown();
            namespaceInsertCursor.CharRight();
            namespaceInsertCursor.Insert(Environment.NewLine);

            //Sort the using code items in ascending string order, with system usings on top.
            usingStatementsItems.Sort((usingStatement1Item, usingStatement2Item) =>
            {
                string textOfUsingStatement1 = usingStatement1Item.StartPoint.GetText(usingStatement1Item.EndPoint);
                string textOfUsingStatement2 = usingStatement2Item.StartPoint.GetText(usingStatement2Item.EndPoint);

                var referenceNameOfStatement1 = ExtractUsingStatementReferenceName(textOfUsingStatement1);
                var referenceNameOfStatement2 = ExtractUsingStatementReferenceName(textOfUsingStatement2);

                if (IsSystemReference(referenceNameOfStatement1) && !IsSystemReference(referenceNameOfStatement2))
                {
                    return -1;
                }
                else if (!IsSystemReference(referenceNameOfStatement1) && IsSystemReference(referenceNameOfStatement2))
                {
                    return 1;
                }
                else
                {
                    return string.Compare(referenceNameOfStatement1, referenceNameOfStatement2);
                }
            });

            foreach (var usingStatement in usingStatementsItems)
            {
                var startPoint = usingStatement.StartPoint;
                var endPoint = usingStatement.EndPoint;

                string text = startPoint.GetText(usingStatement.EndPoint);
                startPoint.Delete(usingStatement.EndPoint);

                namespaceInsertCursor.Insert(text);
                namespaceInsertCursor.Indent(Count: 1);
                namespaceInsertCursor.Insert(Environment.NewLine);
            }
        }

        /// <summary>
        /// In a using statement like "using System.Threading.Tasks", extracts the reference name, i.e. "System.Threading.Tasks" in this case.
        /// </summary>
        /// <param name="text">the using statement complete text</param>
        /// <returns>As described above</returns>
        private static string ExtractUsingStatementReferenceName(string text)
        {
            string firstNonUsingToken = "";
            foreach (string token in text.Split(' '))
            {
                if (token.Equals("using") || token.Length == 0)
                {
                    continue;
                }
                else
                {
                    firstNonUsingToken = token;
                    break;
                }
            }

            return firstNonUsingToken.TrimEnd(';');
        }

        /// <summary>
        /// Determines if the using statement reference is a system reference.
        /// </summary>
        /// <param name="referenceNameStatement">Reference name string</param>
        /// <returns>boolean value determining above</returns>
        private static bool IsSystemReference(string referenceNameStatement)
        {
            return referenceNameStatement.StartsWith("System.") || referenceNameStatement.Equals("System");
        }

        #endregion Methods
    }
}