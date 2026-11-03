' MailHeaderFooter.vb
' Outlook VSTO add-in (VB.NET)
'
' Inserts and updates a header and footer in the message being composed.
'
' - Each block lives in a borderless one-cell "container" table whose cell
'   shading is a near-white fingerprint color. The fingerprint survives
'   Outlook's HTML round-trip, so the block is found even after text edits.
' - Block content can be plain text, HTML (including tables), or built in code,
'   so a footer that itself contains a table is supported. Only TOP-LEVEL tables
'   are checked for the fingerprint; nested tables never interfere.
' - In replies/forwards, everything is limited to the ACTIVE message (above the
'   quoted history). The footer is placed just above the reply separator, and
'   fingerprinted blocks inside the quoted chain are never found or changed.
'
' References: Microsoft.Office.Interop.Outlook, Microsoft.Office.Interop.Word

Imports System.IO
Imports System.Text
Imports Outlook = Microsoft.Office.Interop.Outlook
Imports Word = Microsoft.Office.Interop.Word

Public Class MailHeaderFooter

    Public Enum BlockKind
        Header
        Footer
    End Enum

    ' Word colors are BGR integers: R + G*256 + B*65536
    Private Const HeaderColor As Integer = &HFEFFFF   ' #FFFFFE
    Private Const FooterColor As Integer = &HFFFFFE   ' #FEFFFF

    ' Last-applied source, kept for form prefill and for repair on send
    Private Const PropBase As String =
        "http://schemas.microsoft.com/mapi/string/{00020329-0000-0000-C000-000000000046}/"
    Private Const HeaderSourceProp As String = PropBase & "MyAddin.HeaderSource"
    Private Const FooterSourceProp As String = PropBase & "MyAddin.FooterSource"

    Private Const TextPrefix As String = "text:"
    Private Const HtmlPrefix As String = "html:"
    Private Const CustomMarker As String = "custom:"

    ' Outlook's hidden bookmark marking the quoted original in replies/forwards
    Private Const OriginalBookmark As String = "_MailOriginal"

    Private ReadOnly _mail As Outlook.MailItem
    Private ReadOnly _doc As Word.Document

    Public Sub New(mail As Outlook.MailItem, doc As Word.Document)
        If mail Is Nothing Then Throw New ArgumentNullException(NameOf(mail))
        If doc Is Nothing Then Throw New ArgumentNullException(NameOf(doc))
        _mail = mail
        _doc = doc
    End Sub

    ' =====================================================================
    ' Factories
    ' =====================================================================

    ''' <summary>For a message open in its own window.</summary>
    Public Shared Function ForInspector(insp As Outlook.Inspector) As MailHeaderFooter
        If insp Is Nothing Then Return Nothing
        Dim mail = TryCast(insp.CurrentItem, Outlook.MailItem)
        If mail Is Nothing OrElse insp.EditorType <> Outlook.OlEditorType.olEditorWord Then
            Return Nothing
        End If
        Dim doc = TryCast(insp.WordEditor, Word.Document)
        Return If(doc Is Nothing, Nothing, New MailHeaderFooter(mail, doc))
    End Function

    ''' <summary>For an inline reply in the reading pane.</summary>
    Public Shared Function ForInlineResponse(exp As Outlook.Explorer) As MailHeaderFooter
        If exp Is Nothing Then Return Nothing
        Dim mail = TryCast(exp.ActiveInlineResponse, Outlook.MailItem)
        If mail Is Nothing Then Return Nothing
        Dim doc = TryCast(exp.ActiveInlineResponseWordEditor, Word.Document)
        Return If(doc Is Nothing, Nothing, New MailHeaderFooter(mail, doc))
    End Function

    ''' <summary>
    ''' Finds the open editor for a specific mail item. Use this in ItemSend,
    ''' where the focused window isn't always the message being sent.
    ''' </summary>
    Public Shared Function ForMailItem(app As Outlook.Application,
                                       mail As Outlook.MailItem) As MailHeaderFooter
        ' Popped-out compose window
        For Each insp As Outlook.Inspector In app.Inspectors
            If insp.CurrentItem Is mail Then Return ForInspector(insp)
        Next

        ' Inline reply in a reading pane
        For Each exp As Outlook.Explorer In app.Explorers
            If exp.ActiveInlineResponse Is mail Then Return ForInlineResponse(exp)
        Next

        Return Nothing   ' no open editor (sent from code, Outbox, etc.)
    End Function

    ''' <summary>For whichever compose window has focus.</summary>
    Public Shared Function ForActiveCompose(app As Outlook.Application) As MailHeaderFooter
        Dim w = app.ActiveWindow()
        Dim insp = TryCast(w, Outlook.Inspector)
        If insp IsNot Nothing Then Return ForInspector(insp)
        Dim exp = TryCast(w, Outlook.Explorer)
        If exp IsNot Nothing Then Return ForInlineResponse(exp)
        Return Nothing
    End Function

    ' =====================================================================
    ' Public API
    ' =====================================================================

    ''' <summary>Sets the block to plain text (line breaks allowed).</summary>
    Public Sub SetText(kind As BlockKind, text As String)
        If String.IsNullOrEmpty(text) Then
            Remove(kind)
            Return
        End If
        Apply(kind, Sub(r) r.Text = NormalizeBreaks(text), TextPrefix & text)
    End Sub

    ''' <summary>Sets the block to HTML. Tables, formatting and links are kept.</summary>
    Public Sub SetHtml(kind As BlockKind, html As String)
        If String.IsNullOrEmpty(html) Then
            Remove(kind)
            Return
        End If
        Apply(kind, Sub(r) InsertHtml(r, html), HtmlPrefix & html)
    End Sub

    ''' <summary>
    ''' Builds the block content in code. The range passed in is the empty cell
    ''' content. Content built this way can't be re-created by RepairMissing().
    ''' </summary>
    Public Sub SetContent(kind As BlockKind, builder As Action(Of Word.Range))
        If builder Is Nothing Then Throw New ArgumentNullException(NameOf(builder))
        Apply(kind, builder, CustomMarker)
    End Sub

    ''' <summary>Removes the block (active message only) and its spacer lines.</summary>
    Public Sub Remove(kind As BlockKind)
        Dim tbl = FindBlock(kind)
        If tbl IsNot Nothing Then
            Dim undo = _doc.Application.UndoRecord
            undo.StartCustomRecord("Remove " & kind.ToString().ToLower())
            Try
                Dim pos = tbl.Range.Start
                tbl.Delete()
                RemoveSpacers(kind, pos)
            Finally
                undo.EndCustomRecord()
            End Try
        End If
        SetSource(kind, "")
    End Sub

    ''' <summary>True if the block is in the active message.</summary>
    Public Function Exists(kind As BlockKind) As Boolean
        Return FindBlock(kind) IsNot Nothing
    End Function

    ''' <summary>Current visible text of the block, including any user edits.</summary>
    Public Function GetCurrentText(kind As BlockKind) As String
        Dim tbl = FindBlock(kind)
        If tbl Is Nothing Then Return ""
        Return CellContentRange(tbl.Cell(1, 1)).Text.Replace(ChrW(7), "")
    End Function

    ''' <summary>
    ''' The text or HTML last applied by the add-in (for prefilling an edit
    ''' form). Returns "" if nothing was stored or it was built in code.
    ''' </summary>
    Public Function GetLastApplied(kind As BlockKind, ByRef isHtml As Boolean) As String
        Dim s = GetSource(kind)
        isHtml = False
        If s.StartsWith(HtmlPrefix) Then
            isHtml = True
            Return s.Substring(HtmlPrefix.Length)
        End If
        If s.StartsWith(TextPrefix) Then Return s.Substring(TextPrefix.Length)
        Return ""
    End Function

    ''' <summary>
    ''' Re-inserts a block the user deleted, using the last-applied source.
    ''' Call from Application.ItemSend. Returns True if anything was repaired.
    ''' </summary>
    Public Function RepairMissing() As Boolean
        Dim repaired = False
        For Each kind In {BlockKind.Header, BlockKind.Footer}
            If Exists(kind) Then Continue For
            Dim isHtml As Boolean
            Dim source = GetLastApplied(kind, isHtml)
            If source.Length = 0 Then Continue For
            If isHtml Then SetHtml(kind, source) Else SetText(kind, source)
            repaired = True
        Next
        Return repaired
    End Function

    ''' <summary>Optional: strip the hidden properties before the mail leaves.</summary>
    Public Sub ClearStoredSources()
        Try
            _mail.PropertyAccessor.DeleteProperties({HeaderSourceProp, FooterSourceProp})
        Catch
        End Try
    End Sub

    ' =====================================================================
    ' Core
    ' =====================================================================

    Private Sub Apply(kind As BlockKind, fill As Action(Of Word.Range), source As String)
        Dim undo = _doc.Application.UndoRecord
        undo.StartCustomRecord("Update " & kind.ToString().ToLower())
        Try
            Dim tbl = FindBlock(kind)
            If tbl Is Nothing Then tbl = CreateBlock(kind)

            ClearCell(tbl.Cell(1, 1))
            fill(CellContentRange(tbl.Cell(1, 1)))

            ' Re-find: inserted content can shift ranges
            tbl = FindBlock(kind)
            If tbl IsNot Nothing Then TrimTrailingEmptyParagraph(tbl.Cell(1, 1))

            SetSource(kind, source)
        Finally
            undo.EndCustomRecord()
        End Try
    End Sub

    ''' <summary>
    ''' Finds the block in the ACTIVE message only. Tables at or below the start
    ''' of the quoted history are ignored, so old headers/footers in the chain
    ''' are never matched. Header = first match, footer = last match above the
    ''' history.
    ''' </summary>
    Private Function FindBlock(kind As BlockKind) As Word.Table
        Dim color = ColorFor(kind)
        Dim limit = ActiveMessageEnd()
        Dim found As Word.Table = Nothing

        Dim count = _doc.Tables.Count
        For i = 1 To count
            Dim t = _doc.Tables(i)
            If t.Range.Start >= limit Then Exit For      ' tables are in document order
            If IsFingerprinted(t, color) Then
                found = t
                If kind = BlockKind.Header Then Exit For
            End If
        Next
        Return found
    End Function

    Private Shared Function IsFingerprinted(t As Word.Table, color As Integer) As Boolean
        Try
            If t.Rows.Count <> 1 Then Return False
            If t.Rows(1).Cells.Count <> 1 Then Return False
            Return CInt(t.Cell(1, 1).Shading.BackgroundPatternColor) = color
        Catch
            Return False   ' irregular user tables can throw; they're not ours
        End Try
    End Function

    Private Function CreateBlock(kind As BlockKind) As Word.Table
        Dim anchor As Word.Range
        If kind = BlockKind.Header Then
            ' [header table][spacer][message...]
            anchor = PrepareAnchorBefore(0, spacerBefore:=False)
        Else
            ' [message + signature][spacer][footer table][spacer][quoted history]
            anchor = PrepareAnchorBefore(FooterInsertPosition(), spacerBefore:=True)
        End If

        Dim tbl = _doc.Tables.Add(anchor, 1, 1)
        tbl.Borders.Enable = 0
        tbl.PreferredWidthType = Word.WdPreferredWidthType.wdPreferredWidthPercent
        tbl.PreferredWidth = 100
        tbl.TopPadding = 0
        tbl.BottomPadding = 0
        tbl.LeftPadding = 0
        tbl.RightPadding = 0
        tbl.Cell(1, 1).Shading.BackgroundPatternColor = CType(ColorFor(kind), Word.WdColor)
        Return tbl
    End Function

    ''' <summary>
    ''' Where a new footer goes: just above the quoted history if there is one,
    ''' otherwise at the end of the message.
    ''' </summary>
    Private Function FooterInsertPosition() As Integer
        Dim boundary = FindQuoteBoundary()
        If boundary >= 0 Then Return boundary

        ' No history: use an empty paragraph at the very end
        Dim last = _doc.Paragraphs.Last.Range
        If last.Text <> vbCr Then
            _doc.Content.InsertParagraphAfter()
            last = _doc.Paragraphs.Last.Range
        End If
        Return last.Start
    End Function

    ''' <summary>
    ''' Inserts empty paragraphs at a paragraph-start position and returns the one
    ''' that will become the table. Spacers keep the container from touching (and
    ''' merging with) neighbouring tables, and the new paragraphs are reset so
    ''' they don't inherit quote borders/indents from the reply separator.
    ''' </summary>
    Private Function PrepareAnchorBefore(pos As Integer, spacerBefore As Boolean) As Word.Range
        Dim total = If(spacerBefore, 3, 2)
        Dim needed = total

        If CBool(_doc.Range(pos, pos).Information(Word.WdInformation.wdWithInTable)) Then
            Dim top = TopLevelTableAt(pos)
            If top IsNot Nothing Then
                pos = top.Range.Start
                top.Split(1)          ' inserts an empty paragraph above the table
                needed -= 1
            End If
        End If

        For i = 1 To needed
            _doc.Range(pos, pos).InsertParagraphBefore()
        Next

        Try
            Dim block = _doc.Range(pos, pos + total)
            block.Style = _doc.Styles(Word.WdBuiltinStyle.wdStyleNormal)
            block.ParagraphFormat.Borders.Enable = 0
            block.ParagraphFormat.LeftIndent = 0
            block.ParagraphFormat.RightIndent = 0
        Catch
        End Try

        Dim tablePara = pos + If(spacerBefore, 1, 0)
        Return _doc.Range(tablePara, tablePara).Paragraphs(1).Range
    End Function

    Private Function TopLevelTableAt(pos As Integer) As Word.Table
        Dim count = _doc.Tables.Count
        For i = 1 To count
            Dim t = _doc.Tables(i)
            If t.Range.Start <= pos AndAlso pos < t.Range.End Then Return t
        Next
        Return Nothing
    End Function

    ' =====================================================================
    ' Quoted-history detection
    ' =====================================================================

    ''' <summary>End of the active message (start of quoted history, or doc end).</summary>
    Private Function ActiveMessageEnd() As Integer
        Dim b = FindQuoteBoundary()
        Return If(b >= 0, b, _doc.Content.End)
    End Function

    ''' <summary>
    ''' Returns the character position where the quoted history starts, or -1
    ''' for a new message. Checks, in order, and takes the earliest hit:
    '''   1. Outlook's hidden _MailOriginal bookmark
    '''   2. The reply separator: a paragraph with a top border containing a
    '''      colon (the "From: ... Sent: ..." block), or "-----Original Message-----"
    '''   3. A horizontal line (&lt;hr&gt;) used as a separator
    ''' Anything inside the add-in's own container tables is skipped, so a
    ''' footer that contains borders, lines or tables can't be mistaken for it.
    ''' </summary>
    Private Function FindQuoteBoundary() As Integer
        Dim boundary = Integer.MaxValue

        ' 1. Outlook's own marker (present while the reply is open)
        Try
            _doc.Bookmarks.ShowHidden = True
            If _doc.Bookmarks.Exists(OriginalBookmark) Then
                Dim start = _doc.Bookmarks.Item(OriginalBookmark).Range.Start
                If start > 0 Then boundary = start
            End If
        Catch
        End Try

        Dim ours = GetContainerRanges()

        ' 2. Separator paragraph (stop at the first one)
        For Each p As Word.Paragraph In _doc.Paragraphs
            Dim r = p.Range
            If r.Start >= boundary Then Exit For
            If InsideAny(r.Start, ours) Then Continue For
            If IsSeparatorParagraph(r) Then
                boundary = r.Start
                Exit For
            End If
        Next

        ' 3. Horizontal line (stop at the first one)
        For Each s As Word.InlineShape In _doc.InlineShapes
            Dim pos = s.Range.Start
            If pos >= boundary Then Exit For
            If InsideAny(pos, ours) Then Continue For
            If s.Type = Word.WdInlineShapeType.wdInlineShapeHorizontalLine Then
                boundary = s.Range.Paragraphs(1).Range.Start
                Exit For
            End If
        Next

        Return If(boundary = Integer.MaxValue, -1, boundary)
    End Function

    Private Shared Function IsSeparatorParagraph(r As Word.Range) As Boolean
        Try
            Dim text = r.Text
            If text.Contains("-----") AndAlso
               text.IndexOf("Original Message", StringComparison.OrdinalIgnoreCase) >= 0 Then
                Return True
            End If

            Dim topBorder = r.ParagraphFormat.Borders(Word.WdBorderType.wdBorderTop).LineStyle
            If topBorder <> Word.WdLineStyle.wdLineStyleNone AndAlso text.Contains(":") Then
                Return True
            End If
        Catch
        End Try
        Return False
    End Function

    ''' <summary>Start/end of every top-level container table (header or footer).</summary>
    Private Function GetContainerRanges() As List(Of Tuple(Of Integer, Integer))
        Dim list As New List(Of Tuple(Of Integer, Integer))
        Dim count = _doc.Tables.Count
        For i = 1 To count
            Dim t = _doc.Tables(i)
            If IsFingerprinted(t, HeaderColor) OrElse IsFingerprinted(t, FooterColor) Then
                list.Add(Tuple.Create(t.Range.Start, t.Range.End))
            End If
        Next
        Return list
    End Function

    Private Shared Function InsideAny(pos As Integer,
                                      ranges As List(Of Tuple(Of Integer, Integer))) As Boolean
        For Each rg In ranges
            If pos >= rg.Item1 AndAlso pos < rg.Item2 Then Return True
        Next
        Return False
    End Function

    ' =====================================================================
    ' Cell content
    ' =====================================================================

    ''' <summary>Removes everything in the cell, including nested tables.</summary>
    Private Shared Sub ClearCell(cell As Word.Cell)
        For i = cell.Tables.Count To 1 Step -1
            cell.Tables(i).Delete()
        Next
        Dim r = CellContentRange(cell)
        If r.End > r.Start Then r.Delete()
    End Sub

    ''' <summary>The cell's content without the end-of-cell marker.</summary>
    Private Shared Function CellContentRange(cell As Word.Cell) As Word.Range
        Dim r = cell.Range
        r.MoveEnd(Word.WdUnits.wdCharacter, -1)
        Return r
    End Function

    ''' <summary>
    ''' HTML insertion usually leaves an extra blank paragraph at the end of the
    ''' cell. Remove it, except right after a nested table (Word requires one).
    ''' </summary>
    Private Sub TrimTrailingEmptyParagraph(cell As Word.Cell)
        Try
            Dim paras = cell.Range.Paragraphs
            If paras.Count < 2 Then Return

            Dim lastText = paras(paras.Count).Range.Text.Replace(ChrW(7), "").Replace(vbCr, "")
            If lastText.Length > 0 Then Return

            Dim prev = paras(paras.Count - 1).Range
            If prev.Cells.NestingLevel > cell.NestingLevel Then Return

            _doc.Range(prev.End - 1, prev.End).Delete()
        Catch
        End Try
    End Sub

    ''' <summary>Removes the empty spacer paragraphs left after deleting a block.</summary>
    Private Sub RemoveSpacers(kind As BlockKind, tablePos As Integer)
        If kind = BlockKind.Header Then
            DeleteIfEmpty(tablePos)
        Else
            DeleteIfEmpty(tablePos)                       ' spacer after
            If tablePos > 0 Then DeleteIfEmpty(tablePos - 1) ' spacer before
        End If
    End Sub

    Private Sub DeleteIfEmpty(pos As Integer)
        Try
            Dim p = _doc.Range(pos, pos).Paragraphs(1).Range
            If p.Text = vbCr AndAlso
               Not CBool(p.Information(Word.WdInformation.wdWithInTable)) Then
                p.Delete()
            End If
        Catch
        End Try
    End Sub

    ''' <summary>Inserts HTML (tables included) via a temporary .htm file.</summary>
    Private Shared Sub InsertHtml(target As Word.Range, html As String)
        Dim page = html
        If page.IndexOf("<html", StringComparison.OrdinalIgnoreCase) < 0 Then
            page = "<html><head><meta charset=""utf-8""></head><body>" & html & "</body></html>"
        End If

        Dim path = IO.Path.Combine(IO.Path.GetTempPath(),
                                   "hf_" & Guid.NewGuid().ToString("N") & ".htm")
        File.WriteAllText(path, page, New UTF8Encoding(True))
        Try
            target.InsertFile(FileName:=path, ConfirmConversions:=False)
        Finally
            Try : File.Delete(path) : Catch : End Try
        End Try
    End Sub

    ' =====================================================================
    ' Helpers
    ' =====================================================================

    Private Shared Function ColorFor(kind As BlockKind) As Integer
        Return If(kind = BlockKind.Header, HeaderColor, FooterColor)
    End Function

    Private Shared Function PropFor(kind As BlockKind) As String
        Return If(kind = BlockKind.Header, HeaderSourceProp, FooterSourceProp)
    End Function

    Private Shared Function NormalizeBreaks(s As String) As String
        Return s.Replace(vbCrLf, vbCr).Replace(vbLf, vbCr)
    End Function

    Private Function GetSource(kind As BlockKind) As String
        Try
            Return If(TryCast(_mail.PropertyAccessor.GetProperty(PropFor(kind)), String), "")
        Catch
            Return ""
        End Try
    End Function

    Private Sub SetSource(kind As BlockKind, value As String)
        Try
            If String.IsNullOrEmpty(value) Then
                _mail.PropertyAccessor.DeleteProperty(PropFor(kind))
            Else
                _mail.PropertyAccessor.SetProperty(PropFor(kind), value)
            End If
        Catch
            ' Very large HTML can exceed what PropertyAccessor allows; the block
            ' still works, only prefill/repair is unavailable for this mail.
        End Try
    End Sub

End Class
