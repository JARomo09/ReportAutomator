'Option Strict On
Imports Microsoft.Office.Interop.Word
Imports Microsoft.Office.Core
Imports System.IO

Module ImgSrtr
    'By: Javier Romo
    'https://www.linkedin.com/in/jaromo/
    Public WdApp As New Microsoft.Office.Interop.Word.Application
    Public Shps As InlineShapes
    Public Tbls As Tables
    Public Klr As WdColorIndex

    Sub Main()
        'Originator: Javier Romo
        'Wishes, bugs, support and/or annotations:
        'https://www.linkedin.com/in/jaromo/
        Dim Dktp As String = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
        Dim Loc1 As String = Dktp & "\Pictures"
        Dim SubLoc As String = Nothing
        Dim Mde As Integer = 3 '1 or 2: expenses, 3 or 4: general 
        Dim LstOfImgs As New List(Of String)
        Dim Flg As Boolean = False
        Dim Title As String = "Field Observations"

        'Test()
        Select Case Mde
            Case 1 'Expenses report, portrait
                SubLoc = Loc1 & "Source Expenses"
            Case 2 'Expenses report, landscape
                SubLoc = Loc1 & "Source Expenses"
            Case 3 'General report, portrait
                SubLoc = Loc1 & "\Source " & Title
            Case 4 'General report, portrait
                SubLoc = Loc1 & "\Source " & Title
        End Select

        'Create directory for screenshots if it doesn't exist
        If System.IO.Directory.Exists(Loc1) = False Then
            System.IO.Directory.CreateDirectory(Loc1)
            Flg = True
        Else : End If
        If System.IO.Directory.Exists(SubLoc) = False Then
            System.IO.Directory.CreateDirectory(SubLoc)
        Else : End If

        'Check if the directory is empty or not
        If Flg = False Then 'Proceed with image reporting algorithm
            LstOfImgs = ImgType(Loc1, Mde)
            If LstOfImgs.Count > 0 Then
                ImgHndlrSte.CreateReport(Mde, Title, Loc1, SubLoc, LstOfImgs)
            ElseIf LstOfImgs.Count = 0 Then
                MsgBox("Please store images in the upcoming directory, and try again")
                Process.Start(Loc1)
            End If
        ElseIf Flg = True Then 'Instruct user to feed images to sort
            MsgBox("Please store images in the upcoming directory, and try again")
            Process.Start(Loc1)
        End If
        WdApp.Quit()
    End Sub

    Sub CreateReport(mode As Integer, title As String, location As String, location2 As String, Grp As List(Of String))
        'Originator: Javier Romo
        'https://www.linkedin.com/in/jaromo/
        'Usage:
        'Mode 1: Expense sorter, vertical
        Dim InstanceID1 As String
        Dim Left_margin As Double
        Dim Right_margin As Double
        Dim Page_width As Double
        Dim proportion As Double = 1.2
        Dim HeaderHeight As Double
        Dim UsableHeight As Double
        Dim UsableWidth As Double
        Dim UsableProportion As Double
        Dim TableX As Integer
        Dim CycleNumb As Integer
        Dim Counter As Integer
        Dim Item As Integer
        Dim Limit As Integer
        Dim AppRow As Integer
        Dim AppCol As Integer
        Dim PicHeight As Double
        Dim PicWidth As Double
        Dim PicProportion As Double
        Dim PicName As String
        Dim Group1 As List(Of String) = Grp
        Dim TbC As Integer = Nothing
        Dim TbR As Integer = Nothing
        Dim TbHdr As String = Nothing
        Dim CycLmt As Integer = Nothing
        Dim j As Integer = Nothing
        Dim LstK As New List(Of Integer)

        'Proceed with report creation
        With WdApp
            'Display MS Word instance
            .ScreenUpdating = True
            .Visible = True
            'MS Word document set up and acquisition of relevant parameters for automated handling
            .Documents.Add()
            InstanceID1 = .ActiveDocument.Name
            TableX = 1
            Select Case mode
                Case 1 'Expenses portrait
                    CycleNumb = 1
                    Counter = 0
                    Item = 1
                    TbC = 2
                    TbR = 4
                    TbHdr = "13"
                    CycLmt = 4
                    With .ActiveDocument.PageSetup
                        .Orientation = WdOrientation.wdOrientPortrait
                        Left_margin = .LeftMargin
                        Right_margin = .RightMargin
                        Page_width = .PageWidth
                        HeaderHeight = .HeaderDistance
                        UsableHeight = ((.PageHeight - HeaderHeight) - (.TopMargin + .BottomMargin) - (18 * 10)) / 2
                        HeaderHeight = .HeaderDistance
                    End With
                Case 2 'Expenses landscape
                    CycleNumb = 1
                    Counter = 0
                    Item = 1
                    TbC = 2
                    TbR = 2
                    TbHdr = "1"
                    CycLmt = 2
                    With .ActiveDocument.PageSetup
                        .Orientation = WdOrientation.wdOrientLandscape
                        Left_margin = .LeftMargin
                        Right_margin = .RightMargin
                        Page_width = .PageWidth
                        HeaderHeight = .HeaderDistance
                        UsableHeight = ((.PageHeight - HeaderHeight) - (.TopMargin + .BottomMargin) - (18 * 10))
                        HeaderHeight = .HeaderDistance
                    End With
                Case 3 'Generic report
                    CycleNumb = 1
                    Counter = 0
                    Item = 1
                    TbC = 2
                    TbR = 4
                    TbHdr = "13"
                    CycLmt = 4
                    With .ActiveDocument.PageSetup
                        .Orientation = WdOrientation.wdOrientPortrait
                        Left_margin = .LeftMargin
                        Right_margin = .RightMargin
                        Page_width = .PageWidth
                        HeaderHeight = .HeaderDistance
                        UsableHeight = ((.PageHeight - HeaderHeight) - (.TopMargin + .BottomMargin) - (18 * 10))
                        HeaderHeight = .HeaderDistance
                    End With
                Case 4 'Generic report landscape
                    CycleNumb = 1
                    Counter = 0
                    Item = 1
                    TbC = 2
                    TbR = 2
                    TbHdr = "1"
                    CycLmt = 2
                    With .ActiveDocument.PageSetup
                        .Orientation = WdOrientation.wdOrientLandscape
                        Left_margin = .LeftMargin
                        Right_margin = .RightMargin
                        Page_width = .PageWidth
                        HeaderHeight = .HeaderDistance
                        UsableHeight = ((.PageHeight - HeaderHeight) - (.TopMargin + .BottomMargin) - (18 * 10))
                        HeaderHeight = .HeaderDistance
                    End With
            End Select
            Header(title, location, InstanceID1)
            Limit = Group1.Count

            'Table insertion algorithm
            If Limit >= 1 Then 'Conditional to kill program in the event the number of screenshots is less than 1
                .Selection.TypeParagraph()
                While Counter < Limit
                    InsrTbl(TableX, TbR, TbC, TbHdr)
                    With .ActiveDocument.Tables(TableX)
                        UsableWidth = .Columns(1).Width * 0.95
                    End With
                    UsableProportion = UsableHeight / UsableWidth
                    'Iterating cycle to insert screenshots in one table
                    While CycleNumb <= CycLmt
                        If Counter <= Limit Then
                            If CycleNumb = 1 Then
                                AppRow = 2
                                AppCol = 1
                            ElseIf CycleNumb = 2 Then
                                AppRow = 2
                                AppCol = 2
                            ElseIf CycleNumb = 3 Then
                                AppRow = 4
                                AppCol = 1
                            ElseIf CycleNumb = 4 Then
                                AppRow = 4
                                AppCol = 2
                            End If
                            .ActiveDocument.Tables(TableX).Cell(AppRow, AppCol).Select()
                            .Selection.MoveLeft(WdUnits.wdCharacter, 1)
                            If Counter = Limit Then
                                'do nothing
                            Else
                                'Picture insertion and formating
                                Shps = CType(.ActiveDocument.InlineShapes, InlineShapes)
                                With Shps.AddPicture(location & "\" & Group1(Counter), False, True)
                                    System.IO.File.Move(location & "\" & Group1(Counter), location2 & "\" & Group1(Counter))
                                    PicHeight = .Height
                                    PicWidth = .Width
                                    PicProportion = PicHeight / PicWidth
                                    .LockAspectRatio = MsoTriState.msoTrue
                                    Select Case mode
                                        Case 1
                                            .PictureFormat.Brightness = 0.5
                                            .PictureFormat.Contrast = 0.9
                                        Case 2
                                            .PictureFormat.Brightness = 0.5
                                            .PictureFormat.Contrast = 0.9
                                        Case Else
                                    End Select
                                    'Picture size & proportion setup
                                    If UsableProportion <> PicProportion Then
                                        If UsableProportion > PicProportion Then
                                            .Width = Convert.ToSingle(UsableWidth)
                                        ElseIf UsableProportion < PicProportion Then
                                            .Height = Convert.ToSingle(UsableHeight * proportion)
                                        End If
                                    ElseIf UsableProportion = PicProportion Then
                                        .Height = Convert.ToSingle(UsableHeight * proportion)
                                    End If
                                    .Select()
                                End With
                                With .Selection
                                    .ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter
                                    .Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter
                                    .MoveLeft(WdUnits.wdCharacter, 1)
                                    .TypeParagraph()
                                    .MoveRight(WdUnits.wdCharacter, 1)
                                    .TypeParagraph()
                                End With
                                Tbls = CType(.ActiveDocument.Tables, Tables)
                                'Add picture header
                                With Tbls(TableX).Cell(AppRow - 1, AppCol)
                                    Select Case mode
                                        Case 1
                                            PicName = Mid(Group1(Counter), 1, Len(Group1(Counter)) - 4)
                                            .Range.Text = "Item " & CStr(Item) & ": " & CStr(FrmtFleNm(PicName, 1)) 'Add expense date to table cell
                                        Case 2
                                            PicName = Mid(Group1(Counter), 1, Len(Group1(Counter)) - 4)
                                            .Range.Text = "Item " & CStr(Item) & ": " & CStr(FrmtFleNm(PicName, 1)) 'Add expense date to table cell
                                        Case 3
                                            PicName = Group1(Counter)
                                            For Each i As Char In PicName
                                                If i = "." Then
                                                    j = j + 1
                                                    LstK.Add(j)
                                                Else
                                                    j = j + 1
                                                End If
                                            Next
                                            PicName = Mid(PicName, 1, LstK.Last - 1)
                                            .Range.Text = PicName
                                        Case 4
                                            PicName = Group1(Counter)
                                            For Each i As Char In PicName
                                                If i = "." Then
                                                    j = j + 1
                                                    LstK.Add(j)
                                                Else
                                                    j = j + 1
                                                End If
                                            Next
                                            PicName = Mid(PicName, 1, LstK.Last - 1)
                                            .Range.Text = PicName
                                    End Select
                                End With
                                Counter = Counter + 1
                            End If
                        Else : End If
                        CycleNumb = CycleNumb + 1
                        Item = Item + 1
                        LstK.Clear()
                        j = Nothing

                    End While
                    CycleNumb = 1
                    TableX = TableX + 1
                    'Insertion of a page break to insert a new table
                    If Counter < Limit Then
                        With .Selection
                            .EndKey(WdUnits.wdStory)
                            .InsertBreak(WdBreakType.wdPageBreak)
                        End With
                    ElseIf Counter = Limit Then
                        With .Selection
                            .EndKey(WdUnits.wdStory)
                        End With
                    End If
                End While
                .ActiveDocument.SaveAs2(location & "\" & title & ".docx", WdSaveFormat.wdFormatXMLDocument)
                .ActiveDocument.Paragraphs(1).Range.Delete()
                .ActiveDocument.Save()
                .ActiveDocument.ExportAsFixedFormat(location & "\" & title & ".pdf", WdExportFormat.wdExportFormatPDF)
            Else : End If
            Beep()
        End With
        WdApp.Quit(0) 'exit app
    End Sub

    Sub InsrTbl(tabla As Integer, filas As Long, columnas As Long, FilEncabezado As String)
        'Originator: Javier Romo
        'https://www.linkedin.com/in/jaromo/
        'Summoned to create a header table, consistent with corporate style standards

        Dim Longitud As Long
        Dim contador As Integer
        Dim X As String

        WdApp.ActiveDocument.Tables.Add(WdApp.Selection.Range, CInt(filas), CInt(columnas), WdDefaultTableBehavior.wdWord9TableBehavior, WdDefaultTableBehavior.wdWord8TableBehavior)
        Longitud = Len(FilEncabezado)
        contador = 1
        If FilEncabezado <> "0" Then
            While contador <= CInt(Longitud)
                X = Mid(FilEncabezado, contador, 1)
                Tbls = CType(WdApp.ActiveDocument.Tables, Tables)

                With Tbls(tabla).Rows(CInt(X)).Range
                    With .Font
                        .TextColor.RGB = WdColor.wdColorWhite
                        .Bold = 1
                        .AllCaps = 1
                    End With
                    .ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter
                    .Shading.BackgroundPatternColor = RGB(0, 154, 118)
                End With
                contador = contador + 1
            End While
        Else : End If

        Longitud = Nothing
        contador = Nothing
        X = Nothing
    End Sub

    Function Header(title As String, location As String, InstanceID As String) As Double
        'Originator: Javier Romo
        'https://www.linkedin.com/in/jaromo/
        'Function dedicated to set page style according to corporate standards

        Dim TotalWidth As Double
        Dim LeftMargin As Double
        Dim RightMargin As Double
        Dim HeaderWidth As Double 'points
        Dim HeaderWidth2 As Double 'inches
        Dim TempLogo As String = location & "\logo.png"

        WdApp.Documents(InstanceID).Activate()

        With WdApp.ActiveDocument.PageSetup
            TotalWidth = .PageWidth
            LeftMargin = .LeftMargin
            RightMargin = .RightMargin
            HeaderWidth = TotalWidth - LeftMargin - RightMargin
            HeaderWidth2 = WdApp.PointsToInches(Convert.ToSingle(HeaderWidth))
        End With

        If WdApp.ActiveWindow.View.SplitSpecial <> 0 Then
            WdApp.ActiveWindow.Panes(2).Close()
        End If
        If WdApp.ActiveWindow.ActivePane.View.Type = 1 Or WdApp.ActiveWindow.ActivePane.View.Type = 2 Then
            WdApp.ActiveWindow.ActivePane.View.Type = WdViewType.wdPrintView
        End If
        WdApp.ActiveWindow.ActivePane.View.SeekView = WdSeekView.wdSeekCurrentPageHeader
        Call InsrTbl(1, 1, 3, "0")
        With WdApp.Selection.Tables(1)
            .Columns(1).Width = WdApp.InchesToPoints(1.49)
            .Columns(3).Width = WdApp.InchesToPoints(1.53)
            .Columns(2).Width = WdApp.InchesToPoints(Convert.ToSingle(HeaderWidth2 - 1.53 - 1.49))
            .Borders.Enable = 0 'False
            .Columns(1).Select()

            With WdApp.Selection
                .ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter
                .Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter
            End With
            .Cell(1, 1).Select()
            WdApp.Selection.MoveLeft(WdUnits.wdCharacter, 1)
            ImgHndlrSte.My.Resources.Logo.Save(TempLogo)
            With WdApp.Selection.InlineShapes.AddPicture(TempLogo, False, True)
                .Width = WdApp.InchesToPoints(1.49)
                .Height = WdApp.InchesToPoints(1.49 * (4.62 / 6.13))
            End With
            .Columns(2).Select()

            With WdApp.Selection
                .ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter
                .Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter
            End With
            .Cell(1, 2).Select()
            WdApp.Selection.MoveLeft(WdUnits.wdCharacter, 1)
            With WdApp.Selection
                .Text = title
                .Font.Bold = 1 'True
                .Font.Size = 20
                .Font.TextColor.RGB = RGB(253, 129, 3)
            End With
            .Columns(3).Select()

            With WdApp.Selection
                .ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft
                .Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter
            End With
            .Cell(1, 3).Select()

            With WdApp.Selection 'header details
                .Font.Bold = 1 'True
                .Font.TextColor.RGB = RGB(0, 154, 118)
                .TypeText("Author: ")
                .Font.Bold = 0
                .Font.ColorIndex = WdColorIndex.wdAuto
                .TypeText(WdApp.Application.UserName)
                .TypeParagraph()
                .Font.Bold = 1
                .Font.TextColor.RGB = RGB(0, 154, 118)
                .TypeText("Date: ")
                .Font.Bold = 0
                .Font.ColorIndex = WdColorIndex.wdAuto
                .TypeText(DateString)
                .Font.Size = 11
            End With
        End With
        WdApp.ActiveWindow.ActivePane.View.SeekView = WdSeekView.wdSeekMainDocument
        My.Computer.FileSystem.DeleteFile(TempLogo)

        Header = HeaderWidth

        TotalWidth = Nothing
        LeftMargin = Nothing
        RightMargin = Nothing
        HeaderWidth = Nothing
        HeaderWidth2 = Nothing
        TempLogo = Nothing
    End Function

    Function FrmtFleNm(titulo As String, mode As Integer) As Date
        'Originator: Javier Romo
        'Function dedicated to re-format a string of characters to improve reading ease
        Dim FullDte As Date

        Select Case mode
            Case 1 'Expense sorter
                FullDte = StrDate(titulo)
            Case 2
                FullDte = StrDate(titulo)
        End Select

        Return FullDte

        FullDte = Nothing
    End Function

    Function ImgType(folder As String, mode As Integer) As List(Of String)
        'Originator: Javier Romo
        'Creates an array of image files, in function of the location specified and desired output.
        Dim Location As New DirectoryInfo(folder)
        Dim DocuArr As FileInfo() = Location.GetFiles()
        Dim Docu As FileInfo
        Dim SSList As List(Of String) = New List(Of String)

        Select Case mode
            Case 1 'Output 1 or 2: expenses report portrait
                For Each Docu In DocuArr
                    If UCase(Docu.Extension) = ".PNG" Or UCase(Docu.Extension) = ".JPG" Or UCase(Docu.Extension) = ".JPEG" Or UCase(Docu.Extension) = ".TIFF" Then
                        SSList.Add(Docu.Name)
                    Else : End If
                Next
            Case 2 'Output 1 or 2: expenses report landscape
                For Each Docu In DocuArr
                    If UCase(Docu.Extension) = ".PNG" Or UCase(Docu.Extension) = ".JPG" Or UCase(Docu.Extension) = ".JPEG" Or UCase(Docu.Extension) = ".TIFF" Then
                        SSList.Add(Docu.Name)
                    Else : End If
                Next
            Case 3 'Output 3: general report portrait
                For Each Docu In DocuArr
                    If UCase(Docu.Extension) = ".PNG" Or UCase(Docu.Extension) = ".JPG" Or UCase(Docu.Extension) = ".JPEG" Or UCase(Docu.Extension) = ".TIFF" Then
                        SSList.Add(Docu.Name)
                    Else : End If
                Next
            Case 4 'Output 4: general report landscape
                For Each Docu In DocuArr
                    If UCase(Docu.Extension) = ".PNG" Or UCase(Docu.Extension) = ".JPG" Or UCase(Docu.Extension) = ".JPEG" Or UCase(Docu.Extension) = ".TIFF" Then
                        SSList.Add(Docu.Name)
                    Else : End If
                Next
        End Select
        If Not IsNothing(SSList) Then
            SSList.Sort()
        Else : End If

        Return SSList
        Location = Nothing
        DocuArr = Nothing
        Docu = Nothing
        SSList = Nothing
    End Function

    Function StrDate(str As String) As Date 'Assuming Office Lens Format: i.e. 2020_04_22 8_10 AM Office Lens.jpg
        'Originator: Javier Romo
        'https://www.linkedin.com/in/jaromo/
        Dim OutDate As Date = Nothing
        Dim hr As String = ""
        Dim min As String = ""

        If Len(Mid(str, 11)) = 20 Then 'i.e. 2020_04_22 9_10 AM Office Lens.jpg
            hr = "0" & Mid(str, 12, 1)
            min = Mid(str, 14, 2)
            OutDate = DateTime.Parse(Mid(str, 1, 4) & "-" & Mid(str, 6, 2) & "-" & Mid(str, 9, 2) & "T" & hr & ":" & min) '& ":00.0000000Z")
        ElseIf Len(Mid(str, 11)) = 21 Then 'i.e. 2020_04_22 21_10 PM Office Lens.jpg
            hr = Mid(str, 12, 2)
            min = Mid(str, 15, 2)
            OutDate = DateTime.Parse(Mid(str, 1, 4) & "-" & Mid(str, 6, 2) & "-" & Mid(str, 9, 2) & "T" & hr & ":" & min) '& ":00.0000000Z")
        Else
            OutDate = Nothing
        End If

        Return OutDate

        OutDate = Nothing
        hr = Nothing
        min = Nothing
    End Function

    Sub Test()
        WdApp.Documents.Add()
        Shps = CType(WdApp.ActiveDocument.InlineShapes, InlineShapes)
        Shps.AddPicture("C:\Users\j.romo\Desktop\Pictures\2020_05_14 10_17 AM Office Lens.jpg")
        MsgBox(Shps.Count)
        Shps.Item(1).LockAspectRatio = MsoTriState.msoTrue
        Shps.Item(1).Height = 20
        WdApp.Visible = True
        WdApp.ScreenUpdating = True
    End Sub
End Module
