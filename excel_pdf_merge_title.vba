'if have more than 2 pages, macro create two PDF, the first with titles and the second without titles, macro merge the file using pdftkk and exclude part 1 and 2
files neededs:

Excel
PDFTK (COULD BE PORTABLE)

Sub pdf()

    Dim objShell As Object
    Set objShell = CreateObject("WScript.Shell")
    
    Set myarray = CreateObject("System.Collections.ArrayList")
    Max = Sheets("C_PL_LESER_Ltda").Range("A" & Rows.Count).End(xlUp).Row
    'MsgBox Max
    Z = 0
    For X = 1 To Max
    
        If Cells(X, 1) = "-" Then
            Rows(X).Hidden = True
            Z = Z + 1
            myarray.Add X
        End If
      
    Next
    
    Cells(7, 17) = "  PROPOSTA COMERCIAL"
    
    ultima = WorksheetFunction.Match("#@", Range("A:A"), 0)
    ultima1 = WorksheetFunction.Match("information to match.", Range("A:A"), 0)
    
    ocultar1 = WorksheetFunction.Match("CONDIÇÕES COMERCIAIS", Range("A:A"), 0) - 5
    
    'ocultar bybye tecnico
    Rows((ocultar1) & ":" & (ocultar1 + 4)).Select
    Selection.EntireRow.Hidden = True
    
    ActiveSheet.PageSetup.PrintArea = "A1:Y" & ultima - 1
    
    Call config_pagina
    
    folhas = Application.ExecuteExcel4Macro("GET.DOCUMENT(50)")
    
    If folhas = 1 Then
        FileName = RDB_Create_PDF(Source:=Range("A1:Y" & ultima - 1), _
                                          FixedFilePathName:=ThisWorkbook.path & "\" & Cells(9, 19) & " Proposta Comercial " & Cells(9, 21) & ".pdf", _
                                          OverwriteIfFileExist:=True, _
                                          OpenPDFAfterPublish:=True)
                                          
        Rows((ocultar1) & ":" & (ocultar1 + 4)).Select
        Selection.EntireRow.Hidden = False
    End If
    
    If folhas > 1 Then
        loc_desvios = WorksheetFunction.Match("DESVIOS:", Range("A:A"), 0)
        
        ActiveSheet.PageSetup.PrintTitleRows = "$13:$15"
        
        FileName = RDB_Create_PDF(Source:=Range("A1:Y" & loc_desvios - 1), _
                                          FixedFilePathName:=ThisWorkbook.path & "\" & Cells(9, 19) & " Proposta Comercial " & Cells(9, 21) & "part1.pdf", _
                                          OverwriteIfFileExist:=True, _
                                          OpenPDFAfterPublish:=False)
        
        ActiveSheet.PageSetup.PrintTitleRows = False
        
        FileName = RDB_Create_PDF(Source:=Range("A" & loc_desvios & ":Y" & ultima - 1), _
                                          FixedFilePathName:=ThisWorkbook.path & "\" & Cells(9, 19) & " Proposta Comercial " & Cells(9, 21) & "part2.pdf", _
                                          OverwriteIfFileExist:=True, _
                                          OpenPDFAfterPublish:=False)
        part1_com = ThisWorkbook.path & "\" & Cells(9, 19) & " Proposta Comercial " & Cells(9, 21) & "part1.pdf"
        part2_com = ThisWorkbook.path & "\" & Cells(9, 19) & " Proposta Comercial " & Cells(9, 21) & "part2.pdf"
        final_com = ThisWorkbook.path & "\" & Cells(9, 19) & " Proposta Comercial " & Cells(9, 21) & ".pdf"
    
    
    End If
    
    Rows((ocultar1) & ":" & (ocultar1 + 4)).Select
    Selection.EntireRow.Hidden = False
    
    Cells(7, 17) = "  PROPOSTA TÉCNICA"
    
    ActiveSheet.PageSetup.PrintArea = "A1:W" & ultima1
    
    Call config_pagina
    
    folhas = Application.ExecuteExcel4Macro("GET.DOCUMENT(50)")
    
    If folhas = 1 Then
        FileName = RDB_Create_PDF(Source:=Range("A1:W" & ultima1 - 1), _
                                          FixedFilePathName:=ThisWorkbook.path & "\" & Cells(9, 19) & " Proposta Técnica " & Cells(9, 21) & ".pdf", _
                                          OverwriteIfFileExist:=True, _
                                          OpenPDFAfterPublish:=True)
                                          
        Rows((ocultar1) & ":" & (ocultar1 + 4)).Select
        Selection.EntireRow.Hidden = False
    End If
    
    If folhas > 1 Then
    
        ActiveSheet.PageSetup.PrintTitleRows = "$13:$15"
        
        loc_desvios = WorksheetFunction.Match("DESVIOS:", Range("A:A"), 0)
        FileName = RDB_Create_PDF(Source:=Range("A1:W" & loc_desvios - 1), _
                                          FixedFilePathName:=ThisWorkbook.path & "\" & Cells(9, 19) & " Proposta Técnica " & Cells(9, 21) & "part1.pdf", _
                                          OverwriteIfFileExist:=True, _
                                          OpenPDFAfterPublish:=False)
                                          
        ActiveSheet.PageSetup.PrintTitleRows = False
                                          
        FileName = RDB_Create_PDF(Source:=Range("A" & loc_desvios & ":W" & ultima1 - 1), _
                                          FixedFilePathName:=ThisWorkbook.path & "\" & Cells(9, 19) & " Proposta Técnica " & Cells(9, 21) & "part2.pdf", _
                                          OverwriteIfFileExist:=True, _
                                          OpenPDFAfterPublish:=False)
                                          
        part1_tec = ThisWorkbook.path & "\" & Cells(9, 19) & " Proposta Técnica " & Cells(9, 21) & "part1.pdf"
        part2_tec = ThisWorkbook.path & "\" & Cells(9, 19) & " Proposta Técnica " & Cells(9, 21) & "part2.pdf"
        final_tec = ThisWorkbook.path & "\" & Cells(9, 19) & " Proposta Técnica " & Cells(9, 21) & ".pdf"
    
    End If
    
    For X = 0 To Z - 1
        a = myarray(X)
        Rows(a).Hidden = False
    Next
    
    If folhas > 1 Then
        Call merge(part1_com, part2_com, final_com)
        Call merge(part1_tec, part2_tec, final_tec)
    End If
    
End Sub

Public Sub merge(part1, part2, saida)

    output = "cat output"
    
    exe = "s:\Vendas\011 COTACOES\004 Documentos para a elaboração de propostas\Planilhas\gzip\pdftk\pdftk.exe"
    shell "cmd.exe /k """"" & exe & """ """ & part1 & """ """ & part2 & """ " & output & " """ & saida & """""", vbHide
    
    While Len(Dir(saida, vbArchive)) = 0
        Application.Wait (Now + TimeValue("00:00:02"))
    Wend
    Kill part1
    Kill part2
End Sub


Sub config_pagina()
'
    ActiveSheet.ResetAllPageBreaks
    Range("B47:U47").Select
    Application.PrintCommunication = False
    With ActiveSheet.PageSetup
        .PrintTitleRows = ""
        .PrintTitleColumns = ""
    End With
    Application.PrintCommunication = True
    'ActiveSheet.PageSetup.PrintArea = "$A$1:$Y$78"
    Application.PrintCommunication = False
    With ActiveSheet.PageSetup
        .LeftHeader = ""
        .CenterHeader = ""
        .RightHeader = ""
        .LeftFooter = ""
        .CenterFooter = ""
        .RightFooter = ""
        .LeftMargin = Application.InchesToPoints(0.196850393700787)
        .RightMargin = Application.InchesToPoints(0.196850393700787)
        .TopMargin = Application.InchesToPoints(0.196850393700787)
        .BottomMargin = Application.InchesToPoints(0.196850393700787)
        .HeaderMargin = Application.InchesToPoints(0.31496062992126)
        .FooterMargin = Application.InchesToPoints(0.31496062992126)
        .PrintHeadings = False
        .PrintGridlines = False
        .PrintComments = False
        .CenterHorizontally = True
        .CenterVertically = False
        .Orientation = xlLandscape
        .Draft = False
        .PaperSize = xlPaperA4
        .FirstPageNumber = xlAutomatic
        .Order = xlDownThenOver
        .BlackAndWhite = False
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 2000
        .PrintErrors = xlPrintErrorsDisplayed
        .OddAndEvenPagesHeaderFooter = False
        .DifferentFirstPageHeaderFooter = False
        .ScaleWithDocHeaderFooter = True
        .AlignMarginsHeaderFooter = True
        .EvenPage.LeftHeader.Text = ""
        .EvenPage.CenterHeader.Text = ""
        .EvenPage.RightHeader.Text = ""
        .EvenPage.LeftFooter.Text = ""
        .EvenPage.CenterFooter.Text = ""
        .EvenPage.RightFooter.Text = ""
        .FirstPage.LeftHeader.Text = ""
        .FirstPage.CenterHeader.Text = ""
        .FirstPage.RightHeader.Text = ""
        .FirstPage.LeftFooter.Text = ""
        .FirstPage.CenterFooter.Text = ""
        .FirstPage.RightFooter.Text = ""
    End With
    Application.PrintCommunication = True
    Application.ExecuteExcel4Macro "PAGE.SETUP(,,,,,,,,,,,,{1,#N/A})" 'volta para a pagina percentual no imprimir
    Application.ExecuteExcel4Macro "PAGE.SETUP(,,,,,,,,,,,,{#N/A,#N/A})"
    folhas = Application.ExecuteExcel4Macro("GET.DOCUMENT(50)") 'quantidade de folhas

    
    'MsgBox folhas
    If folhas > 1 Then  'checa se tem mais de uma folha, se tem corta em documentação
        CORTE = WorksheetFunction.Match("DESVIOS:", Range("A:A"), 0)
        Rows(CORTE & ":" & CORTE).Select
        ActiveWindow.SelectedSheets.HPageBreaks.Add Before:=ActiveCell
    End If
    
    ActiveWindow.SmallScroll Down:=-8
    
End Sub

Public Function NumeroDaPagina(Optional ByRef rng As Excel.Range) As Variant 'checa em que página está a célula
    Dim pbHorizontal As HPageBreak
    Dim pbVertical As VPageBreak
    Dim nHorizontalPageBreaks As Long
    Dim nVerticalPageBreaks As Long
    Dim nNumeroDaPagina As Long
    On Error GoTo ErrHandler
        'Application.Volatile   'Você pode utilizar esta linha de comando para atualizar
                                'automaticamente a cada alteração (TORNA UM POUCO LENTO)
                                'Para utilizá-la basta desmarcar como comentário (')
        If rng Is Nothing Then _
        Set rng = Application.Caller
        With rng
            If .Parent.PageSetup.Order = xlDownThenOver Then
                nHorizontalPageBreaks = .Parent.HPageBreaks.Count + 1
                nVerticalPageBreaks = 1
            Else
                nHorizontalPageBreaks = 1
                nVerticalPageBreaks = .Parent.VPageBreaks.Count + 1
            End If
            nNumeroDaPagina = 1
            For Each pbHorizontal In .Parent.HPageBreaks
                If pbHorizontal.Location.Row > .Row Then Exit For
                nNumeroDaPagina = nNumeroDaPagina + nVerticalPageBreaks
            Next pbHorizontal
            For Each pbVertical In .Parent.VPageBreaks
                If pbVertical.Location.Column > .Column Then Exit For
                nNumeroDaPagina = nNumeroDaPagina + nHorizontalPageBreaks
            Next pbVertical
        End With
        NumeroDaPagina = nNumeroDaPagina
ResumeHere:
        Exit Function
ErrHandler:
        NumeroDaPagina = CVErr(xlErrRef)
        Resume ResumeHere
End Function
