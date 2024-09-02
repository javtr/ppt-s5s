Sub CH4M1()
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
'-----------------------------------------------------
On Error GoTo 346
On Error GoTo -1
Set SelectedCell_primera = ActiveCell
On Error GoTo 1
filaoriginal_prim = Selection.Row
columnaoriginal_prim = Selection.Column
ActiveSheet.Cells(filaoriginal_prim, columnaoriginal_prim).Select
Set SelectedCell = ActiveCell
'nombre tabla existente

On Error GoTo 1
TableName = SelectedCell.ListObject.Name

ActiveSheet.ListObjects(TableName).ListColumns(2).Range.Select
tipo = Selection.Column
If tipo = 2 Then GoTo 3 Else GoTo 1

3
ActiveSheet.ListObjects(TableName).ListRows(1).Range.Select
Set ActiveTable = ActiveSheet.ListObjects(TableName)

' primera ultima y cantidad filas tabla existente

filap1 = Selection.Row
filap = filap1 - 1
nfilas = ActiveTable.Range.Rows.Count
filau = filap + nfilas - 1

'primera y cantidad filas tabla nueva

filapn = filau + 4
nfilasnueva = 2
ncolumnasnueva = 8
filaun = filapn + nfilasnueva + 1

'insertar filas

For c = 1 To nfilasnueva + 6
  ActiveSheet.Cells(filau + 2, 1).EntireRow.Select
    Selection.Insert Shift:=xlDown
Next

'crear tabla

'numtabla = ActiveSheet.Cells(1, 18).value
Range(ActiveSheet.Cells(filapn, 1), ActiveSheet.Cells(filaun - 1, ncolumnasnueva)).Select
ActiveSheet.ListObjects.Add
'ActiveSheet.Cells(1, 18) = numtabla + 1
ActiveSheet.Cells(filapn + 1, 1).Select
Set SelectedCell_2 = ActiveCell
tablename_2 = SelectedCell_2.ListObject.Name
ActiveSheet.ListObjects(tablename_2).TableStyle = "S5S Blue"
ActiveSheet.ListObjects(tablename_2).ShowTotals = True
ActiveSheet.ListObjects(tablename_2).ShowAutoFilterDropDown = False
Range(ActiveSheet.Cells(filaun + 2, 1), ActiveSheet.Cells(filaun + 2, ncolumnasnueva)).Select
Selection.Delete Shift:=xlUp

color2 = 13404160

' nombre
    Range(ActiveSheet.Cells(filapn - 1, 1), ActiveSheet.Cells(filapn - 1, 8)).Select
    With Selection.Borders(xlEdgeLeft)
        .Color = color2
        .Weight = xlHairline
    End With
    With Selection.Borders(xlEdgeTop)
        .Color = color2
        .Weight = xlHairline
    End With
    With Selection.Borders(xlEdgeBottom)
        .Color = color2
        .Weight = xlHairline
    End With
    With Selection.Borders(xlEdgeRight)
        .Color = color2
        .Weight = xlHairline
    End With
    With Selection.Borders(xlInsideVertical)
        .Color = color2
        .Weight = xlHairline
    End With
    With Selection.Borders(xlInsideHorizontal)
        .Color = color2
        .Weight = xlHairline
    End With
    
    
    With Selection.Borders(xlEdgeLeft)
        .Color = color2
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .Color = color2
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .Color = color2
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .Color = color2
        .Weight = xlMedium
    End With
    With Selection.Borders(xlInsideVertical)
        .Color = color2
        .Weight = xlHairline
    End With
    Selection.Font.Bold = True

' titulos

Range(ActiveSheet.Cells(filapn, 1), ActiveSheet.Cells(filapn, 8)).Select
    Selection.Font.Bold = True

'total

ActiveSheet.Cells(filaun + 1, 1).Select
    Selection.Font.Bold = True
ActiveSheet.Cells(filaun + 1, 8).Select
    Selection.Font.Bold = True


'agrupar

Range(ActiveSheet.Cells(filapn, 1), ActiveSheet.Cells(filaun, ncolumnasnueva)).Select
Selection.Rows.Group

'poner espacio de titulo general

Range(ActiveSheet.Cells(filapn - 1, 2), ActiveSheet.Cells(filapn - 1, ncolumnasnueva - 1)).Select
With Selection
   .HorizontalAlignment = xlLeft
   .VerticalAlignment = xlTop
   
    End With
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
      
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
       
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
       
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
       
    End With
    
Selection.Merge

ActiveSheet.Cells(filapn - 1, 1).Select
With Selection
   .HorizontalAlignment = xlLeft
   .VerticalAlignment = xlTop
   
    End With
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
      
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
       
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
       
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
       
    End With
    
ActiveSheet.Cells(filapn - 1, ncolumnasnueva).Select
With Selection
   .HorizontalAlignment = xlLeft
   .VerticalAlignment = xlTop
   
    End With
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
      
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
       
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
       
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
       
    End With

'poner titulo en cada columna y sumar total

ActiveSheet.Cells(filapn, 1) = "TIPO"
ActiveSheet.Cells(filapn, 2) = "CÓDIGO"
ActiveSheet.Cells(filapn, 3) = "DESCRIPCIÓN"
ActiveSheet.Cells(filapn, 4) = "UNIDAD"
ActiveSheet.Cells(filapn, 5) = "FACTOR"
ActiveSheet.Cells(filapn, 6) = "CANTIDAD"
ActiveSheet.Cells(filapn, 7) = "PRECIO"
ActiveSheet.Cells(filapn, 8) = "TOTAL"
ActiveSheet.Cells(filaun + 1, ncolumnasnueva) = "=SUM([TOTAL])"
Range(ActiveSheet.Cells(filapn - 1, 2), ActiveSheet.Cells(filapn - 1, ncolumnasnueva - 1)) = "NOMBRE A.P.U."
ActiveSheet.Cells(filapn - 1, 1) = "CÓD"
ActiveSheet.Cells(filapn - 1, ncolumnasnueva) = "UND"

'poner formula en columnas tipo

TIPOS3 = "MATERIALES,EQUIPOS,MANO DE OBRA,TRANSPORTE,SUBCONTRATOS,ACTIVIDADES,OTROS,APU BÁSICOS"

contador1 = 0

For c = 0 To nfilasnueva

    ActiveSheet.Cells(filapn + 1 + contador1, 1).Select
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:=TIPOS3
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
    
contador1 = contador1 + 1
Next

'poner formula en columnas total

contador2 = 0

For c = 0 To nfilasnueva

    ActiveSheet.Cells(filapn + 1 + contador2, ncolumnasnueva).Select
    Selection.Formula = "=[@CANTIDAD]*[@PRECIO]*[@FACTOR]"
    
    ActiveSheet.Cells(filapn + 1 + contador2, 5).Select
    Selection.Formula = "= 1"
    
    
contador2 = contador2 + 1
Next
'-----------------------------------------------------
Range(ActiveSheet.Cells(filapn + 1, 4), ActiveSheet.Cells(filaun, 4)).Select
    With Selection
        .HorizontalAlignment = xlCenter
    End With

Range(ActiveSheet.Cells(filapn + 1, 5), ActiveSheet.Cells(filaun, 6)).Select
    Selection.NumberFormat = "General"

Range(ActiveSheet.Cells(filapn + 1, 7), ActiveSheet.Cells(filaun + 1, 8)).Select
Selection.NumberFormat = "_(* #,##0_);_(* (#,##0);_(* ""-""_);_(@_)"
'-----------------------------------------------------
ActiveSheet.Cells(filapn + 1, 1).Select

GoTo 2
1
OutPut = MsgBox("Es necesario ubicarse sobre una tabla de A.P.U.", vbCritical, "Error")
2
346
'-----------------------------------------------------
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True

End Sub



Public Sub CH2M2(control As IRibbonControl)
'control As IRibbonControl
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
'-----------------------------------------------------
On Error GoTo 346
On Error GoTo -1
On Error GoTo 2
filaini = ActiveCell.Row
columnaini = Selection.Column
ActiveSheet.Cells(filaini, columnaini).Select
Set SelectedCell_3 = ActiveCell

'nombre tabla existente
On Error GoTo 1
TableName_3 = SelectedCell_3.ListObject.Name
ActiveSheet.ListObjects(TableName_3).ListRows(1).Range.Select
Set ActiveTable = ActiveSheet.ListObjects(TableName_3)

' primera ultima y cantidad filas tabla existente

filap1_2 = Selection.Row
filap_2 = filap1_2 - 2
nfilas_2 = ActiveTable.Range.Rows.Count
filau_2 = filap_2 + nfilas_2 - 1
ncolumnasnueva_2 = 8

ActiveSheet.ListObjects(TableName_3).ListColumns(2).Range.Select
tipo = Selection.Column


If tipo = 2 Then GoTo 3 Else GoTo 1

3
Confirmacion = MsgBox("¿Desea eliminar el A.P.U.?", vbCritical + vbYesNoCancel, "Eliminar A.P.U.")
If Confirmacion = vbYes Then


Range(ActiveSheet.Cells(filap_2, 1), ActiveSheet.Cells(filau_2, ncolumnasnueva_2)).Select
Selection.Rows.Ungroup
Range(ActiveSheet.Cells(filap_2 - 1, 1), ActiveSheet.Cells(filau_2 + 2, ncolumnasnueva_2)).EntireRow.Select
Selection.Delete
ActiveSheet.Cells(filap_2 + 2, 1).Select

GoTo 2
1
OutPut = MsgBox("Es necesario ubicarse sobre una tabla de A.P.U.", vbCritical, "Error")
On Error Resume Next
ActiveSheet.Cells(filaini, columnaini).Select
2
Else
End If
ActiveSheet.Cells(filaini, columnaini).Select
346
'-----------------------------------------------------
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
End Sub


Public Sub GR1M3(control As IRibbonControl)
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
'-----------------------------------------------------
On Error GoTo 346
On Error GoTo -1
hojaactiva = ActiveSheet.Name
    If hojaactiva = "BÁSICOS" Then
        Application.Run ("CH2M4")
    Else
    If hojaactiva = "APU" Then
        Application.Run ("CH4M3")
    Else
    End If
    End If
346
'-----------------------------------------------------
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
End Sub


Sub CH2M4()
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
Application.AutoCorrect.AutoFillFormulasInLists = True
Application.DisplayAlerts = False
'-----------------------------------------------------
On Error GoTo 346
On Error GoTo -1
On Error GoTo 1
Set SelectedCell_4 = ActiveCell
filaoriginal_init = Selection.Row
columnaoriginal_init = Selection.Column

'nombre tabla existente
On Error GoTo 1
TableName_4 = SelectedCell_4.ListObject.Name
ActiveSheet.ListObjects(TableName_4).ListRows(1).Range.Select
Set ActiveTable = ActiveSheet.ListObjects(TableName_4)

'primera ultima y cantidad filas tabla existente

filap1_4 = Selection.Row
filap_4 = filap1_4 - 2
nfilas_4 = ActiveTable.Range.Rows.Count
filau_4 = filap_4 + nfilas_4 - 1
ncolumnasnueva_3 = 8
ActiveSheet.ListObjects(TableName_4).ListColumns(2).Range.Select
columna2 = Selection.Column

tipotabla = ActiveSheet.Cells(filap1_4 - 1, columna2)
filatabla = filaoriginal_init - filap1_4 + 1

If tipotabla = "CÓDIGO" Then GoTo 3 Else GoTo 1

3
valora = ActiveSheet.Cells(filaoriginal_init, 1).Value

If valora = "TIPO" Then GoTo 2 Else GoTo 4
4:
'desagrupar

Range(ActiveSheet.Cells(filap_4, 1), ActiveSheet.Cells(filau_4, ncolumnasnueva_3)).Select
Selection.Rows.Ungroup
'insertar

ActiveSheet.Cells(filaoriginal_init, 1).Select
Selection.ListObject.ListRows.Add (filatabla)
Range(ActiveSheet.Cells(filaoriginal_init, 10), ActiveSheet.Cells(filaoriginal_init, 12)).Select
Selection.Insert Shift:=xlDown
Range(ActiveSheet.Cells(filaoriginal_init, 15), ActiveSheet.Cells(filaoriginal_init, 16)).Select
Selection.Insert Shift:=xlDown
'agregar y eliminar filas para ajustar

Range(ActiveSheet.Cells(filau_4 + 3, 1), ActiveSheet.Cells(filau_4 + 3, ncolumnasnueva_3)).EntireRow.Select
Selection.Insert Shift:=xlDown
Range(ActiveSheet.Cells(filau_4 + 3, 1), ActiveSheet.Cells(filau_4 + 3, ncolumnasnueva_3)).Select
Selection.Delete
Range(ActiveSheet.Cells(filau_4 + 3, 10), ActiveSheet.Cells(filau_4 + 3, 12)).Select
Selection.Delete
Range(ActiveSheet.Cells(filau_4 + 3, 15), ActiveSheet.Cells(filau_4 + 3, 16)).Select
Selection.Delete
'poner factor

ActiveSheet.Cells(filaoriginal_init, 5).FormulaR1C1 = "=1"
'reagrupar

Range(ActiveSheet.Cells(filap_4 + 1, 1), ActiveSheet.Cells(filau_4 + 1, ncolumnasnueva_3)).Select
Selection.Rows.Group
SelectedCell_4.Offset(1, 0).Select
' validar celda A
TIPOS4 = "MATERIALES,EQUIPOS,MANO DE OBRA,TRANSPORTE,SUBCONTRATOS,ACTIVIDADES,OTROS"
ActiveSheet.Cells(filaoriginal_init, 1).Select
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:=TIPOS4
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With


GoTo 2
1
OutPut = MsgBox("Es necesario ubicarse sobre una tabla de A.P.U.", vbCritical, "Error")
2
On Error Resume Next
ActiveSheet.Cells(filaoriginal_init, columnaoriginal_init).Select
346
'-----------------------------------------------------
Application.DisplayAlerts = True
Application.Calculation = xlCalculationAutomatic
Cells(Rows.Count, 2).End(xlUp).Select
ultima = Selection.Row

Dim Celda2 As Range
For Each Celda2 In Range(ActiveSheet.Cells(2, 2), ActiveSheet.Cells(ultima, 8))
If IsError(Celda2.Value) Then Celda2.ClearContents
Next Celda2
Application.EnableEvents = True
ActiveSheet.Cells(filaoriginal_init, columnaoriginal_init).Select
End Sub


Sub ApuListaNormall()
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
Application.AutoCorrect.AutoFillFormulasInLists = False
'---------------------------------------------------------------------
On Error GoTo 346
On Error GoTo -1
Dim tbl As ListObject
Sheets("APU").Activate
For Each tbl2 In Sheets("APU").ListObjects
nomtbl2 = tbl2.Name
Sheets("APU").ListObjects(nomtbl2).ListColumns(1).Range.Select
columna1 = Selection.Column
If columna1 = 1 Then
cont = cont + 1
Else
End If
Next tbl2


'eliminar tabla de BBDD
On Error Resume Next
Sheets("LISTA_APU").Activate
ActiveSheet.ListObjects("T_apus").DataBodyRange.Delete
ActiveSheet.ListObjects("T_apus").Resize Range("$B$2:$E$" & cont + 4 & "")
cont1 = 0

'obtener nombre de cada tabla
For Each tbl In Sheets("APU").ListObjects

'para cada tabla en la hoja hacer:
      
Sheets("APU").Activate
TableName_5 = tbl.Name
Sheets("APU").ListObjects(TableName_5).ListRows(1).Range.Select
Set ActiveTable_5 = Sheets("APU").ListObjects(TableName_5)
' primera ultima y cantidad filas tabla existente

filap1_5 = Selection.Row
filap_5 = filap1_5 - 1
nfilas_5 = ActiveTable_5.Range.Rows.Count
filau_5 = filap_5 + nfilas_5 - 1
ncolumnasnueva_4 = 8
'obtener la primera columna

Sheets("APU").ListObjects(TableName_5).ListColumns(1).Range.Select
columna1 = Selection.Column
'Filtrar solo tablas de APU

If columna1 = 1 Then

'valor y nombre de cada apu
nombre = Sheets("APU").Cells(filap1_5 - 2, 2).Value
If nombre = "" Then
GoTo 44

Else
'variables

filavi = (filau_5) - (3 + cont1)
filai = (filap1_5 - 2) - (3 + cont1)
cvalor = (ncolumnasnueva_4) - (5)
cnomb = (3) - (2)
ccod = (0) - (0)
cund = (ncolumnasnueva_4) - (4)
'enlazar celdas

Sheets("LISTA_APU").Activate
'ActiveSheet.ListObjects("T_basico").ListRows.Add AlwaysInsert:=True
Sheets("LISTA_APU").Cells(3 + cont1, 2).FormulaR1C1 = "=APU!R[" & filai & "]C[0]"
Sheets("LISTA_APU").Cells(3 + cont1, 3).FormulaR1C1 = "=APU!R[" & filai & "]C[-2]"
Sheets("LISTA_APU").Cells(3 + cont1, 4).FormulaR1C1 = "=APU!R[" & filai & "]C[" & cund & "]"
Sheets("LISTA_APU").Cells(3 + cont1, 5).FormulaR1C1 = "=APU!R[" & filavi & "]C[" & cvalor & "]"

End If

Else
GoTo 44
End If

cont1 = cont1 + 1
44:
Next tbl

'-----------------------------------------------------
    ActiveWorkbook.Worksheets("LISTA_APU").ListObjects("T_apus").Sort.SortFields. _
        Clear
    ActiveWorkbook.Worksheets("LISTA_APU").ListObjects("T_apus").Sort.SortFields. _
        Add Key:=Range("T_apus[[#All],[DESCRIPCIÓN]]"), SortOn:=xlSortOnValues, _
        Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("LISTA_APU").ListObjects("T_apus").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
'-----------------------------------------------------

Application.Run ("CH1M2")
Application.ScreenUpdating = True
Sheets("LISTA_APU").Activate
Columns("C:E").EntireColumn.AutoFit
Columns("B:B").ColumnWidth = 40
Sheets("LISTA_APU").Cells(2, 1).Select
346
'---------------------------------------------------------------------
Application.ScreenUpdating = False
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
Application.AutoCorrect.AutoFillFormulasInLists = True
End Sub



Sub ApuListaBasicos()
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
Application.AutoCorrect.AutoFillFormulasInLists = False
'---------------------------------------------------------------------
On Error GoTo 346
On Error GoTo -1
Dim tbl As ListObject
filaini_5 = ActiveCell.Row
columnaini_5 = ActiveCell.Column
Sheets("BÁSICOS").Activate

For Each tbl2 In Sheets("BÁSICOS").ListObjects
nomtbl2 = tbl2.Name
Sheets("BÁSICOS").ListObjects(nomtbl2).ListColumns(1).Range.Select
columna1 = Selection.Column
If columna1 = 1 Then
cont = cont + 1
Else
End If
Next tbl2


'eliminar tabla de BBDD
On Error Resume Next
Sheets("LISTA_BÁSICOS").Activate
ActiveSheet.ListObjects("T_basico").DataBodyRange.Delete
ActiveSheet.ListObjects("T_basico").Resize Range("$B$2:$E$" & cont + 4 & "")
cont1 = 0

'obtener nombre de cada tabla
For Each tbl In Sheets("BÁSICOS").ListObjects

'para cada tabla en la hoja hacer:
      
Sheets("BÁSICOS").Activate
TableName_5 = tbl.Name
Sheets("BÁSICOS").ListObjects(TableName_5).ListRows(1).Range.Select
Set ActiveTable_5 = Sheets("BÁSICOS").ListObjects(TableName_5)
' primera ultima y cantidad filas tabla existente

filap1_5 = Selection.Row
filap_5 = filap1_5 - 1
nfilas_5 = ActiveTable_5.Range.Rows.Count
filau_5 = filap_5 + nfilas_5 - 1
ncolumnasnueva_4 = 8
'obtener la primera columna

Sheets("BÁSICOS").ListObjects(TableName_5).ListColumns(1).Range.Select
columna1 = Selection.Column
'Filtrar solo tablas de APU

If columna1 = 1 Then

'valor y nombre de cada apu
nombre = Sheets("BÁSICOS").Cells(filap1_5 - 2, 2).Value
If nombre = "" Then
GoTo 44

Else
'variables

filavi = (filau_5) - (3 + cont1)
filai = (filap1_5 - 2) - (3 + cont1)
cvalor = (ncolumnasnueva_4) - (5)
cnomb = (3) - (2)
ccod = (1) - (1)
cund = (ncolumnasnueva_4) - (4)
'enlazar celdas

Sheets("LISTA_BÁSICOS").Activate
'ActiveSheet.ListObjects("T_basico").ListRows.Add AlwaysInsert:=True
Sheets("LISTA_BÁSICOS").Cells(3 + cont1, 2).FormulaR1C1 = "=BÁSICOS!R[" & filai & "]C[0]"
Sheets("LISTA_BÁSICOS").Cells(3 + cont1, 3).FormulaR1C1 = "=BÁSICOS!R[" & filai & "]C[-2]"
Sheets("LISTA_BÁSICOS").Cells(3 + cont1, 4).FormulaR1C1 = "=BÁSICOS!R[" & filai & "]C[" & cund & "]"
Sheets("LISTA_BÁSICOS").Cells(3 + cont1, 5).FormulaR1C1 = "=BÁSICOS!R[" & filavi & "]C[" & cvalor & "]"

End If

Else
GoTo 44
End If

cont1 = cont1 + 1
44:
Next tbl
'-----------------------------------------------------
    ActiveWorkbook.Worksheets("LISTA_BÁSICOS").ListObjects("T_basico").Sort.SortFields. _
        Clear
    ActiveWorkbook.Worksheets("LISTA_BÁSICOS").ListObjects("T_basico").Sort.SortFields. _
        Add Key:=Range("T_basico[[#All],[DESCRIPCIÓN]]"), SortOn:=xlSortOnValues, _
        Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("LISTA_BÁSICOS").ListObjects("T_basico").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
'-----------------------------------------------------

Application.Run ("CH1M2")
Application.ScreenUpdating = True
Sheets("LISTA_BÁSICOS").Activate
Columns("C:E").EntireColumn.AutoFit
Columns("B:B").ColumnWidth = 40
Sheets("LISTA_BÁSICOS").Cells(2, 1).Select
346
'---------------------------------------------------------------------
Application.ScreenUpdating = False
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
Application.AutoCorrect.AutoFillFormulasInLists = True
End Sub



Public Sub CH6M28V2()
'GUARDAR COMO SIN BOTON
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
'-----------------------------------------------------
On Error GoTo 346
On Error GoTo -1
'Declaramos las variables.
Dim VentanasProtegidas As Boolean
Dim EstructuraProtegida As Boolean
Dim NombreHoja As String
Dim Confirmacion As String
Dim NombreArchivo As String
Dim GuardarComo As Variant
Dim Extension As String
Dim librooriginal As String
'-----------------------------------------------------
ThisWorkbook.Unprotect "123"
'-----------------------------------------------------

'
'En caso de error.
On Error GoTo ErrorHandler
'
'Validamos si la ventana o la estructura del archivo están protegidos.
VentanasProtegidas = ActiveWorkbook.ProtectWindows
EstructuraProtegida = ActiveWorkbook.ProtectStructure
'
'En caso de estar protegidas mostramos mensaje.
If VentanasProtegidas = True Or EstructuraProtegida = True Then
    MsgBox "No se puede ejecutar el comando cuando la estructura del archivo está protegida.", _
           vbExclamation, "S5S"
Else
    '
    'Copiamos la hoja y guardamos.
    librooriginal = ActiveWorkbook.Name
    hojaoriginal = ActiveSheet.Name
    Sheets("INSUMOS").Activate

        Sheets("INSUMOS").Activate
        ActiveSheet.Select
        ActiveSheet.Copy
        NombreArchivo = ActiveWorkbook.Name
        Workbooks(librooriginal).Activate
        On Error Resume Next
            Application.DisplayAlerts = False
        Sheets("BÁSICOS").Copy After:=Workbooks(NombreArchivo).Sheets(1)
        Workbooks(librooriginal).Activate
        Sheets("APU").Copy After:=Workbooks(NombreArchivo).Sheets(2)
        Workbooks(librooriginal).Activate
        Sheets("PRESUPUESTO").Copy After:=Workbooks(NombreArchivo).Sheets(3)
            Application.DisplayAlerts = True
        Workbooks(NombreArchivo).Activate
        GuardarComo = Application.GetSaveAsFilename(InitialFileName:=NombreHoja, _
            FileFilter:="Libro de Excel(*.xlsx), *.xlsx", Title:="S5S - guadar copia del proyecto.")
        If GuardarComo = False Then
            Workbooks(NombreArchivo).Close SaveChanges:=False
        Else
            With Application.WorksheetFunction
                Extension = .Trim(Right(.Substitute(GuardarComo, ".", .Rept(" ", 650)), 650))
            End With
            
Workbooks(librooriginal).Activate
Sheets("LISTA_BÁSICOS").Activate
Sheets("LISTA_BÁSICOS").Cells(2, 16200) = GuardarComo
Sheets("LISTA_BÁSICOS").Cells(2, 16200).Select
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
5555555

'----------------------------------------
 'impiar nombres

Workbooks(NombreArchivo).Activate
Dim nName As Name
Dim lReply As Long
    For Each nName In Names
        nName.Delete
    Next nName
'------------------------------------------------------------------------
'modificar botones

Sheets("INSUMOS").Activate
ActiveSheet.Shapes.SelectAll
Selection.Delete

Sheets("BÁSICOS").Activate
ActiveSheet.Shapes.SelectAll
Selection.Delete

Sheets("APU").Activate
ActiveSheet.Shapes.SelectAll
Selection.Delete

Sheets("PRESUPUESTO").Activate
ActiveSheet.Shapes.SelectAll
Selection.Delete
'----------------------------------------
'Quitar vinculos

Workbooks(NombreArchivo).Activate
nombuscar = librooriginal & "!"


Sheets("INSUMOS").Activate
    Cells.Replace What:=nombuscar, Replacement:="", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
            ActiveSheet.Cells(2, 1).Select
        
Sheets("BÁSICOS").Activate
    Cells.Replace What:=nombuscar, Replacement:="", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Columns("J:AA").Select
        Selection.Delete Shift:=xlToLeft
            ActiveSheet.Cells(2, 1).Select
        
Sheets("APU").Activate
    Cells.Replace What:=nombuscar, Replacement:="", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Columns("J:AA").Select
        Selection.Delete Shift:=xlToLeft
            ActiveSheet.Cells(2, 1).Select
               
nombuscar = "[" & librooriginal & "]"
Sheets("PRESUPUESTO").Activate

    Columns("J:AA").Select
        Selection.Delete Shift:=xlToLeft


    Cells.Replace What:=nombuscar, Replacement:="", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
            ActiveSheet.Cells(2, 2).Select
        
'------------------------------------------------------------------------
   Dim ExternalLinks As Variant
Dim wb As Workbook
Dim x As Long
    Set wb = ActiveWorkbook
    On Error GoTo 1001
    On Error GoTo -1
    ExternalLinks = wb.LinkSources(Type:=xlLinkTypeExcelLinks)
        For x = 1 To UBound(ExternalLinks)
            wb.BreakLink Name:=ExternalLinks(x), Type:=xlLinkTypeExcelLinks
        Next x
'------------------------------------------------------------------------
 'guardar
1001
            
Application.DisplayAlerts = False
            Select Case Extension
            Case Is = "xlsx"
                ActiveWorkbook.SaveAs GuardarComo
            End Select
nombrelibro2 = ActiveWorkbook.Name

 '------------------------------------------------------------------------
End If
End If
Workbooks(nombrelibro2).Close SaveChanges:=False

Sheets(hojaoriginal).Activate
ThisWorkbook.Protect "123"
'----------------------------------
Application.DisplayAlerts = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
Exit Sub
'
'En caso de error mostramos un mensaje.
ErrorHandler:
Workbooks(hojaoriginal).Close SaveChanges:=False
Sheets(librooriginal).Activate
346
'-----------------------------------------------------
Application.DisplayAlerts = True
ThisWorkbook.Protect "123"
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
End Sub


