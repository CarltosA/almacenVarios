Private Sub CommandButton1_Click()
  Range("C2:F7").ClearContents
  Range("C8").ClearContents
  Range("F8").ClearContents
  Range("A11:G20").ClearContents

  Set MIRANGO = Worksheets("CABECERA").Range("A2:A10000")
  var_filasocupadas = WorksheetFunction.CountA(MIRANGO) + 1
  Range("E2").Value = var_filasocupadas
  
  Range("C4").NumberFormat = "dd/mm/yyyy"
  Range("C4").Value = CDate(Date)

  Range("C8").Value = "ACT"
  Range("F8").Value = "NUEVO"
  
End Sub

Private Sub CommandButton2_Click()
  If Worksheets("GUIA").Range("F8").Value <> "NUEVO" Then
    Exit Sub
  End If

  Dim VAR_HOJACABECERA, VAR_HOJADETALLE As Worksheet
  Dim VAR_TABLACABECERA, VAR_TABLADETALLE As ListObject
  Dim NUEVACABECERA, NUEVADETALLE As ListRow
  Dim Pregunta As Byte
  
  Set VAR_HOJACABECERA = ThisWorkbook.Sheets("CABECERA")
  Set VAR_TABLACABECERA = VAR_HOJACABECERA.ListObjects("TABLACABECERA")
  Set NUEVACABECERA = VAR_TABLACABECERA.ListRows.Add
  
  VAR_CANTIDAD_ARTICULO = 0
  Set VAR_MIRANGO = Worksheets("GUIA").Range("D11:D20")
  VAR_CANTIDAD_ARTICULO = WorksheetFunction.CountA(VAR_MIRANGO)
  
  With NUEVACABECERA
    .Range(1) = "C" & WorksheetFunction.Text(Worksheets("GUIA").Range("E2").Value, "00000")
    .Range(2) = Worksheets("GUIA").Range("E2").Value
    .Range(3) = Worksheets("GUIA").Range("C8").Value
    .Range(4) = Worksheets("GUIA").Range("C4").Value
    .Range(5) = Worksheets("GUIA").Range("C5").Value
    .Range(6) = Worksheets("GUIA").Range("C6").Value
    .Range(7) = VAR_CANTIDAD_ARTICULO
    .Range(8) = Worksheets("GUIA").Range("C7").Value
  End With
  
  Set VAR_HOJADETALLE = ThisWorkbook.Sheets("DETALLE")
  Set VAR_TABLADETALLE = VAR_HOJADETALLE.ListObjects("TABLADETALLE")
   
  For i = 1 To VAR_CANTIDAD_ARTICULO
    Set NUEVADETALLE = VAR_TABLADETALLE.ListRows.Add
    With NUEVADETALLE
      .Range(1) = Worksheets("GUIA").Range("C4").Value
      .Range(2) = "C" & WorksheetFunction.Text(Worksheets("GUIA").Range("E2").Value, "00000") & "D" & WorksheetFunction.Text(i, "00")
      .Range(3) = "C" & WorksheetFunction.Text(Worksheets("GUIA").Range("E2").Value, "00000")
      .Range(4) = "GR"
      .Range(5) = Worksheets("GUIA").Range("E2").Value
      .Range(6) = i
      .Range(7) = Worksheets("GUIA").Range("C8").Value
      .Range(8) = Worksheets("GUIA").Range("C" & i + 10).Value
      .Range(9) = Worksheets("GUIA").Range("D" & i + 10).Value
      .Range(10) = Worksheets("GUIA").Range("A" & i + 10).Value
      .Range(11) = "ENT"
    End With
  Next
  
  Pregunta = MsgBox("Se guardaron los datos Deseas nuevo ingreso?", vbYesNo + vbQuestion)

  If Pregunta = vbNo Then Exit Sub
  CommandButton1_Click

End Sub
Private Sub CommandButton3_Click()
  UserForm3.Show
End Sub

Private Sub CommandButton4_Click()
'  Set VAR_RANGOELIMINAR = ThisWorkbook.Sheets("CABECERA").ListObjects("TABLACABECERA")
'
'  Dim i As Integer
'
'  With VAR_RANGOELIMINAR
'    For i = .ListRows.Count To 1 Step -1
'''      If Len(.ListRows(i).Range.Cells(1)) <= 3 Then
''        If .ListRows.Count <= 1 Then
''          MsgBox ("You must enter at least 4 characters of the DeptID"), vbCritical
''          Exit Sub
''        End If
'        .ListRows(i).Delete
''      End If
'    Next i
'  End With
'
'  With ActiveWorkbook.Worksheets("CABECERA")
'    With .ListObjects(1).DataBodyRange
'      .AutoFilter
'      .AutoFilter Field:=6, Criteria1:="5"
'      .EntireRow.Delete
'      .AutoFilter
'    End With
'  End With
'
'  With ActiveWorkbook.Worksheets("DETALLE")
'    With .ListObjects(1).DataBodyRange
'      .AutoFilter
'      .AutoFilter Field:=6, Criteria1:="5"
'      .EntireRow.Delete
'      .AutoFilter
'    End With
'  End With


UserForm4.Show
End Sub
Private Sub CommandButton5_Click()
  UserForm1.Show
End Sub
Private Sub CommandButton6_Click()
 UserForm2.Show
End Sub

Private Sub CommandButton8_Click()
  With ActiveWorkbook.Worksheets("DETALLE").ListObjects(1).ListColumns(2).DataBodyRange
    .FormatConditions.AddUniqueValues
    .FormatConditions(1).DupeUnique = xlDuplicate
    .FormatConditions(1).Interior.Color = vbRed
  End With
End Sub
