Sub UpdateCoucher()

Dim aWB As Workbook, agenda As Workbook, RowDate, ColDate, ColDatel, ag, qs, qr, hj, ai, ae, couche, leve, boiss, cboiss, lboiss, csieste, lsieste, dsieste, c, l

Set aWB = ActiveWorkbook
Workbooks.Open (ActiveWorkbook.Path & "\Agenda du Sommeil PRO-SP.xlsx")
Set agenda = ActiveWorkbook

'Calculer la taille du tableau
Dim myWorkSheet As Worksheet, myTable As ListObject, tabLength As Long
Set myWorkSheet = ActiveWorkbook.Worksheets("Form")
Set myTable = myWorkSheet.ListObjects("DataFromForm")
tabLength = myTable.DataBodyRange.Rows.Count + 1
        

    
'Pour chaque ligne
    For Each cpt In agenda.Sheets("Form").Range("M74:M" & tabLength)
    
    RowDate = agenda.Sheets("Form").Range("T" & CStr(cpt.Row))
    ColDate = agenda.Sheets("Form").Range("U" & CStr(cpt.Row))

    If cpt.Value = 0 Then
        'Couche
        Set agenda = ActiveWorkbook
        If IsEmpty(agenda.Sheets("Form").Range("B" & CStr(cpt.Row))) = False Then
        RowDate = agenda.Sheets("Form").Range("T" & CStr(cpt.Row))
        ColDate = agenda.Sheets("Form").Range("U" & CStr(cpt.Row))
        'MsgBox (CStr(ColDate) & CStr(RowDate))
        agenda.Sheets("Agenda").Range(CStr(ColDate) & CStr(RowDate)) = "c"
        c = agenda.Sheets("Agenda").Range(CStr(ColDate) & CStr(RowDate)).Column
        couche = CStr(ColDate) & CStr(RowDate)
        End If
        
        'LEVE
        If IsEmpty(agenda.Sheets("Form").Range("C" & CStr(cpt.Row))) = False Then
        RowDate = agenda.Sheets("Form").Range("AA" & CStr(cpt.Row))
        ColDatel = agenda.Sheets("Form").Range("AB" & CStr(cpt.Row))
        'MsgBox (CStr(ColDate) & CStr(RowDate))
        agenda.Sheets("Agenda").Range(CStr(ColDatel) & CStr(RowDate)) = "l"
        l = agenda.Sheets("Agenda").Range(CStr(ColDatel) & CStr(RowDate)).Column
        leve = CStr(ColDatel) & CStr(RowDate)
        End If
        
          

        'Qualité sommeil
        If agenda.Sheets("Form").Range("F" & CStr(cpt.Row)) = "Très bien" Then
        qs = "TB"
        ElseIf agenda.Sheets("Form").Range("F" & CStr(cpt.Row)) = "Bien" Then
        qs = "B"
        ElseIf agenda.Sheets("Form").Range("F" & CStr(cpt.Row)) = "Moyen" Then
        qs = "Moy"
        ElseIf agenda.Sheets("Form").Range("F" & CStr(cpt.Row)) = "Mauvais" Then
        qs = "Ma"
        ElseIf agenda.Sheets("Form").Range("F" & CStr(cpt.Row)) = "Très Mauvais" Then
        qs = "TM"
        End If
        
        If IsEmpty(agenda.Sheets("Form").Range("F" & CStr(cpt.Row))) = False Then
            agenda.Sheets("Agenda").Range("CT" & CStr(RowDate)) = qs
        
        End If
        
        
       
        'Qualité Reveil
        If agenda.Sheets("Form").Range("G" & CStr(cpt.Row)) = "Très bien" Then
        qr = "TB"
        ElseIf agenda.Sheets("Form").Range("G" & CStr(cpt.Row)) = "Bien" Then
        qr = "B"
        ElseIf agenda.Sheets("Form").Range("G" & CStr(cpt.Row)) = "Moyen" Then
        qr = "Moy"
        ElseIf agenda.Sheets("Form").Range("G" & CStr(cpt.Row)) = "Mauvais" Then
        qr = "Ma"
        ElseIf agenda.Sheets("Form").Range("G" & CStr(cpt.Row)) = "Très Mauvais" Then
        qr = "TM"
        End If
        If IsEmpty(agenda.Sheets("Form").Range("G" & CStr(cpt.Row))) = False Then
            agenda.Sheets("Agenda").Range("Cu" & CStr(RowDate)) = qr
        End If
        
         
         'Hummeur Journee
        If agenda.Sheets("Form").Range("H" & CStr(cpt.Row)) = "Très bien" Then
        hj = "TB"
        ElseIf agenda.Sheets("Form").Range("H" & CStr(cpt.Row)) = "Bien" Then
        hj = "B"
        ElseIf agenda.Sheets("Form").Range("H" & CStr(cpt.Row)) = "Moyen" Then
        hj = "Moy"
        ElseIf agenda.Sheets("Form").Range("H" & CStr(cpt.Row)) = "Mauvais" Then
        hj = "Ma"
        ElseIf agenda.Sheets("Form").Range("H" & CStr(cpt.Row)) = "Très Mauvais" Then
        hj = "TM"
        End If
        
        If IsEmpty(agenda.Sheets("Form").Range("H" & CStr(cpt.Row))) = False Then
            agenda.Sheets("Agenda").Range("CV" & CStr(RowDate)) = hj
            
        End If
        
        'Anxiétés
        If IsEmpty(agenda.Sheets("Form").Range("I" & CStr(cpt.Row))) = False Then
        agenda.Sheets("Agenda").Range("CW" & CStr(RowDate)) = agenda.Sheets("Form").Range("I" & CStr(cpt.Row))
        End If
        If IsEmpty(agenda.Sheets("Form").Range("J" & CStr(cpt.Row))) = False Then
        agenda.Sheets("Agenda").Range("CX" & CStr(RowDate)) = agenda.Sheets("Form").Range("J" & CStr(cpt.Row))
        End If
        
       
       
        
         If agenda.Sheets("Form").Range("K" & CStr(cpt.Row)) = "Café" Then
            boiss = "ca"
            ElseIf agenda.Sheets("Form").Range("K" & CStr(cpt.Row)) = "Redbull" Then
            boiss = "t"
        End If
        
        'Boissons
        'UPD boiss
        If IsEmpty(agenda.Sheets("Form").Range("K" & CStr(cpt.Row))) = False Then
        lboiss = agenda.Sheets("Form").Range("AH" & CStr(cpt.Row))
        cboiss = agenda.Sheets("Form").Range("AI" & CStr(cpt.Row))
        agenda.Sheets("Agenda").Range(CStr(cboiss) & CStr(lboiss)) = boiss
        End If
        
        
        
        'Sieste
        If IsEmpty(agenda.Sheets("Form").Range("D" & CStr(cpt.Row))) = False And IsEmpty(agenda.Sheets("Form").Range("E" & CStr(cpt.Row))) = False Then
            lsieste = agenda.Sheets("Form").Range("AP" & CStr(cpt.Row))
            csieste = agenda.Sheets("Form").Range("AQ" & CStr(cpt.Row))
            agenda.Sheets("Agenda").Range(CStr(csieste) & CStr(lsieste)) = "1"
            If agenda.Sheets("Form").Range("AO" & CStr(cpt.Row)) > 15 Then
            'MsgBox (agenda.Sheets("Form").Range("AR" & CStr(cpt.Row)))
                Dim cptS, nSieste, nnSieste As Range
                cptS = 0
            
                          
            End If
                
        End If
      
      
      '??
           ' If agenda.Sheets("Agenda").Range("B738").Value = "c" Then
            'If agenda.Sheets("Agenda").Range(CStr(ColDatel) & CStr(RowDate)).Value = "l" Then
            'For I = c + 1 To l - 1
      
             '   agenda.Sheets("Agenda").Cells(RowDate, I).Value = 1
      
     '       Next I
      '
       '     End If
       '
       ' End If
       
       
' END STEP -------------------------
cpt.Value = 1
   End If
    Next cpt

ActiveWorkbook.Save
Workbooks("Agenda du Sommeil PRO-SP.xlsx").Close
End Sub

