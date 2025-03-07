Option Explicit
' Written by Henry

Sub CabNotSetup()

    ' ################################################################################
    ' ####### CHANGE THIS IF THE CABINET NOTIFICATION FILE MOVES OR IS RENAMED #######
    
    Const strCabinetNotificationsFile = "H:\Project Current\Templates\FTTP/Cabinet Notification Template v2.7.xlsm"
    
    ' ####### CHANGE THIS IF THE CABINET NOTIFICATION FILE MOVES OR IS RENAMED #######
    ' ################################################################################

    Dim wkbBuster As Workbook
    Dim wkbCabsTemp As Workbook
    Dim wsTombstone As Worksheet
    Dim wsSetup As Worksheet
    
    
    '''''''''''Set This workbook as wkbbuster''''''''''''''''''''''
    Set wkbBuster = ThisWorkbook
    
    
  'Check L3 exists
   If Range("D15").Value = 0 Then
        MsgBox "You have No L3 in this Buster"
        GoTo Errhandling2
    End If

    
    ' Open the cabinet notifications template
    Workbooks.Open strCabinetNotificationsFile
        
    If Sheets(1).Name = "Setup" Then
        Set wkbCabsTemp = ActiveWorkbook
    Else
        MsgBox "The file has moved"
        GoTo Exit_CabNotSetup
    End If
    
    Sheets(1).Activate
    If Range("M4").Value > 0 Then
        MsgBox "This document aleady has data in"
        GoTo Exit_CabNotSetup
    End If

    '''''''''''''''''''''''''''''''''''''''''''''Copy Basic Data & L3''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim VrL3 As Variant
    Dim StrOpp As String
    Dim StrL3 As String
    Dim StrNBU As String
    'Dim StrDate As Date 'Muted as doesn't match utilities date


    wkbBuster.Activate



    Sheets(1).Activate
    Set wsTombstone = ActiveSheet
    Range("D3").Select  'Set Oppurtunity
    StrOpp = Selection
    Range("D4").Select  'Set L3 code
    StrL3 = Selection
    Range("C15:L15").Select 'Set L3 details
    VrL3 = Selection
    Range("D5").Select  'Set NBU
    StrNBU = Selection
    'Muted as doesn't match utilities date
    'Range("D7").Select  ' set date
    'StrDate = Selection
    
    



    ''''''''''''''''''''''''''''''''''''''Save & Test''''''''''''''''''''''''''''''''''

    'Set CabTemp Sheets to allow for early Save

    wkbCabsTemp.Activate
    Sheets(1).Activate
    Set wsSetup = ActiveSheet

    wsSetup.Activate

    Dim folderpath As String

    folderpath = wkbBuster.Path
    On Error GoTo Errhandling
    wkbCabsTemp.SaveAs folderpath & "/" & StrL3 & " Cabinet Notifications v1" & ".xlsm"


    If wkbCabsTemp.Name = StrL3 & " Cabinet Notifications v1" & ".xlsm" Then
        GoTo Continue
        'If save has failed close Cab Template from FTTP file without saving to prevent changes to template. If error happens later, will be saved to Local Support info.
    Else
        MsgBox "Failed to save the workbook. Make sure the file name is valid and not in use.", vbExclamation
        wkbCabsTemp.Close Savechanges:=False
        GoTo Exit_CabNotSetup
    End If


Continue:


    ''''''''''''''''''''''''''''''''''''''Paste Basic Data in Cab Nots''''''''''''''''''''''''''''''''''
    Range("C3").Value = StrOpp
    Range("C4").Value = StrNBU
    Range("L4:U4").Value = VrL3
    'Range("C17").Value = StrDate 'Muted as doesn't match utilities date
    'Range("C17").NumberFormat = "dd/mm/yyyy"
    
    ''''''''''''''''''Create Loop to find other cabs within Buster''''''''''''''''''''''''''''''''''''''
    wsTombstone.Activate

    Dim i As Long
    Dim Cabtable As Range
    
    ' Loop through the rows starting from F16 to F135
    For i = 16 To 135
        If IsNumeric(wsTombstone.Cells(i, "F").Value) Then
            If wsTombstone.Cells(i, "F").Value > 0 Then
                wsTombstone.Range("C" & i & ":L" & i).Copy ' Copy the entire row from Column C to Column L
                
                Set Cabtable = wsSetup.Cells(wsSetup.Rows.Count, "L").End(xlUp).Offset(1, 0)
                Cabtable.PasteSpecial xlPasteValues
                Application.CutCopyMode = False
            Else
                Exit For ' Exit the loop when F = 0
            End If
        End If
    Next i

    '''''''''''''''''''''''''''''Rename Tabs based on Cab Names''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    wkbCabsTemp.Activate

    Dim wsCab As Worksheet
    Dim Cabname As String
    Dim hideSheets As Boolean
    
   
    For Each wsCab In Worksheets(Array(3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21))
        If wsCab.Range("AK24").Value <> 0 Then
            Cabname = wsCab.Range("AK24").Value
            wsCab.Name = Cabname
        Else
            wsCab.Visible = xlSheetHidden
            hideSheets = True
        End If
        
        If hideSheets Then
            wsCab.Visible = xlSheetHidden
        End If
    Next wsCab

    'Hide OLT sheet
    Sheets(2).Visible = xlSheetHidden


    '''''''''''''' Count Number of Cabs''''''''''''''''''''''''''''''''''''''''''''
    wsSetup.Activate

    Dim Countrange As Range
    Dim cell As Range
    Dim Notempty As Long

    Set Countrange = Range("L5:L23")

    Notempty = 0

    ' Loop through each cell in the target range and check if it is not empty
    For Each cell In Countrange
        If Not IsEmpty(cell.Value) And cell.Value <> "" Then
            Notempty = Notempty + 1
        End If
    Next cell
    
    MsgBox "You have 1 x L3 Cabinet & " & Notempty & " x L4 Cabinet. Workbook saved in Support Info as: " & StrL3 & " Cabinet Notifications v1", vbInformation


Exit_CabNotSetup:
    Set wkbBuster = Nothing
    Set wkbCabsTemp = Nothing

    Exit Sub


Errhandling:
    MsgBox "Failed to save the workbook. Make sure the file name is not in use within your support info.", vbExclamation
    wkbCabsTemp.Close Savechanges:=False
    GoTo Exit_CabNotSetup

Errhandling2:
   ' wkbCabsTemp.Close savechanges:=False
    GoTo Exit_CabNotSetup

End Sub
