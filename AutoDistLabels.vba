Option Explicit


   Sub AutoDistLabels()


'Workbooks & Worksheets
    Dim wkbBuster As Workbookh
    Dim wsTombstones As Worksheet
    
'Find Dist Labels Sheet
    Dim wsDistLabels As Worksheet
    Dim shtSearcDistLabels As Boolean
    Dim shtDistlabel As String

'Find the position of the Comma in the DP XY coords
    Dim Jointcoords As String
    Dim Jointcommacheck As String
    Dim L4Coords As String
    Dim commaPosition As Long

'Move block attribtues table with an Offset
    Dim BlockAtttype As Range
    Dim TargetRange As Range

'loop to put cable label name in block
    Dim i As Long
    Dim strID As String 'Label ID
    Dim rngIDTarget As Range
    Dim strBaseP As Variant 'Basepoint
    Dim rngBasePTarget As Range
    Dim strcabletype As String 'Cable Type
    Dim rngcabletypetarget As Range
    Dim strLength As String 'Length
    Dim rngLengthtarget As Range
    Dim strMeasure As String 'Measure (Y/N)
    Dim rngMeasuretarget As Range
    Dim strDuctType As String 'Duct Type
    Dim rngDuctTypeTarget As Range
    Dim strFibre As String 'Fibre count
    Dim rngFibreTarget As Range

'Loop for Dist Joints
    Dim o As Long

'Coords Loop
    Dim m As Long

'create loop to look for N in measure column and add a circle if N with that basepoint for L4s
    Dim u As Long
    Dim rngCircle As Range
    Dim rngCircleTarget As Range
    Dim rngcircletargetresize As Range
    Dim rngCircleBasePTarget As Range

'Repeat Circle Loop for uncosted Joints
    Dim v As Long

' Reset System Variable
    Dim LastCell As Range
    Dim LastRow As Long
    Dim rngCopy As Range

'Yes/No for Joints
    Dim YesNo As VbMsgBoxResult

'PIAm/NEXm
    Dim rngPIAmTarget As Range 'PIAm
    Dim intPIAm As Integer
    Dim rngNEXmTarget As Range 'NEXm
    Dim intNEXm As Integer

'Declaring dynamic Columns

    Dim rngL4Coords As Range 'L4 Co-ords COlumn
    Dim lngL4CoordsCol As Long
    Dim rngstartCoord As Range 'Joint Co-ords Column
    Dim lngJntCoordscol As Long

'L4 dynamic columns

    Dim rngCADlabel As Range 'Full CAD cable Label column
    Dim lngCADlabelcol As Long
    Dim rngCableType As Range 'Cable Type column
    Dim lngCableTypecol As Long
    Dim rnglength As Range 'Length column
    Dim lngLengthcol As Long
    Dim rngDuctType As Range 'Duct Type Column
    Dim lngDuctTypecol As Long
    Dim rngMes As Range 'Measure Column
    Dim lngMesCol As Long
    Dim rngFbrCount As Range 'FibreCount Column
    Dim lngFbrCountcol As Long
    Dim rngL4PIAm As Range 'L4 PIAm Column
    Dim lngL4PIAmcol As Long
    Dim rngL4NEXm As Range 'L4 NEXm Column
    Dim lngL4NEXmcol As Long

'Joint dynamic columns

    Dim rngJntCADLabel As Range 'Joint CAD label Column
    Dim lngJntCADLabelcol As Long
    Dim rngJntCblType As Range 'Joint Cable Type Column
    Dim lngJntCblTypecol As Long
    Dim rngJntLength As Range 'Joint Length Column
    Dim lngJntLengthcol As Long
    Dim rngJntDuctType As Range 'Joint Duct Type Column
    Dim lngJntDuctTypecol As Long
    Dim rngJntMes As Range 'Joint Measure Column
    Dim lngJntMescol As Long
    Dim rngJntFbrCount As Range 'Joint Fibre Count Column
    Dim lngJntFbrCountcol As Long
    Dim rngJntPIAm As Range 'Joint PIAm Column
    Dim lngJntPIAmcol As Long
    Dim rngJntNEXm As Range 'Joint NEXm Column
    Dim lngJntNEXmcol As Long

'Declare Columns for Circle Loop
    Dim rngCreateCircle As Range 'Create Circle Column
    Dim lngCreateCirclecol As Long
    
'Selected Labels only
    Dim rngSelectedL4Lbl As Range
    Dim lngSelectedL4Lbl As Long
    Dim rngSelectedJntLbl As Range
    Dim lngSelectedJntLbl As Long
    Dim blAllLabelsLink As Boolean

'Set Block Attributes & Non-dynamic attributes table in columns A & B - Loop to do this.

    Dim labels As Variant
    Dim values As Variant
    Dim d As Integer


'Produce All Dist Labels True/False?
    blAllLabelsLink = Cells.Find(what:="Produce All Labels  Linked Cell:", LookIn:=xlValues, Lookat:=xlWhole).Offset(3, 0)


    Set wkbBuster = ActiveWorkbook

'Application.ScreenUpdating = False


'Find & Set Dist Labels within Buster


    shtDistlabel = "Distribution Labels"

        If shtDistlabel = "" Then Exit Sub
            On Error Resume Next
            wkbBuster.Sheets(shtDistlabel).Select
            shtSearcDistLabels = (Err = 0)
            On Error GoTo 0
                If shtSearcDistLabels Then
                'MsgBox "Sheet '" & ShtNameUPRN & "' has been found!"
                'Set wkbRsheet = ActiveWorkbook
                    Set wsDistLabels = wkbBuster.Sheets("Distribution Labels")
                'MsgBox "Distribution Labels tab has been found"
                    Else
                    MsgBox "You have selected the wrong documents"
                    Exit Sub
        End If
    

'Open Distlabels Tab

    wsDistLabels.Activate
    Range("A1").Select


'change number formate to general so scale isnt rounded up
    Columns("A").NumberFormat = "General"



'Check if Dist labels have already been created.


    If Range("A1").Value > 0 Then
        Columns("A:B").Clear
        Else
    End If

'''''''''''''''''''''''''''''''''''''Set Dynamic Columns'''''''''''''''''''''''''''''''''''''''''''''''''''''


'L4 & Joint Co-ords
    Set rngL4Coords = Cells.Find(what:="Coordinates", LookIn:=xlValues, Lookat:=xlWhole) 'L4 Co-ords
    lngL4CoordsCol = rngL4Coords.Column

    Set rngstartCoord = Cells.Find(what:="Start Coord", LookIn:=xlValues, Lookat:=xlWhole)
    lngJntCoordscol = rngstartCoord.Column

'L4 Columns
    Set rngCADlabel = Cells.Find(what:="Full CAD Cable Label:", LookIn:=xlValues, Lookat:=xlWhole) 'Full CAD cable Label column
    lngCADlabelcol = rngCADlabel.Column

    Set rngCableType = Cells.Find(what:="Cable Type:", LookIn:=xlValues, Lookat:=xlWhole) 'Cable Type column
    lngCableTypecol = rngCableType.Column

    Set rnglength = Cells.Find(what:="Length:", LookIn:=xlValues, Lookat:=xlWhole) 'Length Column
    lngLengthcol = rnglength.Column

    Set rngDuctType = Cells.Find(what:="Duct Type:", LookIn:=xlValues, Lookat:=xlWhole) 'Duct Type Column
    lngDuctTypecol = rngDuctType.Column

    Set rngMes = Cells.Find(what:="Measure (Y/N):", LookIn:=xlValues, Lookat:=xlWhole) 'Measure COlumn
    lngMesCol = rngMes.Column

    Set rngFbrCount = Cells.Find(what:="4f or 12f:", LookIn:=xlValues, Lookat:=xlWhole) 'FibreCount Col
    lngFbrCountcol = rngFbrCount.Column

    Set rngL4PIAm = Cells.Find(what:="L4 PIAm:", LookIn:=xlValues, Lookat:=xlWhole) 'L4 PIAm Column
    lngL4PIAmcol = rngL4PIAm.Column

    Set rngL4NEXm = Cells.Find(what:="L4 Actual NEXm (NEXm+(Design Fibre Length - C-C Length)):", LookIn:=xlValues, Lookat:=xlWhole) 'L4 NEXm Column
    lngL4NEXmcol = rngL4NEXm.Column

    Set rngJntCADLabel = Cells.Find(what:="Joint Full CAD Cable Label:", LookIn:=xlValues, Lookat:=xlWhole) 'Joint CAD label Column
    lngJntCADLabelcol = rngJntCADLabel.Column

    Set rngJntCblType = Cells.Find(what:="Joint Cable Type:", LookIn:=xlValues, Lookat:=xlWhole) 'Joint Cable Type Column
    lngJntCblTypecol = rngJntCblType.Column

    Set rngJntLength = Cells.Find(what:="Joint Length:", LookIn:=xlValues, Lookat:=xlWhole) 'Joint Length Column
    lngJntLengthcol = rngJntLength.Column

    Set rngJntDuctType = Cells.Find(what:="Joint Duct Type:", LookIn:=xlValues, Lookat:=xlWhole) 'Joint Duct Type Column
    lngJntDuctTypecol = rngJntDuctType.Column

    Set rngJntMes = Cells.Find(what:="Joint Measure (Y/N):", LookIn:=xlValues, Lookat:=xlWhole) 'Joint Measure Column
    lngJntMescol = rngJntMes.Column

    Set rngJntFbrCount = Cells.Find(what:="Joint 12f or 24f:", LookIn:=xlValues, Lookat:=xlWhole) 'Joint Fibre Count Column
    lngJntFbrCountcol = rngJntFbrCount.Column

    Set rngJntPIAm = Cells.Find(what:="Joint PIAm:", LookIn:=xlValues, Lookat:=xlWhole) 'Joint PIAm Column
    lngJntPIAmcol = rngJntPIAm.Column

    Set rngJntNEXm = Cells.Find(what:="Joint Actual NEXm (NEXm+(Design Fibre Length - C-C Length)):", LookIn:=xlValues, Lookat:=xlWhole) 'Joint NEXm Column
    lngJntNEXmcol = rngJntNEXm.Column

'Selected Labels Only
    Set rngSelectedL4Lbl = Cells.Find(what:="Produce L4 Label:", LookIn:=xlValues, Lookat:=xlWhole) 'Produce specific L4 Labels
    lngSelectedL4Lbl = rngSelectedL4Lbl.Column

    Set rngSelectedJntLbl = Cells.Find(what:="Produce Joint Label:", LookIn:=xlValues, Lookat:=xlWhole) 'Produce Specific Joint Labels
    lngSelectedJntLbl = rngSelectedJntLbl.Column

''''''Find the position of the Comma in the DP XY coords''''''''''''


    L4Coords = wsDistLabels.Cells(16, lngL4CoordsCol).Value

    commaPosition = InStr(1, L4Coords, ",")


'Loop through Joint Coords and add comma into position.

    For m = 16 To 41

        If wsDistLabels.Cells(m, lngJntCoordscol).Value > 0 Then
    
        
        
            Jointcoords = wsDistLabels.Cells(m, lngJntCoordscol).Value
            If InStr(1, Jointcoords, ",") > 0 Then
            'msgbox "Joint Coords already contain comma"
                Exit For
                Else
                wsDistLabels.Cells(m, lngJntCoordscol).Value = Mid(Jointcoords, 1, commaPosition - 1) & "," & Mid(Jointcoords, commaPosition)
            End If
        
            
            Else
            Exit For
        End If
    
    Next m




'Set System Variable & System Variable Type to allow copy into CAD. Need to reset at the end
    Range("A1").Value = "ATTDIA"
    Range("A2").Value = "0"
    Range("B1").Value = "System Variable"
    Range("B2").Value = "System Variable Type"


'Set Layer to insert blocks onto
    Range("A3").Value = "'-LAYER"
    Range("A4").Value = "S"
    Range("A5").Value = "Cable-Fibre-Dist"
    

'If All Labels = False Add Colour
    
    If blAllLabelsLink = False Then
        Range("A7").Value = "'-Color"
        Range("A8").Value = "210"
        Else
        Range("A7").Value = "'-Color"
        Range("A8").Value = "ByLayer"
    End If

'Set Block Attributes & Non-dynamic attributes table in columns A & B - Loop to do this.
    
    labels = Array("Insert block", "Block Name", "Basepoint", "Scale", "Rotation", "Length", "ID", _
                   "Fibre Count", "Cable Type", "Duct Type", "Measure (Y/N)", "Alternative Supplier", _
                   "OwnedBy", "PIA(m)", "NEXfibre(m)")

    values = Array("-INSERT", "Cable-Fibre", "", "0.6", "0", "", "", "", "", "", "", "", "NexFibre", "", "")

    For d = LBound(labels) To UBound(labels)
            Dim RowOffset As Integer 'Offset for All Labels Check box
            If blAllLabelsLink = True Then 'if not needed now colour change is needed for both true and false, but left in incase it's needed in future.
            RowOffset = 9
        Else
            RowOffset = 9 'Offset to Allow for colour change of selected labels - left in incase its needed.
        End If
        
        Range("B" & d + RowOffset).Value = labels(d)
        If values(d) <> "" Then
            Range("A" & d + RowOffset).Value = values(d)
        End If
    Next d
 
    

'''''''''''''''''''''''''''''''''''''''''''''''''DP Dist Labels Loop'''''''''''''''''''''''''''''''''''''''''''''''''''

'Muted as colour change happens on both selected and all labels
'Move block attribtues table with an Offset
'Set Target ranges for all dynamic attributes.

    'Set rngIDTarget = Range("A13")
    'Set rngBasePTarget = Range("A9")
    'Set rngcabletypetarget = Range("A15")
    'Set TargetRange = Range("A7:B21")
    'Set BlockAtttype = Range("A7:B21")
    'Set rngLengthtarget = Range("A12")
    'Set rngMeasuretarget = Range("A17")
    'Set rngDuctTypeTarget = Range("A16")
    'Set rngFibreTarget = Range("A14")
    'Set rngPIAmTarget = Range("A20")
    'Set rngNEXmTarget = Range("A21")
    
'Offset if produce all labels is false

    'If blAllLabelsLink = False Then  'Muted as colour change happens on both selected and all labels
        Set rngIDTarget = Range("A13").Offset(2, 0)
        Set rngBasePTarget = Range("A9").Offset(2, 0)
        Set rngcabletypetarget = Range("A15").Offset(2, 0)
        Set TargetRange = Range("A7:B21").Offset(2, 0)
        Set BlockAtttype = Range("A7:B21").Offset(2, 0)
        Set rngLengthtarget = Range("A12").Offset(2, 0)
        Set rngMeasuretarget = Range("A17").Offset(2, 0)
        Set rngDuctTypeTarget = Range("A16").Offset(2, 0)
        Set rngFibreTarget = Range("A14").Offset(2, 0)
        Set rngPIAmTarget = Range("A20").Offset(2, 0)
        Set rngNEXmTarget = Range("A21").Offset(2, 0)

    'End If
    
    
    
    For i = 16 To 135
        
        
       If (blAllLabelsLink And wsDistLabels.Cells(i, lngCADlabelcol).Value > 0) Or (Not blAllLabelsLink And wsDistLabels.Cells(i, lngSelectedL4Lbl).Value > 0) Then
            
                'Copy all non dynamic values within the block
                TargetRange.Value = BlockAtttype.Value
    
                strBaseP = "'" & Mid(wsDistLabels.Cells(i, lngL4CoordsCol).Value, 4) 'set & copy Basepoint
                rngBasePTarget.Value = strBaseP 'paste basepoint
       
                strcabletype = wsDistLabels.Cells(i, lngCableTypecol).Value 'Set and copy cable type
                rngcabletypetarget.Value = strcabletype
       
                strLength = wsDistLabels.Cells(i, lngLengthcol).Value 'set and copy length
                rngLengthtarget.Value = strLength
       
                strID = wsDistLabels.Cells(i, lngCADlabelcol).Value 'Set & Copy Label ID
                rngIDTarget.Value = strID 'paste Label ID
       
                strDuctType = wsDistLabels.Cells(i, lngDuctTypecol).Value 'set & copy Duct Type
                rngDuctTypeTarget.Value = strDuctType
              
       
                strMeasure = wsDistLabels.Cells(i, lngMesCol).Value 'Set & copy Measure (Y/N)
                rngMeasuretarget.Value = strMeasure
      
                strFibre = wsDistLabels.Cells(i, lngFbrCountcol).Value 'Set & Copy fibre type, 4F or 12F
                rngFibreTarget.Value = strFibre & "F"
       
                intPIAm = wsDistLabels.Cells(i, lngL4PIAmcol).Value 'set & copy PIAm
                rngPIAmTarget.Value = intPIAm
      
                intNEXm = wsDistLabels.Cells(i, lngL4NEXmcol).Value 'Set & copy NEXm
                rngNEXmTarget.Value = intNEXm
       
       
                'Offset all target ranges
                Set rngIDTarget = rngIDTarget.Offset(15, 0) 'offset the target range for the Label ID
                Set TargetRange = TargetRange.Offset(15, 0) 'Offset the target range for the non-dynamic values
                Set rngBasePTarget = rngBasePTarget.Offset(15, 0) ' offset the target range for Basepoints
                Set rngcabletypetarget = rngcabletypetarget.Offset(15, 0) 'offset the target range for cable types
                Set rngLengthtarget = rngLengthtarget.Offset(15, 0) 'offset the target range for lengths
                Set rngMeasuretarget = rngMeasuretarget.Offset(15, 0) 'offset target range for Measure(Y/N)
                Set rngDuctTypeTarget = rngDuctTypeTarget.Offset(15, 0) 'offset target range for duct type
                Set rngFibreTarget = rngFibreTarget.Offset(15, 0) 'offset target range for Fibre count.
                Set rngPIAmTarget = rngPIAmTarget.Offset(15, 0) 'Offset target range for PIAm
                Set rngNEXmTarget = rngNEXmTarget.Offset(15, 0) 'Offset target range for NEXm
       
       
            ElseIf Not blAllLabelsLink Then ' If produce all labels is false then continue loop, else exit loop
Nexti:
               
            Else
            Exit For
        End If
            
    Next i



''''''''''''''''''''''''''''''''''''''''''''''''''''Loop to create Joint dist labels'''''''''''''''''''''''''''''''''''''''''''''''''

'Uncosted Circle
        Set rngCreateCircle = Cells.Find(what:="Create Circle:", LookIn:=xlValues, Lookat:=xlWhole) 'Find Circle attributed Column
        Set rngCircle = rngCreateCircle.Offset(3, 0) 'Offset from the found cell to find the circle attributes
        Set rngCircle = rngCircle.Resize(3, 2) 'Resize rngcircle to include all circle attributes.




       

    YesNo = MsgBox("Would you like to create Joint Distribution Labels?", vbYesNo)
        Select Case YesNo
            Case vbYes
        
            For o = 16 To 41

                If (blAllLabelsLink And wsDistLabels.Cells(o, lngJntCADLabelcol).Value > 0) Or (Not blAllLabelsLink And wsDistLabels.Cells(o, lngSelectedJntLbl).Value > 0) Then
                
  
  
                    TargetRange.Value = BlockAtttype.Value 'All non Dynamic values
       
    
                    strBaseP = "'" & Mid(wsDistLabels.Cells(o, lngJntCoordscol).Value, 4) 'set & copy Basepoint
                    rngBasePTarget.Value = strBaseP 'paste basepoint
    

                    strcabletype = wsDistLabels.Cells(o, lngJntCblTypecol).Value 'Set and copy cable type
                    rngcabletypetarget.Value = strcabletype


                    strLength = wsDistLabels.Cells(o, lngJntLengthcol).Value 'set and copy length
                    rngLengthtarget.Value = strLength
        
                    strID = wsDistLabels.Cells(o, lngJntCADLabelcol).Value 'Set & Copy Label ID
                    rngIDTarget.Value = strID 'paste Label ID
        
                    strDuctType = wsDistLabels.Cells(o, lngJntDuctTypecol).Value 'set & copy Duct Type
                    rngDuctTypeTarget.Value = strDuctType
        
                    strMeasure = wsDistLabels.Cells(o, lngJntMescol).Value 'Set & copy Measure (Y/N)
                    rngMeasuretarget.Value = strMeasure
        
                    strFibre = wsDistLabels.Cells(o, lngJntFbrCountcol).Value
                    rngFibreTarget.Value = strFibre & "F"
               
                    intPIAm = wsDistLabels.Cells(o, lngJntPIAmcol).Value 'set & copy PIAm
                    rngPIAmTarget.Value = intPIAm
      
                    intNEXm = wsDistLabels.Cells(o, lngJntNEXmcol).Value 'Set & copy NEXm
                    rngNEXmTarget.Value = intNEXm
               
               
               
        
                    'Offset all target ranges
                    Set rngIDTarget = rngIDTarget.Offset(15, 0) 'offset the target range for the Label ID
                    Set TargetRange = TargetRange.Offset(15, 0) 'Offset the target range for the non-dynamic values
                    Set rngBasePTarget = rngBasePTarget.Offset(15, 0) ' offset the target range for Basepoints
                    Set rngcabletypetarget = rngcabletypetarget.Offset(15, 0) 'offset the target range for cable types
                    Set rngLengthtarget = rngLengthtarget.Offset(15, 0) 'offset the target range for lengths
                    Set rngMeasuretarget = rngMeasuretarget.Offset(15, 0) 'offset target range for Measure(Y/N)
                    Set rngDuctTypeTarget = rngDuctTypeTarget.Offset(15, 0) 'offset target range for duct type
                    Set rngFibreTarget = rngFibreTarget.Offset(15, 0) 'offset target range for Fibre count.
                    Set rngPIAmTarget = rngPIAmTarget.Offset(15, 0) 'Offset target range for PIAm
                    Set rngNEXmTarget = rngNEXmTarget.Offset(15, 0) 'Offset target range for NEXm
               
               
               
                ElseIf Not blAllLabelsLink Then ' If produce all labels is false then continue loop, else exit loop
Nexto:
               
                Else
                Exit For
                End If
    
            Next o
    
    
    
'''''''''''''''''''''''Repeat Circle Loop for uncosted Joints''''''''''''''''''
        
        Set rngCircleTarget = Cells(wsDistLabels.Rows.Count, "A").End(xlUp).Offset(1, 0)
        Set rngcircletargetresize = rngCircleTarget.Resize(3, 2)
        Set rngCircleBasePTarget = Cells(wsDistLabels.Rows.Count, "A").End(xlUp).Offset(2, 0)


    For v = 16 To 135
    If (blAllLabelsLink And wsDistLabels.Cells(v, lngJntMescol).Value > 0) Or (Not blAllLabelsLink And wsDistLabels.Cells(v, lngSelectedJntLbl).Value > 0) Then
            If wsDistLabels.Cells(v, lngJntMescol).Value = "N" Then
                rngcircletargetresize.Value = rngCircle.Value 'If Measure is set to no then put a circle at the bottom
           
           
                strBaseP = "'" & Mid(wsDistLabels.Cells(v, lngJntCoordscol).Value, 4) 'set & copy Basepoint
                rngCircleBasePTarget.Value = strBaseP 'paste basepoint
            
            
                Set rngcircletargetresize = rngcircletargetresize.Offset(3, 0)
                Set rngCircleBasePTarget = rngCircleBasePTarget.Offset(3, 0)
                Else
        
            End If
        
        
        
            ElseIf Not blAllLabelsLink Then ' If produce all labels is false then continue loop, else exit loop
Nextv:
               
                Else
                Exit For
                End If
       
    Next v

    Case vbNo
        
        Set rngCircleTarget = Cells(wsDistLabels.Rows.Count, "A").End(xlUp).Offset(1, 0)
        Set rngcircletargetresize = rngCircleTarget.Resize(3, 2)
        Set rngCircleBasePTarget = Cells(wsDistLabels.Rows.Count, "A").End(xlUp).Offset(2, 0)
    
    End Select

'''''create loop to look for N in measure column and add a circle if N with that basepoint for L4s

    For u = 16 To 135
        If (blAllLabelsLink And wsDistLabels.Cells(u, lngMesCol).Value > 0) Or (Not blAllLabelsLink And wsDistLabels.Cells(u, lngSelectedL4Lbl).Value > 0) Then
            If wsDistLabels.Cells(u, lngMesCol).Value = "N" Then
            rngcircletargetresize.Value = rngCircle.Value 'If Measure is set to no then put a circle at the bottom
           
           
            strBaseP = "'" & Mid(wsDistLabels.Cells(u, lngL4CoordsCol).Value, 4) 'set & copy Basepoint
            rngCircleBasePTarget.Value = strBaseP 'paste basepoint
            
            
           Set rngcircletargetresize = rngcircletargetresize.Offset(3, 0)
           Set rngCircleBasePTarget = rngCircleBasePTarget.Offset(3, 0)
           Else
        
        End If
        
        ElseIf Not blAllLabelsLink Then ' If produce all labels is false then continue loop, else exit loop
Nextu:
               
                Else
                Exit For
                End If
       
    Next u
  
' Reset System Variable

    
    Set LastCell = Cells(wsDistLabels.Rows.Count, "A").End(xlUp).Offset(1, 0)
    
    LastCell.Value = "'-Color" 'Reset Colour to Dist Layer colour
    LastCell.Offset(1, 0).Value = "ByLayer"
    LastCell.Offset(2, 0).Value = "ATTDIA"
    LastCell.Offset(3, 0).Value = "1"
    LastCell.Offset(2, 1).Value = "System Variable"
    LastCell.Offset(3, 1).Value = "System Variable Type"


    LastRow = Cells(Rows.Count, 1).End(xlUp).Offset(2, 0).Row


'Copy Column A to clipbard for User to Paste

'Application.ScreenUpdating = True

    MsgBox "Cable Blocks have been copied onto your clipboard. Paste into the Command bar of the CAD file you're working in"





    Range("A1:A" & LastRow).Copy


    End Sub
