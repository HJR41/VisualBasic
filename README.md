# Overview
The Visual Basics projects are two of my resent projects undertaken to remove repetative tasks & reduce time for completion. 

## 1. Automatic Distribution Labels
This project is one that I am most proud of. Taking user input data from a complex Spreadsheet type database developed by a collegue, where the data is transformed and manipulated for a vast array of use cases. 

One such usecase is within the AutoDistLabels script.

View the code here: [AutoDistLabels](https://github.com/HJR41/VisualBasic/blob/main/AutoDistLabels.vba)

The script takes multiple intricate data points, and places them into an AutoCAD compatible text string within Excel for the user to simply copy the string into the AutoCAD textbar. Consequently, CAD instantly generates between 30-100 highly data dense poly-line labels.

The script reduced the average task completion time from approximately 4-8+ hours to approximately 1-2 hours saving countless hours of valuable resource & delivering pin-point accuracy.

```vb
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
```
