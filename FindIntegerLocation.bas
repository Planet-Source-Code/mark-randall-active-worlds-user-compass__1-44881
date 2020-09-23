Attribute VB_Name = "FindIntegerLocation"
Global NSV As Long
Global EWV As Long
Global NST As Long
Global EWT As Long

Public Function FindIntergerCoordinates(InputCoords As String)
    'This section splits the coords into integer values
        For i = 0 To 0
            SplitDirect = Split(Coords, " ")
            SplitCoordsNS = SplitDirect(0)
            SplitCoordsEW = SplitDirect(1)
        Next i
    
    'Now the values are found, assemble the values so that they can be displayed as intergers
        Dim NSLength As Integer, EWLength As Integer, NS As Integer, EW As Integer
        
        NSLength = Len(SplitCoordsNS) - 1
        If Right$(SplitCoordsNS, 1) = "N" Then
            NS = Val(Left$(SplitCoordsNS, NSLength))
        ElseIf Right$(SplitCoordsNS, 1) = "S" Then
            NS = Val(Left$(SplitCoordsNS, NSLength))
            NS = NS - NS - NS
        End If

        EWLength = Len(SplitCoordsEW) - 1
        If Right$(SplitCoordsEW, 1) = "E" Then
            EW = Val(Left$(SplitCoordsEW, EWLength))
        ElseIf Right$(SplitCoordsEW, 1) = "W" Then
            EW = Val(Left$(SplitCoordsEW, EWLength))
            EW = EW - EW - EW
        End If
        
        NSV = NS
        EWV = EW
End Function

Public Function FindIntergerTargetCoordinates(InputCoords As String)
    'This section splits the coords into integer values
    
        Coords = InputCoords
        For i = 0 To 0
            SplitDirect = Split(Coords, " ")
            SplitCoordsNS = SplitDirect(0)
            SplitCoordsEW = SplitDirect(1)
        Next i
    
    'Now the values are found, assemble the values so that they can be displayed as intergers
        Dim NSLength As Integer, EWLength As Integer, NS As Integer, EW As Integer
        
        NSLength = Len(SplitCoordsNS) - 1
        If Right$(SplitCoordsNS, 1) = "N" Then
            NS = Val(Left$(SplitCoordsNS, NSLength))
        ElseIf Right$(SplitCoordsNS, 1) = "S" Then
            NS = Val(Left$(SplitCoordsNS, NSLength))
            NS = NS - NS - NS
        End If

        EWLength = Len(SplitCoordsEW) - 1
        If Right$(SplitCoordsEW, 1) = "E" Then
            EW = Val(Left$(SplitCoordsEW, EWLength))
        ElseIf Right$(SplitCoordsEW, 1) = "W" Then
            EW = Val(Left$(SplitCoordsEW, EWLength))
            EW = EW - EW - EW
        End If
        
        NST = NS
        EWT = EW
End Function
