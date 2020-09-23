VERSION 5.00
Begin VB.Form FrmMain 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Tracker"
   ClientHeight    =   2430
   ClientLeft      =   10470
   ClientTop       =   9615
   ClientWidth     =   4710
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2430
   ScaleWidth      =   4710
   Begin VB.TextBox txtCurrent 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2057
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Left            =   4200
      TabIndex        =   12
      Top             =   2040
      Width           =   375
   End
   Begin VB.CheckBox ChkIsScanning 
      BackColor       =   &H00FF8080&
      Caption         =   "Select to initiate"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2760
      TabIndex        =   8
      Top             =   720
      Width           =   1815
   End
   Begin VB.TextBox txtTargetCoords 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   2760
      TabIndex        =   7
      Top             =   360
      Width           =   1815
   End
   Begin VB.TextBox txtY 
      Alignment       =   1  'Right Justify
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2057
         SubFormatType   =   1
      EndProperty
      Enabled         =   0   'False
      Height          =   285
      Left            =   3840
      TabIndex        =   1
      Top             =   1560
      Width           =   735
   End
   Begin VB.TextBox txtX 
      Alignment       =   1  'Right Justify
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2057
         SubFormatType   =   1
      EndProperty
      Enabled         =   0   'False
      Height          =   285
      Left            =   3840
      TabIndex        =   0
      Top             =   1200
      Width           =   735
   End
   Begin VB.Timer Timer1 
      Interval        =   20
      Left            =   4800
      Top             =   720
   End
   Begin VB.Line Line4 
      X1              =   2760
      X2              =   4560
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Label lblWaypoint 
      BackStyle       =   0  'Transparent
      Caption         =   "Waypoint"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2760
      TabIndex        =   11
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label lblYRange 
      BackStyle       =   0  'Transparent
      Caption         =   "Y Range"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2760
      TabIndex        =   10
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label lblXRange 
      BackStyle       =   0  'Transparent
      Caption         =   "X Range"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2760
      TabIndex        =   9
      Top             =   1200
      Width           =   855
   End
   Begin VB.Line Line3 
      X1              =   4680
      X2              =   4680
      Y1              =   120
      Y2              =   2280
   End
   Begin VB.Line Line2 
      X1              =   2760
      X2              =   4560
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Label lblTargetCoords 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Target Coordinates"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2760
      TabIndex        =   6
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label lblE 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      Caption         =   "E"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2280
      TabIndex        =   5
      Top             =   1080
      Width           =   255
   End
   Begin VB.Label lblW 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "W"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   255
   End
   Begin VB.Label lblSouth 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "S"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   3
      Top             =   2040
      Width           =   255
   End
   Begin VB.Label lblNorth 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "N"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   2
      Top             =   120
      Width           =   255
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   3
      Height          =   975
      Left            =   840
      Shape           =   3  'Circle
      Top             =   720
      Width           =   975
   End
   Begin VB.Line NWLine 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      X1              =   960
      X2              =   1320
      Y1              =   840
      Y2              =   1200
   End
   Begin VB.Line SWLine 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      X1              =   960
      X2              =   1320
      Y1              =   1560
      Y2              =   1200
   End
   Begin VB.Line SELine 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      X1              =   1680
      X2              =   1320
      Y1              =   1560
      Y2              =   1200
   End
   Begin VB.Line NELine 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      X1              =   1680
      X2              =   1320
      Y1              =   840
      Y2              =   1200
   End
   Begin VB.Line WLine 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      X1              =   840
      X2              =   1320
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Line ELine 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      X1              =   1800
      X2              =   1320
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Line SLine 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      X1              =   1320
      X2              =   1320
      Y1              =   1680
      Y2              =   1200
   End
   Begin VB.Line Nline 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      X1              =   1320
      X2              =   1320
      Y1              =   1200
      Y2              =   720
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FF8080&
      BackStyle       =   1  'Opaque
      Height          =   2415
      Left            =   2640
      Top             =   0
      Width           =   2055
   End
   Begin VB.Line Line1 
      X1              =   1320
      X2              =   1335
      Y1              =   1200
      Y2              =   1215
   End
   Begin VB.Line Direction 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      X1              =   1320
      X2              =   1320
      Y1              =   1200
      Y2              =   0
   End
   Begin VB.Shape Radar 
      BackColor       =   &H00FF8080&
      BackStyle       =   1  'Opaque
      Height          =   2415
      Left            =   0
      Top             =   0
      Width           =   2655
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'declarations (needed)
Private ontop As New clsOnTop

Private Sub CmdOnTop_Click()
    ontop.MakeTopMost hWnd
End Sub

Private Sub Command1_Click()
Call Timer1_Timer
End Sub

'form load procedures
Private Sub Form_Load()
    Set ontop = New clsOnTop
    'make normal, just in case
    ontop.MakeNormal hWnd
    ontop.MakeTopMost hWnd
End Sub

'make top most button
Private Sub OntopMode_Click()
    ontop.MakeTopMost hWnd
End Sub

'make normal button
Private Sub NormalMode()
    ontop.MakeNormal hWnd
End Sub

Private Function FindWindowLike(ByVal hWndStart As Long, WindowText As String, Classname As String) As Long

   Dim hWnd As Long
   Dim sWindowText As String
   Dim sClassname As String
   Dim r As Long
   Static level As Integer
   If level = 0 Then
      If hWndStart = 0 Then hWndStart = GetDesktopWindow()
   End If
    level = level + 1

   hWnd = GetWindow(hWndStart, GW_CHILD)

   Do Until hWnd = 0
      
     'Search children by recursion
      Call FindWindowLike(hWnd, WindowText, Classname)
      
     'Get the window text and class name
      sWindowText = Space$(255)
      r = GetWindowText(hWnd, sWindowText, 255)
      sWindowText = Left(sWindowText, r)
        
      sClassname = Space$(255)
      r = GetClassName(hWnd, sClassname, 255)
      sClassname = Left(sClassname, r)
              
     'Check if window found matches the search parameters
      If (sWindowText Like WindowText) And _
         (sClassname Like Classname) Then
        
         CurrentPosition = sWindowText
         DeterminRange
         FindWindowLike = hWnd
                     
        'uncommenting the next line causes the routine to
        'only return the first matching window.
        'Exit Do
           
      End If
    
     'Get next child window
      hWnd = GetWindow(hWnd, GW_HWNDNEXT)
  
   Loop
 
  'Reduce the recursion counter
   level = level - 1

End Function


Private Sub Timer1_Timer()
   Dim TitleToFind As String, ClassToFind As String, txtTitle As String, txtClass As Variant
    
  'Set the FindWindowLike text values from
  'the strings entered into the textboxes
   TitleToFind = ("Active Worlds - ") & "*"
   ClassToFind = ("Alphaworld")
  
    If Len(txtTargetCoords.Text) < 4 Or ChkIsScanning.Value = vbUnchecked Then
        Exit Sub
    End If
    
    Dim lbcount As String
    Call FindWindowLike(0, TitleToFind, ClassToFind)
End Sub

Public Sub DeterminRange()

    'This section finds if the browser is facing North, South, East or West
    Dim FaceNSEC As String, i As Integer
    For i = 0 To 1
        SplitFacing = Split(CurrentPosition, " facing ")
        Facing = SplitFacing(i)
    Next i
    
    If Facing = "N" Then
        Nline.Visible = True
        NELine.Visible = False
        ELine.Visible = False
        SELine.Visible = False
        SLine.Visible = False
        SWLine.Visible = False
        WLine.Visible = False
        NWLine.Visible = False
    ElseIf Facing = "NE" Then
        Nline.Visible = False
        NELine.Visible = True
        ELine.Visible = False
        SELine.Visible = False
        SLine.Visible = False
        SWLine.Visible = False
        WLine.Visible = False
        NWLine.Visible = False
    ElseIf Facing = "E" Then
        Nline.Visible = False
        NELine.Visible = False
        ELine.Visible = True
        SELine.Visible = False
        SLine.Visible = False
        SWLine.Visible = False
        WLine.Visible = False
        NWLine.Visible = False
    ElseIf Facing = "SE" Then
        Nline.Visible = False
        NELine.Visible = False
        ELine.Visible = False
        SELine.Visible = True
        SLine.Visible = False
        SWLine.Visible = False
        WLine.Visible = False
        NWLine.Visible = False
    ElseIf Facing = "S" Then
        Nline.Visible = False
        NELine.Visible = False
        ELine.Visible = False
        SELine.Visible = False
        SLine.Visible = True
        SWLine.Visible = False
        WLine.Visible = False
        NWLine.Visible = False
    ElseIf Facing = "SW" Then
        Nline.Visible = False
        NELine.Visible = False
        ELine.Visible = False
        SELine.Visible = False
        SLine.Visible = False
        SWLine.Visible = True
        WLine.Visible = False
        NWLine.Visible = False
    ElseIf Facing = "W" Then
        Nline.Visible = False
        NELine.Visible = False
        ELine.Visible = False
        SELine.Visible = False
        SLine.Visible = False
        SWLine.Visible = False
        WLine.Visible = True
        NWLine.Visible = False
    ElseIf Facing = "NW" Then
        Nline.Visible = False
        NELine.Visible = False
        ELine.Visible = False
        SELine.Visible = False
        SLine.Visible = False
        SWLine.Visible = False
        WLine.Visible = False
        NWLine.Visible = True
    End If
    
    'This section is to find the coordinates of the browser
    For i = 0 To 1
        SplitCoords = Split(CurrentPosition, " at ")
        PartCoords = SplitCoords(1)
    Next i
        
    'This section is a contuniation of the last, it removes the facing points
    For i = 0 To 0
        SplitWithout = Split(PartCoords, " facing ")
        Coords = SplitWithout(0)
    Next i
    
    FindIntergerCoordinates (Coords)
    'Now to find the values for the target
    
    FindIntergerTargetCoordinates (UCase(txtTargetCoords.Text))
    
    'Now for the complicated part - Finding the distances
    
    NSD = NST - NSV
    EWD = EWT - EWV
    
    EWD = 0 - EWD
    txtY = NSD
    txtX = EWD
    
    'Now it gets rediclously Complex
    
    'Find Radar Center
    RadarX = Line1.X1
    RadarY = Line1.Y1
   
    ChangeDistanceLR = Val(txtX.Text) * 100
    ChangeDistanceUD = Val(txtY.Text) * 100
    Direction.X1 = (Line1.X1 - 10)
    Direction.Y1 = (Line1.X1 - 110)
    
    Direction.X2 = (Line1.X1 - ChangeDistanceLR)
    Direction.Y2 = (Line1.X1 - ChangeDistanceUD) - 100
End Sub
