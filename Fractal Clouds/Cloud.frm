VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Fractal Clouds by Dolac"
   ClientHeight    =   8010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8040
   LinkTopic       =   "Form1"
   ScaleHeight     =   8010
   ScaleWidth      =   8040
   StartUpPosition =   3  'Windows Default
   Begin Project1.AutoSizer AutoSizer1 
      Height          =   495
      Left            =   0
      TabIndex        =   21
      Top             =   6840
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   873
   End
   Begin VB.CommandButton cmdSaveImage 
      Caption         =   "Save Image"
      Height          =   375
      Left            =   3240
      TabIndex        =   20
      Top             =   6720
      Width           =   1095
   End
   Begin VB.TextBox txtInput 
      Height          =   285
      Index           =   4
      Left            =   4080
      TabIndex        =   18
      Text            =   "1"
      Top             =   7560
      Width           =   735
   End
   Begin VB.CommandButton cmdSavePreset 
      Caption         =   "Save"
      Height          =   255
      Left            =   7200
      TabIndex        =   17
      Top             =   7560
      Width           =   615
   End
   Begin VB.ComboBox cmbPreset 
      Height          =   315
      Left            =   4800
      Style           =   2  'Dropdown List
      TabIndex        =   16
      Top             =   6720
      Width           =   1215
   End
   Begin VB.TextBox txtInput 
      Height          =   285
      Index           =   6
      Left            =   6000
      TabIndex        =   9
      Text            =   "1"
      Top             =   7560
      Width           =   735
   End
   Begin VB.TextBox txtInput 
      Height          =   285
      Index           =   5
      Left            =   5040
      TabIndex        =   8
      Text            =   "1"
      Top             =   7560
      Width           =   735
   End
   Begin VB.TextBox txtInput 
      Height          =   285
      Index           =   3
      Left            =   3120
      TabIndex        =   7
      Text            =   "1"
      Top             =   7560
      Width           =   735
   End
   Begin VB.TextBox txtInput 
      Height          =   285
      Index           =   2
      Left            =   2160
      TabIndex        =   6
      Text            =   "1"
      Top             =   7560
      Width           =   735
   End
   Begin VB.TextBox txtInput 
      Height          =   285
      Index           =   1
      Left            =   1200
      TabIndex        =   5
      Text            =   "1"
      Top             =   7560
      Width           =   735
   End
   Begin VB.TextBox txtInput 
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   4
      Text            =   "1"
      Top             =   7560
      Width           =   735
   End
   Begin VB.CommandButton cmdAbort 
      Caption         =   "Stop"
      Height          =   375
      Left            =   7080
      TabIndex        =   3
      Top             =   6720
      Width           =   735
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "Go"
      Height          =   375
      Left            =   6240
      TabIndex        =   2
      Top             =   6720
      Width           =   735
   End
   Begin VB.PictureBox picDisplay 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H000000FF&
      ForeColor       =   &H00000000&
      Height          =   6600
      Left            =   0
      ScaleHeight     =   6540
      ScaleWidth      =   7980
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   8040
   End
   Begin VB.Label lblParam 
      Alignment       =   2  'Center
      Caption         =   "Span"
      Height          =   255
      Index           =   4
      Left            =   4080
      TabIndex        =   19
      Top             =   7320
      Width           =   735
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   7800
      Y1              =   7200
      Y2              =   7200
   End
   Begin VB.Label lblParam 
      Alignment       =   2  'Center
      Caption         =   "Cycles"
      Height          =   255
      Index           =   6
      Left            =   6000
      TabIndex        =   15
      Top             =   7320
      Width           =   735
   End
   Begin VB.Label lblParam 
      Alignment       =   2  'Center
      Caption         =   "Scale"
      Height          =   255
      Index           =   5
      Left            =   5040
      TabIndex        =   14
      Top             =   7320
      Width           =   735
   End
   Begin VB.Label lblParam 
      Alignment       =   2  'Center
      Caption         =   "Y"
      Height          =   255
      Index           =   3
      Left            =   3120
      TabIndex        =   13
      Top             =   7320
      Width           =   735
   End
   Begin VB.Label lblParam 
      Alignment       =   2  'Center
      Caption         =   "X"
      Height          =   255
      Index           =   2
      Left            =   2160
      TabIndex        =   12
      Top             =   7320
      Width           =   735
   End
   Begin VB.Label lblParam 
      Alignment       =   2  'Center
      Caption         =   "B"
      Height          =   255
      Index           =   1
      Left            =   1200
      TabIndex        =   11
      Top             =   7320
      Width           =   735
   End
   Begin VB.Label lblParam 
      Alignment       =   2  'Center
      Caption         =   "A"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   10
      Top             =   7320
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Ready"
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   6720
      Width           =   2655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Fractal - DUST CLOUD by Dolac

'picture is derived from single user defined starting point P0(A,B,X,Y).
'by numerous repeating of some transformations we get some cool pics.
'Because of the fact that we can follow the trail start point is leaving
'behined (P0, P1, P2, .... Px - dots on picture) this pics are also called
'orbitals of dynamic system

'User interface by Paul Bahlawan March 2011

Option Explicit

Dim Xs As Double    'picturebox start point
Dim Ys As Double

Dim N As Long     'number of cycles
Dim K As Long     'number of iterative steps in one cycle

Dim A As Double     'user adjustable parametar
Dim B As Double     'user adjustable parametar
Dim X As Double     'user adjustable parametar
Dim Y As Double     'user adjustable parametar

Dim W As Double     'transformation variable - some function
Dim Z As Double     'hold previous step of X
Dim Span As Double  'New paramater (I didn't know what to call it)

Dim ScalePic As Long  'scale graphics in picturebox - rude zoom
Dim bAbort As Boolean 'Abort flag incase we want to get out of long iterations
Dim pList() As String 'List of preset values


Private Sub cmdGo_Click()
    'Reset
    Xs = picDisplay.ScaleWidth / 2: Ys = picDisplay.ScaleHeight / 2
    N = 0: W = 0: Z = 0
    picDisplay.Cls
    bAbort = False
    
    cmdGo.Enabled = False
    
    'Get user input
    A = CDbl(txtInput(0))
    B = CDbl(txtInput(1))
    X = CDbl(txtInput(2))
    Y = CDbl(txtInput(3))
    Span = CDbl(txtInput(4))
    ScalePic = CLng(txtInput(5))
    K = CLng(txtInput(6))
    
    'Main loop
    Do While N < K
        N = N + 1
        Label1.Caption = "Wait - Iteration cycle " & N & " / " & K
        Draw
        DoEvents
        If bAbort Then Exit Do
    Loop
    
    cmdGo.Enabled = True
    Label1.Caption = "Iteration ended"
End Sub

Private Sub cmdAbort_Click()
    bAbort = True
End Sub

Private Function Draw()
    Dim i As Long
    
    For i = 1 To 2000       'used to speed up making of drawing, and also make computer accessible
        Z = X
        X = Y + W
        NextOrbit
        Y = W - Z
        'draw dot on display
        picDisplay.PSet (Xs + X * ScalePic, Ys + Y * ScalePic), vbBlack
    Next i
End Function

Private Function NextOrbit()
    If X > Span Then
        W = A * X + B * (X - 1)     'can u even imagine how many other/different functions can we use here...
    End If
    
    If X < -Span Then
        W = A * X + B * (X + 1)     'maybe some combination involving sin or cos...
    End If
    
    If X < Span And X > -Span Then
        W = A * X                   'or ln/log or exp... just to give you a hint..
    End If
End Function

Private Sub cmdSaveImage_Click()
    Dim PicFile As String
    
    If cmdGo.Enabled = False Then Exit Sub
    
    'save bmp of picture
    PicFile = App.Path & "\Pics\A_" & txtInput(0) & " B_" & txtInput(1) & " X_" & txtInput(2) & " Y_" & txtInput(3) & " Span_" & txtInput(4) & " Scale_" & txtInput(5) & " Cycles_" & txtInput(6) & ".bmp"
    SavePicture picDisplay.Image, PicFile
    MsgBox "Picture saved.." & vbCrLf & PicFile
End Sub

Private Sub cmdSavePreset_Click()
Dim i As Long
Dim P As String
    
    'Put all values into one string
    For i = 0 To 6
        P = P & txtInput(i)
        If i <> 6 Then P = P & ","
    Next i

    'Add to preset file
    Open App.Path & "\Cloud.preset" For Append As #1
    Print #1, P
    Close #1
    
    MsgBox "New preset saved"
    
    LoadPresets
    cmbPreset.ListIndex = cmbPreset.ListCount - 1 'select the last item in the combo box
End Sub

Private Sub Form_Activate()
    AutoSizer1.GetProps ' Initalize AutoSizer control
    LoadPresets
    cmbPreset.ListIndex = 0 'select the first item in the combo box
End Sub

Private Sub Form_Resize()
    'Stretch the PictureBox to fit the Form
    If Form1.Width > 125 Then
        picDisplay.Width = Form1.Width - 125
    End If
    If Form1.Height > 1875 Then
        picDisplay.Height = Form1.Height - 1875
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    bAbort = True
    DoEvents
End Sub

Private Sub cmbPreset_Click()
Dim S() As String
Dim i As Long
    
    'Transfer the Preset values to the user input boxes
    S = Split(pList(cmbPreset.ListIndex), ",")
    For i = 0 To 6
        txtInput(i) = S(i)
    Next i
End Sub

Private Sub LoadPresets()
Dim i As Long
    'Reset
    ReDim pList(0)
    cmbPreset.Clear
    
    'Get the Presets from disk
    Open App.Path & "\Cloud.preset" For Input As #1
    Do While Not EOF(1)
        ReDim Preserve pList(i)
        Line Input #1, pList(i)
        i = i + 1
        cmbPreset.AddItem "Preset " & i
    Loop
    Close #1
End Sub

