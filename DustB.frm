VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Click on picture to save image"
   ClientHeight    =   7125
   ClientLeft      =   3000
   ClientTop       =   3600
   ClientWidth     =   11970
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7125
   ScaleWidth      =   11970
   Begin VB.Frame Frame1 
      Caption         =   "Root Node"
      Height          =   1695
      Left            =   10080
      TabIndex        =   7
      Top             =   720
      Width           =   1815
      Begin VB.OptionButton Option2 
         Caption         =   "X2 (1,0)"
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   840
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         Caption         =   "X1 (0,0)"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Value           =   -1  'True
         Width           =   1575
      End
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   4200
      TabIndex        =   5
      Top             =   6600
      Width           =   800
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   3360
      TabIndex        =   4
      Top             =   6600
      Width           =   800
   End
   Begin VB.CommandButton cmdChoose 
      Caption         =   "&Choose Pattern"
      Height          =   500
      Left            =   10080
      TabIndex        =   3
      Top             =   120
      Width           =   1300
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   500
      Left            =   10080
      TabIndex        =   2
      Top             =   6600
      Width           =   1300
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   6375
      Left            =   120
      ScaleHeight     =   6315
      ScaleWidth      =   9795
      TabIndex        =   0
      Top             =   120
      Width           =   9855
   End
   Begin VB.Label Label2 
      Caption         =   "Press BACKSPACE  to Enable Bactracking"
      Height          =   495
      Left            =   5040
      TabIndex        =   6
      Top             =   6600
      Width           =   2055
   End
   Begin VB.Label Label1 
      Height          =   405
      Left            =   240
      TabIndex        =   1
      Top             =   6645
      Width           =   3015
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Pattern As String   'pattern options [1 to 8]
Dim KeyAscii As Integer 'keytrap

'transformation parameters - see tutorial
Dim A As Single:        Dim B As Single:        Dim C As Single:            Dim D As Single

'order of iteration     View magnification on fractal
Dim Order As Integer:   Dim PicScale As Single

'Level of tree = order +1
Dim Level As Integer

'Group Counter
Dim Group As Double:        Dim N As Double:        Dim K As Double

Dim J As Integer

Dim X1() As Double:     Dim X2() As Double:     Dim X As Double:           Dim Xs As Double
Dim Y1() As Double:     Dim Y2() As Double:     Dim y As Double:           Dim Ys As Double

Private Sub Form_Load()
        
    Pattern = InputBox("Selet Pattern 1 - 8", 100, 1, 300)
        If StrPtr(Pattern) = 0 Then     'MsgBox "Application Ending..":
            End
        End If
        
    Form1.Show
    
    Select Case Pattern
        Case 1  'Levy C Curve
            A = 0.5:    B = 0.5:    C = 0.5:    D = -0.5:   PicScale = 4700:   Order = 15
            Xs = Picture1.Width / 20 * 5    'x-axis fractal screen centering
            Ys = Picture1.Height / 200 * 45 'y-axis fractal centering
        Case 2  'Rotating Levy
            A = 1:      B = 1:      C = 1:      D = -1:     PicScale = 6:   Order = 16
            Xs = Picture1.Width / 2
            Ys = Picture1.Height / 200 * 101
        Case 3  'Leaves - IFS (Iterated Function System)
            A = 0.6:    B = 0.6:    C = 0.53:   D = 0:      PicScale = 5200:   Order = 15
            Xs = Picture1.Width / 20 * 7
            Ys = Picture1.Height / 20 * 7
        Case 4  'Right-angle branching
            A = 0:      B = 0.7:    C = 0.7:    D = 0:      PicScale = 5800:   Order = 15
            Xs = Picture1.Width / 20 * 7
            Ys = Picture1.Height / 200 * 65
        Case 5  'Dragon
            A = 0.5:    B = 0.55:   C = -0.4:   D = 0.27:    PicScale = 3500:   Order = 15
            Xs = Picture1.Width / 20 * 5
            Ys = Picture1.Height / 20 * 9
        Case 6  'Dragon curve 2
            A = 0.5:    B = 0.5:    C = 0.5:    D = 0.5:   PicScale = 4400:   Order = 15
            Xs = Picture1.Width / 200 * 45
            Ys = Picture1.Height / 2
        Case 7  'Dragon curve - aka Jurassic Park Fractal
            A = 0.38:   B = 0.38:   C = 0.38:   D = 0.68:   PicScale = 4300:   Order = 16
            Xs = Picture1.Width / 5 * 1
            Ys = Picture1.Height / 500 * 275
        Case 8  'Squares
            A = 0:      B = 1:   C = 0.45:   D = 0:      PicScale = 2800:   Order = 15
            Xs = Picture1.Width / 50 * 23
            Ys = Picture1.Height / 500 * 245
    End Select
    
    K = 2 ^ (Order + 1) - 1     'total number of nodes in a tree - points of fractal
    '2 ^ (Order + 1)        = Node names = 1, 2, 3, 4, 5, 6....
    '2 ^ (Order + 1) - 1    = Node names = 0, 1, 2, 3, 4, 5....
    
    'number of levels - from root(zero) to Order.
    'if Order is zero only root node exist & only 1 point should be drawn, there is only root level - Level 1
    'if Order is 1, 3 points - root + 2 leaves nodes, there are Level 1 & 2
    'if Order is 2, 7 points - root + 2 mid + 4 leaves (each root's child have 2), Level 1,2&3     etc..
    
    'assign size to tree location array
    ReDim X1(Order): ReDim Y1(Order): ReDim X2(Order): ReDim Y2(Order)
    ReDim X0(Order): ReDim Y0(Order)
    
    'In-depth Search - from root to order
    Level = 1:              Call CheckOption:       Call Transformation
    Label1.Caption = "Wait to backtrack"
End Sub

Private Sub CheckOption()
    'XXXXXXX - The start point is always either (0,0) or (1,0) - will be explained in future lesson
    'Note: Both points (X1,Y1) & (X2,Y2) are part of fractal
    
    If Option1.Value = True Then
        '(1,0) - center of R-transform
        X2(0) = 1:          Y2(0) = 0:      X1(0) = X2(0):          Y1(0) = Y2(0)
        Picture1.Circle (X1(0) * PicScale + Xs, Y1(0) * PicScale + Ys), 60, vbCyan
        Picture1.Circle (X2(0) * PicScale + Xs, Y2(0) * PicScale + Ys), 60, vbCyan
    Else
        '(0,0) - center of L-transform
        X1(0) = 0:          Y1(0) = 0:      X2(0) = X1(0):          Y2(0) = X1(0)
        Picture1.Circle (X1(0) * PicScale + Xs, Y1(0) * PicScale + Ys), 60, vbCyan
        Picture1.Circle (X2(0) * PicScale + Xs, Y2(0) * PicScale + Ys), 60, vbCyan
    End If
End Sub

Private Function Transformation()
    Dim xx As Single:               Dim yy As Single
    
    J = Level
    
    Do Until J = Order + 1
        
        If Option1.Value = True Then
            'Root location - both tree branches intersection - if L transform is choosen
            X1(Level - 1) = X2(Level - 1)       'X(0) is at level 1 which is Order 0
            Y1(Level - 1) = Y2(Level - 1)       'both branches start from the same point
            'for L-transform
            X = X1(J - 1)
            y = Y1(J - 1)
        Else
            'Root location - both tree branches intersection - if R transform is choosen
            X2(Level - 1) = X1(Level - 1)       'X(0) is at level 1 which is Order 0
            Y2(Level - 1) = Y1(Level - 1)       'both branches start from the same point
            'for R-transform
            X = X2(J - 1)
            y = Y2(J - 1)
        End If
                        
        'calculate child nodes - recursion step
            X1(J) = A * X - B * y         'transformation of blue dots
            Y1(J) = B * X + A * y

            X2(J) = C * X - D * y + 1 - C  'transformation of red dots
            Y2(J) = D * X + C * y - D
            
            Picture1.PSet (X1(J) * PicScale + Xs, Picture1.Height - (Y1(J) * PicScale + Ys)), vbBlue
            Picture1.PSet (X2(J) * PicScale + Xs, Picture1.Height - (Y2(J) * PicScale + Ys)), vbRed
            
        J = J + 1
    Loop
    
End Function

Private Function Backtrackin()
    'remove start circle points but also m"0" group
    'Picture1.Cls
    Label1.Caption = "Iteration in progress.."
    
    'in each step 2 points are calculated - m"0" is drawn again
    For Group = 0 To (2 ^ (Order - 1) - 1)    'move along each tree group - Group is group name, not size
        Level = Order + 1                   'number of levels from zero to order. That equals order+1 level in total
        N = Group
line1:
                                        'Note on vb6 arithmetic operations
        If N Mod 2 = 0 And N > 0 Then   ' 0 Mod 2=0,   1 Mod 2=1  ...   9 mod 2 = 1
            N = N / 2                   ' "\" Operator returns the integer quotient of a division
            Level = Level - 1           ' 9\2 = 4 vs. 9/2 = 4.5         but 0\2=0
            DoEvents
            'Call Transformation
            GoTo line1
        End If
        
        Call Transformation
    Next Group
    
    'Note on groups: in example - for order = 4, we get 8 groups [0-7](see tutorial)
    'apart for group 0 where level is 1, level always equal fractal order minus number of 2 in group level
    'm=0  level is 1
    'm=1  no 2 in composition - level=4-0=4
    'm=2  1 x 2 in composition - level=4-1=3
    'm=3  no 2 in composition - level=4-0=4
    'm=4  2 x 2 in composition - level=4-2=2
    'm=5  no 2 in composition - level=4-0=4
    'm=6  1 x 2 in composition (2 x 3)- level=4-1=3
    'm=7  no 2 in composition - level=4-0=4
    Label1.Caption = "Done"
End Function

Private Sub Option1_Click()
    Picture1.Cls    'clear picture
    Level = 1       'repeat in-deapth search
    CheckOption
    Call Transformation
    'NOTE:  Press BackSpace to BackTrack
End Sub
Private Sub Option2_Click()
    Picture1.Cls
    Level = 1
    Call CheckOption
    Call Transformation
End Sub



'**************************************************************************************
'****** NO connection to fractal algorithm - just added to simplify usage and checking..
'**************************************************************************************

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
    'another way - also Ok.
    'If KeyCode = 27 Then Unload Me
    
    If KeyCode = vbKeyBack Then
        'if you backtrackin is disabled, you will see single root to leave transformation for group M=0
        Call CheckOption
        Call Backtrackin
    End If
End Sub
Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
    'used to check results - from root to leaves. XXXXXXX line
    'depth search - disable bactrackin call in next line
    Text1.Text = Format((X - Xs) / PicScale, "0.00")
    Text2.Text = Format((y - Ys) / PicScale, "0.00")
End Sub
Private Sub LocatePoint(X As Single, y As Single)
    Picture1.CurrentX = X
    Picture1.CurrentY = y
End Sub
Private Sub Picture1_Click()
    Dim PicFile As String
    'save bmp of picture
    PicFile = App.Path & "\Pics\Pattern-" & Pattern & " Level" & Level & ".bmp"
        SavePicture Picture1.Image, PicFile
    MsgBox "Picture saved.."
End Sub
Private Sub Form_Click()
    Call cmdChoose_Click
End Sub
Private Sub cmdChoose_Click()
    Picture1.Cls:               Call Form_Load
End Sub
Private Sub cmdExit_Click()
    Unload Me
End Sub


