VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Snake"
   ClientHeight    =   6870
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9000
   DrawWidth       =   5
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   458
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   600
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   1455
      Left            =   3060
      TabIndex        =   11
      Top             =   0
      Width           =   5835
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   $"frmMain.frx":0000
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1170
         Left            =   2880
         TabIndex        =   16
         Top             =   180
         Width           =   2850
         WordWrap        =   -1  'True
      End
      Begin VB.Line Line1 
         X1              =   2640
         X2              =   2640
         Y1              =   390
         Y2              =   1200
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000005&
         X1              =   2655
         X2              =   2655
         Y1              =   390
         Y2              =   1200
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Web:      www.coderpost.com"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   300
         MousePointer    =   10  'Up Arrow
         TabIndex        =   15
         Top             =   900
         Width           =   2145
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Email:     eric@coderpost.com"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   300
         MousePointer    =   10  'Up Arrow
         TabIndex        =   14
         Top             =   660
         Width           =   2115
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Author:  Eric J. Griffin"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   300
         TabIndex        =   13
         Top             =   420
         Width           =   1560
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   " About "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   90
         TabIndex        =   12
         Top             =   0
         Width           =   525
      End
   End
   Begin VB.ListBox lstNextSeg 
      Height          =   450
      ItemData        =   "frmMain.frx":00ED
      Left            =   1110
      List            =   "frmMain.frx":00EF
      TabIndex        =   10
      Top             =   6210
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Timer tmrMove 
      Enabled         =   0   'False
      Interval        =   45
      Left            =   150
      Top             =   6210
   End
   Begin VB.Timer tmrTime 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   599
      Top             =   6210
   End
   Begin VB.Frame Frame1 
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   2865
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Food Eaten:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   435
         TabIndex        =   9
         Top             =   540
         Width           =   885
      End
      Begin VB.Label lblEatenX 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2295
         TabIndex        =   8
         Top             =   540
         Width           =   90
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Time:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   930
         TabIndex        =   7
         Top             =   315
         Width           =   390
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1905
         TabIndex        =   6
         Top             =   300
         Width           =   60
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2145
         TabIndex        =   5
         Top             =   300
         Width           =   60
      End
      Begin VB.Label lblm 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1965
         TabIndex        =   4
         Top             =   315
         Width           =   195
      End
      Begin VB.Label lblh 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1725
         TabIndex        =   3
         Top             =   315
         Width           =   195
      End
      Begin VB.Label lbls 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2205
         TabIndex        =   2
         Top             =   315
         Width           =   195
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   " Statistics "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   90
         TabIndex        =   1
         Top             =   0
         Width           =   735
      End
   End
   Begin VB.Shape shpFood 
      FillColor       =   &H0000FF00&
      FillStyle       =   0  'Solid
      Height          =   150
      Left            =   2430
      Shape           =   1  'Square
      Top             =   6210
      Width           =   150
   End
   Begin VB.Shape Seg 
      FillStyle       =   0  'Solid
      Height          =   150
      Index           =   0
      Left            =   0
      Shape           =   1  'Square
      Top             =   1500
      Width           =   150
   End
   Begin VB.Shape shpArena 
      BackColor       =   &H00E0E0E0&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   4500
      Left            =   0
      Top             =   1500
      Width           =   9000
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'#########################################################################
'##
'##  Snake
'##  by Eric J. Griffin
'##     eric@coderpost.com
'##     www.coderpost.com
'##
'##  Here is the source for the very popular, very classic game 'Snake.'
'##  In an IRC conversation(yes I'm an IRC addict), somebody brought up
'##  the subject of the game, and I thought, "I'll make my own," like I
'##  don't have thousands of other better things I could be doing with my
'##  time.
'##
'##  Now let me tell you this snake game isn't complete. It was meant to be
'##  purely an example and it still needs a little work that I'll leave up to
'##  you to finish on your own... I have many other things I have to get to.
'##
'##  Here is a list of stuff I never bothered to finish:
'##  1. The snake is free to slither right thru itself.
'##  2. After you hit the wall and continue, it's possible to move
'##     outside the arena.
'##  3. After eating a snack, food can reappear behind the snakes body.
'##  4. There is a lack of speed adjustment and levels/obstacles.
'##  5. I'm sure you'll find more if you're a hardcore snake fan. sssssss!
'##
'#########################################################################


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim Dir%
    
    '## Check if game is in progress, and if not - start it
    If tmrMove.Enabled = False Then
        Call SetFood
        tmrMove.Enabled = True
        tmrTime.Enabled = True
    End If
    
    '## Find which directional arrow was pressed and simplify it
    '## Left=1;  Up=2;  Right=3;  Down=4;
    Dir% = KeyCode - 36
    
    '## If user presses the direction we are already going, exit
    If Dir% = Direction Then Exit Sub
    
    '## If a snake is heading due left, and changes direction immediately to the
    '## right, consequently it would be slithering through itself. Therefor,
    '## we make sure the key pressed isn't opposite the current direction before
    '## we set the new direction.
    Select Case KeyCode
        Case vbKeyLeft
            If Direction <> 3 Then Direction = Dir%
        Case vbKeyUp
            If Direction <> 4 Then Direction = Dir%
        Case vbKeyRight
            If Direction <> 1 Then Direction = Dir%
        Case vbKeyDown
            If Direction <> 2 Then Direction = Dir%
    End Select
End Sub

Private Sub Form_Load()

    '## For the sake of you being able to visually see what other objects
    '## are being used on the form, I've raised the height a little bit. We
    '## will fix that here, and add our first segment index to our listbox.
    frmMain.Height = 6375
    lstNextSeg.AddItem "0"
End Sub


Private Sub tmrMove_Timer()
    Dim sCount%
    Dim NextSeg, LastSeg, nNextSeg%, nLastSeg%
    
    '## Now I've seen other snake sources that use For/Next loops to move
    '## every single segment of the snake. This is really unnecessary and slows
    '## down everything when the snake's length starts to get long. A better
    '## method is to move only the last segment of the snake to the front.
    
    '## See how many segments we have (one segment is present at startup)
    sCount% = lstNextSeg.ListCount
    
    '## Find the first and last segments from the listbox and store
    '## their values into an integer variable
    NextSeg = lstNextSeg.List(0)
    LastSeg = lstNextSeg.List(sCount% - 1)
    nNextSeg% = CInt(NextSeg)
    nLastSeg% = CInt(LastSeg)
    
    '## Move the top item to the bottom
    lstNextSeg.RemoveItem 0
    lstNextSeg.AddItem nNextSeg
    
    '## Note the Form_Keydown sub and you'll see the direction var
    '## changes when you hit one of the arror buttons. Based on this,
    '## move the last segment in front of the first segment. If you go
    '## beyond the wall, stop the game.
    Select Case Direction
        Case 1
            Seg(nNextSeg).Left = Seg(nLastSeg).Left - 10
            Seg(nNextSeg).Top = Seg(nLastSeg).Top
            If Seg(nNextSeg).Left = -10 Then Call StopGame("the wall")
        Case 2
            Seg(nNextSeg).Top = Seg(nLastSeg).Top - 10
            Seg(nNextSeg).Left = Seg(nLastSeg).Left
            If Seg(nNextSeg).Top = 90 Then Call StopGame("the wall")
        Case 3
            Seg(nNextSeg).Left = Seg(nLastSeg).Left + 10
            Seg(nNextSeg).Top = Seg(nLastSeg).Top
            If Seg(nNextSeg).Left = 600 Then Call StopGame("the wall")
        Case 4
            Seg(nNextSeg).Top = Seg(nLastSeg).Top + 10
            Seg(nNextSeg).Left = Seg(nLastSeg).Left
            If Seg(nNextSeg).Top = 400 Then Call StopGame("the wall")
    End Select
    
    
    '## Here we check to see if we've landed on top of a piece of food. If
    '## so, first we'll place a new piece of food on the arena, then we'll
    '## add a new index of Seg() and stamp it to the back of the snake.
    If (Seg(nNextSeg).Left = FoodL%) And (Seg(nNextSeg).Top = FoodT%) Then
        Call SetFood
        Load Seg(sCount%)
        Seg(sCount%).Left = Seg(nNextSeg).Left
        Seg(sCount%).Top = Seg(nNextSeg).Top
        Seg(sCount%).Visible = True
        Seg(sCount%).ZOrder 0
        lstNextSeg.AddItem sCount%
        lblEatenX.Caption = lblEatenX.Caption + 1
    End If
    
End Sub

Private Sub tmrTime_Timer()
    Dim s1 As Byte, s2 As Byte
    Dim m1 As Byte, m2 As Byte
    Dim h1 As Byte, h2 As Byte
    
    '## Parse the individual digits
    s1 = Right(lbls.Caption, 1)
    s2 = Left(lbls.Caption, 1)
    m1 = Right(lblm.Caption, 1)
    m2 = Left(lblm.Caption, 1)
    h1 = Right(lblh.Caption, 1)
    h2 = Left(lblh.Caption, 1)
    
    '## Tick every second and increment if needed
    s1 = s1 + 1
    If s1 = 10 Then
        s1 = 0
        s2 = s2 + 1
        If s2 = 6 Then
            s2 = 0
            m1 = m1 + 1
            If m1 = 10 Then
                m1 = 0
                m2 = m2 + 1
                If m2 = 6 Then
                    m2 = 0
                    h1 = h1 + 1
                    If h1 = 10 Then
                        h1 = 0
                        h2 = h2 + 1
                        If h2 = 10 Then h2 = 0
                    End If
                End If
            End If
        End If
    End If
    
    '## Display the new time
    lbls.Caption = s2 & s1
    lblm.Caption = m2 & m1
    lblh.Caption = h2 & h1
End Sub
