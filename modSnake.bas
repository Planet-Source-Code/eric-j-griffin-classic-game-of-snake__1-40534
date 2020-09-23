Attribute VB_Name = "modSnake"
Option Explicit

Public Direction%
Public FoodT%
Public FoodL%


Public Sub StopGame(sWhy$)
    
    '## Here we will stop the movement and timer
    With frmMain
        .tmrMove.Enabled = False
        .tmrTime.Enabled = False
    End With
    MsgBox "oops! ... you bit " & sWhy$ & "!!" & vbCrLf & vbCrLf & "what were you thinking?"
End Sub

Public Sub SetFood()
    Dim tPos%, lPos%
    
    '## Produce a random Top value divisible by 10
    Randomize
    tPos% = Int((Rnd * 289) + 100)
    Do Until tPos% Mod 10 = 0
        DoEvents
        tPos% = tPos% + 1
    Loop
    
    '## Produce a random Left value divisible by 10
    Randomize
    lPos% = Int((Rnd * 589))
    Do Until lPos% Mod 10 = 0
        DoEvents
        lPos% = lPos + 1
    Loop
    
    '## Set the food in its new position
    With frmMain
        .shpFood.Top = tPos%
        .shpFood.Left = lPos%
    End With
    
    '## Assign values to our global variables
    FoodT% = tPos%
    FoodL% = lPos%
End Sub
