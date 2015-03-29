VERSION 5.00
Begin VB.Form frmTicTacToe 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tic Tac Toe"
   ClientHeight    =   1575
   ClientLeft      =   9660
   ClientTop       =   3690
   ClientWidth     =   3930
   Icon            =   "tic_tac_toe_2015.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1575
   ScaleWidth      =   3930
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmd 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   8
      Left            =   2400
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1080
      Width           =   375
   End
   Begin VB.CommandButton cmd 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   7
      Left            =   1920
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1080
      Width           =   375
   End
   Begin VB.CommandButton cmd 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   6
      Left            =   1440
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1080
      Width           =   375
   End
   Begin VB.CommandButton cmd 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   5
      Left            =   2400
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   600
      Width           =   375
   End
   Begin VB.CommandButton cmd 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   4
      Left            =   1920
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   600
      Width           =   375
   End
   Begin VB.CommandButton cmd 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   3
      Left            =   1440
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   600
      Width           =   375
   End
   Begin VB.CommandButton cmd 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   2
      Left            =   2400
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton cmd 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   1
      Left            =   1920
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton cmd 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   0
      Left            =   1440
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton cmdReset 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Reset"
      Height          =   1335
      Left            =   2880
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      UseMaskColor    =   -1  'True
      Width           =   975
   End
   Begin VB.Label lblHelp 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      Caption         =   "Click on the squares to place your piece. You will be facing against the computer."
      ForeColor       =   &H8000000B&
      Height          =   1335
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmTicTacToe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim endgame As Boolean
Dim winner As Integer

Option Explicit

Private Sub cmd_Click(Index As Integer)
    
    If Not endgame And Len(cmd(Index).Caption) < 1 Then
        cmd(Index).Caption = "x"
        
        If horwin() Or verwin() Or diawin() Then
            endgame = True
            gameover
        End If
        
        cmd(Index).Enabled = False
    End If
    
End Sub

Public Function horwin()
    
    horwin = IIf(horpos = -2, True, False)
    
End Function

Public Function verwin()

    verwin = IIf(verpos = -2, True, False)
    
End Function

Public Function diawin()

    diawin = IIf(diapos = -2, True, False)
    
End Function

Public Function horpos()
    
    horpos = dir(1, cmd.Count, 3)
    
End Function

Public Function verpos()
    
    verpos = dir(2, cmd.Count / 3, 1)
    
End Function

Public Function diapos()
    
    diapos = dir(3, 2, 1)
    
End Function

Public Function dir(ByVal sdir As Integer, ByVal cond As Integer, ByVal inc As Integer)

    Dim i As Integer, j As Integer, k As Integer, x As String, y As String, z As String, res As Integer

    i = 1
    res = -1
    
    For j = 1 To 3
        k = IIf(sdir = 3, IIf(i = 1, 1, -1), 1)
        x = IIf(sdir = 1 Or sdir = 2, cmd(i - 1).Caption, cmd(1 - k).Caption)
        y = IIf(sdir = 1, cmd(i).Caption, IIf(sdir = 2, cmd(i + IIf(sdir = 2, 0, -3) + 2).Caption, cmd(4).Caption))
        z = IIf(sdir = 1, cmd(i + 1).Caption, IIf(sdir = 2, cmd(i + IIf(sdir = 2, 0, -6) + 5).Caption, cmd(7 + k).Caption))
    
        If Len(x) > 0 And Len(y) > 0 And Len(z) > 0 Then If x = y And x = z Then res = -2: winner = IIf(x = "x", 1, 2)
        
        i = i + inc
    Next j
    
    i = 1
    
    Do While res < 0 And i <= cond And res <> -2
        k = IIf(sdir = 3, IIf(i = 1, 1, -1), 1)
        x = IIf(sdir = 1 Or sdir = 2, cmd(i - 1).Caption, cmd(1 - k).Caption)
        y = IIf(sdir = 1, cmd(i).Caption, IIf(sdir = 2, cmd(i + IIf(sdir = 2, 0, -3) + 2).Caption, cmd(4).Caption))
        z = IIf(sdir = 1, cmd(i + 1).Caption, IIf(sdir = 2, cmd(i + IIf(sdir = 2, 0, -6) + 5).Caption, cmd(7 + k).Caption))
        
        If Len(x) > 0 And Len(y) > 0 Then If x = y Then res = IIf(sdir = 1, i + 1, IIf(sdir = 2, i + 5, 7 + k)): If Len(cmd(res).Caption) > 0 Then res = -1
        If Len(x) > 0 And Len(z) > 0 Then If x = z Then res = IIf(sdir = 1, i, IIf(sdir = 2, i + 2, 4)): If Len(cmd(res).Caption) > 0 Then res = -1
        If Len(y) > 0 And Len(z) > 0 Then If y = z Then res = IIf(sdir = 1 Or sdir = 2, i - 1, 1 - k): If Len(cmd(res).Caption) > 0 Then res = -1
        
        i = i + inc
    Loop
    
    dir = res

End Function

Public Sub nextmove()
    
    Dim i As Integer
    
    i = -1
    
    Randomize
    
    If Len(cmd(4).Caption) < 1 Then
        i = 4
        cmd(i).Caption = "o"
    Else
        If horpos > -1 Then i = horpos
        If verpos > -1 Then i = verpos
        If diapos > -1 Then i = diapos
        
        If i < 0 Then
            Do
                i = Int(Rnd() * 9)
            Loop While Len(cmd(i).Caption) > 0
        End If
        
        cmd(i).Caption = "o"
    End If
    
    If horwin() Or verwin() Or diawin() Then endgame = True
    
    gameover
    
    cmd(i).Enabled = False
    
End Sub

Public Function gridfilled()

    Dim i As Integer
    Dim tmp As Boolean
    
    tmp = True
    
    For i = 1 To cmd.Count
        If Len(cmd(i - 1).Caption) < 1 Then tmp = False
    Next i
    
    If tmp Then winner = -1
    
    gridfilled = tmp
    
End Function

Public Sub gameover()

    Dim i As Integer
    
    If endgame Then
        MsgBox IIf(winner = -1, "Game ended in draw!", IIf(winner = 1, "Player has won!", "Computer has won!")), vbOKOnly, "Result"
        
        For i = 1 To cmd.Count
            cmd(i - 1).Enabled = False
        Next i
    End If
    
End Sub

Private Sub cmd_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    
    Dim i As Integer
    
    If Not endgame Then
        If horwin() Or verwin() Or diawin() Or gridfilled Then endgame = True: gameover
        
        If Not endgame Then nextmove
    End If
    
    If Not endgame Then If horwin() Or verwin() Or diawin() Or gridfilled Then endgame = True: gameover
    
End Sub

Private Sub cmdReset_Click()
    
    reset

End Sub

Public Sub reset()

    Dim i As Integer
    
    Randomize
    
    endgame = False
    winner = -1
    
    For i = 1 To cmd.Count
        cmd(i - 1).Caption = ""
        cmd(i - 1).Enabled = True
    Next i
    
    If (Int(Rnd() * 100 + 1) - 1) / 50 = 0 Then nextmove

End Sub

Private Sub Form_Load()

    reset
    
End Sub
