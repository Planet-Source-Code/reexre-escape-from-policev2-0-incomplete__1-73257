VERSION 5.00
Begin VB.Form frmMAIN 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Escape from Police !"
   ClientHeight    =   9780
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   ScaleHeight     =   652
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1016
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.PictureBox MinPIC 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      ForeColor       =   &H00C00000&
      Height          =   810
      Left            =   360
      ScaleHeight     =   54
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   57
      TabIndex        =   10
      Top             =   6720
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.PictureBox BCKGRNDpic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      ForeColor       =   &H00C00000&
      Height          =   810
      Left            =   5280
      ScaleHeight     =   54
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   57
      TabIndex        =   9
      Top             =   7200
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.HScrollBar sFPS 
      Height          =   255
      Left            =   13920
      Max             =   100
      TabIndex        =   7
      Top             =   2880
      Value           =   75
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   13920
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   6
      Text            =   "frmMAIN.frx":0000
      Top             =   3960
      Width           =   1335
   End
   Begin VB.TextBox INFO 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   5
      Text            =   "frmMAIN.frx":002A
      Top             =   8760
      Width           =   12855
   End
   Begin VB.Timer TimerFPS 
      Interval        =   2000
      Left            =   12840
      Top             =   6720
   End
   Begin VB.CommandButton cmdGravYesNO 
      Caption         =   "Wind Yes_No"
      Height          =   615
      Left            =   14160
      TabIndex        =   3
      Top             =   3240
      Width           =   855
   End
   Begin VB.CommandButton CmdAddOBJ 
      Caption         =   "Add Police"
      Height          =   375
      Left            =   13920
      TabIndex        =   2
      Top             =   1680
      Width           =   1335
   End
   Begin VB.PictureBox PIC 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      ForeColor       =   &H00C00000&
      Height          =   6375
      Left            =   120
      MouseIcon       =   "frmMAIN.frx":0134
      MousePointer    =   99  'Custom
      ScaleHeight     =   425
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   913
      TabIndex        =   1
      Top             =   120
      Width           =   13695
   End
   Begin VB.CommandButton cmdSTART 
      Caption         =   "RE-START"
      Height          =   615
      Left            =   13920
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Here are dispalied Steer, Gas, Damage. and speed. Damage slowly auto repairs. Collisions make a strong gas decrement."
      Height          =   1815
      Index           =   1
      Left            =   13920
      TabIndex        =   13
      Top             =   7920
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "D"
      Height          =   255
      Left            =   14760
      TabIndex        =   12
      Top             =   6360
      Width           =   255
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "G"
      Height          =   255
      Index           =   0
      Left            =   14040
      TabIndex        =   11
      Top             =   6360
      Width           =   255
   End
   Begin VB.Shape ShapeDAMAGE 
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   14760
      Top             =   5760
      Width           =   255
   End
   Begin VB.Line LineSPEED 
      BorderWidth     =   3
      X1              =   928
      X2              =   968
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Shape ShapeGAS 
      FillColor       =   &H0000FF00&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   14040
      Top             =   5760
      Width           =   255
   End
   Begin VB.Shape ShapeSteer 
      FillColor       =   &H0000FF00&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   14520
      Top             =   5040
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Try to Keep Constant FPS"
      Height          =   615
      Left            =   13920
      TabIndex        =   8
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label lFPS 
      Height          =   615
      Left            =   13920
      TabIndex        =   4
      Top             =   840
      Width           =   1335
   End
End
Attribute VB_Name = "frmMAIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Author :Roberto Mior
'     reexre@gmail.com
'--------------------------------------------------------------------------------

'*********************** 2D Collision:
'Original C++ code written by Benedikt Bitterli Copyright (c) 2009 [The code is released under the ZLib/LibPNG license]
'Original C++ code and tutorial available at links http://www.gamedev.net/reference/programming/features/verletPhys/default.asp
'http://www.gamedev.net/reference/articles/article2714.asp
'Forum:
'http://www.gamedev.net/community/forums/topic.asp?topic_id=553845
'Conversion from C++ to Java done by Craig Mitchell Copyright (c) 2010.
'Conversion from "C++ & Java" to VB6 done by Roberto Mior Copyright (c) 2010.
'******************************

'GAME by Roberto Mior

Option Explicit

Dim X
Dim Y

Dim I                  As Long

Public Sub CmdAddOBJ_Click()


    ADDBox (WorldW - 80), (WorldH - 80), _
           17, 32, True, True

    Npolices = Npolices + 1


End Sub

Private Sub cmdGravYesNO_Click()
    If GravityY = 0 Then
        GravityY = 0.02
        GravityX = (Rnd - 0.5) * 0.04
    Else
        GravityY = 0: GravityX = 0
    End If
End Sub

Private Sub cmdSTART_Click()
    Npolices = 0

    INITSound frmMAIN.hWnd, 1

    BCKGRNDpic = LoadPicture(App.Path & "\BG_big.bmp")
    BCKGRNDpic.Refresh


    FIllwithTiles App.Path & "\BG_big.bmp", App.Path & "\Textures\TilesPlain0021_2_thumbhuge.jpg"

    'WorldW = PIC.Width * 2 - 1
    'WorldH = PIC.Height * 2 - 1

    WorldW = BCKGRNDpic.Width - 1
    WorldH = BCKGRNDpic.Height - 1



    MinPIC.Cls
    MinPIC.Width = WorldW \ MD
    MinPIC.Height = WorldH \ MD
    MinPIC.Refresh
    SetStretchBltMode frmMAIN.MinPIC.Hdc, vbPaletteModeNone
    StretchBlt frmMAIN.MinPIC.Hdc, 0, 0, WorldW \ MD, WorldH \ MD, _
               frmMAIN.BCKGRNDpic.Hdc, 0, 0, frmMAIN.BCKGRNDpic.Width - 1, frmMAIN.BCKGRNDpic.Height - 1, vbSrcCopy
    frmMAIN.MinPIC.Refresh



    NB = 2
    ReDim B(NB)
    NB = 0



    'add player
    ADDBox 100, WorldH - 200, 17, 32, False    '18, 40
    B(1).color = RGB(230, 255, 60)

    '    ADDBox 100, WorldH - 100, 30, 50, False
    '    ADDjoint 1, 2, 3, 2, 20
    '    ADDjoint 1, 2, 4, 1, 20
    If B(1).HasWheels Then
        LoadAndPlayEngine App.Path & "\boatoutboard_2.wav", 1
        LoadDRIFT App.Path & "\carscreech2.wav", 1
    End If

    CmdAddOBJ_Click


    'WORLD BOUNDARIES
    ADDBox 0, 0, WorldW, 50, False, False
    B(NB).SetNotMovable: B(NB).color = RGB(180, 50, 0)
    ADDBox 0, WorldH - 50, WorldW, 50, False, False
    B(NB).SetNotMovable: B(NB).color = RGB(180, 50, 0)
    ADDBox 0, 50, 50, WorldH - 100, False, False
    B(NB).SetNotMovable: B(NB).color = RGB(180, 50, 0)
    ADDBox WorldW - 50, 50, 50, WorldH - 100, False, False
    B(NB).SetNotMovable: B(NB).color = RGB(180, 50, 0)
    '****************


    For I = 1 To 10
        ' ADDBox (WorldW - 150) * Rnd, (WorldH - 50) * Rnd, 18, 18, False, False
        ADDBox (WorldW \ 2 - 150) + Rnd * 300, (WorldH \ 2 - 150) + Rnd * 300, 12, 48, False, False
        B(NB).color = RGB(0, 100, 0)
    Next


    '*** Big Boxes
    ADDBox 400, 300, 140, 140, False, False
    B(NB).NotMovable = True
    B(NB).color = RGB(0, 100, 0)
    B(NB).SetNotMovable
    For I = 1 To 5
        DuplicateOBJ NB, Rnd * (WorldW - 150), Rnd * (WorldH - 150), True
    Next
    '***********


    GravityY = 0             ' 0.1
    UpDateCamera



    '    LoadAndPlayEngine App.Path & "\engine.wav", 1
    'LoadAndPlayEngine App.Path & "\boatoutboard.wav", 1
    MAINLOOP

End Sub

Private Sub Form_Activate()
    sFPS_Change

    PIC.Cls
    PIC.Height = PIC.Width * 0.618
    PIC.Refresh
    INFO.Width = PIC.Width



    ScreenW = PIC.Width
    ScreenH = PIC.Height



    cmdSTART_Click
    DoEvents
End Sub

Private Sub Form_Load()
    Randomize Timer
    Me.Caption = Me.Caption & " V" & App.Major & "." & App.Minor


End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim I              As Long

    '    For I = 1 To UBound(ENGINEsound)
    '        ENGINEsound(I).Stop
    '        Set ENGINEsound(I) = Nothing
    '    Next

    For I = 1 To NB - 1
        If B(I).HasWheels Then B(I).ENGINEsound.Stop
        Set B(I).ENGINEsound = Nothing
    Next


    CleanupSounds

    End

End Sub


Private Sub PIC_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim D              As Single
    Dim MinD           As Single
    MinD = 99999999999#
    For I = 1 To NB
        D = Distance(B(I).CenterX - X, B(I).CenterY - Y)
        If D < MinD Then
            MinD = D
            Omove = I
        End If
    Next
    Xmouse = X
    Ymouse = Y
End Sub

Private Sub PIC_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)


    If Omove <> 0 Then
        For I = 1 To B(Omove).NP
            B(Omove).SetPointX(I) = B(Omove).getPointX(I) + (X - B(Omove).getPointX(I)) * 0.01
            B(Omove).SetPointY(I) = B(Omove).getPointY(I) + (Y - B(Omove).getPointY(I)) * 0.01
        Next
    End If
    'Xmouse = X
    'Ymouse = Y

End Sub

Private Sub PIC_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Omove = 0
End Sub

Private Sub sFPS_Change()
    If sFPS <> 0 Then
        Label1 = "Try to Keep " & sFPS & " Constant FPS"
    Else
        Label1 = "Running at Max Speed"
    End If

End Sub

Private Sub sFPS_Scroll()
    If sFPS <> 0 Then
        Label1 = "Try to Keep " & sFPS & " Constant FPS"
    Else
        Label1 = "Running at Max Speed"
    End If

End Sub

Private Sub TimerFPS_Timer()

    lFPS = "FPS = " & (CNT - OldCNT) / (TimerFPS.Interval / 1000)
    lFPS = lFPS & "   Objs=" & NB
    lFPS = lFPS & "   Polices=" & Npolices
    OldCNT = CNT
    DoEvents

End Sub



Public Sub FIllwithTiles(FileBig As String, FileTile As String)
    Dim X              As Long
    Dim Y              As Long
    Dim Xstep          As Long
    Dim Ystep          As Long
    MinPIC.AutoSize = True

    MinPIC = LoadPicture(FileTile)
    MinPIC.Refresh
    Xstep = MinPIC.Width
    Ystep = MinPIC.Height


    For X = 0 To BCKGRNDpic.Width Step Xstep
        For Y = 0 To BCKGRNDpic.Height Step Ystep
            BitBlt BCKGRNDpic.Hdc, X, Y, Xstep, Ystep, MinPIC.Hdc, 0, 0, vbSrcCopy
        Next
    Next
    BCKGRNDpic.Refresh

    MinPIC.AutoSize = False
End Sub
