Attribute VB_Name = "modWorld"
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

Public Type tPoint
    X                  As Single
    Y                  As Single
    OldX               As Single
    OldY               As Single
    AccX               As Single
    AccY               As Single


    Mass               As Single


    IsNOTJoint         As Boolean

    DrawX              As Single
    DrawY              As Single

End Type

Public Type tEdge
    V1                 As Long
    V2                 As Long
    MainLength         As Single
    Boundary           As Boolean
End Type



Public Const TimeStep  As Single = 1
Public Const TimeStep2 As Single = TimeStep * TimeStep

Public NB              As Long
Public B()             As New clsBody


Public WorldW          As Single
Public WorldH          As Single
Public ScreenW         As Single
Public ScreenH         As Single

Public GravityX        As Single
Public GravityY        As Single

Public Const AirFriction = 0.999    '0.9999

Public Const Pi = 3.14159265358979

Public Const InfiniteMASS As Single = 1E+15

Public CNT             As Long
Public OldCNT          As Long

Public Omove           As Long
Public Xmouse          As Single
Public Ymouse          As Single

Private pTime          As Double

Public Const MD = 20         'Miniature Divisor

Public Npolices        As Long

Public Function Distance(DX As Single, Dy As Single) As Single
    Distance = Sqr(DX * DX + Dy * Dy)
End Function
Public Function DistanceSQ(DX As Single, Dy As Single) As Single
    DistanceSQ = (DX * DX + Dy * Dy)
End Function
Public Sub Normalize(ByRef X As Single, ByRef Y As Single)
    Dim L              As Single
    L = Sqr(X * X + Y * Y)
    If L <> 0 Then L = 1 / L    'Else: Stop

    X = X * L
    Y = Y * L
End Sub


'Public Function Projection(V As cls2DVector) As cls2DVector
'    ' def projection(self, vector):
'    '        k = (self.dot(vector)) / vector.length()
'    '        return k * vector.unit()
'
'    Dim K As Single
'    Set Projection = New cls2DVector
'    K = (X * V.X + Y * V.Y) / Sqr(V.X * V.X + V.Y * V.Y)
'    V.Normalize
'    V.MUL K
'
'    Set Projection = V
'End Function


Public Sub VectorProject(ByVal Vx As Single, ByVal Vy As Single, _
                         PtoX, PtoY, ByRef RVx As Single, ByRef RVy As Single)

'Vx,Vy Vector to Project
'PtoX,PtoY Vector to Project TO
'PtoX and PtoY must be a vector of lenght=1
'RVX,RVY Returned resule

    Dim K              As Single
    K = (Vx * PtoX + Vy * PtoY)    '(/1 len Pto)
    'Normalize Vx, Vy
    RVx = PtoX * K
    RVy = PtoY * K


End Sub


Public Function MathMIN(ByRef A As Single, ByRef B As Single) As Single
    MathMIN = IIf(A < B, A, B)
End Function
Public Function MathMAX(ByRef A As Single, ByRef B As Single) As Single
    MathMAX = IIf(A > B, A, B)
End Function




Public Function IntervalDistance(ByRef MinA As Single, ByRef MaxA As Single, _
                                 ByRef MinB As Single, ByRef MaxB As Single) As Single
    If MinA < MinB Then
        IntervalDistance = MinB - MaxA
    Else
        IntervalDistance = MinA - MaxB
    End If


    '    IntervalDistance( float MinA, float MaxA, float MinB, float MaxB ) {
    '    if( MinA < MinB )
    '        return MinB - MaxA;
    '    Else
    '        return MinA - MaxB;'


End Function



Public Sub MAINLOOP()
    Const OneMillisec  As Long = 1

    Dim I              As Long
    Dim InvFPS         As Double


    Timing = 0
    pTime = Timing
    Do


        '***** Keep Constant FPS
        If frmMAIN.sFPS <> 0 Then
            InvFPS = 1 / frmMAIN.sFPS
            Do
                'Sleep OneMillisec
            Loop While (Timing < (pTime + InvFPS))
            pTime = Timing
        End If
        '****************




        'BitBlt frmMAIN.PIC.hdc, 0, 0, frmMAIN.PIC.ScaleWidth, frmMAIN.PIC.ScaleHeight, frmMAIN.PIC.hdc, 0, 0, vbBlack    'ness
        DRAWworld frmMAIN.PIC.Hdc
        For I = NB To 1 Step -1
            B(I).DRAW frmMAIN.PIC.Hdc
        Next
        frmMAIN.PIC.Refresh
        DoEvents


        '****Commands and Police Movement
        UpDateCamera
        DoEvents
        KEYBOARD
        For I = 2 To NB
            If B(I).IsPolice Then MovePOLICE I
        Next
        '****


        For I = 1 To NB

            B(I).DoWheels
            B(I).UpDateForces
            B(I).UpDateVerlet


        Next


        'MsgBox B(1).Angle

        IterateCollisions
        UpDateJoints

        DoEvents

        CNT = CNT + 1

        If CNT Mod 3000 = 0 Then frmMAIN.CmdAddOBJ_Click

        If CNT Mod 13 = 0 Then UpdateSOUNDs

        If CNT Mod 451 = 0 Then
            For I = 1 To NB
                B(I).CheckAndRestoreFlipped
            Next
        End If


        '        If CNT Mod 20 = 0 Then
        '            For I = 1 To NB
        '                B(I).Steer = B(I).Steer + B(I).Damage * IIf(Rnd < 0.5, 1, -1)
        '            Next
        '        End If

        UpDateShapes

    Loop While True

End Sub


Public Sub ADDBox(X, Y, W, H, IsPolice As Boolean, Optional HasWheels As Boolean = True)
    NB = NB + 1
    ReDim Preserve B(NB)
    If HasWheels Then
        LoadAndPlayEngine App.Path & "\boatoutboard_2.wav", NB
        LoadDRIFT App.Path & "\carscreech2.wav", NB
    End If

    With B(NB)

        .ADDPoint X, Y
        .ADDPoint X + W, Y
        .ADDPoint X + W, Y + H
        .ADDPoint X, Y + H

        .ADDEdge 1, 2
        .ADDEdge 2, 3
        .ADDEdge 3, 4
        .ADDEdge 4, 1
        .ADDEdge 2, 4, False
        .ADDEdge 1, 3, False
        .color = RGB(80 + Rnd * 175, 80 + Rnd * 175, 80 + Rnd * 175)
        .IsPolice = IsPolice

        If IsPolice Then
            .color = RGB(250 - Rnd * 80, 250 - Rnd * 80, 255)
            .PoliceSteerSkill = 0.75 + Rnd * 1    '0.5 + Rnd * 1  '1.25
            .PoliceSkill2 = 20 + Rnd * 15    '25
        End If


        .HasWheels = HasWheels


    End With

End Sub

Public Sub ADDTriangle(X, Y, W, H)
    NB = NB + 1
    ReDim Preserve B(NB)
    If B(NB).HasWheels Then
        LoadAndPlayEngine App.Path & "\boatoutboard_2.wav", NB
        LoadDRIFT App.Path & "\carscreech2.wav", NB
    End If
    With B(NB)
        .ADDPoint X, Y
        .ADDPoint X + W + Rnd * 10, Y
        .ADDPoint X + W + Rnd * 10, Y + H + Rnd * 10

        .ADDEdge 1, 2
        .ADDEdge 2, 3
        .ADDEdge 3, 1
        .color = RGB(80 + Rnd * 175, 80 + Rnd * 175, 80 + Rnd * 175)

    End With
End Sub

Public Sub DuplicateOBJ(ByVal wO, DX, Dy, Optional Absolute = False)
    Dim I              As Long

    NB = NB + 1
    ReDim Preserve B(NB)

    If B(NB).HasWheels Then
        LoadAndPlayEngine App.Path & "\boatoutboard_2.wav", NB
        LoadDRIFT App.Path & "\carscreech2.wav", NB
    End If

    B(wO).CalculateCenter

    With B(NB)

        For I = 1 To B(wO).NP
            If Absolute Then
                .ADDPoint B(wO).getPointX(I) - B(wO).MinX + DX, B(wO).getPointY(I) - B(wO).MinY + Dy
            Else
                .ADDPoint B(wO).getPointX(I) + DX, B(wO).getPointY(I) + Dy
            End If
        Next

        For I = 1 To B(wO).NE
            .ADDEdge B(wO).getEdgeV1(I), B(wO).getEdgeV2(I), B(wO).getEdgeIsBoundary(I)

        Next


        'For inifinitemass
        For I = 1 To B(wO).NP
            .SetMASS(I) = B(wO).getMASS(I)
        Next

        .color = B(wO).color
        .IsPolice = B(wO).IsPolice
        .HasWheels = B(wO).HasWheels
        .PoliceSteerSkill = B(wO).PoliceSteerSkill
        .PoliceSkill2 = B(wO).PoliceSkill2
        .NotMovable = B(wO).NotMovable

    End With
End Sub



Public Function Atan2(ByVal DX As Single, ByVal Dy As Single) As Single

    Dim theta          As Single

    If (Abs(DX) < 0.0000001) Then
        If (Abs(Dy) < 0.0000001) Then
            theta = 0#
        ElseIf (Dy > 0#) Then
            theta = 1.5707963267949
            'theta = PI / 2
        Else
            theta = -1.5707963267949
            'theta = -PI / 2
        End If
    Else
        theta = Atn(Dy / DX)

        If (DX < 0) Then
            If (Dy >= 0#) Then
                theta = Pi + theta
            Else
                theta = theta - Pi
            End If
        End If
    End If

    Atan2 = theta

    If Atan2 < 0 Then Atan2 = Atan2 + Pi * 2

End Function
Public Function AngleDiff(A1 As Single, A2 As Single) As Single
'double difference = secondAngle - firstAngle;
'while (difference < -180) difference += 360;
'while (difference > 180) difference -= 360;
'return difference;

    AngleDiff = A2 - A1
    While AngleDiff < -Pi
        AngleDiff = AngleDiff + Pi * 2
    Wend
    While AngleDiff > Pi
        AngleDiff = AngleDiff - Pi * 2
    Wend

    '''' this is to have values between 0 and 1
    'AngleDiff = AngleDiff + PI
    'AngleDiff = AngleDiff / (PI * 2)

End Function






Public Sub MovePOLICE(wB As Long)
'by Roberto Mior

    Dim A              As Single
    Dim ADiff          As Single
    Dim AbsADiff       As Single

    Dim TargetFeelSpeed As Single

    With B(wB)

        TargetFeelSpeed = B(1).Speed * IIf(B(1).GoingForward, 1, -1) * B(wB).PoliceSkill2    '25

        A = Atan2(B(1).CenterX + Cos(B(1).Angle) * TargetFeelSpeed - .CenterX, _
                  B(1).CenterY + Sin(B(1).Angle) * TargetFeelSpeed - .CenterY)

        ADiff = AngleDiff(.Angle, A)
        AbsADiff = Abs(ADiff)

        .cDoSteer Sgn(ADiff) * IIf((AbsADiff < Pi * 0.46) Or (.GoingForward), (AbsADiff * .PoliceSteerSkill), -(AbsADiff * .PoliceSteerSkill))

        If AbsADiff < Pi * 0.46 Then .cAccellerate Else: .cBrake


    End With


End Sub
Public Sub MovePOLICEAvoiding(wB As Long)
    Dim A              As Single
    Dim ADiff          As Single
    Dim AbsADiff       As Single

    Dim MyCarSpeed     As Single
    Dim NearOdist      As Single
    Dim ODist          As Single
    Dim NO             As Long
    Dim I              As Long
    Dim SteerAmout     As Single

    With B(wB)

        NearOdist = 999999999
        For I = 2 To NB
            If I <> wB Then
                If Not (B(I).IsPolice) Then

                    ODist = Distance(.CenterX + Cos(.Angle * .Speed * 30) - B(I).CenterX, _
                                     .CenterY + Sin(.Angle * .Speed * 30) - B(I).CenterY)
                    If ODist < NearOdist Then
                        '                    Stop

                        NearOdist = ODist
                        NO = I
                    End If
                End If
            End If

        Next

        If NearOdist > 80 + .Speed * IIf(.GoingForward, 1, -1) * 30 Then
            'STANDARD
            '        A = Atan2(B(1).CenterX - .CenterX , B(1).CenterY - .CenterY)

            MyCarSpeed = B(1).Speed * IIf(B(1).GoingForward, 1, -1) * B(wB).PoliceSkill2    '25

            A = Atan2(B(1).CenterX + Cos(B(1).Angle) * MyCarSpeed - .CenterX, _
                      B(1).CenterY + Sin(B(1).Angle) * MyCarSpeed - .CenterY)

            ADiff = AngleDiff(.Angle, A)
            AbsADiff = Abs(ADiff)

            .cDoSteer Sgn(ADiff) * IIf((AbsADiff < Pi * 0.36) Or (.GoingForward), (AbsADiff * .PoliceSteerSkill), -(AbsADiff * .PoliceSteerSkill))

            If AbsADiff < Pi * 0.36 Then .cAccellerate Else: .cBrake

        Else
            'AVOID

            MyCarSpeed = .Speed * IIf(.GoingForward, 1, -1) * 30    ' * B(wB).PoliceSkill2    '25

            A = Atan2(-B(NO).CenterX + Cos(.Angle) * MyCarSpeed + .CenterX, _
                      -B(NO).CenterY + Sin(.Angle) * MyCarSpeed + .CenterY)

            ADiff = AngleDiff(.Angle, A)
            AbsADiff = Abs(ADiff)

            SteerAmout = ((Pi - AbsADiff) * .PoliceSteerSkill) * 0.1

            .cDoSteer Sgn(ADiff) * IIf((AbsADiff < Pi * 0.12) Or (.GoingForward), SteerAmout, -SteerAmout)

            If AbsADiff < Pi * 0.12 Then .cBrake Else: .cAccellerate

        End If

    End With


End Sub
Public Sub DRAWworld(PicHdc As Long)
    Dim I              As Long

    Dim X1             As Long
    Dim Y1             As Long

    Dim Y              As Long

    X1 = CamX \ 1
    Y1 = CamY \ 1

    If Rnd < 0.0001 Then
        SetStretchBltMode frmMAIN.MinPIC.Hdc, vbPaletteModeNone
        StretchBlt frmMAIN.MinPIC.Hdc, 0, 0, WorldW \ MD, WorldH \ MD, _
                   frmMAIN.BCKGRNDpic.Hdc, 0, 0, frmMAIN.BCKGRNDpic.Width - 1, frmMAIN.BCKGRNDpic.Height - 1, vbSrcCopy
    End If

    'Black
    BitBlt frmMAIN.PIC.Hdc, 0, 0, frmMAIN.PIC.ScaleWidth, frmMAIN.PIC.ScaleHeight, frmMAIN.PIC.Hdc, 0, 0, vbBlackness

    BitBlt PicHdc, X1, Y1, _
           frmMAIN.PIC.Width - X1, frmMAIN.PIC.Height - Y1, frmMAIN.BCKGRNDpic.Hdc, 0, 0, vbSrcCopy


    'MINPIC
    BitBlt frmMAIN.PIC.Hdc, 0, 0, WorldW \ MD, WorldH \ MD, frmMAIN.MinPIC.Hdc, 0, 0, vbSrcCopy
    'FastLine PicHdc, WorldW \ MD, 0, WorldW \ MD, WorldH \ MD, 1, vbWhite
    'FastLine PicHdc, 0, WorldH \ MD, WorldW \ MD, WorldH \ MD, 1, vbWhite, False


    '    For I = 1 To NB
    '        With B(I)
    '
    '            X1 = .CenterX \ MD
    '            Y1 = .CenterY \ MD
    '            MyCircle PicHdc, X1, Y1, 2, 2, .color
    '            For Y = .MinY \ MD To .MaxY \ MD Step 1
    '            FastLine PicHdc, .MinX \ MD - 1, Y, .MaxX \ MD, Y, 1, .color '(.MaxY - .MinY) \ MD
    '            Next
    '
    '        End With
    '    Next


End Sub



Public Sub UpdateSOUNDs()
''' '  ENGINEsound(1).SetFrequency CLng((B(1).Speed * 9500) Mod 50000 + 35000)
' ENGINEsound(1).SetFrequency CLng((B(1).Speed * 4500) + 6000)
    Dim I              As Long
    Dim DX             As Single
    Dim Dy             As Single
    Dim D              As Long


    B(1).ENGINEsound.SetFrequency CLng((B(1).Speed * 4500) + 6000)
    For I = 2 To NB
        With B(I)
            If .HasWheels Then
                DX = .CenterX - B(1).CenterX
                Dy = .CenterY - B(1).CenterY
                D = Sqr(DX * DX + Dy * Dy)
                .ENGINEsound.SetVolume -10 - D
                .ENGINEsound.SetPan CLng(DX * 1.5)
                .ENGINEsound.SetFrequency CLng((B(I).Speed * 4500) + 6000)

                If .DRIFTsound.GetStatus = DSBSTATUS_PLAYING Then
                    .DRIFTsound.SetVolume .DriftVOL - D - 10
                    .DRIFTsound.SetPan CLng(DX * 1.5)
                End If


            End If
        End With
    Next


End Sub

Public Sub UpDateShapes()
    With frmMAIN
        If B(1).Steer >= 0 Then
            .ShapeSteer.Left = 968
            .ShapeSteer.Width = B(1).Steer * 100
        Else
            .ShapeSteer.Left = 968 + B(1).Steer * 100
            .ShapeSteer.Width = -B(1).Steer * 100

        End If

        If B(1).GAS <= 0 Then
            .ShapeGAS.Top = 420
            .ShapeGAS.Height = -B(1).GAS * 1000
        Else
            .ShapeGAS.Top = 420 - B(1).GAS * 1000
            .ShapeGAS.Height = B(1).GAS * 1000

        End If


        .LineSPEED.X1 = .LineSPEED.X2 + Cos(-Pi + B(1).Speed * 0.12) * 45
        .LineSPEED.Y1 = .LineSPEED.Y2 + Sin(-Pi + B(1).Speed * 0.12) * 45


        .ShapeDAMAGE.Top = 420 - B(1).Damage * 50
        .ShapeDAMAGE.Height = B(1).Damage * 50

    End With

End Sub
