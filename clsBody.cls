VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBody"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
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

'Cars By Roberto Mior



Option Explicit


Private P()            As tPoint
Public NP              As Long


Private E()            As tEdge
Public NE              As Long

Public CenterX         As Single
Public CenterY         As Single
Public MinX            As Single
Public MinY            As Single
Public MaxX            As Single
Public MaxY            As Single

Public Area            As Single
Public Perim           As Single
Public color           As Long

Public Angle           As Single
Public GAS             As Single
Public Steer           As Single
Private Const MaxSteer As Single = Pi * 0.25
'Public Braking         As Single
'Public InvBrake        As Single

Public GoingForward    As Boolean
Public IsPolice        As Boolean
Public HasWheels       As Boolean
Public Speed           As Single

Public PoliceSteerSkill As Single
Public PoliceSkill2    As Single

Public NotMovable      As Boolean

Public ENGINEsound     As DirectSoundSecondaryBuffer8
Public DRIFTsound      As DirectSoundSecondaryBuffer8
Public DriftVOL        As Long

Public MaxDRIFT        As Single

Public Damage          As Single
Private Const DamageRepair As Single = 0.996    '0.998




Public Property Get IsNOTJoint(WP As Long) As Boolean
    IsNOTJoint = P(WP).IsNOTJoint
End Property
Public Property Let IsNOTJoint(WP As Long, V As Boolean)
    P(WP).IsNOTJoint = V
End Property

Public Property Get getEdgeV1(wE As Long) As Long
    getEdgeV1 = E(wE).V1
End Property
Public Property Get getEdgeV2(wE As Long) As Long
    getEdgeV2 = E(wE).V2
End Property
Public Property Get getEdgeIsBoundary(wE As Long) As Boolean
    getEdgeIsBoundary = E(wE).Boundary
End Property
Public Property Get getPointX(WP As Long) As Single
    getPointX = P(WP).X
End Property
Public Property Get getPointY(WP As Long) As Single
    getPointY = P(WP).Y
End Property
Public Property Get getPointOldX(WP As Long) As Single
    getPointOldX = P(WP).OldX
End Property
Public Property Get getPointOldY(WP As Long) As Single
    getPointOldY = P(WP).OldY
End Property

Public Property Let SetPointX(WP As Long, V As Single)
    P(WP).X = V
End Property
Public Property Let SetPointY(WP As Long, V As Single)
    P(WP).Y = V
End Property

Public Property Get getMASS(WP As Long) As Single
    getMASS = P(WP).Mass
End Property
Public Property Let SetMASS(WP As Long, V As Single)
    P(WP).Mass = V
End Property
Public Sub UpDateVerlet()
'**** For Cars by Roberto Mior

    Dim I              As Long
    Dim tX             As Single
    Dim tY             As Single
    Dim Vx             As Single
    Dim Vy             As Single

    Dim XtoAng         As Single
    Dim YtoAng         As Single
    Dim XPerpAng       As Single
    Dim YPerpAng       As Single

    Dim Ang            As Single
    Dim Aperp          As Single


    Dim DRIFT          As Single
    Dim drColor        As Long


    Speed = 0
    If NotMovable Then Exit Sub

    MaxDRIFT = 0
    For I = 1 To NP

        With P(I)

            tX = .X
            tY = .Y
            Vx = (.X - .OldX)
            Vy = (.Y - .OldY)


            If HasWheels Then
                If I > 2 Then
                    'Rear Wheel
                    Aperp = Angle + Pi * 0.5
                    VectorProject Vx, Vy, Cos(Angle), Sin(Angle), XtoAng, YtoAng
                    VectorProject Vx, Vy, Cos(Aperp), Sin(Aperp), XPerpAng, YPerpAng
                    Speed = Speed + Distance(XtoAng, YtoAng)
                Else
                    'Front Wheel
                    Ang = Steer + Angle
                    Aperp = Ang + Pi * 0.5
                    VectorProject Vx, Vy, Cos(Ang), Sin(Ang), XtoAng, YtoAng
                    VectorProject Vx, Vy, Cos(Aperp), Sin(Aperp), XPerpAng, YPerpAng
                    Speed = Speed + Distance(XtoAng, YtoAng)
                End If


                .X = .X + (XtoAng + XPerpAng * 0.9) * AirFriction + .AccX * TimeStep2
                .Y = .Y + (YtoAng + YPerpAng * 0.9) * AirFriction + .AccY * TimeStep2

                DRIFT = Distance(XPerpAng, YPerpAng)

                If DRIFT > 1.8 Then
                    drColor = 128 - (DRIFT - 1.8) * 30
                    If drColor < 0 Then drColor = 0
                    drColor = RGB(drColor, drColor, drColor)

                    FastLine frmMAIN.BCKGRNDpic.Hdc, .X \ 1, .Y \ 1, tX \ 1, tY \ 1, 2, drColor
                End If

                If DRIFT > MaxDRIFT Then MaxDRIFT = DRIFT


            Else
                Speed = Speed + Distance(Vx, Vy)
                '0.99 def friction
                .X = .X + (Vx * 0.995) * AirFriction + .AccX * TimeStep2
                .Y = .Y + (Vy * 0.995) * AirFriction + .AccY * TimeStep2
            End If


            .OldX = tX
            .OldY = tY
        End With
    Next

    Speed = Speed / NP



    '*********SOUNDS
    If HasWheels Then
        If MaxDRIFT > 1.8 Then

            If DRIFTsound.GetStatus <> DSBSTATUS_PLAYING Then
                DRIFTsound.Play DSBPLAY_LOOPING

            Else

            End If
            DriftVOL = CLng(-4000 + DRIFT * 800)
            If DriftVOL > 0 Then DriftVOL = 0
            DRIFTsound.SetVolume DriftVOL

        Else
            DriftVOL = -9999999
            DRIFTsound.Stop

        End If

        Damage = Damage * DamageRepair
    End If
    '**********

End Sub



Public Sub UpDateEdges()

    Dim I              As Long
    Dim v1v2X          As Single
    Dim v1v2Y          As Single
    Dim V1V2Length     As Single
    Dim Diff           As Single

    If NotMovable Then Exit Sub

    For I = 1 To NE
        With E(I)

            v1v2X = P(.V2).X - P(.V1).X
            v1v2Y = P(.V2).Y - P(.V1).Y
            V1V2Length = Distance(v1v2X, v1v2Y)
            Diff = (V1V2Length - .MainLength) * 0.5    '*.5

            Normalize v1v2X, v1v2Y

            'If .Boundary Then
            '    Diff = Diff * 1
            'Else
            '    Diff = Diff * 1
            'End If

            v1v2X = v1v2X * Diff
            v1v2Y = v1v2Y * Diff

            P(.V1).X = P(.V1).X + v1v2X    '* Diff
            P(.V1).Y = P(.V1).Y + v1v2Y    '* Diff
            P(.V2).X = P(.V2).X - v1v2X    '* Diff
            P(.V2).Y = P(.V2).Y - v1v2Y    '* Diff


        End With
    Next




End Sub
Public Sub ProjectToAxis(AxisX As Single, AxisY As Single, ByRef Min As Single, ByRef Max As Single)
    Dim DotP           As Single
    Dim I              As Long

    DotP = AxisX * P(1).X + AxisY * P(1).Y
    Min = DotP
    Max = DotP
    For I = 2 To NP
        With P(I)
            DotP = AxisX * .X + AxisY * .Y
            Min = MathMIN(DotP, Min)
            Max = MathMAX(DotP, Max)
        End With

    Next

End Sub

Public Sub CalculateCenter()
    Dim I              As Long

    CenterX = 0
    CenterY = 0

    MinX = 100000
    MinY = 100000
    MaxX = -100000
    MaxY = -100000

    For I = 1 To NP
        With P(I)
            CenterX = CenterX + .X
            CenterY = CenterY + .Y
            MinX = MathMIN(MinX, .X)
            MinY = MathMIN(MinY, .Y)
            MaxX = MathMAX(MaxX, .X)
            MaxY = MathMAX(MaxY, .Y)
        End With

    Next

    CenterX = CenterX / NP
    CenterY = CenterY / NP

End Sub


Public Sub ADDPoint(ByVal X, ByVal Y)
    NP = NP + 1
    ReDim Preserve P(NP)
    With P(NP)
        .X = X
        .Y = Y
        .OldX = X
        .OldY = Y
        .AccX = 0
        .AccY = 0
        .IsNOTJoint = True
    End With
    CalculateCenter
End Sub
Public Sub ADDEdge(ByVal P1 As Long, ByVal P2 As Long, Optional IsBoundary = True)
'Stop

    NE = NE + 1
    ReDim Preserve E(NE)
    With E(NE)
        .V1 = P1
        .V2 = P2
        .Boundary = IsBoundary
        .MainLength = Distance(P(P2).X - P(P1).X, P(P2).Y - P(P1).Y)
    End With

    ComputeMasses

    NotMovable = False

End Sub

Public Sub UpDateForces()
    Dim I              As Long

    For I = 1 To NP
        If I > 2 Then
            'Rear Wheel
            P(I).AccX = GravityX    '+ Cos(Angle) * GAS
            P(I).AccY = GravityY    '+ Sin(Angle) * GAS
        Else
            'Front Wheel
            P(I).AccX = GravityX
            P(I).AccY = GravityY
        End If
    Next



End Sub

Public Sub DRAW(PicHdc As Long)
    Dim I              As Long
    Dim X1             As Long
    Dim Y1             As Long
    Dim X2             As Long
    Dim Y2             As Long

    Dim pts()          As POINTAPI
    Dim ptsMin()       As POINTAPI
    Dim PolyRegion     As Long
    Dim Brush          As Long

    Dim CosA           As Single
    Dim SinA           As Single

    Dim LI             As Long
    Dim VL             As Long



    '    For I = 1 To NE
    '        With E(I)
    '               If .Boundary Then
    '                X1 = P(.V1).drawx \ 1
    '                Y1 = P(.V1).drawy \ 1
    '                X2 = P(.V2).drawx \ 1
    '                Y2 = P(.V2).drawy \ 1
    '                FastLine PicHdc, X1, Y1, X2, Y2, 1, color    'vbWhite
    '                '            Else
    '                'X1 = P(.V1).drawx \ 1
    '                'Y1 = P(.V1).drawy \ 1
    '                'X2 = P(.V2).drawx \ 1
    '                'Y2 = P(.V2).drawy \ 1
    '                '
    '                '                FastLine PicHdc, X1, Y1, X2, Y2, 1, vbBlue
    '            End If
    '        End With
    '    Next

    For I = 1 To NP
        P(I).DrawX = P(I).X + CamX
        P(I).DrawY = P(I).Y + CamY
    Next


    If HasWheels Then

        'DrawWheels
        CosA = Cos(Angle + Steer) * 4
        SinA = Sin(Angle + Steer) * 4
        FastLine PicHdc, P(1).DrawX - CosA, P(1).DrawY - SinA, P(1).DrawX + CosA, P(1).DrawY + SinA, 4, vbBlack    'RGB(150, 150, 150)
        FastLine PicHdc, P(2).DrawX - CosA, P(2).DrawY - SinA, P(2).DrawX + CosA, P(2).DrawY + SinA, 4, vbWhite, False
        CosA = Cos(Angle) * 4
        SinA = Sin(Angle) * 4
        FastLine PicHdc, P(3).DrawX - CosA, P(3).DrawY - SinA, P(3).DrawX + CosA, P(3).DrawY + SinA, 4, vbWhite, False
        FastLine PicHdc, P(4).DrawX - CosA, P(4).DrawY - SinA, P(4).DrawX + CosA, P(4).DrawY + SinA, 4, vbWhite, False

    End If

    '*****Draw Chassis
    ReDim pts(0 To NP - 1)
    ReDim ptsMin(0 To NP - 1)
    For I = 1 To NP
        pts(I - 1).X = P(I).DrawX \ 1
        pts(I - 1).Y = P(I).DrawY \ 1
        ptsMin(I - 1).X = P(I).X \ MD
        ptsMin(I - 1).Y = P(I).Y \ MD
    Next
    'Polygon PicHdc, pts(0), NP
    PolyRegion = CreatePolygonRgn(pts(0), NP, 1&)
    Brush = CreateSolidBrush(color)
    FillRgn PicHdc, PolyRegion, Brush
    DeleteObject PolyRegion
    '**********

    'Draw Minatures
    If HasWheels Then
        'MyCircle PicHdc, CenterX \ MD, CenterY \ MD, 2, 2, color
        FastLine PicHdc, ((P(1).X + P(2).X) * 0.5 \ MD), ((P(1).Y + P(2).Y) * 0.5 \ MD), _
                 ((P(3).X + P(4).X) * 0.5 \ MD), ((P(3).Y + P(4).Y) * 0.5 \ MD), 2, color, True
    Else
        PolyRegion = CreatePolygonRgn(ptsMin(0), NP, 1&)
        FillRgn PicHdc, PolyRegion, Brush
    End If
    '---
    DeleteObject PolyRegion
    DeleteObject Brush


    If HasWheels Then
        'Draw Glass
        '        ReDim pts(0 To 2) 'Useless
        '        pts(0).X = (P(1).DrawX + P(2).DrawX) * 0.5
        '        pts(0).Y = (P(1).DrawY + P(2).DrawY) * 0.5
        pts(1).X = (P(2).DrawX + P(3).DrawX) * 0.5
        pts(1).Y = (P(2).DrawY + P(3).DrawY) * 0.5
        pts(2).X = (P(4).DrawX + P(1).DrawX) * 0.5
        pts(2).Y = (P(4).DrawY + P(1).DrawY) * 0.5

        pts(0).X = (P(1).DrawX + P(2).DrawX + pts(1).X + pts(2).X) * 0.25
        pts(0).Y = (P(1).DrawY + P(2).DrawY + pts(1).Y + pts(2).Y) * 0.25

        DeleteObject (SelectObject(PicHdc, CreatePen(vbSolid, 1, vbBlue)))
        Polygon PicHdc, pts(0), 3

        'If IsPolice Then
        '    VL = (30 - Speed) * 20
        '    LI = CNT Mod (VL)
        '    VL = VL * 0.5
        '    MyCircle PicHdc, pts(1).X, pts(1).Y, 3, 3, IIf(LI < VL, vbBlue, vbWhite)
        '    MyCircle PicHdc, pts(2).X, pts(2).Y, 3, 3, IIf(LI > VL, vbBlue, vbWhite)
        'End If
    End If

End Sub




Private Sub ComputeMasses()


    Dim I              As Long

    CalculateCenter

    Area = 0
    For I = 1 To NP - 1
        Area = Area + ((P(I).X - MinX) * (P(I + 1).Y - MinY) _
                       - (P(I + 1).X - MinX) * (P(I).Y - MinY))
    Next
    Area = Area * 0.5

    Perim = 0
    For I = 1 To NE
        If E(I).Boundary Then Perim = Perim + E(I).MainLength
    Next


    For I = 1 To NP
        'P(I).Mass = Area / NP
        'P(I).Mass = Perim / NP

        P(I).Mass = Area
    Next

End Sub

Public Sub KeepInWorld()
    Dim I              As Long
    If NotMovable Then Exit Sub

    For I = 1 To NP
        With P(I)
            .X = MathMAX(MathMIN(.X, WorldW), 0)
            .Y = MathMAX(MathMIN(.Y, WorldH), 0)
        End With
    Next

End Sub
Public Sub cAccellerate()
    If Damage < 1 Then GAS = GAS + 0.005 * (1 - Damage)
    If GAS > 0.05 Then GAS = 0.05

End Sub
Public Sub cBrake()


    If GAS < 0 Then
        GAS = GAS - 0.02
        If GAS < -0.025 Then GAS = -0.025
    Else
        GAS = GAS - 0.05
    End If



End Sub
Public Sub cDoSteer(V As Single)
    Steer = Steer + V * 0.035    ' 0.045    ' 0.03
    If Abs(Steer) > MaxSteer Then Steer = Sgn(Steer) * MaxSteer
End Sub

Public Sub DoWheels()
    Dim X              As Single
    Dim Y              As Single
    Dim I              As Long


    X = P(1).X - P(4).X + P(2).X - P(3).X
    Y = P(1).Y - P(4).Y + P(2).Y - P(3).Y

    Angle = Atan2(X, Y)

    '    For I = 1 To 2
    '        With P(I)
    '            .X = .X + Cos(Angle + Steer) * GAS
    '            .Y = .Y + Sin(Angle + Steer) * GAS
    '        End With
    '    Next
    For I = 3 To NP
        With P(I)
            .X = .X + Cos(Angle) * GAS * 2
            .Y = .Y + Sin(Angle) * GAS * 2
        End With
    Next


    GAS = GAS * 0.95
    Steer = Steer * 0.9

    GoingForward = IIf(GAS >= 0, True, False)

End Sub

Private Sub LinesIntersect(aX1 As Single, aY1 As Single, Ax2 As Single, Ay2 As Single, _
                           Bx1 As Single, By1 As Single, Bx2 As Single, By2 As Single, ByRef RetX As Single, ByRef RetY As Single)

    Dim R              As Single
    Dim S              As Single
    Dim D              As Single

    Dim AX2mAX1        As Single
    Dim AY2mAY1        As Single

    AX2mAX1 = (Ax2 - aX1)
    AY2mAY1 = (Ay2 - aY1)
    RetX = -999
    RetY = -999
    D = AX2mAX1 * (By2 - By1) - AY2mAY1 * (Bx2 - Bx1)

    If D <> 0 Then

        R = ((aY1 - By1) * (Bx2 - Bx1) - (aX1 - Bx1) * (By2 - By1)) / D

        If R > 0 And R < 1 Then

            S = ((aY1 - By1) * AX2mAX1 - (aX1 - Bx1) * AY2mAY1) / D

            If S < 1 Then
                If S > 0 Then
                    RetX = aX1 + R * AX2mAX1
                    RetY = aY1 + R * AY2mAY1
                End If
            End If

        End If

    End If



End Sub



Public Sub CheckAndRestoreFlipped()
'Works with Body of only 4 points

    Dim I              As Long
    Dim J              As Long

    Dim Inter          As Boolean

    Dim iX1            As Single
    Dim iY1            As Single
    Dim iX2            As Single
    Dim iY2            As Single
    Dim jX1            As Single
    Dim jY1            As Single
    Dim jX2            As Single
    Dim jY2            As Single

    Dim RX             As Single
    Dim RY             As Single

    If NP <> 4 Then Exit Sub
    If NotMovable Then Exit Sub

    For I = 1 To NE - 1
        If E(I).Boundary Then
            iX1 = P(E(I).V1).X
            iY1 = P(E(I).V1).Y
            iX2 = P(E(I).V2).X
            iY2 = P(E(I).V2).Y

            For J = I + 1 To NE
                If E(J).Boundary Then
                    jX1 = P(E(J).V1).X
                    jY1 = P(E(J).V1).Y
                    jX2 = P(E(J).V2).X
                    jY2 = P(E(J).V2).Y

                    LinesIntersect iX1, iY1, iX2, iY2, _
                                   jX1, jY1, jX2, jY2, RX, RY

                    If RX <> -999 Then Inter = True: Exit For
                End If
            Next

        End If

    Next

    If Inter Then

        P(1).X = CenterX - (E(1).MainLength / 2)    '* Cos(Angle)
        P(1).Y = CenterY - (E(2).MainLength / 2)    '* Sin(Angle)

        P(2).X = P(1).X + E(1).MainLength
        P(2).Y = P(1).Y
        P(3).X = P(2).X
        P(3).Y = P(2).Y + E(2).MainLength

        P(4).X = P(1).X
        P(4).Y = P(3).Y
        For I = 1 To NP
            P(I).OldX = P(I).X
            P(I).OldY = P(I).Y
        Next

    End If

End Sub

Public Sub SetNotMovable()
    Dim I              As Long

    NotMovable = True
    For I = 1 To NP
        P(I).Mass = InfiniteMASS
    Next
End Sub


Private Function OnePointInScreen() As Boolean
    Dim I              As Long

    OnePointInScreen = False
    For I = 1 To NP
        With P(I)
            If .DrawX > 0 Then
                If .DrawY > 0 Then
                    If .DrawX < ScreenW Then
                        If .DrawY < ScreenH Then
                            OnePointInScreen = True: Exit For
                        End If
                    End If
                End If
            End If
        End With
    Next
End Function
