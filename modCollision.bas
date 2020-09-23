Attribute VB_Name = "modCollision"
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


Option Explicit


Public Type TCollisionInfo

    Depth              As Single
    NormalX            As Single
    NormalY            As Single
    E                  As tEdge
    P                  As tPoint

    OE                 As Long    'BodyObject on Edge
    OP                 As Long    'BodyObject on Point

    WichPointOP        As Long    'wich Point on OP
    WichEdgeOE         As Long    'wich Edges on OE

End Type

Public CollisionInfo   As TCollisionInfo


Public Function bodiesOverlap(ByRef B1 As Long, ByRef B2 As Long) As Boolean

'return ( B1->MinX <= B2->MaxX ) && ( B1->MinY <= B2->MaxY ) && ( B1->MaxX >= B2->MinX ) && ( B2->MaxY >= B1->MinY );

    bodiesOverlap = (B(B1).MinX <= B(B2).MaxX) And _
                    (B(B1).MinY <= B(B2).MaxY) And _
                    (B(B1).MaxX >= B(B2).MinX) And _
                    (B(B2).MaxY >= B(B1).MinY)
End Function


Public Function bodiesOverlapMY(ByRef B1 As Long, ByRef B2 As Long) As Boolean
'**** Roberto Mior
'**** With bodiesOverlap[MY] it's useless to
'**** do [for b1=1 to NB] and [for b2=1 to NB]
'**** in IterateCollision SUB

    Dim T1             As Boolean
    Dim T2             As Boolean

    'return ( B1->MinX <= B2->MaxX ) && ( B1->MinY <= B2->MaxY ) && ( B1->MaxX >= B2->MinX ) && ( B2->MaxY >= B1->MinY );

    T1 = (B(B1).MinX <= B(B2).MaxX) And _
         (B(B1).MinY <= B(B2).MaxY) And _
         (B(B1).MaxX >= B(B2).MinX) And _
         (B(B2).MaxY >= B(B1).MinY)

    If Not (T1) Then
        T2 = (B(B2).MinX <= B(B1).MaxX) And _
             (B(B2).MinY <= B(B1).MaxY) And _
             (B(B2).MaxX >= B(B1).MinX) And _
             (B(B1).MaxY >= B(B2).MinY)

        bodiesOverlapMY = T2
    Else
        bodiesOverlapMY = True
    End If

End Function
Public Function DetectCollision(ByVal B1 As Long, ByVal B2 As Long) As Boolean

    Dim I              As Long
    Dim MinDistance    As Single
    Dim Distance       As Single
    Dim SmallestD      As Single

    Dim E              As tEdge
    Dim AxisX          As Single
    Dim AxisY          As Single

    Dim EV1            As tPoint
    Dim EV2            As tPoint
    Dim MinA           As Single
    Dim MinB           As Single
    Dim MaxA           As Single
    Dim MaxB           As Single

    Dim OE             As Long
    Dim OP             As Long

    Dim T              As Long

    Dim xx             As Single
    Dim YY             As Single
    Dim mULT           As Single
    Dim II             As Long

    Dim TotEdges       As Long


    MinDistance = 10000

    DetectCollision = False

    TotEdges = B(B1).NE + B(B2).NE

    For I = 1 To TotEdges

        '  for( int I = 0; I < B1->EdgeCount + B2->EdgeCount; I++ ) { //Same old
        '    Edge* E;''
        '
        '    if( I < B1->EdgeCount )
        '      E = B1->Edges[ I ];
        '    Else
        '      E = B2->Edges[ I - B1->EdgeCount ];'


        If I <= B(B1).NE Then
            II = I
            E.V1 = B(B1).getEdgeV1(I)
            E.V2 = B(B1).getEdgeV2(I)
            E.Boundary = B(B1).getEdgeIsBoundary(I)

            EV1.X = B(B1).getPointX(E.V1)
            EV1.Y = B(B1).getPointY(E.V1)
            EV2.X = B(B1).getPointX(E.V2)
            EV2.Y = B(B1).getPointY(E.V2)
            OE = B1
            OP = B2
        Else
            II = I - B(B1).NE
            E.V1 = B(B2).getEdgeV1(II)
            E.V2 = B(B2).getEdgeV2(II)
            E.Boundary = B(B2).getEdgeIsBoundary(II)

            EV1.X = B(B2).getPointX(E.V1)
            EV1.Y = B(B2).getPointY(E.V1)
            EV2.X = B(B2).getPointX(E.V2)
            EV2.Y = B(B2).getPointY(E.V2)
            OE = B2
            OP = B1
        End If

        If E.Boundary = False Then GoTo ContinueNext





        'axis.x = e.v1.position.y - e.v2.position.y;
        'axis.y = e.v2.position.x - e.v1.position.x;

        AxisX = EV1.Y - EV2.Y
        AxisY = EV2.X - EV1.X

        Normalize AxisX, AxisY

        B(B1).ProjectToAxis AxisX, AxisY, MinA, MaxA
        B(B2).ProjectToAxis AxisX, AxisY, MinB, MaxB

        Distance = IntervalDistance(MinA, MaxA, MinB, MaxB)

        If Distance > 0 Then Exit Function
        If (Abs(Distance) < MinDistance) Then
            MinDistance = Abs(Distance)

            CollisionInfo.NormalX = AxisX
            CollisionInfo.NormalY = AxisY
            CollisionInfo.E = E
            CollisionInfo.OE = OE
            CollisionInfo.OP = OP
            CollisionInfo.WichEdgeOE = II
        End If

ContinueNext:

    Next

    CollisionInfo.Depth = MinDistance

    '//Ensure that the body containing the collision edge lies in
    '//B2 and the one containing the collision vertex in B1
    If (CollisionInfo.OE <> B2) Then

        T = B2
        B2 = B1
        B1 = T

        CollisionInfo.OE = B2
        CollisionInfo.OP = B1

    End If


    '//This is needed to make sure that the collision normal is pointing at B1
    'int Sign = SGN( CollisionInfo.Normal*( B1->Center - B2->Center ) );

    '//Remember that the line equation is N*( R - R0 ). We choose B2->Center
    '//as R0; the normal N is given by the collision normal

    'if( Sign != 1 )
    'CollisionInfo.Normal = -CollisionInfo.Normal; //Revert the collision normal if it points away from B1


    '// int Sign = SGN( CollisionInfo.Normal.multiplyVal( B1.Center.minus(B2.Center) ) ); //This is needed to make sure that the collision normal is pointing at B1
    'float xx = b1.center.x - b2.center.x;
    'float yy = b1.center.y - b2.center.y;
    'float mult = CollisionInfo.normal.x * xx + CollisionInfo.normal.y * yy;
    xx = B(B1).CenterX - B(B2).CenterX
    YY = B(B1).CenterY - B(B2).CenterY
    mULT = CollisionInfo.NormalX * xx + CollisionInfo.NormalY * YY

    '  // Remember that the line equation is N*( R - R0 ). We choose B2->Center as R0; the normal N is given by the collision normal
    'if (mult < 0) {
    '// Revert the collision normal if it points away from B1
    'CollisionInfo.normal.x = 0-CollisionInfo.normal.x;
    'CollisionInfo.normal.y = 0-CollisionInfo.normal.y;
    '}

    If mULT < 0 Then
        CollisionInfo.NormalX = -CollisionInfo.NormalX
        CollisionInfo.NormalY = -CollisionInfo.NormalY
    End If


    SmallestD = 10000000     ' //Initialize the smallest distance to a high value

    For I = 1 To B(B1).NP



        xx = B(B1).getPointX(I) - B(B2).CenterX
        YY = B(B1).getPointY(I) - B(B2).CenterY
        Distance = CollisionInfo.NormalX * xx + CollisionInfo.NormalY * YY
        If Distance < SmallestD Then
            SmallestD = Distance

            With B(B1)
                CollisionInfo.P.X = .getPointX(I)
                CollisionInfo.P.Y = .getPointY(I)
                CollisionInfo.WichPointOP = I
            End With

        End If


    Next

    DetectCollision = True

End Function


Public Sub IterateCollisions()
'Optimized by Roberto Mior

    Dim I              As Long

    Dim BB             As Long
    Dim B1             As Long
    Dim B2             As Long

    Dim Iterations     As Long
    Dim NBm1           As Long

    NBm1 = NB - 1

    Iterations = 2           '2           '2'3          '10



    For BB = 1 To NB
        With B(BB)
            '        '*****************************************
            '        '** In Origin these were inside the Iteration Cycle
            '        '** but I decided to put them outside for increase Computing Speed
            '        .KeepInWorld
            '        '.CalculateCenter
            '        '*****************************************
        End With
    Next

    For I = 1 To Iterations

        For BB = 1 To NB

            With B(BB)
                '.KeepInWorld
                .UpDateEdges
                .CalculateCenter
            End With

        Next

        For B1 = 1 To NBm1
            For B2 = B1 + 1 To NB
                If Not (B(B1).NotMovable And B(B2).NotMovable) Then

                    'If B1 <> B2 Then
                    '**** With bodiesOverlap[MY] it's useless to
                    '**** do [for b1=1 to NB] and [for b2=1 to NB]
                    If bodiesOverlapMY(B1, B2) Then
                        If DetectCollision(B1, B2) Then


                            B(CollisionInfo.OE).Damage = B(CollisionInfo.OE).Damage + CollisionInfo.Depth * 0.005
                            B(CollisionInfo.OP).Damage = B(CollisionInfo.OP).Damage + CollisionInfo.Depth * 0.005

                            B(CollisionInfo.OE).GAS = B(CollisionInfo.OE).GAS * 0.5
                            B(CollisionInfo.OP).GAS = B(CollisionInfo.OP).GAS * 0.5

                            ProcessCollision

                        End If
                    End If
                    'End If
                End If
            Next
        Next
    Next

End Sub


Public Sub ProcessCollision()

    Dim collisionVectorX As Single
    Dim collisionVectorY As Single
    Dim T              As Single
    Dim Lambda         As Single
    Dim EdgeMass       As Single
    Dim InvCollisionMass As Single
    Dim Ratio1         As Single
    Dim Ratio2         As Single


    Dim OEV1Px         As Single
    Dim OEV1Py         As Single
    Dim OEV2Px         As Single
    Dim OEV2Py         As Single
    Dim OpPx           As Single
    Dim OpPy           As Single
    Dim PointMASS      As Single

    Dim OEV1mass       As Single
    Dim OEV2mass       As Single
    Dim OPmass         As Single


    PointMASS = 0.1

    With CollisionInfo

        OEV1Px = B(.OE).getPointX(.E.V1)
        OEV1Py = B(.OE).getPointY(.E.V1)
        OEV2Px = B(.OE).getPointX(.E.V2)
        OEV2Py = B(.OE).getPointY(.E.V2)


        OEV1mass = B(.OE).getMASS(B(.OE).getEdgeV1(.WichEdgeOE))
        OEV2mass = B(.OE).getMASS(B(.OE).getEdgeV2(.WichEdgeOE))
        OPmass = B(.OP).getMASS(.WichPointOP)

        'OEV1Px = B(.OE).getPointX(B(.OE).getEdgeV1(.WichEdgeOE))
        'OEV1Py = B(.OE).getPointY(B(.OE).getEdgeV1(.WichEdgeOE))
        'OEV2Px = B(.OE).getPointX(B(.OE).getEdgeV2(.WichEdgeOE))
        'OEV2Py = B(.OE).getPointY(B(.OE).getEdgeV2(.WichEdgeOE))

        '*** DEBUG
        'FastLine frmMAIN.PIC.Hdc, OEV1Px \ 1, OEV1Py \ 1, OEV2Px \ 1, OEV2Py \ 1, 3, vbRed
        'MyCircle frmMAIN.PIC.Hdc, .P.X \ 1, .P.Y \ 1, 4, 2, vbRed
        'frmMAIN.PIC.Refresh
        'DoEvents
        '******

        'OpPx = B(.OP).getPointX(.WichPointOP)
        'OpPy = B(.OP).getPointY(.WichPointOP)
        'E1.Position.X = CollisionInfo.E.V1x
        'E1.Position.Y = CollisionInfo.E.V1y
        'E2.Position.X = CollisionInfo.E.V2x
        'E2.Position.Y = CollisionInfo.E.V2y
    End With

    collisionVectorX = CollisionInfo.NormalX * CollisionInfo.Depth
    collisionVectorY = CollisionInfo.NormalY * CollisionInfo.Depth

    If (Abs(OEV1Px - OEV2Px) > Abs(OEV1Py - OEV2Py)) Then
        T = (CollisionInfo.P.X - collisionVectorX - OEV1Px) / (OEV2Px - OEV1Px)
    Else
        T = (CollisionInfo.P.Y - collisionVectorY - OEV1Py) / (OEV2Py - OEV1Py)
    End If

    Lambda = 1 / (T * T + (1 - T) * (1 - T))
    'edgeMass = t*e2.parent.mass + ( 1f - t )*e1.parent.mass; //Calculate the mass at the intersection point

    '    EdgeMass = T * PointMASS + (1 - T) * PointMASS    ' //Calculate the mass at the intersection point
    EdgeMass = T * OEV2mass + (1 - T) * OEV1mass    ' //Calculate the mass at the intersection point


    'invCollisionMass = 1.0f/( edgeMass + CollisionInfo.v.parent.mass );
    'InvCollisionMass = 1 / (EdgeMass + PointMASS)
    InvCollisionMass = 1 / (EdgeMass + OPmass)


    'ratio1 = CollisionInfo.v.parent.mass*invCollisionMass;
    'ratio2 = edgeMass*invCollisionMass;

    'Ratio1 = PointMASS * InvCollisionMass
    Ratio1 = OPmass * InvCollisionMass
    Ratio2 = EdgeMass * InvCollisionMass


    Ratio1 = Ratio1 * Lambda

    OEV1Px = OEV1Px - collisionVectorX * ((1 - T) * Ratio1)
    OEV1Py = OEV1Py - collisionVectorY * ((1 - T) * Ratio1)
    OEV2Px = OEV2Px - collisionVectorX * (T * Ratio1)
    OEV2Py = OEV2Py - collisionVectorY * (T * Ratio1)

    CollisionInfo.P.X = CollisionInfo.P.X + collisionVectorX * Ratio2
    CollisionInfo.P.Y = CollisionInfo.P.Y + collisionVectorY * Ratio2

    With CollisionInfo

        B(.OE).SetPointX(.E.V1) = OEV1Px
        B(.OE).SetPointY(.E.V1) = OEV1Py
        B(.OE).SetPointX(.E.V2) = OEV2Px
        B(.OE).SetPointY(.E.V2) = OEV2Py

        B(.OP).SetPointX(.WichPointOP) = CollisionInfo.P.X
        B(.OP).SetPointY(.WichPointOP) = CollisionInfo.P.Y

    End With

End Sub

