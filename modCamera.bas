Attribute VB_Name = "modCamera"
Public CamX            As Single
Public CamY            As Single


Public Sub UpDateCamera()

    With B(1)

        CamX = -.CenterX + ScreenW \ 2    '- Cos(.Angle) * .GAS * 1025
        CamY = -.CenterY + ScreenH \ 2    '- Sin(.Angle) * .GAS * 1025

        'CamX = CamX - Cos(B(1).Angle) * B(1).Speed
        'CamY = CamY - Sin(B(1).Angle) * B(1).Speed
    End With


End Sub
