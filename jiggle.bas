' Nur für Windows – 64bit-kompatibel
Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Declare PtrSafe Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
Declare PtrSafe Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Type POINTAPI
    x As Long
    y As Long
End Type

Dim jiggleActive As Boolean
Dim jiggleDirection As Boolean

Sub StartVisibleMouseJiggler()
    jiggleActive = True
    jiggleDirection = True ' Start mit Bewegung nach unten

    Do While jiggleActive
        Dim currentPos As POINTAPI
        GetCursorPos currentPos

        If jiggleDirection Then
            SetCursorPos currentPos.x, currentPos.y + 10
        Else
            SetCursorPos currentPos.x, currentPos.y - 10
        End If

        jiggleDirection = Not jiggleDirection ' Richtung wechseln
        Sleep 1000 ' Alle 1 Sekunde bewegen
        DoEvents ' Rechner bleibt ansprechbar
    Loop
End Sub

Sub StopVisibleMouseJiggler()
    jiggleActive = False
End Sub
