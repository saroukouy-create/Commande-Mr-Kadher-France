Attribute VB_Name = "modUtil"
Option Explicit

' Mode silencieux (désactive les MsgBox) - utilisé par les tests headless
' pour piloter la macro d'import sans boîte de dialogue bloquante.
Public TestMode As Boolean

Public Sub SetTestMode(ByVal v As Boolean)
    TestMode = v
End Sub

Public Sub NotifyUser(ByVal msg As String, ByVal buttons As VbMsgBoxStyle, ByVal title As String)
    If Not TestMode Then MsgBox msg, buttons, title
End Sub

' Convertit une valeur en nombre, tolère le format français ("3 200,50" -> 3200.5)
' et la notation scientifique - reproduit E5Controller::convertToFloat
Public Function ParseFrenchNumber(ByVal v As Variant) As Double
    Dim s As String
    Dim c As Long
    Dim cleaned As String
    Dim ch As String

    If IsNumeric(v) And Not (VarType(v) = vbString) Then
        ParseFrenchNumber = CDbl(v)
        Exit Function
    End If

    s = CStr(v)
    cleaned = ""
    For c = 1 To Len(s)
        ch = Mid$(s, c, 1)
        If (ch >= "0" And ch <= "9") Or ch = "-" Or ch = "." Or ch = "," Or LCase$(ch) = "e" Or ch = "+" Then
            If ch = "," Then
                cleaned = cleaned & "."
            Else
                cleaned = cleaned & ch
            End If
        End If
    Next c

    If Len(cleaned) = 0 Or Not IsNumeric(cleaned) Then
        ParseFrenchNumber = 0
    Else
        ParseFrenchNumber = CDbl(cleaned)
    End If
End Function

' Lit la valeur d'une plage nommée via sa cellule (1,1) plutôt que .Value directement.
' Nécessaire pour les plages fusionnées : lire .Value sur une Range multi-cellules
' provoque un blocage lorsque la macro est invoquée via Application.Run externe
' (automatisation COM) - lire via .Cells(1,1) évite ce blocage et reste correct
' pour les cellules simples (Cells(1,1) d'une plage à 1 cellule = la cellule elle-même).
Public Function ReadRangeValue(ByVal name As String) As Variant
    ReadRangeValue = ThisWorkbook.Names(name).RefersToRange.Cells(1, 1).Value
End Function

Public Function SafeRange(ByVal name As String) As Range
    Set SafeRange = ThisWorkbook.Names(name).RefersToRange
End Function
