Attribute VB_Name = "PeriodicTable"
Option Explicit
Const Super1 As String = "¹", Super2 As String = "²", Super3 As String = "³"

Public Sub WriteClipboard(Optional Text As String, Optional Filename As String)
    Dim tempstr As String, tempfile As Long
    tempstr = ParseClipboard(Text)
    If Len(tempstr) > 0 Then
        If Len(Filename) = 0 Then Filename = "C:\Users\Techni\Documents\VB4A\LCAR\ptoe.ini"
        tempfile = FreeFile
        Open Filename For Append As tempfile
            Print #tempfile, vbNewLine & tempstr
        Close tempfile
    End If
End Sub

Public Function ParseClipboard(Optional Text As String) As String
    Dim tempstr() As String, tempstr2() As String, tempstr3() As String, temp As Long, tempstr4 As String
    If Len(Text) = 0 Then Text = Clipboard.GetText
    tempstr = Split(Text, vbNewLine)
    For temp = 0 To UBound(tempstr)
        If Len(tempstr(temp)) > 0 Then
            tempstr2 = Split(tempstr(temp), vbTab)
            tempstr2(0) = Trim(tempstr2(0))
            If UBound(tempstr2) > 0 Then
                tempstr2(1) = Trim(RemoveBrackets(tempstr2(1)))  'Replace(Replace(tempstr2(1), "[", ""), "]", "")
                tempstr3 = Split(tempstr2(1), ", ")
                
                'Debug.Print tempstr2(0) & "=" & tempstr2(1)
                Select Case tempstr2(0)
                    Case "Name, symbol, number"
                        tempstr4 = "[" & tempstr3(2) & "]"
                        tempstr4 = MakeLine(tempstr4, "Name", CapFirstLetter(tempstr3(0)))
                        tempstr4 = MakeLine(tempstr4, "Symbol", tempstr3(1))
                    Case "Element category"
                        tempstr4 = MakeLine(tempstr4, "Category", tempstr2(1))
                    Case "Group, period, block"
                        tempstr4 = MakeLine(tempstr4, "Loc", tempstr3(0) & "," & tempstr3(1))
                    Case "Standard atomic weight"
                        tempstr4 = MakeLine(tempstr4, "Weight", tempstr2(1))
                    Case "Electron configuration"
                        tempstr4 = MakeLine(tempstr4, "Electron", HandleElectrons(tempstr2(1)))
                    Case "Electrons per shell"
                        tempstr4 = MakeLine(tempstr4, "PerShell", Replace(tempstr2(1), " (Image)", Empty))
                    Case "Color", "Phase", "Electronegativity", "Covalent radius", "Van der Waals radius", "Crystal structure", "CAS registry number", "Atomic radius", "Young's modulus", "Shear modulus", "Bulk modulus", "Mohs hardness", "Poisson ratio", "Vickers hardness", "Brinell hardness", "Electrical resistivity", "Band gap energy at 300 K"
                        tempstr4 = MakeLine(tempstr4, tempstr2(0), tempstr2(1))
                    Case "Density"
                        tempstr4 = MakeLine(tempstr4, "Density", tempstr2(1))
                        If Left(tempstr2(1), 1) = "(" And Right(tempstr2(1), 1) = ")" Then tempstr4 = tempstr4 & " " & RemoveBrackets(Trim(tempstr(temp + 1)))
                    Case "Density (near r.t.)"
                        tempstr2(1) = "(near r.t.) " & tempstr2(1)
                        tempstr4 = MakeLine(tempstr4, "Density", tempstr2(1))
                        If Left(tempstr2(1), 1) = "(" And Right(tempstr2(1), 1) = ")" Then tempstr4 = tempstr4 & " " & RemoveBrackets(Trim(tempstr(temp + 1)))
                    Case "Liquid density at m.p."
                        tempstr4 = MakeLine(tempstr4, "AtMP", Replace(tempstr2(1), "cm-3", "cm-" & Super3))
                    Case "Liquid density at b.p."
                        tempstr4 = MakeLine(tempstr4, "AtBP", Replace(tempstr2(1), "cm-3", "cm-" & Super3))
                    Case "Melting point", "Boiling point", "Critical point", "Sublimation point"
                        If InStr(tempstr3(0), "(") = 0 Then
                            tempstr4 = MakeLine(tempstr4, tempstr2(0), tempstr3(0))
                        Else
                            tempstr4 = MakeLine(tempstr4, tempstr2(0), tempstr2(1))
                        End If
                    Case "Heat of fusion", "Heat of vaporization", "Molar heat capacity", "Magnetic ordering", "Thermal conductivity", "Speed of sound", "Thermal expansion", "Speed of sound (thin rod)"
                        tempstr4 = MakeLine(tempstr4, tempstr2(0), Replace(Replace(Replace(Replace(RemoveBrackets(tempstr2(1), "(", ")"), "mol-1", "mol-" & Super1), "K-1", "K-" & Super1), "m-1", "m-" & Super1), "s-1", "s-" & Super1))
                    Case "Oxidation states"
                        tempstr4 = MakeLine(tempstr4, tempstr2(0), Replace(GetAllLines(tempstr, temp, tempstr2(1)), "mol-1", "mol-" & Super1))
                    Case "(more)", "Ionization energies"
                        tempstr4 = MakeLine(tempstr4, "Ionization energies", Replace(GetAllLines(tempstr, temp, tempstr2(1)), "mol-1", "mol-" & Super1))
                End Select
            End If
        End If
    Next
    ParseClipboard = tempstr4
End Function
Public Function MakeLine(tempstr As String, Key As String, Value As String) As String
    MakeLine = tempstr & vbNewLine & Key & "=" & Trim(RemoveBrackets(Value))
End Function
Public Function CapFirstLetter(Name As String) As String
    CapFirstLetter = UCase(Left(Name, 1)) & LCase(Right(Name, Len(Name) - 1))
End Function
Public Function HandleElectrons(Electrons As String) As String
    Electrons = Replace(Electrons, "s1", "s" & Super1)
    Electrons = Replace(Electrons, "s2", "s" & Super2)
    Electrons = Replace(Electrons, "s3", "s" & Super3)
    Electrons = Replace(Electrons, "p1", "p" & Super1)
    Electrons = Replace(Electrons, "p2", "p" & Super2)
    Electrons = Replace(Electrons, "p3", "p" & Super3)
    HandleElectrons = Electrons
End Function

Public Function RemoveBrackets(Text As String, Optional L As String = "[", Optional R As String = "]") As String
    Dim temp As Long, temp2 As Long
    temp = InStr(Text, L)
    Do While temp > 0
        If temp > 0 Then
            temp2 = InStr(temp + 1, Text, R)
            Text = Left(Text, temp - 1) & Right(Text, Len(Text) - temp2)
            temp = InStr(Text, L)
        End If
    Loop
    RemoveBrackets = Text
End Function
Public Function GetAllLines(tempstr, ByVal temp As Long, ByRef Text As String) As String
    Dim temp2 As Long
    For temp2 = temp + 1 To UBound(tempstr)
        If InStr(tempstr(temp2), vbTab) Then
            Exit For
        Else
            Text = Text & " " & tempstr(temp2)
        End If
    Next
    GetAllLines = Text
End Function
