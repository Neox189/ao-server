Option Explicit

Public MAX_VIAJES As Byte

Private Type tViajes
    Nombre As String
    Nivel As Byte
    Costo As Long
    Tiempo As Integer
    Pos As WorldPos
    SkillNavegacion As Byte
End Type

Public Viajes() As tViajes

Public Sub CargarViajes()

    Dim Leer As clsIniManager
    Set Leer = New clsIniManager

    Call Leer.Initialize(DatPath & "Viajes.dat")

    MAX_VIAJES = Val(Leer.GetValue("INIT", "MAX_VIAJES"))

    If MAX_VIAJES > 0 Then

        Dim LoopC As Long, Cadena As String
        ReDim Viajes(1 To MAX_VIAJES) As tViajes

        For LoopC = 1 To MAX_VIAJES

            '++ Funciona 10 pts xD
            With Viajes(LoopC)

                .Nombre = CStr(Leer.GetValue("VIAJE#" & LoopC, "Nombre"))
                .Nivel = Val(Leer.GetValue("VIAJE#" & LoopC, "Nivel"))

                If .Nivel < 1 Then
                    .Nivel = 1
                    Call MsgBox("El viaje n° " & LoopC & " estaba dateado con un nivel menor a (1) revisar el VIAJES.DAT" & vbNewLine & "Se ha seteado el nivel minimo en respectivo viaje.")
                End If

                If .Nivel > STAT_MAXELV Then
                    .Nivel = STAT_MAXELV
                    Call MsgBox("Exediste el limite de nivel (" & STAT_MAXELV & ") en el viaje n° " & LoopC & " revisar el VIAJES.DAT" & vbNewLine & "Se ha seteado el nivel maximo en respectivo viaje.")
                End If

                .Costo = Val(Leer.GetValue("VIAJE#" & LoopC, "Costo"))
                .Tiempo = Val(Leer.GetValue("VIAJE#" & LoopC, "Tiempo"))
                .SkillNavegacion = Val(Leer.GetValue("VIAJE#" & LoopC, "SkillReq"))

                If .SkillNavegacion > MAXSKILLPOINTS Then
                    .SkillNavegacion = MAXSKILLPOINTS
                    Call MsgBox("Exediste el limite skill de navegacion en el viaje n° " & LoopC & " revisar el VIAJES.DAT" & vbNewLine & "Se ha seteado el skill en 100 en respectivo viaje.")
                End If

                Cadena = CStr(Leer.GetValue("VIAJE#" & LoopC, "Pos"))

                If LenB(Cadena) > 5 Then
                    .Pos.map = Val(ReadField(1, Cadena, 45))

                    If .Pos.map < 1 Then
                        .Pos.map = 1
                        Call MsgBox("El viaje n° " & LoopC & " estaba dateado con un num de mapa menor a (1) revisar el VIAJES.DAT" & vbNewLine & "Se ha seteado el mapa minimo en respectivo viaje.")
                    End If

                    If .Pos.map > NumMaps Then
                        .Pos.map = NumMaps
                        Call MsgBox("Exediste el limite de mapas en el viaje n° " & LoopC & " revisar el VIAJES.DAT" & vbNewLine & "Se ha seteado el mapa maximo en respectivo viaje.")
                    End If

                    .Pos.X = Val(ReadField(2, Cadena, 45))
                    .Pos.Y = Val(ReadField(3, Cadena, 45))

                End If

            End With

        Next LoopC

    End If

    Set Leer = Nothing

End Sub

Private Function PuedeViajar(ByVal UserIndex As Integer) As Boolean

    PuedeViajar = False

    With UserList(UserIndex)

        If .flags.Muerto > 0 Then
            Call WriteConsoleMsg(UserIndex, "No puedes viajar estando muerto.", FontTypeNames.FONTTYPE_INFO)
            Exit Function
        End If

        If .flags.Comerciando > 0 Then
            Call WriteConsoleMsg(UserIndex, "No puedes viajar estando comerciando.", FontTypeNames.FONTTYPE_INFO)
            Exit Function
        End If

        If .flags.invisible > 0 Then
            Call WriteConsoleMsg(UserIndex, "No puedes viajar estando invisible.", FontTypeNames.FONTTYPE_INFO)
            Exit Function
        End If

        If .flags.Navegando > 0 Then
            Call WriteConsoleMsg(UserIndex, "No puedes viajar estando navegando.", FontTypeNames.FONTTYPE_INFO)
            Exit Function
        End If

        If .flags.Mimetizado > 0 Then
            Call WriteConsoleMsg(UserIndex, "No puedes viajar estando trasformado.", FontTypeNames.FONTTYPE_INFO)
            Exit Function
        End If

    End With

    PuedeViajar = True

End Function

Public Sub Viajar(ByVal UserIndex As Integer, ByVal ID As Byte)

    If Not PuedeViajar(UserIndex) Then Exit Sub

    ' ++ Anti salchichas
    If ID < 0 Then ID = 0
    If ID > MAX_VIAJES Then ID = MAX_VIAJES

    With Viajes(ID)

        ' ++ Requisito de nivel
        If UserList(UserIndex).Stats.ELV < .Nivel Then
            Call WriteConsoleMsg(UserIndex, "Debes ser nivel " & .Nivel & " para viajar.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

        ' ++ Requisito de oro
        If UserList(UserIndex).Stats.GLD < .Costo Then
            Call WriteConsoleMsg(UserIndex, "Debes tener " & .Costo & " monedas de oro para viajar.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

        ' ++ Requisito de oro
        If UserList(UserIndex).Stats.UserSkills(eSkill.Navegacion) < .SkillNavegacion Then
            Call WriteConsoleMsg(UserIndex, "Debes tener " & .SkillNavegacion & " skills de navegacion.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

        If UserList(UserIndex).Counters.Viaje > 0 Then
            Call WriteConsoleMsg(UserIndex, "Ya te encuentras viajando. Intente mas tarde.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

        ' ++ Pense que lo habia puesto xdxd
        UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - .Costo
        Call WriteUpdateGold(UserIndex)

        ' ++ Si el viaje es con tiempo no lo warpeamos enseguida.
        If .Tiempo > 0 Then
            UserList(UserIndex).Counters.Viaje = .Tiempo
            UserList(UserIndex).PosViaje = .Pos
            Call WriteConsoleMsg(UserIndex, "Has emprendido tu viaje hacia " & .Nombre & " llegaras en " & .Tiempo & " segundos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

        ' ++ Era sin tiempo de retardo
        Call WarpUserCharX(UserIndex, .Pos.map, .Pos.X, .Pos.Y, True)
        Call WriteConsoleMsg(UserIndex, "Has llegado a tu destino." & vbNewLine & "Fue un largo viaje, aliméntate y bebe algo para recuperarte", FontTypeNames.FONTTYPE_INFO)
        Call UpdateStatsViaje(UserIndex)

    End With

End Sub

Public Sub UpdateStatsViaje(ByVal UserIndex As Integer)

    With UserList(UserIndex)
        .Stats.MinAGU = 0
        .Stats.MinHam = 0
        .Stats.MinSta = 0
    End With

    Call WriteUpdateHungerAndThirst(UserIndex)
    Call WriteUpdateSta(UserIndex)

End Sub

Public Sub WarpUserCharX(ByVal UserIndex As Integer, ByVal Mapa As Integer, ByVal X As Integer, ByVal Y As Integer, Optional ByVal FX As Boolean = False)

    If MapData(Mapa, X, Y).UserIndex = UserIndex Then
        Exit Sub
    End If

    Dim NuevaPos As WorldPos
    Dim FuturePos As WorldPos

    FuturePos.map = Mapa
    FuturePos.X = X
    FuturePos.Y = Y

    Call ClosestLegalPos(FuturePos, NuevaPos, True)

    If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then
        Call WarpUserChar(UserIndex, NuevaPos.map, NuevaPos.X, NuevaPos.Y, FX)
    End If

End Sub