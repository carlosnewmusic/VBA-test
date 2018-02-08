'Dim nombre
'Dim apellidos
'Dim via
'Dim calle
'Dim ciudad
'Dim privincia
Dim cp As Integer
'Dim fechanacimiento As Date
'Dim dni
'Dim sexo
Dim index
Dim valor
Dim msg
Dim persona()
'vaciar los spinner al cargar el formulario
'Tipo de Documento
Private Sub ComboBox1_Change()
    TextBox4.Visible = True
    Label17.Visible = True
End Sub
'Provincia
Private Sub ComboBox3_Change()
    TextBox5.Visible = True
    cp = ComboBox3.ListIndex + 1
    If (cp < 10) Then
        TextBox9.Value = "0" + Trim(Str(cp))
        Trim (TextBox9.Value)
    Else
        TextBox9.Value = cp
        Trim (TextBox9.Value)
    End If
End Sub
'Spinner de tipo de Via de direccion
Private Sub ComboBox2_Change()
    TextBox6.Visible = True
    TextBox7.Visible = True
    TextBox8.Visible = True
    TextBox9.Visible = True
    TextBox10.Visible = True
    TextBox11.Visible = True
    TextBox12.Visible = True
    TextBox99.Enabled = False
End Sub
'Boton de guardado
Private Sub CommandButton1_Click()
msg = ""
msg = comprobarCampos()
If (msg <> "") Then
    MsgBox (msg)
Else
    linea = ultimaLinea()
    Call dibujarLineas(linea)
    Call recogerCampos
    Call Borrar
End If
End Sub
Private Sub CommandButton2_Click()
    ComboBox1.Clear
    ComboBox2.Clear
    ComboBox3.Clear
    Call Borrar
End Sub


Private Sub Label14_Click()
    
End Sub


Private Sub Label17_Click()
    MsgBox ("Introduzca su documento de identidad sin espacios entre numeros y letras")
End Sub
Private Sub TextBox11_Change()
    Frame2.Visible = True
End Sub


Private Sub TextBox14_Change()
    If Len(TextBox14) > 4 Then
        TextBox14 = Left(TextBox14.Text, 4)
        TextBox14.SelStart = 4
    End If
End Sub

Private Sub TextBox15_Change()
    If Len(TextBox15) > 4 Then
        TextBox15 = Left(TextBox15.Text, 4)
        TextBox15.SelStart = 4
    End If
End Sub

Private Sub TextBox16_Change()
    If Len(TextBox16) > 2 Then
        TextBox16 = Left(TextBox16.Text, 2)
        TextBox16.SelStart = 2
    End If
End Sub
Private Sub TextBox17_Change()
    If Len(TextBox17) > 10 Then
        TextBox17 = Left(TextBox17.Text, 10)
        TextBox17.SelStart = 10
    End If
End Sub

Private Sub TextBox4_Change()
    ComboBox3.Visible = True
End Sub
Private Sub TextBox5_Change()
 ComboBox2.Visible = True
End Sub
Private Sub UserForm_Activate()
Dim linea As Integer
For i = 1 To 15
    Cells(1, i).Interior.Color = RGB(0, 255, 0) '134, 134, 134) '
    'Cells(1, i).Borders(xlEdgeLeft).LineStyle = xlContinuous
Next
    Call dibujarLineas(1)
    linea = ultimaLinea()
    Cells(1, 1).Value = "Codigo"
    Cells(1, 2).Value = "Nombre"
    Cells(1, 3).Value = "Apellido 1"
    Cells(1, 4).Value = "Apellido 2"
    Cells(1, 5).Value = "Sexo"
    Cells(1, 6).Value = "Documento"
    Cells(1, 7).Value = "Provincia"
    Cells(1, 8).Value = "Ciudad"
    Cells(1, 9).Value = "Nombre Via"
    Cells(1, 10).Value = "Piso"
    Cells(1, 11).Value = "Puerta"
    Cells(1, 12).Value = "C.P."
    Cells(1, 13).Value = "Telefono"
    Cells(1, 14).Value = "Fecha"
    Cells(1, 15).Value = "Cuenta Bancaria"
    

    'MsgBox (linea)
    Call Borrar
End Sub
Sub Borrar()
    ComboBox1.Clear
    ComboBox2.Clear
    ComboBox3.Clear
    
    OptionButton1.Value = False
    OptionButton2.Value = False
    
    TextBox1 = ""
    TextBox2 = ""
    TextBox3 = ""
    TextBox4 = ""
    TextBox5 = ""
    TextBox6 = ""
    TextBox7 = ""
    TextBox8 = ""
    TextBox9 = ""
    TextBox10 = ""
    TextBox11 = ""
    TextBox12 = Date
    TextBox14 = ""
    TextBox15 = ""
    TextBox16 = ""
    TextBox17 = ""
    
    ComboBox1.AddItem "DNI"
    ComboBox1.AddItem "NIE"
    'ComboBox1.AddItem "Pasaporte"
    ComboBox2.AddItem "Calle"
    ComboBox2.AddItem "Avenida"
    ComboBox2.AddItem "Otro"
    ComboBox1.Visible = False 'carlosnewmusic para que activas y desactivas?
    ComboBox2.Visible = False
    ComboBox3.Visible = False
    ComboBox1.Visible = True
    
    TextBox4.Visible = False
    TextBox5.Visible = False
    TextBox6.Visible = False
    TextBox7.Visible = False
    TextBox8.Visible = False
    TextBox9.Visible = False
    TextBox10.Visible = False
    TextBox11.Visible = False
    'TextBox11.Visible = False carlosnewmusic que verga es esto?
    
    Label17.Visible = False
    
    ComboBox3.AddItem "Álava"
    ComboBox3.AddItem "Albacete"
    ComboBox3.AddItem "Alicante"
    ComboBox3.AddItem "Almería"
    ComboBox3.AddItem "Ávila"
    ComboBox3.AddItem "Badajoz"
    ComboBox3.AddItem "Baleares"
    ComboBox3.AddItem "Barcelona"
    ComboBox3.AddItem "Burgos"
    ComboBox3.AddItem "Cáceres"
    ComboBox3.AddItem "Cádiz"
    ComboBox3.AddItem "Castellón"
    ComboBox3.AddItem "Ciudad Real"
    ComboBox3.AddItem "Córdoba"
    ComboBox3.AddItem "Coruña"
    ComboBox3.AddItem "Cuenca"
    ComboBox3.AddItem "Gerona"
    ComboBox3.AddItem "Granada"
    ComboBox3.AddItem "Guadalajara"
    ComboBox3.AddItem "Guipúzcoa"
    ComboBox3.AddItem "Huelva"
    ComboBox3.AddItem "Huesca"
    ComboBox3.AddItem "Jaén"
    ComboBox3.AddItem "León"
    ComboBox3.AddItem "Lérida"
    ComboBox3.AddItem "La Rioja"
    ComboBox3.AddItem "Lugo"
    ComboBox3.AddItem "Madrid"
    ComboBox3.AddItem "Málaga"
    ComboBox3.AddItem "Murcia"
    ComboBox3.AddItem "Navarra"
    ComboBox3.AddItem "Orense"
    ComboBox3.AddItem "Asturias"
    ComboBox3.AddItem "Palencia"
    ComboBox3.AddItem "Las Palmas"
    ComboBox3.AddItem "Pontevedra"
    ComboBox3.AddItem "Salamanca"
    ComboBox3.AddItem "Santa Cruz de Tenerife"
    ComboBox3.AddItem "Cantabria"
    ComboBox3.AddItem "Segovia"
    ComboBox3.AddItem "Sevilla"
    ComboBox3.AddItem "Soria"
    ComboBox3.AddItem "Tarragona"
    ComboBox3.AddItem "Teruel"
    ComboBox3.AddItem "Toledo"
    ComboBox3.AddItem "Valencia"
    ComboBox3.AddItem "Valladolid"
    ComboBox3.AddItem "Vizcaya"
    ComboBox3.AddItem "Zamora"
    ComboBox3.AddItem "Zaragoza"
    ComboBox3.AddItem "Ceuta"
    ComboBox3.AddItem "Melilla"
    Frame2.Visible = False
    
    'harley
     'cambiado
    ComboBox4.AddItem "Álava"
    ComboBox4.AddItem "Albacete"
    ComboBox4.AddItem "Alicante"
    ComboBox4.AddItem "Almería"
    ComboBox4.AddItem "Ávila"
    ComboBox4.AddItem "Badajoz"
    ComboBox4.AddItem "Baleares"
    ComboBox4.AddItem "Barcelona"
    ComboBox4.AddItem "Burgos"
    ComboBox4.AddItem "Cáceres"
    ComboBox4.AddItem "Cádiz"
    ComboBox4.AddItem "Castellón"
    ComboBox4.AddItem "Ciudad Real"
    ComboBox4.AddItem "Córdoba"
    ComboBox4.AddItem "Coruña"
    ComboBox4.AddItem "Cuenca"
    ComboBox4.AddItem "Gerona"
    ComboBox4.AddItem "Granada"
    ComboBox4.AddItem "Guadalajara"
    ComboBox4.AddItem "Guipúzcoa"
    ComboBox4.AddItem "Huelva"
    ComboBox4.AddItem "Huesca"
    ComboBox4.AddItem "Jaén"
    ComboBox4.AddItem "León"
    ComboBox4.AddItem "Lérida"
    ComboBox4.AddItem "La Rioja"
    ComboBox4.AddItem "Lugo"
    ComboBox4.AddItem "Madrid"
    ComboBox4.AddItem "Málaga"
    ComboBox4.AddItem "Murcia"
    ComboBox4.AddItem "Navarra"
    ComboBox4.AddItem "Orense"
    ComboBox4.AddItem "Asturias"
    ComboBox4.AddItem "Palencia"
    ComboBox4.AddItem "Las Palmas"
    ComboBox4.AddItem "Pontevedra"
    ComboBox4.AddItem "Salamanca"
    ComboBox4.AddItem "Santa Cruz de Tenerife"
    ComboBox4.AddItem "Cantabria"
    ComboBox4.AddItem "Segovia"
    ComboBox4.AddItem "Sevilla"
    ComboBox4.AddItem "Soria"
    ComboBox4.AddItem "Tarragona"
    ComboBox4.AddItem "Teruel"
    ComboBox4.AddItem "Toledo"
    ComboBox4.AddItem "Valencia"
    ComboBox4.AddItem "Valladolid"
    ComboBox4.AddItem "Vizcaya"
    ComboBox4.AddItem "Zamora"
    ComboBox4.AddItem "Zaragoza"
    ComboBox4.AddItem "Ceuta"
    ComboBox4.AddItem "Melilla"
    
    'carlosnewmusic
    ComboBox7.Clear
    ComboBox6.Clear
    ComboBox5.Clear
    
    OptionButton5.Value = False
    OptionButton6.Value = False
    
    TextBox25 = ""
    TextBox26 = ""
    TextBox27 = ""
    TextBox28 = ""
    TextBox29 = ""
    TextBox30 = ""
    TextBox31 = ""
    TextBox32 = ""
    TextBox33 = ""
    TextBox34 = ""
    TextBox35 = ""
    TextBox36 = Date
    TextBox37 = ""
    TextBox38 = ""
    TextBox39 = ""
    TextBox40 = ""
    
    ComboBox5.AddItem "DNI" 'cambiado carlosnewmusic
    ComboBox5.AddItem "NIE"
    'ComboBox5.AddItem "Pasaporte"
    ComboBox6.AddItem "Calle"
    ComboBox6.AddItem "Avenida"
    ComboBox6.AddItem "Plaza"
    ComboBox6.AddItem "Otro"
    
    ComboBox5.Visible = True
    ComboBox6.Visible = False
    ComboBox7.Visible = False
    
    
    TextBox28.Visible = False
    TextBox29.Visible = False
    TextBox30.Visible = False
    TextBox31.Visible = False
    TextBox32.Visible = False
    TextBox33.Visible = False
    TextBox34.Visible = False
    TextBox35.Visible = False

    ComboBox7.AddItem "Álava"
    ComboBox7.AddItem "Albacete"
    ComboBox7.AddItem "Alicante"
    ComboBox7.AddItem "Almería"
    ComboBox7.AddItem "Ávila"
    ComboBox7.AddItem "Badajoz"
    ComboBox7.AddItem "Baleares"
    ComboBox7.AddItem "Barcelona"
    ComboBox7.AddItem "Burgos"
    ComboBox7.AddItem "Cáceres"
    ComboBox7.AddItem "Cádiz"
    ComboBox7.AddItem "Castellón"
    ComboBox7.AddItem "Ciudad Real"
    ComboBox7.AddItem "Córdoba"
    ComboBox7.AddItem "Coruña"
    ComboBox7.AddItem "Cuenca"
    ComboBox7.AddItem "Gerona"
    ComboBox7.AddItem "Granada"
    ComboBox7.AddItem "Guadalajara"
    ComboBox7.AddItem "Guipúzcoa"
    ComboBox7.AddItem "Huelva"
    ComboBox7.AddItem "Huesca"
    ComboBox7.AddItem "Jaén"
    ComboBox7.AddItem "León"
    ComboBox7.AddItem "Lérida"
    ComboBox7.AddItem "La Rioja"
    ComboBox7.AddItem "Lugo"
    ComboBox7.AddItem "Madrid"
    ComboBox7.AddItem "Málaga"
    ComboBox7.AddItem "Murcia"
    ComboBox7.AddItem "Navarra"
    ComboBox7.AddItem "Orense"
    ComboBox7.AddItem "Asturias"
    ComboBox7.AddItem "Palencia"
    ComboBox7.AddItem "Las Palmas"
    ComboBox7.AddItem "Pontevedra"
    ComboBox7.AddItem "Salamanca"
    ComboBox7.AddItem "Santa Cruz de Tenerife"
    ComboBox7.AddItem "Cantabria"
    ComboBox7.AddItem "Segovia"
    ComboBox7.AddItem "Sevilla"
    ComboBox7.AddItem "Soria"
    ComboBox7.AddItem "Tarragona"
    ComboBox7.AddItem "Teruel"
    ComboBox7.AddItem "Toledo"
    ComboBox7.AddItem "Valencia"
    ComboBox7.AddItem "Valladolid"
    ComboBox7.AddItem "Vizcaya"
    ComboBox7.AddItem "Zamora"
    ComboBox7.AddItem "Zaragoza"
    ComboBox7.AddItem "Ceuta"
    ComboBox7.AddItem "Melilla"
    
End Sub
Sub recogerCampos()
    l = ultimaLinea()
    Cells(l, 1).Value = l - 1 & "-" & Left(TextBox4, 3)
    Hoja2.Cells(1, 1) = Hoja2.Cells(1, 1).Value + 1
    Cells(l, 2).Value = TextBox1
    Cells(l, 3).Value = TextBox2
    Cells(l, 4).Value = TextBox3
    If (OptionButton1.Value = True) Then
        Cells(l, 5).Value = "Hombre"
    Else
        Cells(l, 5).Value = "Mujer"
    End If
    Cells(l, 6).Value = UCase(TextBox4)
    Cells(l, 7).Value = ComboBox3.Value
    Cells(l, 8).Value = TextBox5
    Cells(l, 9).Value = ComboBox2.Value + " - " + TextBox6
    Cells(l, 10).Value = TextBox7
    Cells(l, 11).Value = TextBox8
    Cells(l, 12).Value = TextBox9 + "" + TextBox10
    Cells(l, 13).Value = TextBox11
    Cells(l, 14).Value = TextBox12
    Cells(l, 15).Value = TextBox14 + "-" + TextBox15 + "-" + TextBox16 + "-" + TextBox17
    
End Sub
Sub dibujarLineas(ByVal l As Integer)
    For c = 1 To 15
        Cells(l, c).Borders(xlEdgeRight).LineStyle = xlContinuous
        Cells(l, c).Borders(xlEdgeBottom).LineStyle = xlContinuous
    Next
End Sub
Function ultimaLinea() 'ByVal Hoja As String)
Dim Salir As Boolean
Dim linea As Integer
    Salir = True
    l = 1
    Do
        If (Cells(l, 2).Value = "") Then
            ultimaLinea = l
            Salir = False
        Else
            l = l + 1
        End If
    Loop While (Salir)
End Function
Function comprobarCampos()
Dim arr(23) As String
'RESTO   0   1   2   3   4   5   6   7   8   9   10  11  12  13  14  15  16  17  18  19  20  21  22
'LETRA   T   R   W   A   G   M   Y   F   P   D   X   B   N   J   Z   S   Q   V   H   L   C   K   E
arr(0) = "T"
arr(1) = "R"
arr(2) = "W"
arr(3) = "A"
arr(4) = "G"
arr(5) = "M"
arr(6) = "Y"
arr(7) = "F"
arr(8) = "P"
arr(9) = "D"
arr(10) = "X"
arr(11) = "B"
arr(12) = "N"
arr(13) = "J"
arr(14) = "Z"
arr(15) = "S"
arr(16) = "Q"
arr(17) = "V"
arr(18) = "H"
arr(19) = "L"
arr(20) = "C"
arr(21) = "K"
arr(22) = "E"
    If (TextBox1 = "") Then
        msg = msg + "Falta Nombre " + vbCr
    End If
    If (TextBox2 = "") Then
        msg = msg + "Falta Preimer Apellido" + vbCr
    End If
    If (TextBox3 = "") Then
        msg = msg + "Falta Segundo Apellido" + vbCr
    End If
    If (OptionButton1.Value = False And OptionButton2.Value = False) Then
        msg = msg + "Falta Seleccionar sexo" + vbCr
    End If
    If (TextBox4 = "") Then
        msg = msg + "Falta Documento" + ComboBox1.Value + vbCr
    Else
        If (Len(TextBox4) = 9) Then
            Select Case ComboBox1.Value
                Case "DNI"
                    Numero = Mid(TextBox4, 1, 8)
                    LetraF = UCase(Right(TextBox4, 1))
                    If (IsNumeric(Numero) = True) Then
                        resto = Int(Numero) Mod (23)
                        'MsgBox (arr(resto))
                        If (LetraF <> arr(resto)) Then
                            msg = msg + "DNI es incorrecto" + vbCr
                        End If
                    Else
                        msg = msg + "Formato DNI es incorrecto" + vbCr
                    End If
                Case "NIE"
                    LetraI = UCase(Left(TextBox4, 1))
                    
                    Numero = Mid(TextBox4, 2, 7)
                    If (IsNumeric(Numero) = True) Then
                        'MsgBox (Numero)
                        LetraF = UCase(Right(TextBox4, 1))
                        Select Case LetraI
                            Case "X"
                                nie = "0" + Numero
                                'MsgBox (nie)
                                resto = Int(nie) Mod (23)
                            Case "Y"
                                nie = "1" + Numero
                                'MsgBox (nie)
                                resto = Int(nie) Mod (23)
                            Case "Z"
                                nie = "2" + Numero
                                'MsgBox (nie)
                                resto = Int(nie) Mod (23)
                        End Select
                        'MsgBox (arr(resto))
                        'resto = Int(TextBox4) Mod (23)
                        If (LetraF <> arr(resto)) Then
                            msg = msg + "NIE es incorrecto" + vbCr
                        Else
                        
                        End If
                    Else
                        msg = msg + "Formato NIE es incorrecto" + vbCr
                    End If
            End Select
        Else
            msg = msg + "Formato del Documento no es correcto" + vbCr
        End If
    End If
    If (TextBox5 = "") Then
        msg = msg + "Falta Direccion" + vbCr
    End If
    If (TextBox6 = "") Then
        msg = msg + "Falta Ciudad" + vbCr
    End If
    If (TextBox7 = "") Then
        msg = msg + "Falta Escalera" + vbCr
    End If
    If (TextBox8 = "") Then
        msg = msg + "Falta Puerta" + vbCr
    End If
    If (TextBox10 = "") Then
        msg = msg + "Falta Codigo Postal" + vbCr
    Else
        If (IsNumeric(TextBox10) = False) Then
            msg = msg + "Codigo Postal debe ser un numero" + vbCr
        Else
            If (Len(TextBox10) <> 3) Then
                msg = msg + "Codigo Postal debe ser un numero de 3 digitos" + vbCr
            End If
        End If
    End If
    If (TextBox11 = "") Then
        msg = msg + "Falta Telefono" + vbCr
    Else
        If (IsNumeric(TextBox11) = False) Then
            msg = msg + "Numero Telefono no puede contener letras" + vbCr
        ElseIf (Len(TextBox11) <> 9) Then
            msg = msg + "Numero Telefono no puede contener mas de 9 cifras" + vbCr
        End If
    End If
    If (TextBox12 = "" Or IsDate(TextBox12) = False) Then
        'Comprobacion fecha
        msg = msg + "Falta Fecha" + vbCr
    End If
    If (TextBox14 = "" Or TextBox15 = "" Or TextBox16 = "" Or TextBox17 = "") Then
        msg = msg + "Falta Cuenta Bancaria" + vbCr
    Else
        If (Len(TextBox14) = 4 And Len(TextBox15) = 4 And Len(TextBox16) = 2 And Len(TextBox17) = 10) Then
            If (IsNumeric(TextBox14) = False Or IsNumeric(TextBox15) = False Or IsNumeric(TextBox16) = False Or IsNumeric(TextBox17) = False) Then
                msg = msg + "La cuenta bancaria debe ser numerica" + vbCr
'            Else
'                If (Len(TextBox10) <> 3) Then
'                    msg = msg + "Cuenta bancaria es erronea" + vbCr
'                End If
            End If
        Else
            msg = msg + "Cuenta bancaria es erronea" + vbCr
        End If
    End If
    comprobarCampos = msg
End Function



'desde aqui harley
Private Sub CheckBox1_Click()
    If (CheckBox1.Value) Then
        Frame3.Visible = True
    Else
        Frame3.Visible = False
    End If
End Sub

Private Sub CheckBox2_Click()

End Sub
    If (CheckBox2.Value) Then
        Label21.Visible = True
        ComboBox4.Visible = True
    Else
        Label21.Visible = False
        ComboBox4.Visible = False
    End If
End Sub
Private Sub CommandButton3_Click()
l = ultimaLinea()
nom = TextBox20.Value
sex = ""
prov = ""
    If (nom = "" And CheckBox1.Value = False And CheckBox2.Value = False) Then
        MsgBox ("Ingrese un valor para buscar")
    Else
        If (CheckBox1.Value Or CheckBox2.Value) Then
            If (CheckBox1.Value) Then
                If (OptionButton3.Value) Then
                    sex = "Hombre"
                    'MsgBox ("sexo masculino")
                ElseIf (OptionButton4.Value) Then
                    sex = "Mujer"
                    'MsgBox ("sexo femenino")
                Else
                    MsgBox ("Falta sexo")
                End If
            End If
            If (CheckBox2.Value) Then
                If (ComboBox4.Value = "") Then
                    MsgBox ("Seleccione una provincia")
                Else
                    prov = ComboBox4.Value
                End If
            End If
        'Else
            'persona = buscar(nom, sex, prov)
        End If
        Call buscar(UCase(Trim(nom)), UCase(Trim(sex)), UCase(Trim(prov)))
    ListBox1.List = persona
    End If
End Sub
Sub buscar(ByVal nombre, ByVal sexo, ByVal provincia)
l = ultimaLinea()
Line = 0
Dim p()
ReDim persona(l - 2, 15)
    For i = 1 To l
        For j = 1 To 15
            If (nombre <> "") Then
                    If (sexo = "" And provincia = "") Then
                        If (StrComp(nombre, UCase(Cells(i, j).Value)) = 0) Then
                            For k = 1 To 15
                                persona(Line, k - 1) = Cells(i, k).Value
                                j = j + 1
                            Next
                            Line = Line + 1
                        End If
                    ElseIf (sexo <> "" And provincia = "") Then
                        If (StrComp(nombre, UCase(Cells(i, j).Value)) = 0 And StrComp(UCase(sexo), UCase(Cells(i, 5).Value)) = 0) Then
                            For k = 1 To 15
                                persona(Line, k - 1) = Cells(i, k).Value
                                j = j + 1
                            Next
                            Line = Line + 1
                        End If
                    ElseIf (sexo = "" And provincia <> "") Then
                        If (StrComp(nombre, UCase(Cells(i, j).Value)) = 0 And StrComp(UCase(provincia), UCase(Cells(i, 7).Value)) = 0) Then
                            For k = 1 To 15
                                persona(Line, k - 1) = Cells(i, k).Value
                                j = j + 1
                            Next
                            Line = Line + 1
                        End If
                    ElseIf (sexo <> "" And provincia <> "") Then
                        If (StrComp(nombre, UCase(Cells(i, j).Value)) = 0 And StrComp(UCase(sexo), UCase(Cells(i, 5).Value)) = 0 And StrComp(provincia, Cells(i, 7).Value) = 0) Then
                            For k = 1 To 15
                                persona(Line, k - 1) = Cells(i, k).Value
                                j = j + 1
                            Next
                            Line = Line + 1
                        End If
                    End If
            Else
                    If (sexo <> "" And provincia = "") Then
                        If (StrComp(UCase(sexo), UCase(Cells(i, 5).Value)) = 0) Then
                            For k = 1 To 15
                                persona(Line, k - 1) = Cells(i, k).Value
                                j = j + 1
                            Next
                            Line = Line + 1
                        End If
                    ElseIf (sexo = "" And provincia <> "") Then
                        If (StrComp(UCase(provincia), UCase(Cells(i, 7).Value)) = 0) Then
                            For k = 1 To 15
                                persona(Line, k - 1) = Cells(i, k).Value
                                j = j + 1
                            Next
                            Line = Line + 1
                        End If
                    ElseIf (sexo <> "" And provincia <> "") Then
                        If (StrComp(UCase(sexo), UCase(Cells(i, 5).Value)) = 0 And StrComp(provincia, UCase(Cells(i, 7).Value)) = 0) Then
                            For k = 1 To 15
                                persona(Line, k - 1) = Cells(i, k).Value
                                j = j + 1
                            Next
                            Line = Line + 1
                        End If
                    End If
            End If
        Next
    Next
End Sub
Private Sub ListBox1_Click()
    index = ListBox1.ListIndex
    valor = ListBox1.Column(0)
    MsgBox (CStr(index) + vbCr + CStr(valor))
    'ListBox1.Column(index))
    m = MsgBox("Desea Modificar este usuario", vbOKCancel)
    If (m = 1) Then
        MultiPage1.Pages(2).Visible = True
    Else
        MultiPage1.Pages(2).Visible = False
    End If
End Sub
Private Sub TextBox11_Change()
    Frame2.Visible = True
End Sub
'carlosnewmusic
Function comprobarCampos2()
Dim arr(23) As String
'RESTO   0   1   2   3   4   5   6   7   8   9   10  11  12  13  14  15  16  17  18  19  20  21  22
'LETRA   T   R   W   A   G   M   Y   F   P   D   X   B   N   J   Z   S   Q   V   H   L   C   K   E
arr(0) = "T"
arr(1) = "R"
arr(2) = "W"
arr(3) = "A"
arr(4) = "G"
arr(5) = "M"
arr(6) = "Y"
arr(7) = "F"
arr(8) = "P"
arr(9) = "D"
arr(10) = "X"
arr(11) = "B"
arr(12) = "N"
arr(13) = "J"
arr(14) = "Z"
arr(15) = "S"
arr(16) = "Q"
arr(17) = "V"
arr(18) = "H"
arr(19) = "L"
arr(20) = "C"
arr(21) = "K"
arr(22) = "E"
    If (TextBox25 = "") Then
        msg = msg + "Falta Nombre " + vbCr
    End If
    If (TextBox26 = "") Then
        msg = msg + "Falta Preimer Apellido" + vbCr
    End If
    If (TextBox27 = "") Then
        msg = msg + "Falta Segundo Apellido" + vbCr
    End If
    If (OptionButton5.Value = False And OptionButton6.Value = False) Then
        msg = msg + "Falta Seleccionar sexo" + vbCr
    End If
    If (TextBox28 = "") Then
        msg = msg + "Falta Documento" + ComboBox5.Value + vbCr
    Else
        If (Len(TextBox28) = 9) Then
            Select Case ComboBox5.Value
                Case "DNI"
                    Numero = Mid(TextBox4, 1, 8)
                    LetraF = UCase(Right(TextBox4, 1))
                    If (IsNumeric(Numero) = True) Then
                        resto = Int(Numero) Mod (23)
                        'MsgBox (arr(resto))
                        If (LetraF <> arr(resto)) Then
                            msg = msg + "DNI es incorrecto" + vbCr
                        End If
                    Else
                        msg = msg + "Formato DNI es incorrecto" + vbCr
                    End If
                Case "NIE"
                    LetraI = UCase(Left(TextBox28, 1))
                    
                    Numero = Mid(TextBox28, 2, 7)
                    If (IsNumeric(Numero) = True) Then
                        'MsgBox (Numero)
                        LetraF = UCase(Right(TextBox28, 1))
                        Select Case LetraI
                            Case "X"
                                nie = "0" + Numero
                                'MsgBox (nie)
                                resto = Int(nie) Mod (23)
                            Case "Y"
                                nie = "1" + Numero
                                'MsgBox (nie)
                                resto = Int(nie) Mod (23)
                            Case "Z"
                                nie = "2" + Numero
                                'MsgBox (nie)
                                resto = Int(nie) Mod (23)
                        End Select
                        'MsgBox (arr(resto))
                        'resto = Int(TextBox28) Mod (23)
                        If (LetraF <> arr(resto)) Then
                            msg = msg + "NIE es incorrecto" + vbCr
                        Else
                        
                        End If
                    Else
                        msg = msg + "Formato NIE es incorrecto" + vbCr
                    End If
            End Select
        Else
            msg = msg + "Formato del Documento no es correcto" + vbCr
        End If
    End If
    If (TextBox29 = "") Then
        msg = msg + "Falta Ciudad" + vbCr
    End If
        If (TextBox30 = "") Then
        msg = msg + "Falta Direccion" + vbCr
    End If
    If (TextBox31 = "") Then
        msg = msg + "Falta piso" + vbCr
    End If
    If (TextBox32 = "") Then
        msg = msg + "Falta Puerta" + vbCr
    End If
    If (TextBox33 = "" Or TextBox34 = "") Then
        msg = msg + "Falta Codigo Postal" + vbCr
    Else
        If ((IsNumeric(TextBox33) = False) Or (IsNumeric(TextBox34) = False)) Then
            msg = msg + "Codigo Postal debe ser un numero" + vbCr
        Else
            If (Len(TextBox34) <> 3) Then
                msg = msg + "Codigo Postal debe ser un numero de 3 digitos" + vbCr
            End If
            If (Len(TextBox33) <> 2) Then
                msg = msg + "Codigo Postal de provincia debe ser un numero de 2 digitos" + vbCr
            End If
        End If
    End If
    If (TextBox35 = "") Then
        msg = msg + "Falta Telefono" + vbCr
    Else
        If (IsNumeric(TextBox35) = False) Then
            msg = msg + "Numero Telefono no puede contener letras" + vbCr
        ElseIf (Len(TextBox35) <> 9) Then
            msg = msg + "Numero Telefono no puede contener mas de 9 cifras" + vbCr
        End If
    End If
    If (TextBox36 = "" Or IsDate(TextBox36) = False) Then
        'Comprobacion fecha
        msg = msg + "Falta Fecha" + vbCr
    End If
    If (TextBox38 = "" Or TextBox15 = "" Or TextBox16 = "" Or TextBox17 = "") Then
        msg = msg + "Falta Cuenta Bancaria" + vbCr
    Else
        If (Len(TextBox38) = 4 And Len(TextBox39) = 4 And Len(TextBox40) = 2 And Len(TextBox41) = 10) Then
            If (IsNumeric(TextBox38) = False Or IsNumeric(TextBox39) = False Or IsNumeric(TextBox40) = False Or IsNumeric(TextBox41) = False) Then
                msg = msg + "La cuenta bancaria debe ser numerica" + vbCr
'            Else
'                If (Len(TextBox10) <> 3) Then
'                    msg = msg + "Cuenta bancaria es erronea" + vbCr
'                End If
            End If
        Else
            msg = msg + "Cuenta bancaria es erronea" + vbCr
        End If
    End If
    comprobarCampos = msg
End Function
