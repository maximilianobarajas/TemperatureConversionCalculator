' 
' Welcome to GDB Online.
' GDB online is an online compiler and debugger tool for C, C++, Python, Java, PHP, Ruby, Perl,
' C#, OCaml, VB, Swift, Pascal, Fortran, Haskell, Objective-C, Assembly, HTML, CSS, JS, SQLite, Prolog.
' Code, Compile, Run and Debug online from anywhere in world.
' 
' 

Sub kc()
' Rutina para convertir Kelvin a Celcius
' c=k-273.15
    Dim c As Double
    Dim k As Double
    k = InputBox("Cuántos °K:")
    c = k - 273.15
    MsgBox ("Los " & k & " °K equivalen a " & _
    c & " °C")
End Sub

Sub kf()
' Rutina para convertir Kelvin a Fahrenheit
' f=(9(k-273.15)/5)+32
    Dim f As Double
    Dim k As Double
    k = InputBox("Cuántos °K:")
    f = (9 * (k - 273.15) / 5) + 32
    MsgBox ("Los " & k & " °K equivalen a " & _
    f & " °F")
End Sub

Sub fc()
' Rutina para convertir Fahrenheit a Celcius
' c=5(f-32)/9
    Dim f As Double
    Dim c As Double
    f = InputBox("Cuántos °F:")
    c = 5 * (f - 32) / 9
    MsgBox ("Los " & f & " °F equivalen a " & _
    c & " °C")
End Sub

Sub fk()
' Rutina para convertir Fahrenheit a Kelvin
' k=(5(f-32)/9)+273.15
    Dim f As Double
    Dim k As Double
    f = InputBox("Cuántos °F:")
    k = (5 * (f - 32) / 9) + 273.15
    MsgBox ("Los " & f & " °F equivalen a " & _
    k & " °K")
End Sub

Sub ck()
' Rutina para convertir Celcius a Kelvin
' k=c+273.15
    Dim c As Double
    Dim k As Double
    c = InputBox("Cuántos °C:")
    k = c + 273.15
    MsgBox ("Los " & c & " °C equivalen a " & _
    k & " °K")
End Sub

Sub cf()
' Rutina para convertir Celcius a Fahrenheit
' f=(9c/5)+32
    Dim c As Double
    Dim f As Double
    c = InputBox("Cuántos °C:")
    f = (9 * c / 5) + 32
    MsgBox ("Los " & c & " °C equivalen a " & _
    f & " °F")
End Sub


Function temperaturas(tipoConversion As String, _
                      grados As Double) As String

    Select Case tipoConversion
        Case "KC"
            temperaturas = grados - 273.15
        Case "KF"
            temperaturas = (9 * (grados - 273.15) / 5) + 32
        Case "FC"
            temperaturas = (5 * (grados - 32)) / 9
        Case "FK"
            temperaturas = (5 * (grados - 32) / 9) + 273.15
        Case "CK"
            temperaturas = grados + 273.15
        Case "CF"
            temperaturas = (9 * grados / 5) + 32
        Case Else
            temperaturas = "El tipo de conversión " & _
                           "NO es una opción válida"
    End Select
    
End Function

Sub main()
    Dim grados As Double
    Dim tipoConversion As String
    tipoConversion = InputBox("Selecciona tipo de Conversión:" & _
                     vbCrLf & "CK - Celcius a Kelvin" & _
                     vbCrLf & "CF - Celcius a Fahrenheit" & _
                     vbCrLf & "FC - Fahrenheit a Celcius" & _
                     vbCrLf & "FK - Fahrenheit a Kelvin" & _
                     vbCrLf & "KC - Kelvin a Celcius" & _
                     vbCrLf & "KF - Kelvin a Fahrenheit")
    grados = InputBox("Cuántos Grados?")
    MsgBox (temperaturas(UCase(tipoConversion), grados))
End Sub