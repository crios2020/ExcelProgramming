VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub macro1()



    'Estructura Case

    Dim dia As Integer

    dia = InputBox("Ingrese un número del 1 al 7")

    'dia = Range(a2).Value

    

    [a10].Select

    

    Select Case dia

        Case 1

            ActiveCell.Value = "1-domingo"

        Case 2

            ActiveCell.Value = "2-lunes"

        Case 3

            ActiveCell.Value = "3-martes"

        Case 4

            ActiveCell.Value = "4-miercoles"

        Case 5

            ActiveCell.Value = "5-jueves"

        Case 6

            ActiveCell.Value = "6-viernes"

        Case 7

            ActiveCell.Value = "7-sabado"

        Case Else

            ActiveCell.Value = "no se ingreso valor de 1 a 7"

    End Select

End Sub

Sub mcaro2()

    'estructura do while

    Dim x As Integer

    x = 1

    [a1].Select

    Do While x <= 20

        ActiveCell.Value = x

        ActiveCell.Offset(1, 0).Select

        x = x + 3

    Loop

    

End Sub

Sub macro3()

    'loop infinito

    ' No Hacer

    'Dim x As Integer

    'x = 1

    '[a1].Select

    'Do While x >= 1

    '   ActiveCell.Value = x

    '  ActiveCell.Offset(1, 0).Select

    ' x = x + 1

    'Loop

End Sub

Sub IngresoDatos()
    Dim nombre As String
    Dim ciudad As String
    Dim edad As Integer
    Dim fecha As Date
    [a1].Select
    
    Do While ActiveCell.Value <> ""
        ActiveCell.Offset(1, 0).Select
    Loop
    
    
    nombre = InputBox("Ingrese el nombre (Enter para Terminar);", "nombre")
    Do While nombre <> ""
        ciudad = InputBox("Ingrese la ciudad:", "ciudad")
        edad = Val(InputBox("Ingrese la edad:", "edad"))
        fecha = CDate(InputBox("Ingrese fecha ingreso:", "fecha"))
        With ActiveCell
            .Value = nombre
            .Offset(0, 1).Value = ciudad
            .Offset(0, 2).Value = edad
            .Offset(0, 3).Value = fecha
        End With
        ActiveCell.Offset(1, 0).Select
        nombre = InputBox("Ingrese el nombre (Enter para Terminar);", "nombre")
    Loop
    
    
End Sub

Sub macro5()
    'Estructura For
    [a1].Select
    For a = 1 To 10
        ActiveCell.Value = a
        ActiveCell.Offset(1, 0).Select
    Next
    
    [a1].Select
    For a = 20 To 100 Step 2
        ActiveCell.Value = a
        ActiveCell.Offset(1, 0).Select
    Next
    
    [a1].Select
    For a = 200 To 100 Step -3
        ActiveCell.Value = a
        ActiveCell.Offset(1, 0).Select
    Next
End Sub

Sub macro6()
    'vectores - arrays - arreglos
    Dim numeros(10) As Integer
    
    'Ingreso de números manual
    'numeros(1) = 34
    'numeros(2) = 39
    'numeros(3) = 38
    'numeros(4) = 29
    'numeros(5) = 26
    'numeros(6) = 24
    'numeros(7) = 16
    'numeros(8) = 14
    'numeros(9) = 23
    'numeros(10) = 65
    
    'Ingreso de numeros al vector
    For i = 1 To 10
        numeros(i) = InputBox("Ingrese un número")
    Next
    
    'Impresión de valores
    [a1].Select
    For i = 1 To 10
        ActiveCell.Value = numeros(i)
        ActiveCell.Offset(1, 0).Select
    Next
End Sub
Sub macro7()
    Dim semana(7) As String
    semana(1) = "Domingo"
    semana(2) = "Lunes"
    semana(3) = "Martes"
    semana(4) = "Miércoles"
    semana(5) = "Jueves"
    semana(6) = "Viernes"
    semana(7) = "Sábado"
    
    'Redimensionando el vector
    'ReDim Preserve semana(8)
    'semana(8) = "Osvaldo"
    
    dia = Range("a2").Value
    [a10].Value = semana(dia)
End Sub






