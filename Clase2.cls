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
    'Tipo de datos en VBA
    'Byte numeros de 0 a 255
    'Boolean True o False
    'Integer Entero Simple -32768 a 32767
    'Long Entero -2000000000 a 200000000
    'String texto
    'Double Decimal (64 bits)
    
    'Declaración de variable
    Dim nombre As String
    'Asignación de valor
    nombre = "Karina"
    
    Dim edad As Byte
    edad = 38
    
    'camelCase      letraCamello
    
    Dim sueldoBasico As Long
    sueldoBasico = 80000
    
    ActiveCell.Value = sueldoBasico
    
    'Constantes
    'Las constantes permiten almacenar un valor, No pueden cambiar de valor
    
    Const PI = 3.141592
    
    edad = edad + 1 'sumo 1 a la variable
    
    'PI = PI + 1     'error no se puede cambiar valor a una constante
    
End Sub
Sub macro2()
    'método offset (Desplazamiento)
    Range("D10").Select
    ActiveCell.Value = "OffSet"
    ActiveCell.Offset(2, 3).Value = "OffSet 2f 3c" ' Desplaza 2filas 3 columnas
    
    With ActiveCell
        'código indentado (tabulado, sangria)
        .Offset(4, 5).Value = "Abajo 4 Derecha 5"
        .Offset(-2, -3).Value = "Arriba 2 Izquierda 3"
        '.Offset(-15, -8).Value = "Fuera de Rango"
    End With
End Sub
Sub macro3()
    'Ingresar (desde un InputBox) Base y Altura de un rectangulo
    'Informar en un MSGBOX y en celda A1 y A2 el perimetro y superficie
    '       del rectangulo
    Dim base As Double, alt As Double
    Dim perimetro As Double, superficie As Double
    base = InputBox("Ingrese la base del rectangulo")
    alt = InputBox("Ingrese la altura del rectangulo")
    
    perimetro = base * 2 + alt * 2
    superficie = base * alt
    
    MsgBox ("Perimetro: " & perimetro & " Superficie: " & superficie)
    [a1].Value = "Perimetro: " & perimetro
    [a1].Offset(1, 0).Value = "Superficie: " & superficie
    
End Sub
Sub macro4()
    ' Estructura IF - THEN - END IF
    Dim num As Integer
    num = InputBox("Ingrese un número")
    [d1].Activate
    If num < 100 Then
        ActiveCell.Value = "Numero " & num
        ActiveCell.Offset(1, 0).Value = "El número es menor de 100"
    End If
End Sub
Sub macro5()
    ' Estructura IF - THEN - ELSE - END IF
    Dim num As Integer
    num = InputBox("Ingrese un número")
    [d1].Activate
    If num <= 100 Then
        ActiveCell.Value = "Numero " & num
        ActiveCell.Offset(1, 0).Value = "El número es menor de 100"
        Range("d1:f2").Interior.Color = RGB(255, 0, 0)
        Range("d1:f2").Font.Color = RGB(0, 0, 0)
    Else
        ActiveCell.Value = "Numero " & num
        ActiveCell.Offset(1, 0).Value = "El número es mayor de 100"
        Range("d1:f2").Interior.Color = RGB(0, 0, 255)
        Range("d1:f2").Font.Color = RGB(255, 255, 255)
    End If
End Sub
Sub macro6()
    
    ' Función MsgBox Mensaje, Boton, Titulo
    ' MsgBox "Hola", vbOKCancel, "Martes"
    
    ' MsgBox "A usted del gusta este curso?", vbYesNo, "Pregunta"
    ' MsgBox "Tiene Calor?", vbOKCancel, "Pregunta"
    
    ' Valor Devuelto
    ' 1     OK
    ' 2     Cancel
    ' 3     Abort
    ' 4     Retry
    ' 5     Ignore
    ' 6     Yes
    ' 7     NO
    
    Dim respuesta As Integer
    respuesta = MsgBox("A usted del gusta este curso?", vbYesNo, "Pregunta")
    If respuesta = 6 Then
        [a1].Value = "Muchas Gracias!"
        Range("a1:c2").Interior.Color = RGB(0, 255, 0)
        Range("a1:c2").Font.Color = RGB(0, 0, 0)
    Else
        [a1].Value = "Gracias Igualmente!"
        Range("a1:c2").Interior.Color = RGB(0, 0, 255)
        Range("a1:c2").Font.Color = RGB(255, 255, 255)
    End If
End Sub










