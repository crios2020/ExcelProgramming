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

    ' Profe: Carlos Rios  carlos.rios@educacionit.com
    ' Alumni: alumni.educacionit.com
    '         user: email
    '         pass: dni
    
    ' github: https://github.com/crios2020/ExcelProgramming

    ' VBA: Visual Basic For Application

    ' Comentarios
    
    
End Sub
Sub macro2()
    ActiveCell.Value = "Hola Mundo!!"       ' Imprime en celda activa
End Sub
Sub macro3()
    Range("d10").Select                     ' Activa la celda d10
    ActiveCell.Value = "Hola Mundo!!"
    ActiveCell.Font.Bold = True             ' Setea Negritas al texto
    ' ActiveCell.Font.Color = RGB(255, 0, 0)  ' Red, Green, Blue
    ActiveCell.Font.Color = RGB(7, 36, 245)
    
End Sub
Sub macro4()
    Worksheets("hoja2").Activate            ' Cambia de hoja
    ActiveSheet.Range("e1").Select          ' Activa la celda e1
    [a1].Select                             ' Activa la celda a1
    ActiveCell.Value = "Jueves"
    Range("f1").Value = "Viernes"           ' Cambia valor a celda f1 sin activar celda
    [g1].Value = "Sábado"                   ' Cambia valor a celda g1 sin activar celda
End Sub
Sub macro5()
    Worksheets.Add                          ' Agrega una nueva hoja al libro
    Sheets.Add After:=ActiveSheet           ' Agrega una hoja después de la activa
    Sheets("Hoja6").Select                  ' Cambia de hoja
    Sheets("Hoja6").Move After:=Sheets(6)   ' Mueve la hoja de lugar
    Sheets("Hoja6").Name = "Enero"          ' Cambia el nombre de la hoja
    
End Sub
Sub macro6()
    ' ActiveSheet.Range("a7").Value = "Excel"
    ' ActiveSheet.Range("a7").Font.Bold = True
    ' ActiveSheet.Range("a7").Font.Color = RGB(0, 255, 0)
    ' [a7].Value = "Excel"
    ' [a7].Font.Bold = True
    ' [a7].Font.Color = RGB(0, 255, 0)
    
    'Estructura with - end with
    With [a7]
        .Value = "Excel"
        .Font.Bold = True
        .Font.Color = RGB(0, 255, 0)
    End With
End Sub
Sub macro7()
    ' Método Cells: es igual a Range, pero con referencias numericas
    Cells(5, 7).Value = "Excel"
    ' equivalente a Range("g5").value="Excel"
End Sub
Sub macro8()
    ' Variables
    ' Una variable solo puede declararse una vez
    ' Una variable pueden tener infinitas asignaciones de valor
    
    Dim text As String          ' Declara una variable del tipo String
    text = "Hola"               ' Asigna un valor a la variable
    [a1].Value = text
    
    Dim nombre As String
    nombre = InputBox("Ingrese su nombre")      ' Caja de Entrada
    [a1].Value = "Hola " + nombre
    
    MsgBox "Hola " + nombre                     ' Caja de Salida
    ' Dim text As String         ' Error Variable duplicada
End Sub
Sub macro9()
    Dim text As String
    Dim celda As String
    
    celda = InputBox("Ingrese un nombre de celda")
    text = InputBox("Ingrese su nombre")
    
    Range(celda).Value = text

End Sub

