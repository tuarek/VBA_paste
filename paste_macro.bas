Attribute VB_Name = "Excel_MenExplo"
Public num_f As Integer
Public direccion As String, strClip As String

Public Sub Pegar()
Attribute Pegar.VB_ProcData.VB_Invoke_Func = "w\n14"
Dim rango As Range
Dim Form_fecha As Date, Form_tiempo As Date
Dim frase2 As String, Asunto As String, fecha As String, espacio As String
Dim frase As String, enter As String, cuerpo As String, nada As String
Dim Pos_Asuntoi As Integer, Pos_fecha As Integer
Dim Pos_1Enter As Integer, Pos_Asunto As Integer, num_f As Integer
Dim Tam_TOTAL As Integer, Tam_Cuerpo As Integer, Pos_Cuerpoi As Integer, Tam_Asunto As Integer

'Copiar portapapeles a String
Call portapapeles

'Obtener fila donde esta la seleccion del Usuario
num_f = num_fila(direccion)

'Definicion palabra clave
frase = "Asunto: "
frase2 = "Enviado el: "
'Definiendo enter y espacio
enter = vbCrLf
nada = ""

'Eliminar espacios ambos lados solo
strClip = Trim(strClip)

' Eliminar cracateres no deseados
strClip = Replace(strClip, ">", nada, 1, -1, vbTextCompare)
strClip = Replace(strClip, "<", nada, 1, -1, vbTextCompare)

'Posicion palabra clave "Asunto: "
Pos_Asuntoi = InStr(1, strClip, frase)

'Posicion 1º ENTER tras palabra clave
Pos_1Enter = InStr(Pos_Asuntoi, strClip, enter)

'Numero caracteres asunto
Tam_Asunto = ((Pos_1Enter - Pos_Asuntoi) - 8)

'Eliminar ENTERS
strClip = Replace(strClip, enter, nada, 1, -1, vbTextCompare)

'Volvemos a calcular la posicion palabra clave sin los enters
Pos_Asunto = InStr(1, strClip, frase)

'Numero TOTAL en StringClipboard
Tam_TOTAL = Len(strClip)

'Cortar Asunto del Mensaje
Asunto = Cortar_texto(strClip, (Pos_Asunto + 8), Tam_Asunto)
ActiveSheet.Cells(num_f, 3) = Asunto

'Num carácteres cuerpo del mensaje
Pos_Cuerpoi = Pos_Asunto + Tam_Asunto + 8
Tam_Cuerpo = Tam_TOTAL - (Pos_Cuerpoi)

'Posicion palabra clave "Enviado el: "
Pos_fecha = InStr(1, strClip, frase2)

'Cortar fecha en string
fecha = Cortar_texto(strClip, (Pos_fecha + 11), 21)

'Validar e insertar fecha
Form_fecha = val_fecha(fecha)
ActiveSheet.Cells(num_f, 1) = Form_fecha

'Validar e insertar tiempo
Form_tiempo = Val_tiempo(fecha)
ActiveSheet.Cells(num_f, 2) = Form_tiempo

'Cortar cuerpo del Mensaje
cuerpo = Cortar_texto(strClip, Pos_Cuerpoi, (Tam_Cuerpo + 1))
ActiveSheet.Cells(num_f, 5) = cuerpo

 On Error Resume Next
  For Each rango In ActiveSheet.Cells(num_f, 5)
   rango.Value = Application.WorksheetFunction.Substitute(Trim(rango.Value), "     ", " ")
   rango.Value = Application.WorksheetFunction.Substitute(Trim(rango.Value), "    ", " ")
   rango.Value = Application.WorksheetFunction.Substitute(Trim(rango.Value), "   ", " ")
   rango.Value = Application.WorksheetFunction.Substitute(Trim(rango.Value), "  ", " ")
  Next
 On Error GoTo 0

Explotacion.Show

End Sub

Public Function portapapeles()
Set MyData = New DataObject
MyData.GetFromClipboard
strClip = MyData.GetText
End Function

Public Function num_fila(direccion As String) As Integer

direccion = ActiveCell.Address(False, False)
direccion = Mid(direccion, 2, 3)
num_fila = CInt(direccion)

End Function

Public Function Cortar_texto(texto2 As String, valor_inicial2 As Integer, Num_Caracteres As Integer) As String

Cortar_texto = Mid(texto2, valor_inicial2, Num_Caracteres)

End Function
Public Function val_fecha(text_fecha As String) As Date

'Corta el dia
text_fecha = Trim(text_fecha)
text_fecha = Cortar_texto(text_fecha, 4, 19)

'Validar fecha as date
val_fecha = DateValue(text_fecha)

End Function
Public Function Val_tiempo(text_tiempo As String) As Date

'Recortar hora
text_tiempo = Right(text_tiempo, 5)
'Validar hora
Val_tiempo = TimeValue(text_tiempo)

End Function
