# EXCEL
## Caso #1) Para hacer un código en VBA en el que copies y pegas información de un documento a otro:

Sub CopiaryPegar()
### 1) Declaras las variables:
    Dim LibroOrigen As Workbook
    Dim HojaOrigen As Worksheet
### 2) Estableces la ruta del archivo:
'Ruta
    Dim Ruta As String
    Ruta1 = "  "
### 3) Identificación del libro y hoja origen:
'Datos origen
    Set LibroOrigen1 = Workbooks.Open(Ruta1)
    Set HojaOrigen1 = LibroOrigen1.Worksheets("Hoja 1")
### 4) Libro y Hoja destino:
'Datos Destino
    Set LibroDestino = Workbooks(ThisWorkbook.Name)
    Set HojaDestino = LibroDestino.Worksheets("Hoja 1")
### 5) Copia los datos y los pega en el libro destino
    HojaOrigen.Range("A2:A800").Copy Destination:=HojaDestino.Range("A2:A800")
### 6) Cierra libro y guarda
    Workbooks(LibroOrigen.Name).Close SaveChanges:=False
End Sub
