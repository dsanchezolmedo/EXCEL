# EXCEL
## Caso #1) Para hacer un código en VBA en el que copies y pegas información de un documento a otro:

Sub CopiaryPegar()
### 1) Declaras las variables:
    Dim LibroOrigen As Workbook
    Dim HojaOrigen As Worksheet
### 2) Estableces la ruta del archivo:
    Dim Ruta As String
    Ruta1 = "C:\Users\dsanchezo\Downloads\VBA code.xlsb"
### 3) Identificación del libro y hoja origen:
    Set LibroOrigen1 = Workbooks.Open(Ruta1)
    Set HojaOrigen1 = LibroOrigen1.Worksheets("Hoja 1")
### 4) Libro y Hoja destino:
    Set LibroDestino = Workbooks(ThisWorkbook.Name)
    Set HojaDestino = LibroDestino.Worksheets("Hoja 1")
### 5) Copia los datos y los pega en el libro destino
    HojaOrigen.Range("A2:A800").Copy Destination:=HojaDestino.Range("A2:A800")
### 6) Cierra libro y guarda
    Workbooks(LibroOrigen.Name).Close SaveChanges:=False
End Sub

### Nota: Siempre inicias un proceso con Sub Nombredelproceso(), y terminas con End Sub.

## Caso #2) Aplicar un filtro antes de copiar y pegar la información:

Sub Filtro()
### 1) Variables:
    Dim LibroOrigen3 As Workbook
    Dim HojaOrigen3 As Worksheet
    Dim HojaOrigen32 As Worksheet
    Dim LibroDestino3 As Workbook
    Dim HojaDestino3 As Worksheet
    
    Dim rangoOrigen As Range
    Dim rangoDestino As Range
    Dim rangoOrigen1 As Range
    Dim rangoDestino1 As Range
    Dim filtro As Range
    
    Dim ultimaFila As Long
    Dim ultimaFila2 As Long
    Dim ultimaFila3 As Long
    Dim ultimaFila4 As Long

### 2) Ruta
    Ruta3 = " "

### 3) 'Datos origen
    Set LibroOrigen3 = Workbooks.Open(Ruta3)
    Set HojaOrigen3 = LibroOrigen3.Worksheets("Hoja 1")
    Set HojaOrigen32 = LibroOrigen3.Worksheets("Hoja 2")

### 4) Datos Destino
    Set LibroOrigen3 = Workbooks.Open(Ruta3)
    Set HojaOrigen3 = LibroOrigen3.Worksheets("Hoja 1")
    Set HojaOrigen32 = LibroOrigen3.Worksheets("Hoja 2")
    Set LibroDestino3 = Workbooks(ThisWorkbook.Name)
    Set HojaDestino3 = LibroDestino3.Worksheets("Hoja 1")

### 5) Se tiene que crear una variable para determinar el rango total de celdas activas después de la aplicación del filtro, dado que este número de filas es cambiante, se utiliza una fórmula para determinar la última fila con información en la hoja origen. Se hace lo mismo para la hoja de destino:
    ultimaFila = HojaOrigen3.Cells(Rows.Count, 1).End(xlUp).Row
    ultimaFila2 = HojaDestino3.Cells(Rows.Count, 1).End(xlUp).Row

### 6)Selecciona el rango origen
    Set rangoOrigen = HojaOrigen3.Range("A1:O" & ultimaFila)
    
### 7) Aplica el filtro al rango origen. *Field indica el # de la columna que sse filtra.
    rangoOrigen.AutoFilter Field:=15, Criteria1:=Array("PENDIENTE", "POR CREAR"), Operator:=xlFilterValues

### 8) Selecciona el rango filtrado y pega en la Hoja destino.
    Set rangoOrigen = HojaOrigen3.Range("A2:B" & ultimaFila).Offset(1).Resize(ultimaFila - 1).SpecialCells(xlCellTypeVisible)
    rangoOrigen.Copy Destination:=HojaDestino3.Range("A2:B" & ultimaFila2)

### 9) Cierra libro y guarda
    Workbooks(LibroOrigen3.Name).Close SaveChanges:=False
End Sub

## **Si deseas pegar información ABAJO de la última fila llena con información, agrega este código después del paso 8:

### 9-b) Nuevas posiciones, ultimaFila3 - selecciona celdas activas con el filtro aplicado. UltimaFila4 - Determina la siguiente fila disponible para pegar información: 

    ultimaFila3 = HojaOrigen32.Cells(Rows.Count, 1).End(xlUp).Row
    ultimaFila4 = HojaDestino3.Range("A2:A" & Rows.Count).Cells.SpecialCells(xlCellTypeBlanks).Row

### 10)Selecciona el rango origen 2
    Set rangoOrigen1 = HojaOrigen32.Range("A1:O" & ultimaFila3)
    
### 11) Aplica el filtro al rango origen 2
    rangoOrigen1.AutoFilter Field:=15, Criteria1:=Array("PENDIENTE",  "POR CREAR"), Operator:=xlFilterValues

### 12) Selecciona el rango filtrado 2
    Set rangoOrigen1 = HojaOrigen32.Range("A2:B" & ultimaFila3).Offset(1).Resize(ultimaFila3 - 1).SpecialCells(xlCellTypeVisible)

### 13) Pega la información:
rangoOrigen1.Copy Destination:=HojaDestino3.Range("A" & ultimaFila4 & ":B" & ultimaFila4)

### 14)Cierra libro y guarda
    Workbooks(LibroOrigen3.Name).Close SaveChanges:=False

End Sub

## Caso #3) Pegar datos como valores:

Sub Pegarcomovalores()
### 1)Variables:
    Dim LibroOrigen4 As Workbook
    Dim HojaOrigen4 As Worksheet
    Dim LibroDestino4 As Workbook
    Dim HojaDestino4 As Worksheet

### 2)Ruta:
    Ruta4 = " "

### 3)Datos origen:
    Set LibroOrigen4 = Workbooks.Open(Ruta4)
    Set HojaOrigen4 = LibroOrigen4.Worksheets("Hoja 1")

### 4)Datos Destino:
    Set LibroDestino4 = Workbooks(ThisWorkbook.Name)
    Set HojaDestino4 = LibroDestino1.Worksheets("Hoja 1")

### 5)Copia la información:
    HojaOrigen4.Range("A2:AH13085").Copy
### 6)Pegar como texto:
    HojaDestino4.Range("A2:AH13085").PasteSpecial xlPasteValues

### 7)Cierra libro y guarda
    Workbooks(LibroOrigen4.Name).Close SaveChanges:=False

End Sub



