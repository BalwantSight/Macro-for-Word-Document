    Sub PalabrasNaranja()
    ' Colores de resaltado
    Dim colorAmarillo As Long
    Dim colorNaranjaFluorescente As Long
    colorVioleta = RGB(201, 51, 255) ' Violeta Brillante
    colorNaranja = RGB(255, 102, 0) ' Naranja Brillante
    
    
    ' Palabras a resaltar
    Dim palabrasGenerales() As Variant
    Dim palabrasPersonales() As Variant
    Dim palabrasEspaciales() As Variant
    Dim palabrasNaturales() As Variant
    Dim palabrasObjetos() As Variant
    Dim palabrasPlurales() As Variant
    
    palabrasGenerales = Array("viejo", "nuevo", "árbol", "pájaro", "planta") ' Puedes agregar palabras dentro del paréntesis
    palabrasPersonales = Array("familia", "niño", "adulto", "anciano") ' Puedes agregar palabras dentro del paréntesis
    palabrasEspaciales = Array("piso", "pared", "techo", "jardín", "parque") ' Puedes agregar palabras dentro del paréntesis
    palabrasNaturales = Array("ave", "reptil", "bosque", "oscuro", "pez", "insecto") ' Puedes agregar palabras dentro del paréntesis
    palabrasObjetos = Array("utensilio", "máquina", "espejo", "publicación", "documento") ' Puedes agregar palabras dentro del paréntesis
    palabrasPlurales = Array("animales", "personas", "ciudades", "ropas", "alimentos", "bebidas") ' Puedes agregar palabras dentro del paréntesis
    
    Dim todasLasPalabras As Variant
    todasLasPalabras = JoinArrays(palabrasGenerales, palabrasPersonales, palabrasEspaciales, palabrasNaturales, palabrasObjetos, palabrasPlurales)
    
    ' Mejorar rendimiento desactivando actualizaciones de pantalla
    Application.ScreenUpdating = False
    
    Dim palabra As range
    Dim doc As Document
    Set doc = ActiveDocument
    
   ' Resaltar palabras terminadas en "mente" con el color amarillo
    For Each palabra In doc.Words
        If Right(Trim(palabra.Text), 5) = "mente" Then
            ' Eliminar otros formatos y aplicar el color amarillo
            palabra.Font.Bold = False
            palabra.Font.Italic = False
            palabra.Font.Underline = False
            palabra.Font.Color = colorVioleta
        End If
    Next palabra
    
    ' Resaltar palabras específicas
    Dim buscarPalabra As Variant
    Dim rangoBusqueda As range
    For Each buscarPalabra In todasLasPalabras
        Set rangoBusqueda = doc.Content
        With rangoBusqueda.Find
            .ClearFormatting
            .Text = buscarPalabra
            .Replacement.ClearFormatting
            .Replacement.Font.Color = colorNaranja
            .MatchWholeWord = True
            .Execute Replace:=wdReplaceAll
        End With
    Next buscarPalabra
    
    ' Restaurar actualizaciones de pantalla
    Application.ScreenUpdating = True
    
    MsgBox "Resaltado de Palabras Naranja y Adverbios en ...mente."
End Sub

' Función para combinar arrays
Function JoinArrays(ParamArray arrays() As Variant) As Variant
    Dim totalLength As Long
    Dim arrIndex As Long
    Dim combinedArray() As Variant
    totalLength = 0
    
    ' Calcular longitud total
    For arrIndex = LBound(arrays) To UBound(arrays)
        totalLength = totalLength + UBound(arrays(arrIndex)) - LBound(arrays(arrIndex)) + 1
    Next arrIndex
    
    ' Crear el array combinado
    ReDim combinedArray(totalLength - 1)
    Dim combinedIndex As Long
    combinedIndex = 0
    
    ' Copiar elementos
    For arrIndex = LBound(arrays) To UBound(arrays)
        Dim i As Long
        For i = LBound(arrays(arrIndex)) To UBound(arrays(arrIndex))
            combinedArray(combinedIndex) = arrays(arrIndex)(i)
            combinedIndex = combinedIndex + 1
        Next i
    Next arrIndex
    
    JoinArrays = combinedArray
End Function

