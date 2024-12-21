Sub RestaurarTextoOriginal()
    Dim doc As Document
    Dim rango As range

    ' Desactivar la actualización de pantalla
    Application.ScreenUpdating = False

    ' Asignar el documento activo
    Set doc = ActiveDocument

    ' Recorrer cada rango en el documento
    For Each rango In doc.StoryRanges
        Do
            ' Eliminar formatos de fuente excepto tamaño
            With rango.Font
                .Bold = False
                .Italic = False
                .Underline = wdUnderlineNone
                .Name = "Calibri" ' Fuente predeterminada (puedes ajustarla)
                .Color = wdColorAutomatic ' Color predeterminado
            End With

            ' Eliminar resaltado sin modificar alineación
            rango.highlightColorIndex = wdNoHighlight

            ' Restablecer párrafo sin afectar la alineación
            With rango.ParagraphFormat
                .LeftIndent = 0 ' Sin sangría izquierda
                .RightIndent = 0 ' Sin sangría derecha
                .SpaceBefore = 0 ' Sin espacio antes
                .SpaceAfter = 0 ' Sin espacio después
                .LineSpacingRule = wdLineSpaceSingle ' Interlineado simple
                ' Alineación actual permanece inalterada
            End With

            ' Ir al siguiente rango, si existe
            Set rango = rango.NextStoryRange
        Loop While Not rango Is Nothing
    Next rango

    ' Reactivar la actualización de pantalla
    Application.ScreenUpdating = True

    ' Notificar que se completó la tarea
    MsgBox "Se ha restaurado el formato del texto manteniendo tamaño y alineación.", vbInformation
End Sub

