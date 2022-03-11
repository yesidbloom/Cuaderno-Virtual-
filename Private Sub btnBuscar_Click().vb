Private Sub btnBuscar_Click()
    Formulario.Show
End Sub

Private Sub btnEditar_Click()
      activar_Controles (True)
End Sub

Private Sub btnGuardar_Click()
      ubicacion = de.Cells(1, 2)
      Datos.Cells(ubicacion, 1) = txtaprendiz.Text
      Datos.Cells(ubicacion, 2) = txtFicha.Text
      Datos.Cells(ubicacion, 3) = txtPrograma.Text
      activar_Controles (False)
      MsgBox "Se Guardo Correctamente"
      
      Limpiar
      
      
      
      
End Sub

Private Sub btnNuevo_Click()
     de.Cells(2, 2) = de.Cells(2, 2) + 1
     de.Cells(1, 2) = de.Cells(2, 2)
     activar_Controles (True)
     Limpiar
     txtaprendiz.SetFocus
     
       
End Sub


Function activar_Controles(estado)
txtaprendiz.Enabled = estado
txtFicha.Enabled = estado
txtPrograma.Enabled = estado


End Function

 
 Function Limpiar()
 txtaprendiz.Text = Empty
 txtFicha.Text = ""
 txtPrograma.Text = Empty
 
 
 End Function
 
Private Sub CommandButton1_Click()
Aplication.Visible = True

End Sub


Private Sub btnSubirImagen_Click()
archivo = Application.GetOpenFilename("Imágenes(*.Jpg;*.bmp;),*Jpg;.bmp")
Fotografia.Picture = LoadPicture(archivo)

Datos.Cells(de.Cells(1, 2), 4) = archivo


End Sub


    


