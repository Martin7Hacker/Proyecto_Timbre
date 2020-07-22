Attribute VB_Name = "Inicio"
'***************************************************************************
'*
'*
'* Iniciar Programa con Virtual Martin temporize v1.7
'*
'*
'***************************************************************************
Public Sub Inicio() 'inicio del programa para que se pueda
                     'configurar previamente y luego
                     'armarse para despues crearlo
                     'graficamente
 Lenguage.definir_lenguage_opciones 'carga el lenguage previo
 frmprograma.Show                   'carga el programa
End Sub




Public Sub cargar_Skins(ByVal formulario As Form)
'cargar mascara
On Error GoTo nose
frmprograma.Skin1.LoadSkin App.Path & "\Skins\" _
& frmprograma.File1.FileName
frmprograma.Skin1.ApplySkin formulario.hwnd
nose:
End Sub


Public Sub cargar_Skins_Picture(ByVal Picture As PictureBox)
'cargar mascara
On Error GoTo nose
frmprograma.Skin1.LoadSkin App.Path & "\Skins\" _
& frmprograma.File1.FileName
frmprograma.Skin1.ApplySkin Picture.hwnd
nose:
End Sub


