Imports Microsoft.Office.Interop
Imports System.IO
Imports Microsoft.Office.Interop.Word
Imports System.Drawing
Imports System.Drawing.Imaging
Imports System.Text

Public Class Form1

    'Guarda el nombre de la carpeta selecionada
    Dim selected As String

    Dim marca As String
    Dim modelo As String


    Public Sub genFicha()
        Dim wrdApp As Word.Application
        Dim wrdDoc As Word.Document
        Dim template As String = CurDir() & "\files\ficha.docx"
        Dim row As Integer

        wrdApp = CreateObject("Word.Application")
        wrdDoc = wrdApp.Documents.Add(template)

        For k = 1 To ListBox1.Items.Count - 1
            wrdDoc.Tables(3).Rows.Add()
        Next



        Dim oMeta As metaData = New metaData
        Dim folderImg As String 'Guarda individualmente la ruta de cada imagen
        Dim photo As Image


        'Determina la fila de la tabla en donde comenzará a insertar datos
        row = 3
        Try
            For i = 0 To (ListBox1.Items.Count) - 1 Step 1
                folderImg = selected & "\" & ListBox1.Items.Item(i).ToString()
                photo = Image.FromFile(folderImg)
                wrdDoc.Tables(3).Cell(row, 1).Range.InsertAfter(ListBox1.Items.Item(i).ToString.Replace(".JPG", ""))
                wrdDoc.Tables(3).Cell(row, 2).Range.InsertAfter(oMeta.getFocal(photo))
                wrdDoc.Tables(3).Cell(row, 3).Range.InsertAfter(oMeta.getExposure(photo))
                wrdDoc.Tables(3).Cell(row, 4).Range.InsertAfter(oMeta.getFPoint(photo))
                wrdDoc.Tables(3).Cell(row, 5).Range.InsertAfter(oMeta.getISO(photo))

                row = row + 1
                photo.Dispose()
            Next
        Catch ex As Exception
            MessageBox.Show("Error insertando datos. Restaure una copia de seguridad de la plantilla.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        Status.Text = "Estado: Informe terminado."

        wrdApp.Visible = True
        wrdApp.Activate()
        wrdApp = Nothing

    End Sub

    'Añadir filas a la tabla principal e insertar las imágenes convertidas al informe
    Public Sub imgToReport()
        Status.Text = "Estado: Generando informe fotográfico..."

        Dim wrdApp As Word.Application
        Dim wrdDoc As Word.Document
        Dim row As Integer
        Dim first As Boolean
        Dim folderConvert As String = CurDir() & "\converted"
        Dim template As String = CurDir() & "\files\template.docx"


        wrdApp = CreateObject("Word.Application")
        wrdDoc = wrdApp.Documents.Add(template)



        'Añade datos extra
        Dim oMeta As metaData = New metaData
        Dim photo As Image
        Dim firstImage As String = selected & "\" & ListBox1.Items.Item(0).ToString()
        photo = Image.FromFile(firstImage)
        marca = oMeta.getCamera(photo)
        modelo = oMeta.getModel(photo)

        wrdDoc.Bookmarks.Item("marcam").Range.Text = marca
        wrdDoc.Bookmarks.Item("modelom").Range.Text = modelo
        wrdDoc.Bookmarks.Item("numfotos").Range.Text = ListBox1.Items.Count


        For k = 1 To ListBox1.Items.Count - 1
            wrdDoc.Tables(3).Rows.Add()
        Next


        row = 1
        first = True

        Try
            For i = 0 To (ListBox1.Items.Count) - 1 Step 2
                wrdDoc.Tables(3).Cell(row, 1).Range.InlineShapes.AddPicture(folderConvert & "\" & ListBox1.Items.Item(i).ToString())
                wrdDoc.Tables(3).Cell(row, 2).Range.InlineShapes.AddPicture(folderConvert & "\" & ListBox1.Items.Item(i + 1).ToString())
                row = row + 2
            Next


            'Agregar nueva página enblanco
            ' wrdApp.Selection.InsertBreak(Word.WdBreakType.wdPageBreak)

        Catch ex As Exception

        End Try


        

        Status.Text = "Estado: Informe terminado."

        wrdApp.Visible = True
        wrdApp.Activate()
        wrdApp = Nothing
    End Sub


    'Selecciona las carpeta en donde se encuentras las fotografias y obtiene una cadena
    Public Sub selectFolder()
        Dim folder As New FolderBrowserDialog
        selected = ""

        If folder.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            selected = folder.SelectedPath
            getFileNames()
            Button5.Enabled = True
            Button2.Enabled = True
        End If

    End Sub


    'Imprime el listado de imagenes en el ListBox
    Public Sub getFileNames()
        Dim folder1 As New DirectoryInfo(selected)
        For Each file As FileInfo In folder1.GetFiles()
            ListBox1.Items.Add(file.Name)
        Next
    End Sub





    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Status.Text = "Estado: Generando informe..."
        convertFiles()
    End Sub



    'Redimensionar imágenes y copiar a nueva carpeta
    Public Sub convertFiles()

        Dim folderConvert As String = CurDir() & "\converted\"
        Try
            Dim newbitmap As Bitmap = New Bitmap(CInt(280), CInt(210))

            For i = 0 To ListBox1.Items.Count - 1
                Dim bitmap As Bitmap = New Bitmap(selected & "\" & ListBox1.Items.Item(i).ToString())
                Dim g As Graphics = Graphics.FromImage(newbitmap)
                g.DrawImage(bitmap, 0, 0, newbitmap.Width, newbitmap.Height)
                newbitmap.Save(folderConvert & ListBox1.Items.Item(i).ToString(), Imaging.ImageFormat.Jpeg)

                bitmap.Dispose()
                g.Dispose()

            Next
        Catch ex As Exception

        End Try

        'Función para pasar las imágenes al informe.
        imgToReport()


    End Sub

  
    Private Sub ListBox1_SelectedValueChanged(sender As Object, e As EventArgs) Handles ListBox1.SelectedValueChanged
        Dim folderImg As String = selected & "\" & ListBox1.SelectedItem.ToString()
        PictureBox1.ImageLocation = folderImg
        Status.Text = "Estado: Mostrando imagen " & ListBox1.SelectedItem.ToString()

        Dim objMetada As metaData = New metaData
        Dim photo As Image = Image.FromFile(folderImg)
        lblISO.Text = objMetada.getISO(photo)
        lblFocal.Text = objMetada.getFocal(photo)
        lblPuntoF.Text = objMetada.getFPoint(photo)
        lblExposure.Text = objMetada.getExposure(photo)
        lblDate.Text = objMetada.getDate(photo)
        lblTime.Text = objMetada.getTime(photo)
        lblFlash.Text = objMetada.getFlash(photo)

        photo.Dispose()
    End Sub

    Private Sub AcercaDeToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AcercaDeToolStripMenuItem.Click
        MessageBox.Show("Aplicación diseñada por Fabián Rodríguez. Contacto: fabianrodriguez85@gmail.com", "Contacto", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    Private Sub SeleccionarCarpetaToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SeleccionarCarpetaToolStripMenuItem.Click
        ListBox1.Items.Clear()
        selectFolder()
        Try
            For Each fichero As String In Directory.GetFiles(CurDir() & "\converted", "*.JPG")
                File.Delete(fichero)
            Next
        Catch ex As Exception
            MessageBox.Show("Hubo un error intentando vaciar la carpeta 'converted'. Elimine archivos manualmente.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End Try
        
    End Sub

   
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Status.Text = "Estado: Generando ficha técnica..."
        genFicha()
    End Sub

    Private Sub SalirToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SalirToolStripMenuItem.Click
        Me.Close()
    End Sub

  
   

   

    Private Sub Button2_MouseDown(sender As Object, e As MouseEventArgs) Handles Button2.MouseDown
        Status.Text = "Estado: Procesando imágenes..."
    End Sub

    'Cuando se cierra el programa, se borran las imagenes de la carpeta converted
    Private Sub Form1_FormClosed(sender As Object, e As FormClosedEventArgs) Handles MyBase.FormClosed
        Try
            For Each fichero As String In Directory.GetFiles(CurDir() & "\converted", "*.JPG")
                File.Delete(fichero)
            Next
        Catch ex As Exception
            MessageBox.Show("Hubo un error intentando vaciar la carpeta 'converted'. Elimine archivos manualmente.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End Try
    End Sub
End Class

Public Class metaData

    'References Property Items
    'https://msdn.microsoft.com/en-us/library/windows/desktop/ms534416%28v=vs.85%29.aspx

    Public Function getISO(photo As Image)
        Dim isoProperty As PropertyItem = photo.GetPropertyItem(&H8827)
        Dim iso As UShort = BitConverter.ToUInt16(isoProperty.Value, 0)
        Return iso
    End Function

    Public Function getFocal(photo As Image)
        Dim focalProperty As PropertyItem = photo.GetPropertyItem(&H920A)
        Dim focal As UInt32 = BitConverter.ToUInt16(focalProperty.Value, 0)
        Dim bfocal As UInt32 = BitConverter.ToUInt16(focalProperty.Value, 4)
        Return Convert.ToInt32(focal / bfocal) & "mm"
    End Function

    Public Function getFPoint(photo As Image)
        Dim pfProperty As PropertyItem = photo.GetPropertyItem(&H829D)
        Dim puntof As UInt32 = BitConverter.ToUInt16(pfProperty.Value, 0)
        Dim b As UInt32 = BitConverter.ToUInt16(pfProperty.Value, 4)
        Return (puntof / b)
    End Function

    Public Function getExposure(photo As Image)
        Dim exposureProperty As PropertyItem = photo.GetPropertyItem(&H829A)
        Dim exposicion As UInt32 = BitConverter.ToUInt16(exposureProperty.Value, 4)
        Dim bexpo As UInt32 = BitConverter.ToUInt16(exposureProperty.Value, 0)
        Return "1/" & Convert.ToInt32(exposicion / bexpo) & "s"
    End Function

    Public Function getFlash(photo As Image)

        '16 = Sin Flash
        '24 = Sin Flash
        '25 = Flash automático
        '9 = Flash Obligatorio

        Dim flashProperty As PropertyItem = photo.GetPropertyItem(&H9209)
        Dim flashStatus As UInt16 = BitConverter.ToInt16(flashProperty.Value, 0)
        Dim flash As String
        If flashStatus = 16 Or flashStatus = 24 Then
            flash = "No"
        Else
            flash = "Sí"
        End If
        Return flash
    End Function

    Public Function getDate(photo As Image)
        Dim dateProperty As PropertyItem = photo.GetPropertyItem(&H9003)
        Dim dateTaken As String = (New UTF8Encoding()).GetString(dateProperty.Value)
        Dim dateTemp As String = dateTaken.Substring(0, 11)
        Return dateTemp
    End Function

    Public Function getTime(photo As Image)
        Dim dateProperty As PropertyItem = photo.GetPropertyItem(&H9003)
        Dim dateTaken As String = (New UTF8Encoding()).GetString(dateProperty.Value)
        Dim dateTemp As String = dateTaken.Substring(10, 10)
        Return dateTemp
    End Function


    Public Function getModel(photo As Image)
        Dim modelProperty As PropertyItem = photo.GetPropertyItem(&H110)
        Dim model As String = (New UTF8Encoding()).GetString(modelProperty.Value)
        Return sTrim(model)
    End Function

    Public Function getCamera(photo As Image)
        Dim cameraProperty As PropertyItem = photo.GetPropertyItem(&H10F)
        Dim camera As String = (New UTF8Encoding()).GetString(cameraProperty.Value)
        Return sTrim(camera)
    End Function


    Function sTrim(s As String) As String
        ' this function trims a string of right and left spaces
        ' it recognizes 0 as a string terminator
        'Extraido de http://www.lansa.com/support/tips/t0002.htm
        Dim i As Integer
        i = InStr(s, Chr(0))
        If (i > 0) Then
            sTrim = Trim(Left(s, i - 1))
        Else
            sTrim = Trim(s)
        End If
    End Function

End Class
