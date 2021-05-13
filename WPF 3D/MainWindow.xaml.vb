#Disable Warning
Imports Microsoft.Win32
Class MainWindow
    ReadOnly listTempStlFiles As New List(Of String)
    Dim title As String

    Private Sub Constructor()
        Dim newModel3D As System.Windows.Media.Media3D.Model3D
        Dim newMOdelImporter As HelixToolkit.Wpf.ModelImporter = New HelixToolkit.Wpf.ModelImporter
        newModel3D = newMOdelImporter.Load("C:\Users\Marchiz\Desktop\gun.stl")
        'Launch the model in Viewport.
        Dim device3D As System.Windows.Media.Media3D.ModelVisual3D = New Media3D.ModelVisual3D
        Dim lights As HelixToolkit.Wpf.DefaultLights = New HelixToolkit.Wpf.DefaultLights

        device3D.Content = newModel3D
        ViewPort3D.Children.Clear()
        ViewPort3D.Children.Add(device3D)
        ViewPort3D.Children.Add(lights)
        ViewPort3D.ShowCoordinateSystem = True
        ViewPort3D.ZoomAroundMouseDownPoint = True
        ViewPort3D.ZoomExtentsWhenLoaded = True
        ViewPort3D.ResetCamera()
        'Camera Position
        'ViewPort3D.CameraController.CameraPosition = New Media3D.Point3D(1, 1, 1)

        'Controls
        ViewPort3D.RotateGesture = New MouseGesture(MouseAction.LeftClick)
        ViewPort3D.PanGesture = New MouseGesture(MouseAction.RightClick)
        ViewPort3D.ResetCameraGesture = New MouseGesture(MouseAction.RightDoubleClick)
        Mouse.OverrideCursor = Cursors.Arrow
    End Sub
    Private Sub MainWindow_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded

    End Sub

    Private Sub MenuOpen_Click(sender As Object, e As RoutedEventArgs) Handles MenuOpen.Click
        ConstructorOpenFile()
    End Sub

    Private Sub MenuExit_Click(sender As Object, e As RoutedEventArgs) Handles MenuExit.Click
        Me.Close()
    End Sub

    Private Sub MenuCube_Click(sender As Object, e As RoutedEventArgs) Handles MenuCube.Click
        If MenuCube.IsChecked = True Then
            MenuCube.IsChecked = True
            ViewPort3D.ShowViewCube = True
        Else
            MenuCube.IsChecked = False
            ViewPort3D.ShowViewCube = False
        End If

    End Sub

    Private Sub MenuFPS_Click(sender As Object, e As RoutedEventArgs) Handles MenuFPS.Click
        If MenuFPS.IsChecked = False Then
            MenuFPS.IsChecked = False
            ViewPort3D.ShowFrameRate = False
        Else
            MenuFPS.IsChecked = True
            ViewPort3D.ShowFrameRate = True
        End If
    End Sub

    Private Sub MenuCS_Click(sender As Object, e As RoutedEventArgs) Handles MenuCS.Click
        If MenuCS.IsChecked = False Then
            MenuCS.IsChecked = False
            ViewPort3D.ShowCoordinateSystem = False
        Else
            MenuCS.IsChecked = True
            ViewPort3D.ShowCoordinateSystem = True
        End If
    End Sub

    Private Sub MenuCI_Click(sender As Object, e As RoutedEventArgs) Handles MenuCI.Click
        If MenuCI.IsChecked = False Then
            MenuCI.IsChecked = False
            ViewPort3D.ShowCameraInfo = False
        Else
            MenuCI.IsChecked = True
            ViewPort3D.ShowCameraInfo = True
        End If
    End Sub
    Private Sub MenuTitle_Click(sender As Object, e As RoutedEventArgs) Handles MenuTitle.Click
        If MenuTitle.IsChecked = False Then
            MenuTitle.IsChecked = False
            ViewPort3D.Title = ""
        Else
            MenuTitle.IsChecked = True
            ViewPort3D.Title = title
        End If
    End Sub

    Private Sub MenuRC_Click(sender As Object, e As RoutedEventArgs) Handles MenuRC.Click
        ViewPort3D.ResetCamera()
    End Sub

    Private Sub MenuCP_Click(sender As Object, e As RoutedEventArgs) Handles MenuCP.Click
        Dim x, y, z As Integer
        x = InputBox("Give x", "Hello")
        y = InputBox("Give y", "Hello")
        z = InputBox("Give z", "Hello")
        ViewPort3D.CameraController.CameraPosition = New Media3D.Point3D(x, y, z)
    End Sub

    Private Sub MenuLA_Click(sender As Object, e As RoutedEventArgs) Handles MenuLA.Click
        Dim x, y, z As Integer
        x = InputBox("Give x", "Hello")
        y = InputBox("Give y", "Hello")
        z = InputBox("Give z", "Hello")
        Dim xyz As New Media3D.Point3D(x, y, z)
        ViewPort3D.LookAt(xyz, 1, 3000)
    End Sub

    Sub ConstructorOpenFile()
        Dim fileName As String
        Dim myDialog As New OpenFileDialog()
        Dim strFilePath As String

        Try

            myDialog.Filter = "Model File|*.stl;*.par;*.psm;*.asm"

            If (myDialog.ShowDialog()) Then

                Mouse.OverrideCursor = Cursors.Wait

                fileName = myDialog.FileName

                Dim newModel3D As System.Windows.Media.Media3D.Model3D

                Dim newMOdelImporter As HelixToolkit.Wpf.ModelImporter = New HelixToolkit.Wpf.ModelImporter

                Dim fileInf As New IO.FileInfo(fileName)

                Dim extn As String = fileInf.Extension

                If extn = ".stl" Or extn = ".STL" Then

                    newModel3D = newMOdelImporter.Load(fileName)
                Else
                    ' Generate STL File from File.
                    Dim objSEApp As SolidEdgeFramework.Application
                    Dim objActiveDoc As SolidEdgeFramework.SolidEdgeDocument = Nothing
                    If Process.GetProcessesByName("edge").Count = 0 Then
                        MessageBox.Show("SE is not running, cannot proceed further..",
                                        "App not running", MessageBoxButton.OK,
                                         MessageBoxImage.Error)
                        Return
                    End If
                    objSEApp = GetObject(, "SolidEdge.Application")
                    If objSEApp Is Nothing Then
                        MessageBox.Show("SE is not running, cannot proceed further",
                                        "App not running", MessageBoxButton.OK,
                                         MessageBoxImage.Error)
                        Return
                    End If
                    objSEApp.Visible = False

                    ' Open the file
                    objActiveDoc = objSEApp.Documents.Open(fileName)
                    If objActiveDoc Is Nothing Then
                        MessageBox.Show("Failed to open SE file, cannot proceed further..",
                                        "File not open", MessageBoxButton.OK,
                                         MessageBoxImage.Error)
                        Return
                    End If

                    ' Generate STL File path
                    strFilePath = IO.Path.GetFileNameWithoutExtension(fileName)
                    strFilePath = IO.Path.GetTempPath & strFilePath & ".stl"

                    If System.IO.File.Exists(strFilePath) Then
                        System.IO.File.Delete(strFilePath)
                    End If

                    'Store STL
                    listTempStlFiles.Add(strFilePath)

                    ' Create STL File.
                    objActiveDoc.SaveAs(strFilePath)
                    objActiveDoc.Close()
                    newModel3D = newMOdelImporter.Load(strFilePath)
                    objSEApp.Visible = True

                End If

                ' Launch the model in Viewport.
                Dim device3D As System.Windows.Media.Media3D.ModelVisual3D = New Media3D.ModelVisual3D
                Dim lights As HelixToolkit.Wpf.DefaultLights = New HelixToolkit.Wpf.DefaultLights
                device3D.Content = newModel3D
                ViewPort3D.Children.Clear()
                ViewPort3D.Children.Add(device3D)
                ViewPort3D.Children.Add(lights)

                ' ViewPort3D.ShowCoordinateSystem = True
                ViewPort3D.ZoomAroundMouseDownPoint = True
                ViewPort3D.ZoomExtentsWhenLoaded = True
                ViewPort3D.ResetCamera()
                ViewPort3D.Title = "File : " + fileName
                title = fileName

                'Controls
                ViewPort3D.RotateGesture = New MouseGesture(MouseAction.LeftClick)
                ViewPort3D.PanGesture = New MouseGesture(MouseAction.RightClick)
                ViewPort3D.ResetCameraGesture = New MouseGesture(MouseAction.RightDoubleClick)
                Mouse.OverrideCursor = Cursors.Arrow
            End If
        Catch ex As Exception
            Mouse.OverrideCursor = Cursors.Arrow
        End Try
    End Sub

End Class
