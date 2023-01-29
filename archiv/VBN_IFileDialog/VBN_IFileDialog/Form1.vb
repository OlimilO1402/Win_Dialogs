Public Class Form1

    Private Const LockedFolder As String = "C:\Windows"

    Private WithEvents OFD As New FileDialog
    Private WithEvents PFD_Locked As New FileDialog(FileDialog.DialogType.PicFolderDialog)
    Private ReadOnly OpenFileTypes As FileDialog.COMDLG_FILTERSPEC() = New FileDialog.COMDLG_FILTERSPEC(4) {}
    Private ReadOnly SaveFileTypes As FileDialog.COMDLG_FILTERSPEC() = New FileDialog.COMDLG_FILTERSPEC(3) {}

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles Me.Load

        OpenFileTypes(0) = New FileDialog.COMDLG_FILTERSPEC("JPEG Files", "*.jpg")
        OpenFileTypes(1) = New FileDialog.COMDLG_FILTERSPEC("GIF Files", "*.gif")
        OpenFileTypes(2) = New FileDialog.COMDLG_FILTERSPEC("BITMAP Files", "*.bmp")
        OpenFileTypes(3) = New FileDialog.COMDLG_FILTERSPEC("IMAGE Files", "*.bmp; *.gif; *.jpg")
        OpenFileTypes(4) = New FileDialog.COMDLG_FILTERSPEC("All Files", "*.*")

        SaveFileTypes(0) = New FileDialog.COMDLG_FILTERSPEC("JPEG Files", "*.jpg")
        SaveFileTypes(1) = New FileDialog.COMDLG_FILTERSPEC("GIF Files", "*.gif")
        SaveFileTypes(2) = New FileDialog.COMDLG_FILTERSPEC("BITMAP Files", "*.bmp")
        SaveFileTypes(3) = New FileDialog.COMDLG_FILTERSPEC("All Files", "*.*")

        OFD.SetFileTypes(OpenFileTypes)
        OFD.SetFileTypeIndex(3)

        OFD.SetTitle("Dialog Title")
        OFD.SetOkButtonLabel("OK Button Label")
        OFD.SetCancelButtonLabel("Cancel Button Label")
        OFD.SetFileNameLabel("FileName Label")

        OFD.StartVisualGroup(1, "ComboBox:")
        OFD.AddComboBox(100)
        OFD.AddControlItem(100, 101, "ComboBoxItem 1")
        OFD.AddControlItem(100, 102, "ComboBoxItem 2")
        OFD.AddControlItem(100, 103, "ComboBoxItem 3")
        OFD.AddControlItem(100, 104, "ComboBoxItem 4")
        OFD.SetSelectedControlItem(100, 101)
        OFD.EndVisualGroup()

        OFD.StartVisualGroup(2, "Menu:")
        OFD.AddMenu(200, "Menu")
        OFD.AddControlItem(200, 201, "MenuItem 1")
        OFD.AddControlItem(200, 202, "MenuItem 2")
        OFD.AddControlItem(200, 203, "MenuItem 3")
        OFD.AddControlItem(200, 204, "MenuItem 4")
        OFD.EndVisualGroup()

        OFD.StartVisualGroup(3, "RadioButtonList:")
        OFD.AddRadioButtonList(300)
        OFD.AddControlItem(300, 301, "RadioButton 1")
        OFD.AddControlItem(300, 302, "RadioButton 2")
        OFD.AddControlItem(300, 303, "RadioButton 3")
        OFD.AddControlItem(300, 304, "RadioButton 4")
        OFD.SetSelectedControlItem(300, 304)
        OFD.EndVisualGroup()

        OFD.StartVisualGroup(4, "Other Controls:")
        OFD.AddEditBox(400, "EditBox")
        OFD.AddText(401, "Text 1")
        OFD.AddSeparator(402)
        OFD.AddText(403, "Text 2")
        OFD.EndVisualGroup()

        OFD.StartVisualGroup(5, "CheckBoxes:")
        OFD.AddCheckButton(500, "CheckBox 1")
        OFD.AddCheckButton(501, "CheckBox 2")
        OFD.AddCheckButton(502, "CheckBox 3")
        OFD.AddCheckButton(503, "CheckBox 4")
        OFD.EndVisualGroup()

        OFD.AddPushButton(600, "PushButton")
        OFD.MakeProminent(600) ' macht ein Control oder VisualGroup prominent (links neben den Standard Buttons)


        ' Optional:
        OFD.SetOptions(OFD.GetOptions Or FileDialog.FILEOPENDIALOGOPTIONS.FOS_FORCEPREVIEWPANEON)

        ' Systemordner
        ' "::{645FF040-5081-101B-9F08-00AA002F954E}" = Papierkorb
        ' "::{F02C1A0D-BE21-4350-88B0-7367FC96EF3C}" = Netzwerk
        ' "::{031E4825-7B94-4DC3-B131-E946B44C8DD5}" = Bibliotheken
        ' und andere Sytemordner...

        'OFD.SetFolder("C:") '<- oder auch Systemordner
        'OFD.SetDefaultFolder("C:") '<- oder auch Systemordner
        'OFD.SetNavigationRoot("C:")

        ' zusätzlicher Ordner links im TreeView (Projektname mit dem User.ico) mit zwei Unterordner
        OFD.AddPlace("C:\Windows\System32")
        OFD.AddPlace("C:\Windows\SoftwareDistribution")

        ' ohne diese Zeile werden Einstellungen des Dialoges, wie zB. der zuletzt ausgewählte Pfad, Standardmäßig für diese Anwendung gespeichert oder
        ' mit dieser Zeile und GUID = Einstellungen des Dialoges, wie der zuletzt ausgewählte Pfad, werden Standardmäßig für diese Anwendung zu dieser GUID gespeichert
        ' und stehen nach einem erneutem Start der Applikation dem Dialog wieder zu Verfügung
        'OFD.SetClientGuid([hier eine erzeugte GUID eintragen]) 

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        OFD.SetEditBoxText(400, "EditBox")

        If OFD.Show Then

            Debug.Print(OFD.GetResult)
            'Debug.Print(OFD.GetResult(FileDialog.SIGDN.SIGDN_NORMALDISPLAY))
            Debug.Print(OFD.GetFolder)
            'Debug.Print(OFD.GetFolder(FileDialog.SIGDN.SIGDN_NORMALDISPLAY))

        End If

        ' Optional:
        ' löscht die gespeicherten Einstellungen dieses Dialoges wie den zuletzt gewählten Pfad
        ' entwerder Standard oder für die GUID (SetClientGuid)
        OFD.ClearClientData()

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        Using OFD1 As New FileDialog

            OFD1.SetFileTypes(OpenFileTypes)
            OFD1.SetFileTypeIndex(3)

            If OFD1.Show Then

                Debug.Print(OFD1.GetResult)

            End If

        End Using

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        Using SFD As New FileDialog(FileDialog.DialogType.SaveFileDialog)

            SFD.SetFileTypes(SaveFileTypes)
            SFD.SetFileTypeIndex()

            SFD.SetFileName("NewFile")
            SFD.SetDefaultExtension()

            If SFD.Show Then

                Debug.Print(SFD.GetResult)

            End If

        End Using

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click

        Using PFD As New FileDialog(FileDialog.DialogType.PicFolderDialog)

            If PFD.Show Then

                Debug.Print(PFD.GetResult)

            End If

        End Using

    End Sub


    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click

        PFD_Locked.SetTitle("Hinweis: Es können nur Ordner und Unterordner von '" & LockedFolder & "' ausgwählt werden!")

        PFD_Locked.SetNavigationRoot(LockedFolder)

        If PFD_Locked.Show Then

            Debug.Print(PFD_Locked.GetResult)

        End If

    End Sub

    Private Sub PFD_Locked_FolderChanging(Folder As String) Handles PFD_Locked.FolderChanging

        ' hier setzen wir einfach wieder den LockedFolder wenn es nicht der Ordner selbst
        ' oder ein Unterordner von LockedFolder ist

        If Folder.Length < LockedFolder.Length Then

            PFD_Locked.SetFolder(LockedFolder)

        ElseIf Folder.Substring(0, LockedFolder.Length) <> LockedFolder Then

            PFD_Locked.SetFolder(LockedFolder)

        End If

    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click

        Using OFD1 As New FileDialog

            Dim EditorFileTypes As FileDialog.COMDLG_FILTERSPEC() = New FileDialog.COMDLG_FILTERSPEC(1) {}
            EditorFileTypes(0) = New FileDialog.COMDLG_FILTERSPEC("Textdateien", "*.txt")
            EditorFileTypes(1) = New FileDialog.COMDLG_FILTERSPEC("Alle Dateien", "*.*")

            OFD1.SetFileTypes(EditorFileTypes)
            OFD1.SetFileTypeIndex()

            OFD1.StartVisualGroup(1, "Codierung:")
            OFD1.AddComboBox(100)
            OFD1.AddControlItem(100, 101, "Automatische Erkennung")
            OFD1.AddControlItem(100, 102, "Ansi")
            OFD1.AddControlItem(100, 103, "UTF-16 LE")
            OFD1.AddControlItem(100, 104, "UTF-16 BE")
            OFD1.AddControlItem(100, 105, "UTF-8")
            OFD1.AddControlItem(100, 106, "UTF-8 mit BOM")
            OFD1.SetSelectedControlItem(100, 101)
            OFD1.EndVisualGroup()

            If OFD1.Show Then

                Debug.Print(OFD1.GetResult)

            End If

        End Using

    End Sub

    '----==== Events IFileDialogEvents ====----
    Private Sub OFD_FileOK() Handles OFD.FileOK
        Debug.Print("FileOK")
    End Sub

    Private Sub OFD_FolderChange() Handles OFD.FolderChange
        Debug.Print("FolderChange")
    End Sub

    Private Sub OFD_FolderChanging(Folder As String) Handles OFD.FolderChanging
        Debug.Print("FolderChanging: " & Folder)
    End Sub

    Private Sub OFD_Overwrite(Name As String, Response As FileDialog.FDE_OVERWRITE_RESPONSE) Handles OFD.Overwrite
        Debug.Print("Overwrite: " & Name & " Response: " & Response.ToString)
    End Sub

    Private Sub OFD_SelectionChange() Handles OFD.SelectionChange
        Debug.Print("SelectionChange: " & OFD.GetCurrentSelection)
    End Sub

    Private Sub OFD_ShareViolation(Name As String, Response As FileDialog.FDE_SHAREVIOLATION_RESPONSE) Handles OFD.ShareViolation
        Debug.Print("ShareViolation: " & Name & " Response: " & Response.ToString)
    End Sub

    Private Sub OFD_TypeChange() Handles OFD.TypeChange
        Debug.Print("TypeChange")
    End Sub


    '----==== Events IFileDialogControlEvents ====----
    Private Sub OFD_ButtonClicked(CtlID As Integer) Handles OFD.ButtonClicked
        Debug.Print("ButtonClicked CtlID: " & CtlID.ToString)
    End Sub

    Private Sub OFD_CheckButtonToggled(CtlID As Integer, Checked As Boolean) Handles OFD.CheckButtonToggled
        Debug.Print("CheckButtonToggled CtlID: " & CtlID.ToString & " Checked: " & Checked.ToString)
    End Sub

    Private Sub OFD_ControlActivating(CtlID As Integer) Handles OFD.ControlActivating
        Debug.Print("ControlActivating CtlID: " & CtlID.ToString)
    End Sub

    Private Sub OFD_ItemSelected(CtlID As Integer, ItemID As Integer) Handles OFD.ItemSelected
        Debug.Print("ItemSelected CtlID: " & CtlID.ToString & " ItemID: " & ItemID.ToString)

        If CtlID = 300 Then
            Select Case ItemID
                Case 301
                    OFD.SetEditBoxText(400, "Option 1")
                Case 302
                    OFD.SetEditBoxText(400, "Option 2")
                Case 303
                    OFD.SetEditBoxText(400, "Option 3")
                Case 304
                    OFD.SetEditBoxText(400, "Option 4")
            End Select
        End If
    End Sub

End Class
