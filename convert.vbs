Sub BatchSave()
' Opens each PPT in the target folder and saves as PowerPoint 2007/2010 (.pptx) format

Dim sFolder As String
Dim sPresentationName As String
Dim oPresentation As Presentation

' Select the folder:

Set fDialog = Application.FileDialog(msoFileDialogFolderPicker)
With fDialog
.Title = "Select folder and click OK"
.AllowMultiSelect = False
.InitialView = msoFileDialogViewList
If .Show <> -1 Then
    MsgBox "Cancelled By User", , "List Folder Contents"
    Exit Sub
End If
sFolder = fDialog.SelectedItems.Item(1)
If Right(sFolder, 1) <> "\" Then sFolder = sFolder + "\"
End With

' Make sure the folder name has a trailing backslash
If Right$(sFolder, 1) <> "\" Then
    sFolder = sFolder & "\"
End If

' Are there PPT files there?
If Len(Dir$(sFolder & "*.PPTX")) = 0 Then
    MsgBox "Bad folder name or no PPTX files in folder."
    Exit Sub
End If

' Open and save the presentations
sPresentationName = Dir$(sFolder & "*.PPTX")
While sPresentationName <> ""
    Set oPresentation = Presentations.Open(sFolder & sPresentationName, , , False)
    Call oPresentation.SaveAs(sFolder & sPresentationName & ".wmv", ppSaveAsWMV)
    oPresentation.Close
Wend

MsgBox "DONE"

End Sub