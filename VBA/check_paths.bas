Sub ListAllFolders()
    Dim olNamespace As Outlook.NameSpace
    Dim olFolder As Outlook.Folder
    Dim olSubFolder As Outlook.Folder

    Set olNamespace = Application.GetNamespace("MAPI")
    
    For Each olFolder In olNamespace.Folders
        Debug.Print "Account: " & olFolder.Name
        For Each olSubFolder In olFolder.Folders
            Debug.Print "  Folder: " & olSubFolder.Name
        Next olSubFolder
    Next olFolder
End Sub
