Sub ListAllFolders()
    Dim olNamespace As Outlook.NameSpace
    Dim olFolder As Outlook.Folder
    Dim olSubFolder As Outlook.Folder

    Set olNamespace = Application.GetNamespace("MAPI")
    
    ' Loop through all stores (email accounts) in the profile
    For Each olFolder In olNamespace.Folders
        Debug.Print "Account: " & olFolder.Name
        ' Loop through all folders in each store
        For Each olSubFolder In olFolder.Folders
            Debug.Print "  Folder: " & olSubFolder.Name
        Next olSubFolder
    Next olFolder
End Sub
