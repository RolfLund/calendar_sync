Sub SyncCalendars()
    Dim srcCalendar As Outlook.Folder
    Dim destCalendar As Outlook.Folder
    Dim srcItem As AppointmentItem
    Dim destItem As AppointmentItem
    Dim destItems As Items
    Dim itemExists As Boolean
    Dim i As Integer

    Set srcCalendar = Outlook.Session.Folders("ADD_NAME_OF_INTERNET_CALENDARS_HERE").Folders("ADD_NAME_OF_INTERNET_FOLDER_HERE")
    Set destCalendar = Outlook.Session.Folders("ADD_YOUR_E-MAIL_HERE").Folders("ADD_THE_NAME_OF_YOUR_MAIN_CALENDAR_HERE")
    Set destItems = destCalendar.Items

    For Each srcItem In srcCalendar.Items
        itemExists = False

        For i = destItems.Count To 1 Step -1
            Set destItem = destItems.Item(i)
            If destItem.Subject = srcItem.Subject Then
                With destItem
                    .Start = srcItem.Start
                    .End = srcItem.End
                    .Location = srcItem.Location
                    .Body = srcItem.Body
                    .Save
                End With
                itemExists = True
                Exit For
            End If
        Next i

        If Not itemExists Then
            Set destItem = destCalendar.Items.Add(olAppointmentItem)
            With destItem
                .Subject = srcItem.Subject
                .Start = srcItem.Start
                .End = srcItem.End
                .Location = srcItem.Location
                .Body = srcItem.Body
                .Save
            End With
        End If
    Next srcItem
End Sub
