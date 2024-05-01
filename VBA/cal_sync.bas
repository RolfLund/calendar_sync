Sub SyncCalendars()
    Dim srcCalendar As Outlook.Folder
    Dim destCalendar As Outlook.Folder
    Dim srcItem As AppointmentItem
    Dim destItem As AppointmentItem
    Dim destItems As Items
    Dim itemExists As Boolean
    Dim i As Integer

    Set srcCalendar = Outlook.Session.Folders("Internet Calendars").Folders("Undervisning")
    Set destCalendar = Outlook.Session.Folders("rolfll@socsci.aau.dk").Folders("Kalender")
    Set destItems = destCalendar.Items

    For Each srcItem In srcCalendar.Items
        itemExists = False

        For i = destItems.Count To 1 Step -1
            Set destItem = destItems.Item(i)
            If destItem.Subject = srcItem.Subject Then ' Matching only by title
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
