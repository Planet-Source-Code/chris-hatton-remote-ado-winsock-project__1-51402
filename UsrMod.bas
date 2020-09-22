Attribute VB_Name = "ServerMod"
Public Sub ParseRecv(RX As Variant, sck As Winsock)
Dim i As Integer
Dim ProcessUsr, ProcessPass As String
Dim NewJob As String
Dim SaveJob(12) As String
Dim SaveEditJob(14) As String
Dim EditJob As String
Dim GetJobNumber, RealJobNumber, PortNumber As String
Dim DelJob As Long

Debug.Print RX



If RX = "QUIT" Then 'client is disconnecting
            FrmServer.sckServer(sck.Index).SendData "SeeYA" 'disconnect the client
End If

If Mid(RX, 1, 10) = "VerifyUser" Then
        FrmServer.sckServer(sck.Index).SendData "ShowAuthFrm" & sck.Index  ' Shows User Authentication Screen
 
    Else
    
If Mid(RX, 1, 8) = "UserName" Then
        ProcessUsr = Split(RX, "~~")(1)
        ProcessPass = Split(RX, "~~")(2)
            Call Validate(ProcessUsr, ProcessPass, sck)  'Validates the user against the database
    End If: End If
    
If Mid(RX, 1, 8) = "ShowJobs" Then
        Call SendJobs(False, sck)                    'Sends all jobs that are not finished.
Else

If Mid(RX, 1, 17) = "ShowCompletedJobs" Then
        Call SendJobs(True, sck)                   'Sends all jobs that are Completed
                 
End If: End If

        If Mid(RX, 1, 10) = "GetRsCount" Then
                Call RecordCount(sck)                    'Client has requested how many records are in the recordset.
        
        End If

        If Mid(RX, 1, 9) = "ListUsers" Then
                Call ListedUsers(sck)
                                   
        End If


If Mid(RX, 1, 2) = "~@" Then
    NewJob = Mid(RX, 3)
         For i = 0 To 12
            SaveJob(i) = Split(NewJob, "~~")(i)
            
               If Len(SaveJob(12)) Then FrmServer.InitMax = SaveJob(12)
                    Next i
                          Call RsAddNew(SaveJob(0), SaveJob(1), SaveJob(2), SaveJob(3), SaveJob(4), SaveJob(5), _
                          SaveJob(6), SaveJob(7), SaveJob(8), SaveJob(9), SaveJob(10), SaveJob(11), sck)

End If

If Mid(RX, 1, 9) = "JobNumber" Then
    GetJobNumber = Mid(RX, 10)
            PortNumber = Split(GetJobNumber, "~~")(1)          'Recieves the jobnumber from the frmclient
            RealJobNumber = Split(GetJobNumber, "~~")(0)       'listview control and passes to the server
                FrmServer.JobNumber = RealJobNumber            'to query and send the details back.
                FrmServer.InitMax = PortNumber
                Call SendEditJob(sck)
End If

If Mid(RX, 1, 2) = "~%" Then                'Saves the information from the frmeditjob
    EditJob = Mid(RX, 3)
        For i = 0 To 14
            SaveEditJob(i) = Split(EditJob, "~~")(i)
                                            'Send this to the Database to save the current record
        Next i
            If Len(SaveEditJob(14)) Then FrmServer.InitMax = SaveEditJob(14)
            Call RsEditJob(SaveEditJob(0), SaveEditJob(1), SaveEditJob(2), SaveEditJob(3), SaveEditJob(4), SaveEditJob(5), _
                           SaveEditJob(6), SaveEditJob(7), SaveEditJob(8), SaveEditJob(9), SaveEditJob(10), SaveEditJob(11), _
                           SaveEditJob(12), SaveEditJob(13), sck)

End If


If Mid(RX, 1, 12) = "DeleteRecord" Then             'Deletes the Record
    DelJob = Mid(RX, 13, 13)
          Call RsDel(DelJob, sck)
        
End If

End Sub

Public Sub BroadcastRefresh() 'a record has change, refresh every clients screen.
    On Error Resume Next
    Dim i As Integer
    For i = 0 To FrmServer.sckServer.Count
        DoEvents
        FrmServer.sckServer(i).SendData "Refresh"
    Next i
End Sub


Public Function GetsckState(sck As Winsock) As String
If FrmServer.sckServer(sck.Index).State = sckConnected Then GetsckState = "Connected:  "
If FrmServer.sckServer(sck.Index).State = sckClosed Then GetsckState = "Connection Closed: "
If FrmServer.sckServer(sck.Index).State = sckConnecting Then GetsckState = "Connecting: "
If FrmServer.sckServer(sck.Index).State = sckConnectionPending Then GetsckState = "Connection Pending:' "
If FrmServer.sckServer(sck.Index).State = sckBadState Then GetsckState = "Bad State Connection: "
If FrmServer.sckServer(sck.Index).State = sckError Then GetsckState = "Disconnected:  "
If FrmServer.sckServer(sck.Index).State = sckNotConnected Then GetsckState = "Disconnected:  "
If FrmServer.sckServer(sck.Index).State = sckConnectionReset Then GetsckState = "Disconnected:  "
If FrmServer.sckServer(sck.Index).State = 8 Then GetsckState = "Disconnected:  "
End Function

