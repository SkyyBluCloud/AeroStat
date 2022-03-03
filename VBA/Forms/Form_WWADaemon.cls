VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_WWADaemon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Form_Close()
log "WWADaemon stopped. Bye!", Name & ".Close"
End Sub

Private Sub Form_Open(cancel As Integer)
log "Starting WWADaemon...", Name & ".Open"
Form.TimerInterval = Nz(DLookup("data", "tblSettings", "key = ""WWARefreshInterval"""), 5) * 1000
End Sub

Public Function updateWWAs() As Variant
'Updates tblWWA; returns collection of IDs from that table that are currently active
'https://owsjet26.us.af.mil/rssRequest/wwa.rss?type=wwa&station=KHST
On Error GoTo errtrap
Const cstrXPath As String = "/rss/channel/item"
Dim strURL As String: strURL = WWAFeed
Dim ctUpdates As New Collection
Dim updates As Integer

Dim RS As DAO.Recordset: Set RS = CurrentDb.OpenRecordset("tblWWA")

    log "Fetching WWAs from JET (" & strURL & ")", Name & ".updateWWAs"
    
    Dim xmldoc As New MSXML2.DOMDocument60
    xmldoc.aSync = False
    xmldoc.Load strURL
    
    If xmldoc Is Nothing Then Exit Function
    log "This update was brought to you by '" & xmldoc.SelectSingleNode("/rss/channel/title").Text & "'!", Name & ".updateWWAs"
    
    Dim xmlSelection As MSXML2.IXMLDOMSelection
    Set xmlSelection = xmldoc.SelectNodes(cstrXPath)
    updates = xmlSelection.length
    Dim wwa, wwaID As String
    Dim result(0 To 2) As String
    Dim xmlElement As MSXML2.IXMLDOMElement: For Each xmlElement In xmlSelection
        With RS
            
            Dim cid: cid = DLookup("ID", "tblWWA", "pubDate = #" & Trim(Replace(Mid(xmlElement.ChildNodes(2).Text, InStr(1, xmlElement.ChildNodes(2).Text, ",") + 1), "GMT", "")) & "#")
            If IsNull(cid) Then
                .AddNew
            End If
            
            Dim c: For Each c In xmlElement.ChildNodes
                
                Select Case c.nodeName
                    Case "pubDate"
                        'Example:
                        result(0) = Trim(Replace(Mid(c.Text, InStr(1, c.Text, ",") + 1), "GMT", ""))
                    Case "description"
                        'Example:
                        result(1) = Replace(Replace(Replace(Replace(c.Text, "&nbsp;", " "), "<![CDATA[ ", ""), vbLf & "<br />" & vbLf & "<br />", vbCrLf & "Event: "), " ]]>", "")
                         
                    Case "title"
                        Dim a As Integer
                        result(2) = c.Text
                        a = InStr(1, result(2), "[")
                        wwa = Mid(result(2), a + 1, InStr(a, result(2), "]") - (a + 1))
                        
                        
                End Select
                'log result, Name & ".updateWWAs"
            Next c
            
            If Not IsNull(cid) Then .edit
            .Fields("pubDate") = result(0)
            .Fields("description") = result(1)
            .Fields("title") = result(2)
            
            !wwa = Nz(wwa, !wwa)
            
            Dim wwaIDX: wwaIDX = InStr(1, result(1), wwa) + Len(wwa) + 1
            !wwaID = Mid(result(1), wwaIDX, InStr(wwaIDX, result(1), " ") - wwaIDX)
            wwaID = !wwaID
            
            Dim valid, validstop, validStr
            valid = InStr(1, result(1), "valid") + 6
            validstop = InStr(valid, result(1), vbCrLf)
            validStr = Mid(result(1), valid, validstop - valid)
            
'                validStr = Replace(validStr, Mid(validStr, InStr(1, validStr, "("), InStr(1, validStr, ") ") + 1 - InStr(1, validStr, "(")), "")
'                DoEvents
'                validStr = Replace(validStr, Mid(validStr, InStr(1, validStr, "("), InStr(1, validStr, ") ") + 1 - InStr(1, validStr, "(")), "")

            Dim sgroup: sgroup = Mid(validStr, InStr(1, validStr, "(") + 1, InStr(1, validStr, ")") - InStr(1, validStr, "(") + 1)
            !start = DateSerial(Year(!pubdate), Month(!pubdate), Left(sgroup, 2)) & _
                        " " & LToZ(getTime4Char(Mid(sgroup, InStr(1, sgroup, "/") + 1, 4)))
                        
            If InStr(1, validStr, "UFN") = 0 Then
                Dim egroup: egroup = Mid(validStr, InStr(1, validStr, "to ") + 3)
                !End = DateSerial(Year(!pubdate), Month(!pubdate), Left(egroup, 2)) & _
                        " " & LToZ(getTime4Char(Mid(egroup, InStr(1, egroup, "Z (") + 6, 4)))
            End If
            
'                !start = DateSerial(Year(!pubdate), Month(!pubdate), Left(validStr, InStr(1, validStr, "/") - 1)) _
'                        & " " & getTime4Char(Mid(validStr, InStr(1, validStr, "/") + 1, 4))
'                !End = DateSerial(Year(!pubdate), Month(!pubdate), Left(validStr, InStr(1, validStr, "/") - 1)) _
'                        & " " & getTime4Char(Mid(validStr, InStr(1, validStr, "/") + 1, 4))
                    
            !sdescription = Mid(result(1), InStr(1, result(1), "Event: ") + 7)
            
            .update
            .Bookmark = .LastModified
            ctUpdates.add !ID
'                !wwa = wwa
'                .update
'            Else
'                .edit
'                !wwa = wwa
'                !wwaID = Mid(result(1), wwaIDX, InStr(wwaIDX, result(1), " ") - wwaIDX)
'                !start = DateSerial(Year(!pubdate), Month(!pubdate), Left(validStr, InStr(1, validStr, "/") - 1)) _
'                        & " " & getTime4Char(Mid(validStr, InStr(1, validStr, "/") + 1, 4))
'                '!End = DateSerial(Year(!pubdate), Month(!pubdate), Mid(validStr, InStr(1, validStr, "to") + 3, 3)) _
'                        & " " & getTime4Char(Mid(validStr, InStr(5, validStr, "/") + 1, 4))
'
'                !sdescription = Mid(result(1), InStr(1, result(1), "Event: ") + 1)
'                ctUpdates.add cid
'                .update
'            End If
            log "Found Weather " & wwa & " #" & wwaID, Name & ".updateWWAs"
        End With
    Next xmlElement
    
    Dim idx As Integer
    Dim r() As String
    Dim ctu: For Each ctu In ctUpdates
        ReDim Preserve r(0 To idx) As String
        r(idx) = ctu
        idx = idx + 1
    Next
    
    CurrentDb.Execute "UPDATE tblWWA SET active = False;"
    If ctUpdates.Count <> 0 Then CurrentDb.Execute "UPDATE tblWWA SET active = True WHERE ID = " & join(r, " OR ID = ")
    
fexit:
    log "Processed " & ctUpdates.Count & " updates.", Name & ".updateWWAs"
    updateWWAs = Array(ctUpdates)
    Exit Function
    Resume Next
errtrap:
    ErrHandler err, Error$, Name & ".updateWWAs"
End Function

Private Sub webBrowser_DocumentComplete(ByVal pDisp As Object, URL As Variant)
log CStr(URL), Name & ".webBrowser_DocumentComplete"
If URL <> "about:blank" Then updateWWAs
End Sub
