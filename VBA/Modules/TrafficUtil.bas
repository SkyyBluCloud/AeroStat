Attribute VB_Name = "TrafficUtil"
Option Compare Database
Option Explicit

Public Function getArrDate(ByVal DOF As Date, ByVal ATD As Date, _
                            ByVal ETD As Date, ByVal ETE As Date, _
                            ByVal cETA As Date) As Date
Dim tz As Integer: tz = DLookup("timezone", "tblSettings")

    getArrDate = Format(DateAdd("h", tz, (DOF + (Nz(ATD, ETD) + ETE))), "dd-mmm-yy") & " " & Format(DateAdd("h", tz, cETA), "hh:nn")
    
End Function
