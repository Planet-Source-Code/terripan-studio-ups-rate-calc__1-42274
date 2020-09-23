Attribute VB_Name = "ModFunc"
Public Zone As String
Public CODCharge As Single, EarlyCharge As Single, SaturdayCharge As Single, Handling As Single, LocalZip As Integer


'FindZone
'Used to Find the Dest. Zone based on Dest. Zip and Service Type
Public Function FindZone(InZip As String, InService As Integer)
Dim TempZoneInfo(0 To 6) As String, ZoneInfo(0 To 6) As String

    InZip = Left(InZip, 3)
    If InService > 5 Then InService = 5
    Open App.Path & "\data\Zones.csv" For Input As #1
        Do Until Left(RawRead, 5) = "Dest."
            Line Input #1, RawRead
        Loop
        Line Input #1, RawRead
        
        Do Until ZipCompare(InZip, ZoneInfo(0))
            Input #1, ZoneInfo(0), ZoneInfo(1), ZoneInfo(2), ZoneInfo(3), ZoneInfo(4), ZoneInfo(5), ZoneInfo(6)
        Loop
    Close #1
    
    FindZone = ZoneInfo(InService + 1)
End Function

'ZipCompare
'Used to Compare a Zip with a zip record in the Zones.txt File. Here a Zip is actualy only
'the first 3 digits of the zip. The zip record can either be zip by itself or two zips
'seperated by "-". Anyting between these zips is valid.
Public Function ZipCompare(InZip As String, UpsZip As String) As Boolean
    If InStr(UpsZip, "-") Then
        If (InZip >= Left(UpsZip, 3)) And (InZip <= Right(UpsZip, 3)) Then ZipCompare = True
    Else
        If InZip = UpsZip Then ZipCompare = True
    End If
End Function


Public Function GetBaseCost(InZone As String, InWeight As String, InService As Integer) As Single
Dim CSVFile As String, RawRead As String, Zones() As String, FoundColumn As Integer

    If InService > 5 Then InService = 5
    
    Select Case InService
        Case 0
            CSVFile = "gndcomm.csv"
        Case 1
            CSVFile = "3dscomm.csv"
        Case 2
            CSVFile = "2da.csv"
        Case 3
            CSVFile = "2dam.csv"
        Case 4
            CSVFile = "1dasaver.csv"
        Case 5
            CSVFile = "1da.csv"
    End Select
    
    Open App.Path & "\data\" & CSVFile For Input As #2
        Do Until Left(RawRead, 6) = "Weight"
            Line Input #2, RawRead
        Loop
        
        Zones = Split(RawRead, ", ")
        For i = 1 To UBound(Zones)
            If InStr(Zones(i), InZone) Then FoundColumn = i
        Next i
        
        Do Until Left(RawRead, Len(InWeight)) = InWeight
            Line Input #2, RawRead
        Loop
        
        Zones = Split(RawRead, ",")
        GetBaseCost = Zones(FoundColumn)
   Close #2
End Function
