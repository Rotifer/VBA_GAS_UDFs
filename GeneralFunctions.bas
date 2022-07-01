Attribute VB_Name = "GeneralFunctions"
Option Explicit

'###################################### General Functions ##########################################
'
' General utility VBA user-defined functions.
'
'###############################################################################################

'Return the URL from an input cell hyperlink, return an empy string if the input cell does not contain a hyperlink.
'When stored in Personal.xlsb call as follows: =PERSONAL.XLSB!url(A2).
Public Function URL(cell As Range) As String
    On Error GoTo Err
    URL = cell.Hyperlinks(1).Address
    Exit Function
Err:
    URL = ""
End Function

' If the input cell contains a date value, return the date as a ISO8601 format string (YYYY-MM-DD)
' When stored in Personal.xlsb call as follows: =PERSONAL.XLSB!DATE_AS_ISO8601_STRING(B7)
Public Function DATE_AS_ISO8601_STRING(cell As Range) As String
    If Not VBA.IsDate(cell.Value) Then
        DATE_AS_ISO8601_STRING = ""
        Exit Function
    End If
    DATE_AS_ISO8601_STRING = VBA.Format(cell.Value, "YYYY-MM-DD")
End Function
