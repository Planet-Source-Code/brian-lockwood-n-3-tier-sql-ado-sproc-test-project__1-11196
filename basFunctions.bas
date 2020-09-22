Attribute VB_Name = "basCmd2SQL"
Option Explicit

Public Function Cmd2SQL(objCmd As ADODB.Command) As String
'   Takes and ADO Command object and translates it into a SQL string
'   that you can run in Query-Analyzer to get a better error message or use
'   in your application

 Dim strSQL As String
 Dim n      As Integer
 
    '   Take out all extra characters in CommandText
    strSQL = objCmd.CommandText
    strSQL = Replace(strSQL, "?", "")
    strSQL = Replace(strSQL, "{", "")
    strSQL = Replace(strSQL, "}", "")
    strSQL = Replace(strSQL, " ", "")
    strSQL = Replace(strSQL, "call", "")
    strSQL = Replace(strSQL, "(", "")
    strSQL = Replace(strSQL, ")", "")
    strSQL = Replace(strSQL, ",", "")
    strSQL = Replace(strSQL, "=", "")
    
    '   Convert parameter names to SQL @parameters
    For n = 0 To objCmd.Parameters.Count - 1
        If objCmd.Parameters(n).Name <> "RETURN_VALUE" Then
            strSQL = strSQL & " @" & objCmd.Parameters(n).Name & " = " & _
                WrapWithApos(objCmd.Parameters(n)) & ", "
        End If
    Next n
    
    '    Take off trailing comma
    Cmd2SQL = Left(strSQL, Len(RTrim(strSQL)) - 1)
    
End Function
 
Private Function WrapWithApos(prm As ADODB.Parameter) As String
'   Interrogates parameter for special cases then calls the Quote
'   function to wrap the parameter value with quotes if applicable
 Dim strText As String
 
    If IsNull(prm.Value) Then
        strText = "NULL"
    ElseIf IsDate(prm.Value) Then
        strText = "'" & prm.Value & "'"
    Else
        strText = prm.Value
    End If
    
    If prm.Value <> "NULL" Then
        If Quote(prm.Type) = True Then
            strText = "'" & RTrim(strText) & "'"
        End If
    End If
    WrapWithApos = RTrim(strText)
End Function

Private Function Quote(intPrmType As Integer) As Boolean
'   This function determines if a ADO Command Object Parameter should
'   be wrapped with quotes when it is converted to a SQL string or not

'   Input:  Parameter Type as integer
'   Output: Boolean, True - this is a string param and should be wrapped
'                           with quotes
'                    False - this is a numeric param and should not
 Dim bolVarQuote As Boolean
 
    Select Case intPrmType
        Case Is = adNumeric
        Case Is = adVarBinary
        Case Is = adUnsignedTinyInt
        Case Is = adSmallInt
        Case Is = adBoolean
        Case Is = adSingle
        Case Is = adCurrency
        Case Is = adInteger
        Case Is = adDouble
        Case Is = adBinary
        Case Is = adVarBinary
        Case Is = adLongVarBinary
        
        Case Is = adLongVarWChar
            bolVarQuote = True
        Case Is = adVarChar
            bolVarQuote = True
        Case Is = adWChar
            bolVarQuote = True
        Case Is = adDBTimeStamp
            bolVarQuote = True
        Case Else
            bolVarQuote = True
    End Select
    Quote = bolVarQuote
End Function
    
    


