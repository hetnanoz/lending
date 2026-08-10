VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ErrorManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const CLASS_NAME As String = "ErrorManager"
Private Const ERROR_LOG_SHEET_NAME As String = "Error_Log"
Private Const ERROR_LOG_COLUMN_COUNT As Long = 6

Private mErrors As Collection
Private mFixedMessage As String
Private mLastDisplayMessage As String

'-------------------------------------------------------------------------------
' Author:        Pawel Ligezka
' Creation date: 2026-08-10
' Parameters:    strClassName - source module or class name
'                strMethodName - source procedure name
'                lngErrorNumber - captured VBA error number
'                strDescription - captured VBA error description
'                arrContext - optional parameter-name string followed by values
' Returns:       Boolean - True when the error was added to the internal queue
' Description:   Stores one error and its optional non-sensitive context.
'-------------------------------------------------------------------------------
Public Function addError(ByVal strClassName As String, _
                         ByVal strMethodName As String, _
                         ByVal lngErrorNumber As Long, _
                         ByVal strDescription As String, _
                         ParamArray arrContext() As Variant) As Boolean
    Const METHOD_NAME As String = "addError"
    Dim arrEntry As Variant
    Dim arrNames As Variant
    Dim colNames As Collection
    Dim errDescription As String
    Dim errNumber As Long
    Dim lngIndex As Long
    Dim strContext As String
    Dim strName As String
    Dim strNames As String
    Dim strValue As String
    Dim vContextItem As Variant
    Dim vName As Variant

    On Error GoTo ErrorHandler

    If mErrors Is Nothing Then
        Set mErrors = New Collection
    End If

    Set colNames = New Collection
    lngIndex = 0

    For Each vContextItem In arrContext
        If lngIndex = 0 Then
            If Not IsObject(vContextItem) And _
               Not IsError(vContextItem) And _
               Not IsNull(vContextItem) And _
               Not IsEmpty(vContextItem) Then

                strNames = CStr(vContextItem)

                If Len(strNames) > 0 Then
                    arrNames = Split(strNames, ";")

                    For Each vName In arrNames
                        colNames.Add Trim$(CStr(vName))
                    Next vName
                End If
            End If
        Else
            strName = "param" & CStr(lngIndex)

            If lngIndex <= colNames.Count Then
                If Len(CStr(colNames(lngIndex))) > 0 Then
                    strName = CStr(colNames(lngIndex))
                End If
            End If

            If IsObject(vContextItem) Then
                strValue = "[" & TypeName(vContextItem) & "]"
            ElseIf IsError(vContextItem) Then
                strValue = "[Error value]"
            ElseIf IsNull(vContextItem) Then
                strValue = "Null"
            ElseIf IsEmpty(vContextItem) Then
                strValue = vbNullString
            Else
                strValue = CStr(vContextItem)
            End If

            If Len(strContext) > 0 Then
                strContext = strContext & "; "
            End If

            strContext = strContext & strName & "=" & strValue
        End If

        lngIndex = lngIndex + 1
    Next vContextItem

    arrEntry = Array(Now, _
                     strClassName, _
                     strMethodName, _
                     lngErrorNumber, _
                     strDescription, _
                     strContext)

    mErrors.Add arrEntry
    mLastDisplayMessage = strDescription
    addError = True

ExitFunction:
    Set colNames = Nothing
    Exit Function

ErrorHandler:
    errNumber = VBA.Err.Number
    errDescription = VBA.Err.Description
    addError = False
    Debug.Print CLASS_NAME & "." & METHOD_NAME & " failed: " & _
                CStr(errNumber) & " - " & errDescription
    Resume ExitFunction
End Function

'-------------------------------------------------------------------------------
' Author:        Pawel Ligezka
' Creation date: 2026-08-10
' Parameters:    ---
' Returns:       ---
' Description:   Writes queued errors to Immediate and Error_Log, then clears them.
'-------------------------------------------------------------------------------
Public Sub save()
    Const METHOD_NAME As String = "save"
    Dim arrEntry As Variant
    Dim arrHeaders As Variant
    Dim arrOutput As Variant
    Dim errDescription As String
    Dim errNumber As Long
    Dim lngColumn As Long
    Dim lngLastRow As Long
    Dim lngNextRow As Long
    Dim lngRow As Long
    Dim lngUsedRow As Long
    Dim wkbLog As Excel.Workbook
    Dim wksCandidate As Excel.Worksheet
    Dim wksLog As Excel.Worksheet

    On Error GoTo ErrorHandler

    If mErrors Is Nothing Then GoTo ExitSub
    If mErrors.Count = 0 Then GoTo ExitSub

    For lngRow = 1 To mErrors.Count
        arrEntry = mErrors(lngRow)
        Debug.Print VBA.Format$(CDate(arrEntry(0)), "yyyy-mm-dd hh:nn:ss") & _
                    " | " & CStr(arrEntry(1)) & _
                    " | " & CStr(arrEntry(2)) & _
                    " | " & CStr(arrEntry(3)) & _
                    " | " & CStr(arrEntry(4)) & _
                    IIf(Len(CStr(arrEntry(5))) > 0, " | " & CStr(arrEntry(5)), vbNullString)
    Next lngRow

    Set wkbLog = ThisWorkbook

    For Each wksCandidate In wkbLog.Worksheets
        If StrComp(wksCandidate.Name, ERROR_LOG_SHEET_NAME, vbTextCompare) = 0 Then
            Set wksLog = wksCandidate
            Exit For
        End If
    Next wksCandidate

    If wksLog Is Nothing Then
        Set wksLog = wkbLog.Worksheets.Add(After:=wkbLog.Worksheets(wkbLog.Worksheets.Count))
        wksLog.Name = ERROR_LOG_SHEET_NAME
    End If

    arrHeaders = Array("Timestamp", _
                       "Module/Class", _
                       "Procedure", _
                       "Error Number", _
                       "Description", _
                       "Context")

    For lngColumn = 1 To ERROR_LOG_COLUMN_COUNT
        If StrComp(Trim$(CStr(wksLog.Cells(1, lngColumn).Value2)), _
                   CStr(arrHeaders(lngColumn - 1)), _
                   vbBinaryCompare) <> 0 Then
            wksLog.Cells(1, lngColumn).Value = arrHeaders(lngColumn - 1)
        End If
    Next lngColumn

    lngLastRow = 1
    For lngColumn = 1 To ERROR_LOG_COLUMN_COUNT
        lngUsedRow = wksLog.Cells(wksLog.Rows.Count, lngColumn).End(xlUp).Row
        If lngUsedRow > lngLastRow Then
            lngLastRow = lngUsedRow
        End If
    Next lngColumn

    If lngLastRow < 1 Then
        lngLastRow = 1
    End If

    lngNextRow = lngLastRow + 1
    ReDim arrOutput(1 To mErrors.Count, 1 To ERROR_LOG_COLUMN_COUNT)

    For lngRow = 1 To mErrors.Count
        arrEntry = mErrors(lngRow)
        For lngColumn = 1 To ERROR_LOG_COLUMN_COUNT
            arrOutput(lngRow, lngColumn) = arrEntry(lngColumn - 1)
        Next lngColumn
    Next lngRow

    wksLog.Cells(lngNextRow, 1).Resize(mErrors.Count, ERROR_LOG_COLUMN_COUNT).Value = arrOutput
    wksLog.Columns(1).NumberFormat = "yyyy-mm-dd hh:mm:ss"
    wksLog.Rows(1).Font.Bold = True
    wksLog.Columns("A:F").AutoFit

    arrEntry = mErrors(mErrors.Count)
    mLastDisplayMessage = CStr(arrEntry(4))
    Set mErrors = New Collection

ExitSub:
    Set wksCandidate = Nothing
    Set wksLog = Nothing
    Set wkbLog = Nothing
    Exit Sub

ErrorHandler:
    errNumber = VBA.Err.Number
    errDescription = VBA.Err.Description
    Debug.Print CLASS_NAME & "." & METHOD_NAME & " failed: " & _
                CStr(errNumber) & " - " & errDescription
    Resume ExitSub
End Sub

'-------------------------------------------------------------------------------
' Author:        Pawel Ligezka
' Creation date: 2026-08-10
' Parameters:    strMessage - optional user-facing message
' Returns:       ---
' Description:   Shows one user-facing error message after the error was saved.
'-------------------------------------------------------------------------------
Public Sub display(Optional ByVal strMessage As String = vbNullString)
    Const METHOD_NAME As String = "display"
    Dim errDescription As String
    Dim errNumber As Long
    Dim strDisplayMessage As String

    On Error GoTo ErrorHandler

    strDisplayMessage = Trim$(strMessage)

    If Len(strDisplayMessage) = 0 Then
        strDisplayMessage = Trim$(mFixedMessage)
    End If

    If Len(strDisplayMessage) = 0 Then
        strDisplayMessage = Trim$(mLastDisplayMessage)
    End If

    If Len(strDisplayMessage) = 0 Then
        strDisplayMessage = "An unexpected error occurred. Please contact support."
    End If

    VBA.MsgBox strDisplayMessage, vbExclamation, "Report Helaba"

ExitSub:
    Exit Sub

ErrorHandler:
    errNumber = VBA.Err.Number
    errDescription = VBA.Err.Description
    Debug.Print CLASS_NAME & "." & METHOD_NAME & " failed: " & _
                CStr(errNumber) & " - " & errDescription
    Resume ExitSub
End Sub

'-------------------------------------------------------------------------------
' Author:        Pawel Ligezka
' Creation date: 2026-08-10
' Parameters:    strMessage - fixed message used by display when no message is passed
' Returns:       ErrorManager - current instance for optional method chaining
' Description:   Sets the default user-facing message for this manager instance.
'-------------------------------------------------------------------------------
Public Function setFixedMessage(ByVal strMessage As String) As ErrorManager
    Const METHOD_NAME As String = "setFixedMessage"
    Dim errDescription As String
    Dim errNumber As Long

    On Error GoTo ErrorHandler

    mFixedMessage = strMessage
    Set setFixedMessage = Me

ExitFunction:
    Exit Function

ErrorHandler:
    errNumber = VBA.Err.Number
    errDescription = VBA.Err.Description
    Debug.Print CLASS_NAME & "." & METHOD_NAME & " failed: " & _
                CStr(errNumber) & " - " & errDescription
    Set setFixedMessage = Me
    Resume ExitFunction
End Function

'-------------------------------------------------------------------------------
' Author:        Pawel Ligezka
' Creation date: 2026-08-10
' Parameters:    ---
' Returns:       ---
' Description:   Clears queued errors without writing them.
'-------------------------------------------------------------------------------
Public Sub clear()
    Const METHOD_NAME As String = "clear"
    Dim errDescription As String
    Dim errNumber As Long

    On Error GoTo ErrorHandler

    Set mErrors = New Collection

ExitSub:
    Exit Sub

ErrorHandler:
    errNumber = VBA.Err.Number
    errDescription = VBA.Err.Description
    Debug.Print CLASS_NAME & "." & METHOD_NAME & " failed: " & _
                CStr(errNumber) & " - " & errDescription
    Resume ExitSub
End Sub
