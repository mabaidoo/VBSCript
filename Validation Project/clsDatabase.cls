VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DataBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
 
Private Const CLNG_DEFAULT_TIMEOUT          As Long = 0
Private Const CINT_PARAMETER_FIELDS_COUNT   As Integer = 5
 
Private moConn                  As New ADODB.Connection
Private mstrConnectionString    As String
Private mstrADOErrors           As String
Private mlngCmdTimeOut          As Long
Private mvarParams()            As Variant
 
Private Enum EnuParameters
    Name = 0
    DataType = 1
    Length = 2
    Value = 3
    Direction = 4
End Enum
 
Public Property Get ConnectionErrors() As String
    ConnectionErrors = mstrADOErrors
End Property
 
Private Sub Class_Initialize()
    mstrADOErrors = Empty
    mlngCmdTimeOut = CLNG_DEFAULT_TIMEOUT
    ReDim mvarParams(CINT_PARAMETER_FIELDS_COUNT, 0)
End Sub
 
Public Function OpenDatabase(ByVal objCred As DBCredentials) As Boolean
 
On Error GoTo ERR_HANDLER:
    OpenDatabase = False
    If (moConn Is Nothing) Or (moConn.State = adStateClosed) Then
        mstrConnectionString = objCred.GetConnectionString
        moConn.Open mstrConnectionString
    End If
    OpenDatabase = True
    
EXIT_HERE:
    Exit Function
    
ERR_HANDLER:
    OpenDatabase = False
    mstrADOErrors = GetADOErrorInformation()
    mstrADOErrors = mstrADOErrors & vbCrLf & "Err Description:" & Err.Description
    MsgBox mstrADOErrors
    Set moConn = Nothing
    GoTo EXIT_HERE
End Function
 
Public Function CloseDatabase() As Boolean
 
On Error GoTo ERR_HANDLER:
    CloseDatabase = False
    If (Not moConn Is Nothing) Or (moConn.State = adStateOpen) Then
        moConn.Close
        Set moConn = Nothing
    End If
    CloseDatabase = True
    
EXIT_HERE:
    Exit Function
    
ERR_HANDLER:
    CloseDatabase = False
    GoTo EXIT_HERE
    
End Function
 
Public Sub ClearParamList()
    ReDim mvarParams(CINT_PARAMETER_FIELDS_COUNT, 0)
End Sub
 
Public Sub AddToParamList(ByVal strParamName As String, _
                          ByVal dblDataType As Double, _
                          ByVal dblDataLength As Double, _
                          ByVal varParamValue As Variant, _
                          ByVal intParamDirection As Integer)
 
    Dim intArrayElement         As Integer
    Dim intArrayBound           As Integer
 
    intArrayBound = UBound(mvarParams(), 2)
    intArrayElement = intArrayBound
    intArrayBound = intArrayBound + 1
    
    ReDim Preserve mvarParams(CINT_PARAMETER_FIELDS_COUNT, intArrayBound)
    mvarParams(EnuParameters.Name, intArrayElement) = strParamName
    mvarParams(EnuParameters.DataType, intArrayElement) = dblDataType
    mvarParams(EnuParameters.Length, intArrayElement) = dblDataLength
    mvarParams(EnuParameters.Value, intArrayElement) = varParamValue
    mvarParams(EnuParameters.Direction, intArrayElement) = intParamDirection
  
End Sub
 
Public Function GetParamValue(ByVal strParamName As String) As Variant
    Dim intIndex As Integer
    For intIndex = LBound(mvarParams, 2) To UBound(mvarParams, 2) - 1
        If mvarParams(EnuParameters.Name, intIndex) = strParamName Then
            GetParamValue = mvarParams(EnuParameters.Value, intIndex)
            Exit Function
        End If
    Next intIndex
End Function
 
Private Function GetADOErrorInformation() As String
    Dim lngErrorCount   As Long
    Dim lngErrorIndex   As Long
    Dim oError          As ADODB.Error
    Dim oErrorColl      As ADODB.Errors
    Dim strErr          As String
    
    If moConn Is Nothing Then
        GoTo EXIT_HERE
    Else
        Set oErrorColl = moConn.Errors
    End If
 
    lngErrorCount = oErrorColl.Count
    If (lngErrorCount > 0) Then
        strErr = "Errors reported by ADO" & vbCrLf
    End If
    For lngErrorIndex = 0 To (lngErrorCount - 1)
        Set oError = oErrorColl.Item(lngErrorIndex)
        With oError
            strErr = strErr & "(" & lngErrorIndex + 1 & ") "
            strErr = strErr & "Error#: " & .Number & vbCrLf
            strErr = strErr & vbTab & "Desc : " & .Description & vbCrLf
            strErr = strErr & vbTab & "Source: " & .Source & vbCrLf
            strErr = strErr & vbTab & "Native Error: " & .NativeError & vbCrLf
            strErr = strErr & vbTab & "SQL State: " & .SQLState & vbCrLf
            strErr = strErr & vbTab & "Help Context: " & .HelpContext & vbCrLf
            strErr = strErr & vbTab & "Help File: " & .HelpFile & vbCrLf
        End With
    Next lngErrorIndex
    GetADOErrorInformation = strErr
EXIT_HERE:
    Set oError = Nothing
    Set oErrorColl = Nothing
End Function
 
Public Function IsConnected() As Boolean
    IsConnected = False
    If Not moConn Is Nothing Then
        IsConnected = IIf(moConn.State = adStateOpen, True, False)
    End If
End Function
 


Public Function GetRecordsetFromStoredProc(ByVal strQueryName As String, _
                                           Optional ByVal intCursorType = adOpenKeyset, _
                                           Optional ByVal intLockType = adLockOptimistic, _
                                           Optional ByVal intCursorLocation = adUseClient) As ADODB.Recordset
 
On Error GoTo ERR_HANDLER
 
    Dim intParam        As Integer
    Dim prmParameter    As ADODB.Parameter
    Dim objCommand      As ADODB.Command
    Dim rstOutput       As ADODB.Recordset
    
    Set objCommand = New ADODB.Command
    Set rstOutput = New ADODB.Recordset
    
    Set GetRecordsetFromStoredProc = Nothing
    mstrADOErrors = Empty
    
    If Not IsConnected Then
        GoTo EXIT_HERE
    End If
    
    With objCommand
        .ActiveConnection = moConn
        .CommandTimeout = mlngCmdTimeOut
        .CommandType = adCmdText 'adCmdStoredProc
        .CommandText = strQueryName
        Set rstOutput = .Execute
    End With

    Set GetRecordsetFromStoredProc = rstOutput
    
EXIT_HERE:
    Set objCommand = Nothing
    Set prmParameter = Nothing
    Set rstOutput = Nothing
    Exit Function
ERR_HANDLER:
    mstrADOErrors = GetADOErrorInformation()
    mstrADOErrors = mstrADOErrors & vbCrLf & "Err Description:" & Err.Description
    Set GetRecordsetFromStoredProc = Nothing
    GoTo EXIT_HERE
End Function

Public Function InsertDataintoTable(ByVal promo As String, _
                                          Sqlstr As String, promo_id As String, BEG_DAT As String, END_DAT As String, User_ID As String, WriteLog As Boolean) As String
 
On Error GoTo ERR_HANDLER
 
    Dim intParam        As Integer
    Dim prmParameter    As ADODB.Parameter
    Dim objCommand      As ADODB.Command
    Dim rstOutput       As String
    Dim strQueryName As String
    Dim strDel As String
    Dim strLog As String
    Dim UserName As String
    strQueryName = ""
    strLog = ""
    Set objCommand = New ADODB.Command
    
    mstrADOErrors = Empty
    
    If Not IsConnected Then
        GoTo EXIT_HERE
    End If
    UserName = "(select NAM_UPR from SY_USR where USR_ID ='" & User_ID & "')"
       If promo = "General" Then
            'Promo Table
            strQueryName = "Insert into USER_IM_PROMO(DIVISION,CATEGORY,LINE,SUB_LIN,PROD_COD,STYL_NUM,CLR_CD,DPT_CD,PRC_METH,AMT_OR_PCT,BEG_DAT,END_DAT,STR_ID,PROMO_ID) values("
            strQueryName = strQueryName & Sqlstr
            'Log table
            strLog = "Insert into USER_IM_PROMO_LOG(USER_ID,USER_NAME,PROMO_ID,PROMO_TYP,BEG_DAT,END_DAT) values("
            strLog = strLog & "'" & User_ID & "'," & UserName & ",'" & promo_id & "','" & promo & "','" & BEG_DAT & "','" & END_DAT & "' )"
        ElseIf promo = "Bogo" Then
          ' Promo Table
            strQueryName = "Insert into USER_IM_BOGO_PROMO(DIVISION,CATEGORY,LINE,SUB_LIN,PROD_COD,STYL_NUM,DPT_CD,BUY_QTY,DISC_QTY,PRC_METH,AMT_OR_PCT,BEG_DAT,END_DAT,STR_ID,PROMO_ID) Values("
            strQueryName = strQueryName & Sqlstr
            'Log table
            strLog = "Insert into USER_IM_PROMO_LOG(USER_ID,USER_NAME,PROMO_ID,PROMO_TYP,BEG_DAT,END_DAT) values("
            strLog = strLog & "'" & User_ID & "'," & UserName & ",'" & promo_id & "','" & promo & "','" & BEG_DAT & "','" & END_DAT & "' )"
        '
        ElseIf promo = "Coupon" Then
            ' Promo Table
            strQueryName = "Insert into USER_IM_COUPON_PROMO(DIVISION,CATEGORY,LINE,SUB_LIN,PROD_COD,STYL_NUM,DPT_CD,COUPON_ID,COUPON_DESCR,BAR_CODE,MIN_QTY,PRC_METH,AMT_OR_PCT,BEG_DAT,END_DAT,STR_ID,PROMO_ID) Values("
            strQueryName = strQueryName & Sqlstr
            'Log table
            strLog = "Insert into USER_IM_PROMO_LOG(USER_ID,USER_NAME,PROMO_ID,PROMO_TYP,BEG_DAT,END_DAT) values("
            strLog = strLog & "'" & User_ID & "'," & UserName & ",'" & promo_id & "','" & promo & "','" & BEG_DAT & "','" & END_DAT & "' )"
        '
        End If
        
       moConn.Execute strQueryName
       If WriteLog Then
           moConn.Execute strLog
       End If
        Sqlstr = ""
     rstOutput = ""

EXIT_HERE:
    Set objCommand = Nothing
    Set prmParameter = Nothing
  rstOutput = mstrADOErrors
InsertDataintoTable = rstOutput

    Exit Function
ERR_HANDLER:
    mstrADOErrors = GetADOErrorInformation()
    mstrADOErrors = mstrADOErrors & vbCrLf & "Err Description:" & Err.Description
     rstOutput = mstrADOErrors
    GoTo EXIT_HERE
End Function
Public Function DeletePromoFromTable(ByVal promo As String, promo_id As String) As String
 
On Error GoTo ERR_HANDLER
 
    Dim intParam        As Integer
    Dim prmParameter    As ADODB.Parameter
    Dim objCommand      As ADODB.Command
    Dim rstOutput       As String
    Dim strQueryName As String
    Dim strDel As String
    Dim strLog As String
    strQueryName = ""
    Set objCommand = New ADODB.Command
    
    mstrADOErrors = Empty
    
    If Not IsConnected Then
        GoTo EXIT_HERE
    End If
       If promo = "General" Then
       'Delete
            strDel = "Delete  from USER_IM_PROMO where  PROMO_ID = '" & promo_id & "'" '
        ElseIf promo = "Bogo" Then
             strDel = "Delete  from USER_IM_BOGO_PROMO where  PROMO_ID = '" & promo_id & "'"
        ElseIf promo = "Coupon" Then
           strDel = "Delete  from USER_IM_COUPON_PROMO where  PROMO_ID = '" & promo_id & "'"
         
        End If
        
       moConn.Execute strDel

EXIT_HERE:
    Set objCommand = Nothing
    Set prmParameter = Nothing
    ' rstOutput = ""
    Exit Function
ERR_HANDLER:
    mstrADOErrors = GetADOErrorInformation()
    mstrADOErrors = mstrADOErrors & vbCrLf & "Err Description in Delete:" & Err.Description
     rstOutput = mstrADOErrors
    GoTo EXIT_HERE
End Function

