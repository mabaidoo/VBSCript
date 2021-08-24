VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DBInstance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
 
Private objSharedClass As DataBase
 
Public Function GetNewDatabase(Optional ByVal DBName As String) As DataBase
    If objSharedClass Is Nothing Then
        Set objSharedClass = CreateDatabase(DBName)
    End If
    Set GetNewDatabase = objSharedClass
End Function

 
Public Function CloseSharedDatabase() As Boolean
    CloseSharedDatabase = objSharedClass.CloseDatabase
    Set objSharedClass = Nothing
End Function
 
Private Function CreateDatabase(Optional ByVal Name As String) As DataBase
 
    Dim objDatabase      As DataBase
   Dim objDBCredentials As DBCredentials
    
    Set objDBCredentials = New DBCredentials
    Set objDatabase = New DataBase
    
     Dim File As String
     Dim Str As String
 
     File = App.Path & "\Validation.ini"
     
     Str = ReadIni(File, "DataBase", "Server")
    
     objDBCredentials.SetServer = Str
     

    objDBCredentials.SetDBName = "LandsEnd"
    objDBCredentials.SetUserId = "SQL_DataStage"                'really not used
    objDBCredentials.SetPassword = "oanYEafbu45Mz6Yi5hMO"       'Really not Used

    objDBCredentials.SetTrustedSecurity = True
    objDatabase.OpenDatabase objDBCredentials
    Set CreateDatabase = objDatabase
 
End Function
