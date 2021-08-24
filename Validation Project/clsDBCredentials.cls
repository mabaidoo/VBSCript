VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DBCredentials"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
 
Private Const CSTR_PROVIDER As String = "SQLNCLI11"


 
Private mstrServer              As String
Private mstrDBName              As String
Private mstrPassword            As String
Private mstrUserID              As String
Private mstrConnectionString    As String
Private mblnUsedTrustedSecurity As Boolean
 
Public Property Let SetUserId(ByVal Value As String)
    mstrUserID = Value
End Property
 
Public Property Let SetPassword(ByVal Value As String)
    mstrPassword = Value
End Property
 
Public Property Let SetDBName(ByVal Value As String)
    mstrDBName = Value
End Property
 
Public Property Let SetServer(ByVal Value As String)
    mstrServer = Value
End Property
 
Public Property Let SetTrustedSecurity(ByVal Value As Boolean)
    mblnUsedTrustedSecurity = Value
End Property
 
Public Function GetConnectionString()
    If mblnUsedTrustedSecurity Then
        mstrConnectionString = "Provider=" & CSTR_PROVIDER & ";Server=" & mstrServer & ";Initial Catalog=" & mstrDBName & ";Integrated Security=SSPI"
        
    Else
        mstrConnectionString = _
          "Provider=" & CSTR_PROVIDER & ";Data Source=" & mstrServer & ";Initial Catalog=" & mstrDBName & ";User ID='" & mstrUserID & "';Password='" & mstrPassword & "';"
    End If
    GetConnectionString = mstrConnectionString
End Function
 
 
Private Sub Class_Initialize()
    mblnUsedTrustedSecurity = False
End Sub

