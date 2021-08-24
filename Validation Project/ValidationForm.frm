VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form ValidationForm 
   Caption         =   "Promo Validation"
   ClientHeight    =   11025
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   17325
   Icon            =   "ValidationForm.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   11025
   ScaleWidth      =   17325
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ProgressBar ValProg 
      Height          =   375
      Left            =   4320
      TabIndex        =   16
      Top             =   2880
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.CommandButton Quit 
      Caption         =   "Exit"
      Height          =   855
      Left            =   15000
      TabIndex        =   14
      Top             =   7920
      Width           =   1695
   End
   Begin MSComDlg.CommonDialog FileLoad 
      Left            =   16320
      Top             =   9120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Btn_Move 
      Caption         =   "Move to Counterpoint"
      Enabled         =   0   'False
      Height          =   375
      Left            =   14400
      TabIndex        =   13
      Top             =   4680
      Width           =   2295
   End
   Begin VB.CommandButton Btn_Print 
      Caption         =   "Preview / Print"
      Height          =   375
      Left            =   14400
      TabIndex        =   12
      Top             =   4080
      Width           =   2295
   End
   Begin VB.TextBox Errors 
      ForeColor       =   &H000000FF&
      Height          =   5895
      Left            =   1080
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   10
      Top             =   3360
      Width           =   12975
   End
   Begin VB.CommandButton Btn_Validate 
      Caption         =   "Validate Promo"
      Height          =   495
      Left            =   12480
      TabIndex        =   9
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "Format"
      Height          =   735
      Left            =   8520
      TabIndex        =   5
      Top             =   2040
      Width           =   3735
      Begin VB.OptionButton Coupon_Option 
         Caption         =   "Coupon"
         Height          =   255
         Left            =   2520
         TabIndex        =   8
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton Bogo_Option 
         Caption         =   "Bogo"
         Height          =   255
         Left            =   1560
         TabIndex        =   7
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton General_Option 
         Caption         =   "General"
         Height          =   255
         Left            =   480
         TabIndex        =   6
         Top             =   240
         Value           =   -1  'True
         Width           =   975
      End
   End
   Begin VB.CommandButton Btn_GetFile 
      Caption         =   "Browse"
      Height          =   495
      Left            =   6840
      TabIndex        =   4
      Top             =   2160
      Width           =   1455
   End
   Begin VB.TextBox FileName 
      Height          =   495
      Left            =   1440
      TabIndex        =   2
      Top             =   2160
      Width           =   5295
   End
   Begin VB.PictureBox Picture1 
      Height          =   1365
      Left            =   240
      Picture         =   "ValidationForm.frx":9647
      ScaleHeight     =   1305
      ScaleWidth      =   1785
      TabIndex        =   1
      Top             =   280
      Width           =   1845
   End
   Begin VB.Label versiondetails 
      Caption         =   "Label1"
      Height          =   375
      Left            =   14280
      TabIndex        =   15
      Top             =   10080
      Width           =   2175
   End
   Begin VB.Label Lbl_Errors 
      Caption         =   "Validation Messages"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   3000
      Width           =   2175
   End
   Begin VB.Label LblFilename 
      Alignment       =   1  'Right Justify
      Caption         =   "File:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   3
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label LblHeader 
      Caption         =   "Promotions Validation"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5040
      TabIndex        =   0
      Top             =   480
      Width           =   6375
   End
   Begin VB.Line Line4 
      X1              =   16680
      X2              =   16680
      Y1              =   240
      Y2              =   1650
   End
   Begin VB.Line Line3 
      X1              =   120
      X2              =   120
      Y1              =   240
      Y2              =   1650
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   16680
      Y1              =   1650
      Y2              =   1650
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   16680
      Y1              =   240
      Y2              =   240
   End
End
Attribute VB_Name = "ValidationForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Bogo_Option_Click()
    'Set Bogo Option to True
    If Bogo_Option.Value = True Then
        General_Option.Value = False
        Coupon_Option.Value = False
    End If
End Sub
Private Sub Btn_Print_Click()
    'Open the Print Preview window
    PrintPreview.Show vbModal
End Sub

Private Sub Coupon_Option_Click()
    'Set the Coupon Option to True
    If Coupon_Option.Value = True Then
        General_Option.Value = False
        Bogo_Option.Value = False
    End If
End Sub

Private Sub General_Option_Click()
    'Set the General Option to True
    If General_Option.Value = True Then
        Bogo_Option.Value = False
        Coupon_Option.Value = False
    End If
End Sub

Private Sub Btn_GetFile_Click()
    'This Subroutine gets the file name to be validated
    
    On Error GoTo Exit_GetFile
    'clear the errors
    Errors.Text = ""
    
    'Disable the move button
    Btn_Move.Enabled = False
    
    'Get the file name with full path
    FileLoad.DialogTitle = "Select Promo File to Validate"
    FileLoad.Filter = "CSV Files (*.csv)|*.csv|All " & "Files (*.*)|*.*"
    FileLoad.ShowOpen
    
    If FileLoad.CancelError Then GoTo Exit_GetFile
    
    If FileLoad.FileName <> "" Then
        
        ' save file name to global variable
        GVars.Promo_filename_full = FileLoad.FileName
        
        'Check to see if it is a csv file
        If (InStr(LCase(GVars.Promo_filename_full), LCase(".csv")) = 0) Then
            MsgBox "Only CSV files supported", 16, "Get Promo File Name"
        End If
        
        ' split the file name into name only and name without extension
        GVars.Promo_filename = Right(GVars.Promo_filename_full, Len(GVars.Promo_filename_full) - InStrRev(GVars.Promo_filename_full, "\"))
        GVars.Promo_filename_noext = Left(GVars.Promo_filename, InStr(GVars.Promo_filename, ".csv") - 1)
        GVars.Promo_filepath = Left(GVars.Promo_filename_full, (Len(GVars.Promo_filename_full) - Len(GVars.Promo_filename)))
        'Show the file name on the screen
        FileName.Text = GVars.Promo_filename
    End If
    ' create log file directory if is doesn't exist.
    If Dir(GVars.Log_file_path, vbDirectory) = vbNullString Then
        MkDir GVars.Log_file_path
    End If
    'Set the log file name and path
    GVars.Log_FileName = GVars.Log_file_path + GVars.Promo_filename_noext + ".log"
  

Exit_GetFile:

End Sub

Private Sub Btn_Validate_Click()

    'This subroutine does the actual data validation.  It loops through the promo file data one row at a time.

    On Error GoTo Validate_Error
    'Clear the Errors
    Errors.Text = ""
    'Disable the move button
    Btn_Move.Enabled = False
    
    WeHaveErrors = False
    Success = False
    
    
    Dim Barcode_Str As String
    Dim strQueryName As String
    Dim rstData_store  As New ADODB.Recordset
    Dim rstData1     As New ADODB.Recordset
    
     
    'get a list of stores from db to compare againg promo file
    
    strQueryName = "SELECT STUFF("
    strQueryName = strQueryName & " (SELECT"
    strQueryName = strQueryName & " ', ' + cast(s2.str_no as varchar(200))"
    strQueryName = strQueryName & " FROM PS_STR s2"
    strQueryName = strQueryName & " ORDER BY s2.str_no"
    strQueryName = strQueryName & " FOR XML PATH(''), TYPE"
    strQueryName = strQueryName & " ).value('.','varchar(max)')"
    strQueryName = strQueryName & " ,1,2, ''"
    strQueryName = strQueryName & " ) AS Str_no"
    
    Set rstData_store = oDatabase.GetRecordsetFromStoredProc(strQueryName)
    
    rstData_store.MoveFirst
    Do While Not rstData_store.EOF
      Store_db = Store_db & "|" & rstData_store.Fields("Str_no").Value
        rstData_store.MoveNext
    Loop
    
    Set rstData_store = Nothing
    
    
    ' Make sure we have a file
    If GVars.Promo_filename_full = vbNullString Then
        MsgBox "You Must Specify a Valid File Name", 33, "Promo Validation"
        GoTo Finish:
    End If
    
    ' check to see if the file matchs the format selected
      
    If Not Check_Format() Then
        
        MsgBox "Please choose correct promo validation file format.", 33, "Promo Validation"
        GoTo Finish
        
    End If
    
    
    ' check to see if the file has been uploaded already

    If Check_MovedtoCP() Then
        Dim ReplaceAnsw As Integer
        ReplaceAnsw = MsgBox("Promo already moved to Counterpoint.Do you wish to Continue?", 36, "Promo Validation")
        
        If ReplaceAnsw = 6 Then Replace_Promo = True
        
        If Not Replace_Promo Then GoTo Finish  'if uploaded and not replacing exit
        
    End If


    ' Create the correct schema file for reading the csv
    
    Create_SchemaFile GVars.Promo_filename, GVars.Promo_filepath
    
    
    
    ' prepare to open the CSV File
    
    Dim CsvConStr As String
    Dim CsvCon As New ADODB.Connection
    Dim CsvRec As New ADODB.Recordset
    
    
    
    
' setup index into  record set

    GVars.Division = 0
    GVars.Category = 0
    GVars.Line = 0
    GVars.Sub_Line = 0
    GVars.ProductCode = 0
    GVars.StyleNum = 0
    GVars.ColorCode = 0
    GVars.DeptCode = 0
    GVars.PriceMethod = 0
    GVars.AmtorPct = 0
    GVars.StartDate = 0
    GVars.EndDate = 0
    GVars.Store = 0
    GVars.promo_id = 0
    GVars.CouponID = 0
    GVars.CouponDescr = 0
    GVars.Barcode = 0
    GVars.MinQty = 0
    GVars.BuyQty = 0
    GVars.DiscQty = 0
    
    
    If GVars.promo_format = "General" Then
        Division = 0
        Category = 1
        GVars.Line = 2
        Sub_Line = 3
        ProductCode = 4
        StyleNum = 5
        ColorCode = 6
        DeptCode = 7
        PriceMethod = 8
        AmtorPct = 9
        StartDate = 10
        EndDate = 11
        Store = 12
        promo_id = 13
    
    ElseIf GVars.promo_format = "Bogo" Then
        Division = 0
        Category = 1
        GVars.Line = 2
        Sub_Line = 3
        ProductCode = 4
        StyleNum = 5
        DeptCode = 6
        BuyQty = 7
        DiscQty = 8
        PriceMethod = 9
        AmtorPct = 10
        StartDate = 11
        EndDate = 12
        Store = 13
        promo_id = 14
    
    ElseIf GVars.promo_format = "Coupon" Then
        Division = 0
        Category = 1
        GVars.Line = 2
        Sub_Line = 3
        ProductCode = 4
        StyleNum = 5
        DeptCode = 6
        CouponID = 7
        CouponDescr = 8
        Barcode = 9
        MinQty = 10
        PriceMethod = 11
        AmtorPct = 12
        StartDate = 13
        EndDate = 14
        Store = 15
        promo_id = 16
    
    End If
    
    'setup connection string to the Promo file
    
    CsvConStr = "Driver={Microsoft Text Driver (*.txt; *.csv)};Dbq=" & GVars.Promo_filepath & ";Extensions=txt,csv;HDR=YES"
    
    'Open the connection to the file path
    
    CsvCon.Open CsvConStr
    
    'Open the file and read the records
    
    CsvRec.Open "SELECT * FROM [" & GVars.Promo_filename & "]", CsvCon, adOpenStatic, adLockReadOnly, adCmdText
    
    'Make sure you are at the firs record
    
    CsvRec.MoveFirst
    
    Dim CsvLine As Integer
    Dim LineHasErrors As Boolean
    
    'Set staring line number.
    csvline_no = 1
    
    'Re-Set the progress bar
    ValProg.Value = 0
    
    Do While Not CsvRec.EOF
    
        'set line error flag
        LineHasErrors = False
        
        ' Paint the Progress Bar
        If ValProg.Value <> ValProg.Max Then
            ValProg.Value = ValProg.Value + 1
        End If
        
 ' --------------------------------------------------------------------
 'Validate Division
 
        Set rstData1 = oDatabase.GetRecordsetFromStoredProc("SELECT top 1 categ_cod FROM IM_ITEM where categ_cod = '" & CsvRec.Fields(Division) & "'")
        
        If Len(CsvRec.Fields(Division)) > 2 Or (CsvRec.Fields(Division) <> "" And rstData1.EOF) Then
            Errors.Text = Errors.Text & "Line# " & csvline_no & " Division " & CsvRec.Fields(Division) & " is invalid." & vbCrLf
            WeHaveErrors = True
            LineHasErrors = True
        End If
'
'-------------------------------------------------------------------------
'Validate Category
       
       Set rstData1 = oDatabase.GetRecordsetFromStoredProc("SELECT top 1 subcat_cod FROM IM_ITEM where subcat_cod = '" & CsvRec.Fields(Category) & "'")
       
       If CsvRec.Fields(Category) <> "" And rstData1.EOF Then
           Errors.Text = Errors.Text & "Line# " & csvline_no & " Category " & CsvRec.Fields(Category) & " is invalid." & vbCrLf
           WeHaveErrors = True
            LineHasErrors = True
       End If
       
       If promo_format = "Coupon" Then
            If CsvRec.Fields(Category) <> "" And IsNull(CsvRec.Fields(Division)) Then
                Errors.Text = Errors.Text & "Line# " & csvline_no & " Category " & CsvRec.Fields(Category) & " Must be combined with a Division for Coupons." & vbCrLf
                WeHaveErrors = True
                LineHasErrors = True
            End If
       End If
       
'
'-------------------------------------------------------------------------
'Validate Line
       
       If (IsNull(CsvRec.Fields(Division)) Or IsNull(CsvRec.Fields(Category))) And CsvRec.Fields(GVars.Line) <> "" Then
           Errors.Text = Errors.Text & "Line# " & csvline_no & " Line of " & CsvRec.Fields(GVars.Line) & " and a blank Division or Category is invalid." & vbCrLf
           WeHaveErrors = True
            LineHasErrors = True
       End If
       
       
       Set rstData1 = oDatabase.GetRecordsetFromStoredProc("SELECT top 1 attr_cod_1 FROM IM_ITEM where attr_cod_1 = '" & CsvRec.Fields(GVars.Line) & "'")

       If (Len(CsvRec.Fields(GVars.Line)) > 2 And (CsvRec.Fields(GVars.Line) <> "" And rstData1.EOF)) Then
           Errors.Text = Errors.Text & "Line# " & csvline_no & " Line of " & CsvRec.Fields(GVars.Line) & " is invalid." & vbCrLf
           WeHaveErrors = True
           LineHasErrors = True
        End If
       
'
'-------------------------------------------------------------------------
'Validate Sub_Line
       
       If (IsNull(CsvRec.Fields(Division)) Or IsNull(CsvRec.Fields(Category)) Or IsNull(CsvRec.Fields(GVars.Line))) And CsvRec.Fields(Sub_Line) <> "" Then
           Errors.Text = Errors.Text & "Line# " & csvline_no & " Sub_Line of " & CsvRec.Fields(Sub_Line) & " and a blank Division, Category or Line is invalid." & vbCrLf
           WeHaveErrors = True
           LineHasErrors = True
      End If
       
       
       Set rstData1 = oDatabase.GetRecordsetFromStoredProc("SELECT top 1 attr_cod_2 FROM IM_ITEM where attr_cod_2 = '" & CsvRec.Fields(Sub_Line) & "'")
       
       If Len(CsvRec.Fields(Sub_Line)) > 2 Or (CsvRec.Fields(Sub_Line) <> "" And rstData1.EOF) Then
           Errors.Text = Errors.Text & "Line# " & csvline_no & " Sub_line " & CsvRec.Fields(Sub_Line) & " is invalid." & vbCrLf
           WeHaveErrors = True
           LineHasErrors = True
       End If
'
'-------------------------------------------------------------------------
'Validate product
       
       If (CsvRec.Fields(ProductCode) <> "" And (CsvRec.Fields(Division) <> "" Or CsvRec.Fields(Category) <> "" Or CsvRec.Fields(GVars.Line) <> "" Or CsvRec.Fields(Sub_Line) <> "")) Then
           Errors.Text = Errors.Text & "Line# " & csvline_no & " ProdcutCode " & CsvRec.Fields(ProductCode) & " Cannot be combined with a Division, Category or Line." & vbCrLf
           WeHaveErrors = True
            LineHasErrors = True
       End If
       
       Set rstData1 = oDatabase.GetRecordsetFromStoredProc("SELECT top 1 attr_cod_3 FROM IM_ITEM where attr_cod_3 = '" & CsvRec.Fields(ProductCode) & "'")
       
       If (CsvRec.Fields(ProductCode) <> "" And rstData1.EOF) Then
           Errors.Text = Errors.Text & "Line# " & csvline_no & " ProductCode " & CsvRec.Fields(ProductCode) & " is invalid." & vbCrLf
           WeHaveErrors = True
            LineHasErrors = True
       End If
'
'-------------------------------------------------------------------------
'Validate StyleNum

       If (CsvRec.Fields(StyleNum) <> "" And (CsvRec.Fields(ProductCode) <> "" Or CsvRec.Fields(Division) <> "" Or CsvRec.Fields(Category) <> "" Or CsvRec.Fields(GVars.Line) <> "" Or CsvRec.Fields(Sub_Line) <> "")) Then
           Errors.Text = Errors.Text & "Line# " & csvline_no & " StyleNumber " & CsvRec.Fields(StyleNum) & " Cannot be combined with a Division, Category, Line or Sub_Line." & vbCrLf
           WeHaveErrors = True
            LineHasErrors = True
       End If
       
       Set rstData1 = oDatabase.GetRecordsetFromStoredProc("SELECT top 1 prof_alpha_1 FROM IM_ITEM where prof_alpha_1 = '" & CsvRec.Fields(StyleNum) & "'")
       
       If (CsvRec.Fields(StyleNum) <> "" And rstData1.EOF) Then
           Errors.Text = Errors.Text & "Line# " & csvline_no & " StyleNumber " & CsvRec.Fields(StyleNum) & " is invalid." & vbCrLf
           WeHaveErrors = True
            LineHasErrors = True
       End If
       
'
'-------------------------------------------------------------------------
'Validate ColorCode
       
       If GVars.promo_format = "General" Then
       
           Set rstData1 = oDatabase.GetRecordsetFromStoredProc("SELECT top 1 prof_alpha_2 FROM IM_ITEM where prof_alpha_2 = '" & CsvRec.Fields(ColorCode) & "'")
        
        If (CsvRec.Fields(ColorCode) <> "" And rstData1.EOF) Then
           Errors.Text = Errors.Text & "Line# " & csvline_no & " ColorCode " & CsvRec.Fields(ColorCode) & " is invalid." & vbCrLf
           WeHaveErrors = True
            LineHasErrors = True
        End If
       End If
'
'-------------------------------------------------------------------------
'Validate DeptCode
        
       If IsNull(CsvRec.Fields(DeptCode)) Or (CsvRec.Fields(DeptCode) <> "1" And CsvRec.Fields(DeptCode) <> "2" And CsvRec.Fields(DeptCode) <> "3") Then
           Errors.Text = Errors.Text & "Line# " & csvline_no & " DeptCode " & CsvRec.Fields(DeptCode) & " is invalid." & vbCrLf
           WeHaveErrors = True
            LineHasErrors = True
       End If
'
'-------------------------------------------------------------------------
'Validate BOGO Specific

     'BuyQuantity /DiscountQuantity Check only for Bogo
     
     If GVars.promo_format = "Bogo" Then
     
        If (IsNumeric(CsvRec.Fields(BuyQty)) = False) Then
           Errors.Text = Errors.Text & "Line# " & csvline_no & " BuyQuantity of " & CsvRec.Fields(BuyQty) & " is invalid. BuyQuantity must be a whole number and must not be blank." & vbCrLf
           WeHaveErrors = True
            LineHasErrors = True
        End If
        
        If CsvRec.Fields(PriceMethod) <> "F" And CsvRec.Fields(DiscQty) = "" Then
                 Errors.Text = Errors.Text & "Line# " & csvline_no & " DiscountQuantity of " & CsvRec.Fields(DiscQty) & " is invalid. DiscountQuantity must be a whole number." & vbCrLf
                WeHaveErrors = True
                LineHasErrors = True
        ElseIf CsvRec.Fields(DiscQty) <> "" And (IsNumeric(CsvRec.Fields(DiscQty)) = False) Then
           Errors.Text = Errors.Text & "Line# " & csvline_no & " DiscountQuantity of " & DiscQty & " is invalid. DiscountQuantity must be a whole number." & vbCrLf
           WeHaveErrors = True
            LineHasErrors = True
        End If
      End If
      
      
'
'-------------------------------------------------------------------------
'Validate Coupon Specific
      
    ' check only for Coupon CouponID    CouponDescr BarCode MinQty
    
      If GVars.promo_format = "Coupon" Then
        If CsvRec.Fields(CouponID) = "" Or (Len(CsvRec.Fields(CouponID)) = 10) Then
           Errors.Text = Errors.Text & "Line# " & csvline_no & " CouponID of " & CsvRec.Fields(CouponID) & " is invalid. CouponID must be unique 10 Characters and must not be blank." & vbCrLf
           WeHaveErrors = True
            LineHasErrors = True
        End If
        If CsvRec.Fields(CouponDescr) <> "" And (Len(CsvRec.Fields(CouponDescr)) > 30) Then
           Errors.Text = Errors.Text & "Line# " & csvline_no & " CouponDescr of " & CsvRec.Fields(CouponDescr) & " is invalid. CouponDescr cannot be more than 30 characters." & vbCrLf
           WeHaveErrors = True
            LineHasErrors = True
        End If
        
        If IsNull(CsvRec.Fields(Barcode)) Or (CsvRec.Fields(Barcode) <> "" And ((Len(CsvRec.Fields(Barcode)) < 10 Or Len(CsvRec.Fields(Barcode)) > 20) Or InStr(LCase(Barcode_Str), LCase(CsvRec.Fields(Barcode))) <> 0)) Then
           Errors.Text = Errors.Text & "Line# " & csvline_no & " Barcode of " & CsvRec.Fields(Barcode) & " is invalid. Barcode cannot be less than 10 and more than 20 characters and cannot be blank." & vbCrLf
           WeHaveErrors = True
            LineHasErrors = True
        End If
        
        If InStr(LCase(Barcode_Str), LCase(CsvRec.Fields(Barcode))) = 0 Then
          Barcode_Str = Barcode_Str & "|" & CsvRec.Fields(Barcode)
         End If
         
         If IsNull(CsvRec.Fields(MinQty)) Or (IsNumeric(CsvRec.Fields(MinQty)) = False) Then
           Errors.Text = Errors.Text & "Line# " & csvline_no & " MinQty of " & CsvRec.Fields(MinQty) & " is invalid. MinQty must minimum 1 and must not be blank." & vbCrLf
           WeHaveErrors = True
            LineHasErrors = True
        End If
       End If
'
'-------------------------------------------------------------------------
'Validate Price Method
       
       If CsvRec.Fields(PriceMethod) = "A" And GVars.promo_format <> "General" Then
         Errors.Text = Errors.Text & "Line# " & csvline_no & " CsvRec.Fields(PriceMethod) must be  F or D for Bogo or Coupon Promos." & vbCrLf
           WeHaveErrors = True
            LineHasErrors = True
       End If
       If CsvRec.Fields(PriceMethod) <> "F" And CsvRec.Fields(PriceMethod) <> "D" And CsvRec.Fields(PriceMethod) <> "A" Then
         Errors.Text = Errors.Text & "Line# " & csvline_no & " CsvRec.Fields(PriceMethod) must be A, F or D." & vbCrLf
           WeHaveErrors = True
            LineHasErrors = True
       End If
'
'-------------------------------------------------------------------------
'Validate AmtorPct

       If (IsNumeric(CsvRec.Fields(AmtorPct)) = False) And CsvRec.Fields(PriceMethod) = "F" Then
           Errors.Text = Errors.Text & "Line# " & csvline_no & " " & CsvRec.Fields(AmtorPct) & " Given PriceMethod of F , the AmountOrPercent must be numerical." & vbCrLf
           WeHaveErrors = True
            LineHasErrors = True
       ElseIf CsvRec.Fields(PriceMethod) = "D" Then
         If (IsNumeric(CsvRec.Fields(AmtorPct)) = False) Then
              Errors.Text = Errors.Text & "Line# " & csvline_no & " " & CsvRec.Fields(AmtorPct) & " The AmountOrPercent must be numerical." & vbCrLf
            WeHaveErrors = True
            LineHasErrors = True
         Else
       
           If (CInt(CsvRec.Fields(AmtorPct)) < 0 Or CInt(CsvRec.Fields(AmtorPct)) > 100) Then
                 Errors.Text = Errors.Text & "Line# " & csvline_no & " " & CsvRec.Fields(AmtorPct) & " Given PriceMethod of D (Percentage), the AmountOrPercent must be a percentage greater than 0% but no greater than 100%." & vbCrLf
               WeHaveErrors = True
               LineHasErrors = True
           End If
        End If
       ElseIf CsvRec.Fields(PriceMethod) = "A" Then
         If (IsNumeric(CsvRec.Fields(AmtorPct)) = False) And CInt(AmtorPct) > 0 Then
             Errors.Text = Errors.Text & "Line# " & csvline_no & " " & CsvRec.Fields(AmtorPct) & " Given PriceMethod of A ,  AmountOrPercent is numerical and greater than 0." & vbCrLf
             WeHaveErrors = True
             LineHasErrors = True
         End If
         
       End If
       
'
'-------------------------------------------------------------------------
'Validate Start Date
       
       
       If CsvRec.Fields(StartDate) < Date Or IsDate(CsvRec.Fields(StartDate)) = False Then
           Errors.Text = Errors.Text & "Line# " & csvline_no & " StartDate of " & CsvRec.Fields(StartDate) & " is invalid. StartDate must be a date and must be today or later." & vbCrLf
           WeHaveErrors = True
           LineHasErrors = True
       End If
       
'
'-------------------------------------------------------------------------
'Validate End Date
       
       If (CsvRec.Fields(EndDate) < Date) Or (IsDate(CsvRec.Fields(EndDate)) = False) Or (CsvRec.Fields(EndDate) < CsvRec.Fields(StartDate)) Then
           Errors.Text = Errors.Text & "Line# " & csvline_no & " EndDate of " & CsvRec.Fields(EndDate) & " is invalid. EndDate must be a date and must be today or later.  AND that date is equal to or greater than StartDate." & vbCrLf
           WeHaveErrors = True
           LineHasErrors = True
       End If
       
'
'-------------------------------------------------------------------------
'Validate Store
       
        If CsvRec.Fields(Store) <> "" Then
            Dim i As Integer
            Dim err_exists As Boolean
            err_exists = False
            'check if store is seperated by comma
            If InStr(CsvRec.Fields(Store), ",") > 0 Then
                Dim stores() As String
            ' Split the string at the comma characters
               stores() = Split(CsvRec.Fields(Store), ",")
        
                For i = 0 To UBound(stores)
                  If (InStr(LCase(Store_db), LCase(stores(i))) = 0) Then
                    err_exists = True
                  End If
                 Next i
            Else
                If (InStr(LCase(Store_db), LCase(CsvRec.Fields(Store))) = 0) Then
                 Errors.Text = Errors.Text & "Line# " & csvline_no & " Store " & CsvRec.Fields(Store) & " is invalid. Store must be either a single valid store number or a comma separated list of stores, blank (for all stores)." & vbCrLf
                 WeHaveErrors = True
                 LineHasErrors = True
                End If
            End If
            If err_exists = True Then
                Errors.Text = Errors.Text & "Line# " & csvline_no & " Store " & CsvRec.Fields(Store) & " is invalid. Store must be either a single valid store number or a comma separated list of stores, with no spaces between or blank (for all stores)." & vbCrLf
                WeHaveErrors = True
                LineHasErrors = True
            End If
        End If
       
'
'-------------------------------------------------------------------------
'Validate Promo ID
      
       
        If CsvRec.Fields(promo_id) = "" Then
           Errors.Text = Errors.Text & "Line# " & csvline_no & " Promo ID of " & promo_id & " is invalid. Promo ID cannot be blank." & vbCrLf
           WeHaveErrors = True
            LineHasErrors = True
        ElseIf promo_format = "Coupon" Or promo_format = "Bogo" Then
         If ((CsvRec.Fields(promo_id) <> "" And Len(CsvRec.Fields(promo_id)) > 15) Or InStr(LCase(promo_id_str), LCase(CsvRec.Fields(promo_id))) <> 0) Then
           Errors.Text = Errors.Text & "Line# " & csvline_no & " Promo ID of " & promo_id & " is invalid. Promo ID must be unique, must be 15 characters or less and cannot be blank." & vbCrLf
           WeHaveErrors = True
            LineHasErrors = True
         End If
        ElseIf promo_format = "General" Then
         If (CsvRec.Fields(promo_id) <> "" And Len(CsvRec.Fields(promo_id)) > 15) Then
           Errors.Text = Errors.Text & "Line# " & csvline_no & " Promo ID of " & promo_id & " is invalid. Promo ID must be unique,  must be 15 characters or less and cannot be blank." & vbCrLf
           WeHaveErrors = True
           LineHasErrors = True
         End If
       End If
       
       ' create list of promo id's used to verify that ir hasn't been added already for BOGO and Coupon
        If InStr(LCase(promo_id_str), LCase(CsvRec.Fields(promo_id))) = 0 Then
          promo_id_str = promo_id_str & "|" & CsvRec.Fields(promo_id)
        End If
        

        'If this data row has errors print a blank line.  (This groups all the errors for one data row)
        If (LineHasErrors) Then
            Errors.Text = Errors.Text & vbCrLf
            LineHasErrors = False
        End If
        
        'Move to next record
        CsvRec.MoveNext
        'increment line counter
        csvline_no = csvline_no + 1
    Loop
    
    'Close the Promo Data Object
    CsvRec.Close
    Set CsvRec = Nothing
    
    'Re-Set the progress bar
    ValProg.Value = 0
    
    'Check for Errors
    If WeHaveErrors Then
        Write_log 2, Errors.Text
        MsgBox "Please check errors, correct and re-validate", 0, "Validation Errors"
        GoTo Finish:
    Else
        ' No Errors mark as successful
        Errors.Text = Errors.Text & "Successfully Validated Promo " & vbCrLf
        Errors.Text = Errors.Text & vbCrLf
        Write_log 2, Errors.Text
        Success = True
        Btn_Move.Enabled = True
        GoTo Finish:
    End If
    

Validate_Error:
   
      MsgBox "Error # " & Str(Err.Number) & " was generated by " & Err.Source & Chr(13) & Err.Description
      Resume Finish
Finish:
   
End Sub

Private Function Check_MovedtoCP()

    'This Functions checks the log file to see if it has already been moved to Counterpoint
    
    Dim MovedToCP As Boolean
        
    MovedToCP = False
    
   
   
    If Len(Dir(GVars.Log_FileName, 0)) > 0 Then
        
        Dim FileNo As Integer
        FileNo = FreeFile()
        Dim LogEntry As String
                
        Open GVars.Log_FileName For Input As FileNo Len = -1
        
        While Not EOF(FileNo)
            Line Input #1, LogEntry
        Wend
        Close FileNo
        
        MovedToCP = InStr(1, LCase(LogEntry), LCase("Moved to Counterpoint"))
        
    End If
    
        Check_MovedtoCP = MovedToCP
End Function



Private Function Check_Format()

' This Function checks the format of the Promo file and compares it to the indicated format

    On Error GoTo Format_Error:
    
    Dim RightFormat As Boolean
    
    RightFormat = False
    
    Const ForReading = 1, ForWriting = 2, ForAppending = 8
    Const TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0
    
    Dim fso As New FileSystemObject
    Dim f As File
    Dim fsoStream As TextStream
    Dim strLine As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set f = fso.GetFile(GVars.Promo_filename_full)
    Set fsoStream = f.OpenAsTextStream(ForReading, TristateUseDefault)
    strLine = fsoStream.ReadLine
    fsoStream.Close
    Set fsoStream = Nothing
    Set f = Nothing
    Set fso = Nothing
    
    'Check if header has ColorCode indicating it is a general file
    
    If InStr(strLine, "ColorCode") Then
            
        If (General_Option.Value) Then      'It's a General File make sure they have indicated as such'
            RightFormat = True
            GVars.promo_format = General_Option.Caption
            GoTo Done:
        Else
            RightFormat = False
            GoTo Done:
        End If
    End If
    
    'Check if header has Buy_Qty indicating it is a Bogo file
    
    If InStr(strLine, "Buy_Qty") Then
            
        If (Bogo_Option.Value) Then         'It's a Bogo File make sure they have indicated as such'
            RightFormat = True
            GVars.promo_format = Bogo_Option.Caption
            GoTo Done:
        Else
            RightFormat = False
            GoTo Done:
        End If
    End If
    
    'Check if header has CoupinID indicating it is a Coupon file
    
    If InStr(strLine, "CouponID") Then
            
        If (Coupon_Option.Value) Then        'It's a Coupon File make sure they have indicated as such'
            RightFormat = True
            GVars.promo_format = Coupon_Option.Caption
            GoTo Done:
        Else
            RightFormat = False
            GoTo Done:
        End If
    End If
    
Format_Error:
    RightFormat = False
Done:
    Check_Format = RightFormat
End Function

Public Sub Create_SchemaFile(TXTFile, DirectoryPath)
' This subroutine creates the Schema file for the Promo file.  It is uesd to ensure that the data is formatted correctly

    Dim File As Integer
    
    File = FreeFile
    
    'Delete the existing Scheme file
    If (Len(Dir(DirectoryPath & "\schema.ini", 0)) > 0) Then
        Kill DirectoryPath & "\schema.ini"
    End If

    ' Create General Schema
    If General_Option.Value = True Then
        Open DirectoryPath & "\Schema.ini" For Output As #File Len = -1
        Print #File, "[" & TXTFile & "]"
        Print #File, "Format=CSVDelimited"
        Print #File, "ColNameHeader = true"
        Print #File, "MaxScanRows = 0"
        Print #File, "CharacterSet = ANSI"
        Print #File, "Col1=Division Text"
        Print #File, "Col2=Category Text"
        Print #File, "Col3=Line Text"
        Print #File, "Col4=Sub-line Text"
        Print #File, "Col5=ProductCode Text"
        Print #File, "Col6=StyleNum Text"
        Print #File, "Col7=ColorCode Text"
        Print #File, "Col8=DeptCode Text"
        Print #File, "Col9=PriceMethod Text"
        Print #File, "Col10=AmtorPct Text"
        Print #File, "Col11=StartDate date"
        Print #File, "Col12=EndDate date"
        Print #File, "Col13=Store Text"
        Print #File, "Col14=Promo_ID Text"
        
        Close #File
    End If
      
    'Create Coupon Schema
    If Coupon_Option.Value = True Then
        Open DirectoryPath & "\Schema.ini" For Output As #File Len = -1
        Print #File, "[" & TXTFile & "]"
        Print #File, "Format=CSVDelimited"
        Print #File, "ColNameHeader = true"
        Print #File, "MaxScanRows = 0"
        Print #File, "CharacterSet = ANSI"
        Print #File, "Col1=Division Text"
        Print #File, "Col2=Category Text"
        Print #File, "Col3=Line Text"
        Print #File, "Col4=Sub-line Text"
        Print #File, "Col5=ProductCode Text"
        Print #File, "Col6=StyleNum Text"
        Print #File, "Col7=DeptCode Text"
        Print #File, "Col8=CouponID Text"
        Print #File, "Col9=CouponDescr Text"
        Print #File, "Col10=Barcode Text"
        Print #File, "Col11=MinQty Text"
        Print #File, "Col12=PriceMethod Text"
        Print #File, "Col13=AmtorPct Text"
        Print #File, "Col14=StartDate date"
        Print #File, "Col15=EndDate date"
        Print #File, "Col16=Store Text"
        Print #File, "Col17=Promo_ID Text"
        Close #File
     
    End If
    
    'Create Bogo Schema
    If Bogo_Option.Value = True Then
        Open DirectoryPath & "\Schema.ini" For Output As #File Len = -1
        Print #File, "[" & TXTFile & "]"
        Print #File, "Format=CSVDelimited"
        Print #File, "ColNameHeader = true"
        Print #File, "MaxScanRows = 0"
        Print #File, "CharacterSet = ANSI"
        Print #File, "Col1=Division Text"
        Print #File, "Col2=Category Text"
        Print #File, "Col3=Line Text"
        Print #File, "Col4=Sub-line Text"
        Print #File, "Col5=ProductCode Text"
        Print #File, "Col6=StyleNum Text"
        Print #File, "Col7=DeptCode Text"
        Print #File, "Col8=Buy_Qty Text"
        Print #File, "Col9=Disc_Qty Text"
        Print #File, "Col10=PriceMethod Text"
        Print #File, "Col11=AmtorPct Text"
        Print #File, "Col12=StartDate date"
        Print #File, "Col13=EndDate date"
        Print #File, "Col14=Store Text"
        Print #File, "Col15=Promo_ID Text"
        Close #File
    End If
End Sub

Private Function Write_log(ByVal Mode As Integer, data As String)
    'This function writes the log file
    
    Const TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0
    
    Dim fso As New FileSystemObject
    Dim f As File
    Dim fsoStream As TextStream
    Dim strLine As String
    

    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If Len(Dir(GVars.Log_FileName, 0)) >= 0 And Mode = 2 Then
        If Len(Dir(GVars.Log_FileName, 0)) > 0 Then fso.DeleteFile (GVars.Log_FileName)
        fso.CreateTextFile (GVars.Log_FileName)
    End If
    
    Set f = fso.GetFile(GVars.Log_FileName)
    Set fsoStream = f.OpenAsTextStream(Mode, TristateUseDefault)
        
        fsoStream.Write (data)
    
    fsoStream.Close
    Set fsoStream = Nothing
    Set f = Nothing
    Set fso = Nothing

End Function

Private Sub Btn_Move_Click()
    'This subroutine writes the Validated data found in the Promo File to the Counterpoint Database

    On Error GoTo Move_Error:
    
    Dim Sql_Vals As String
    Dim User_ID As String
    Dim Last_PromoID As String
    Dim WriteLog As Boolean
    
    'Get the user name on the person doing the validation and move
    User_ID = Environ("USERNAME")
    
    Dim Final_Msg As String
    
    Final_Msg = vBCrLg & "Moved to Counterpoint"
    
   
    
    
     ' prepare to open the CSV File
    
    Dim CsvConStr As String
    Dim CsvCon As New ADODB.Connection
    Dim CsvRec As New ADODB.Recordset
    CsvConStr = "Driver={Microsoft Text Driver (*.txt; *.csv)};Dbq=" & GVars.Promo_filepath & ";Extensions=txt,csv;HDR=YES"
    
    'Open the connection to the file path
    
    CsvCon.Open CsvConStr
    
    'Open the file and read the records
    
    CsvRec.Open "SELECT * FROM [" & GVars.Promo_filename & "]", CsvCon, adOpenStatic, adLockReadOnly, adCmdText
    
    'Make sure you are at the firs record
    
    CsvRec.MoveFirst
    
    Last_PromoID = ""
    Dim CsvLine As Integer
    
    'Set the Line counter
    csvline_no = 1
    
    Do While Not CsvRec.EOF
    
        'Check to see if we are replacing the promo and delete it from the staging table before Writing the replacement.
        If Replace_Promo Then
            If promo_format = "General" Then
                oDatabase.DeletePromoFromTable promo_format, CsvRec.Fields(promo_id)
                Replace_Promo = False
            End If
            If promo_format <> "General" Then
                'Bogo and Coupon files can have multiple promo ID's so delete all of then as you hit them
                oDatabase.DeletePromoFromTable promo_format, CsvRec.Fields(promo_id)
                
            End If
        End If
        
        ' set flag as to write a log file entry
        If CsvRec.Fields(promo_id) = Last_PromoID Then
            WriteLog = False
        Else
            WriteLog = True
            Last_PromoID = CsvRec.Fields(promo_id)
        End If
        
        
        
        If promo_format = "General" Then                            ' Setup the General Promo Values to be written to the Database
            Sql_Vals = "'" & CsvRec.Fields(Division) & "'" & "," & "'" & CsvRec.Fields(Category) & "'" & "," & "'" & CsvRec.Fields(GVars.Line) & "'" & "," & _
            "'" & CsvRec.Fields(Sub_Line) & "'" & "," & "'" & CsvRec.Fields(ProductCode) & "'" & "," & "'" & CsvRec.Fields(StyleNum) & "'" & "," & "'" & CsvRec.Fields(ColorCode) & "'" & "," & _
            "'" & CsvRec.Fields(DeptCode) & "'" & "," & "'" & CsvRec.Fields(PriceMethod) & "'" & "," & "'" & CsvRec.Fields(AmtorPct) & "'" & "," & "'" & CsvRec.Fields(StartDate) & "'" & "," & _
            "'" & CsvRec.Fields(EndDate) & "'" & "," & "'" & CsvRec.Fields(Store) & "'" & "," & "'" & CsvRec.Fields(promo_id) & "'" & ")"
            
        ElseIf promo_format = "Bogo" Then                           ' Setup the BOGO Promo Values to be written to the Database
            Sql_Vals = "'" & CsvRec.Fields(Division) & "'" & "," & "'" & CsvRec.Fields(Category) & "'" & "," & "'" & CsvRec.Fields(GVars.Line) & "'" & "," & "'" & CsvRec.Fields(Sub_Line) & "'" & "," & _
            "'" & CsvRec.Fields(ProductCode) & "'" & "," & "'" & CsvRec.Fields(StyleNum) & "'" & "," & "'" & CsvRec.Fields(DeptCode) & "'" & "," & "'" & CsvRec.Fields(BuyQty) & "'" & "," & _
            "'" & CsvRec.Fields(DiscQty) & "'" & "," & "'" & CsvRec.Fields(PriceMethod) & "'" & "," & "'" & CsvRec.Fields(AmtorPct) & "'" & "," & "'" & CsvRec.Fields(StartDate) & "'" & "," & _
            "'" & CsvRec.Fields(EndDate) & "'" & "," & "'" & CsvRec.Fields(Store) & "'" & "," & "'" & CsvRec.Fields(promo_id) & "'" & ")"
        
        ElseIf promo_format = "Coupon" Then                          ' Setup the Coupon Promo Values to be written to the Database
            Sql_Vals = "'" & CsvRec.Fields(Division) & "'" & "," & "'" & CsvRec.Fields(Category) & "'" & "," & "'" & CsvRec.Fields(GVars.Line) & "'" & "," & "'" & CsvRec.Fields(Sub_Line) & "'" & "," & _
            "'" & CsvRec.Fields(ProductCode) & "'" & "," & "'" & CsvRec.Fields(StyleNum) & "'" & "," & "'" & CsvRec.Fields(DeptCode) & "'" & "," & "'" & CsvRec.Fields(CouponID) & "'" & "," & _
            "'" & CsvRec.Fields(CouponDescr) & "'" & "," & "'" & CsvRec.Fields(Barcode) & "'" & "," & "'" & CsvRec.Fields(MinQty) & "'" & "," & "'" & CsvRec.Fields(PriceMethod) & "'" & "," & _
            "'" & CsvRec.Fields(AmtorPct) & "'" & "," & "'" & CsvRec.Fields(StartDate) & "'" & "," & "'" & CsvRec.Fields(EndDate) & "'" & "," & "'" & CsvRec.Fields(Store) & "'" & "," & "'" & CsvRec.Fields(promo_id) & "'" & ")"
        
        End If
        
        'Write the data to the database
        oDatabase.InsertDataintoTable promo_format, Sql_Vals, CsvRec.Fields(promo_id), CsvRec.Fields(StartDate), CsvRec.Fields(EndDate), User_ID, WriteLog
    
        CsvRec.MoveNext
        csvline_no = csvline_no + 1
    Loop
    
    CsvRec.Close
    Set CsvRec = Nothing
    
    ' Update the log file to indicate moved to Counterpoint
    
    Write_log 8, Final_Msg
    
    Errors.Text = Errors.Text & "Moved to Counterpoint" & vbCrLf
    
    'Disable the move button
    
    Btn_Move.Enabled = False
    
    GoTo Finish:
    
Move_Error:
    
    MsgBox "Error # " & Str(Err.Number) & " was generated by " & Err.Source & Chr(13) & Err.Description
    Resume Finish
Finish:
    
  End Sub

Private Sub form_load()
    'This subroutine is ran at startup and makes the connection to the Database
    
    versiondetails.Caption = "Ver 2.0 06/02/2020"
   
   ' Open Connection to Lands End DB
 
    Set oDatabase = oDBInstance.GetNewDatabase()

End Sub

Private Sub Quit_Click() ' Handles Quit.Click

    ' This subroutine exits the program and is ran when the exit button is clicked
   oDBInstance.CloseSharedDatabase
   Set oDBInstance = Nothing
   Set oDatabase = Nothing
   End
End Sub

    
Private Sub Form_unload(Cancel As Integer)
    'This subroutine is ran when the user clicks the X on the form
    Dim answer As Integer
            answer = MsgBox("Are you sure you want to Exit the Application", vbOKCancel + vbQuestion, "Exit")
            If answer = vbCancel Then
                Cancel = True
                Exit Sub
            End If
            
           oDBInstance.CloseSharedDatabase
           Set oDBInstance = Nothing
           Set oDatabase = Nothing
           End
    End Sub

