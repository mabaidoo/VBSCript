Attribute VB_Name = "GVars"
Public Promo_filename_full As String
Public Promo_filename As String
Public Promo_filename_noext As String
Public Promo_filepath As String
Public promo_format As String
Public Const Log_file_path = "C:\Program Files (x86)\LandsEnd\Validation\Logs\"
Public Log_FileName As String
Public Replace_Promo As Boolean
Public Log_Exists As Boolean
Public WeHaveErrors As Boolean
Public Store_db As String
Public Success As Boolean
' vars used as index
Public Division As Integer
Public Category As Integer
Public Line As Integer
Public Sub_Line As Integer
Public ProductCode As Integer
Public StyleNum As Integer
Public ColorCode As Integer
Public DeptCode As Integer
Public PriceMethod As Integer
Public AmtorPct As Integer
Public StartDate As Integer
Public EndDate As Integer
Public Store As Integer
Public promo_id As Integer
Public CouponID As Integer
Public CouponDescr As Integer
Public Barcode As Integer
Public MinQty As Integer
Public BuyQty As Integer
Public DiscQty As Integer
Public oDBInstance As New DBInstance
Public oDatabase   As DataBase




