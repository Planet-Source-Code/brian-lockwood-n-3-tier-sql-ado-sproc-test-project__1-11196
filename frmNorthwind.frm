VERSION 5.00
Begin VB.Form frmTest 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "LockwoodTech Proc-Blaster Test App. (3-tier)"
   ClientHeight    =   4980
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7740
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4980
   ScaleWidth      =   7740
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtCustomerID 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1680
      MaxLength       =   10
      TabIndex        =   25
      Top             =   960
      Width           =   2475
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "&New"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4590
      TabIndex        =   24
      Top             =   990
      Width           =   915
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5550
      TabIndex        =   23
      Top             =   990
      Width           =   915
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6510
      TabIndex        =   22
      Top             =   990
      Width           =   915
   End
   Begin VB.TextBox txtFax 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1680
      MaxLength       =   24
      TabIndex        =   20
      Top             =   4560
      Width           =   2475
   End
   Begin VB.TextBox txtPhone 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1680
      MaxLength       =   24
      TabIndex        =   18
      Top             =   4200
      Width           =   2475
   End
   Begin VB.TextBox txtCountry 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1680
      MaxLength       =   15
      TabIndex        =   16
      Top             =   3840
      Width           =   2475
   End
   Begin VB.TextBox txtPostalCode 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1680
      MaxLength       =   10
      TabIndex        =   14
      Top             =   3480
      Width           =   2475
   End
   Begin VB.TextBox txtRegion 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1680
      MaxLength       =   15
      TabIndex        =   12
      Top             =   3120
      Width           =   2475
   End
   Begin VB.TextBox txtCity 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1680
      MaxLength       =   15
      TabIndex        =   10
      Top             =   2760
      Width           =   2475
   End
   Begin VB.TextBox txtAddress 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1680
      MaxLength       =   60
      TabIndex        =   8
      Top             =   2400
      Width           =   2475
   End
   Begin VB.TextBox txtContactTitle 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1680
      MaxLength       =   30
      TabIndex        =   6
      Top             =   2040
      Width           =   2475
   End
   Begin VB.TextBox txtContactName 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1680
      MaxLength       =   30
      TabIndex        =   4
      Top             =   1680
      Width           =   2475
   End
   Begin VB.TextBox txtCompanyName 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1680
      MaxLength       =   40
      TabIndex        =   2
      Top             =   1320
      Width           =   2475
   End
   Begin VB.ComboBox cboCompanyName 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1680
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   2475
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      Caption         =   "3 tier Data Access"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   4350
      TabIndex        =   28
      Top             =   120
      Width           =   3345
   End
   Begin VB.Label Label11 
      Caption         =   "Note:  This project requires the SQL Stored Procedures in the attached Procs.sql file to be added to the Northwind database 1st"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   795
      Left            =   4350
      TabIndex        =   27
      Top             =   1650
      Width           =   3345
   End
   Begin VB.Label Label10 
      Caption         =   "Customer ID:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   26
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Fax 
      Caption         =   "Postal Code:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   4560
      Width           =   1455
   End
   Begin VB.Label Label9 
      Caption         =   "Phone"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   4200
      Width           =   1455
   End
   Begin VB.Label Label7 
      Caption         =   "Country:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   3840
      Width           =   1455
   End
   Begin VB.Label Label8 
      Caption         =   "Postal Code:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   3480
      Width           =   1455
   End
   Begin VB.Label txtRegionx 
      Caption         =   "Region"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Label Label6 
      Caption         =   "City:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   2760
      Width           =   1455
   End
   Begin VB.Label Label5 
      Caption         =   "Address:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "Contact Title:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Company Name:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Contact Name"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Filter:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   180
      Width           =   1455
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'   Make sure to create 5 procs for the Customer table in the northwind database
'   called prc_ins_Customers, prc_upd_Customers, prc_sel_Customers and prc_del_Customers
'   and one additional proc to allow for two types of Selects prc_sel_Customers_Output
'   See attached file:  Procs.sql

'   To make both type of select procs run the procs 1st (using Proc-Blaster) with the
'   option for Output parameters off, then run another select with it on and rename
'   it prc_sel_Customers_Output

'   You will also need to make an ODBC DSN called Northwind or alter the code in
'   this project for the connection string

'   All Procs and VB Data Access code generated with LockwoodTech Proc-Blaster 2

'                       http://www.lockwoodtech.com


Option Explicit

Dim objCustomer As clsCustomer

Private Sub cboCompanyName_Click()
 Dim strSQL             As String
 Dim rs                 As New ADODB.Recordset
 Dim lngRetVal          As Long
 
 If cboCompanyName.ListIndex = -1 Then Exit Sub
 
    '   Creating a new customer object (in the process destroying any existing one)
    Set objCustomer = New clsCustomer
    
    lngRetVal = objCustomer.Find(RTrim(cboCompanyName))
    
    '   load up the form with the object's data
    With objCustomer
    
        txtCompanyName = .CompanyName
        txtContactName = IIf(IsNull(.ContactName), "", .ContactName)
        txtContactTitle = IIf(IsNull(.ContactTitle), "", .ContactTitle)
        txtAddress = IIf(IsNull(.Address), "", .Address)
        txtCity = IIf(IsNull(.City), "", .City)
        txtRegion = IIf(IsNull(.Region), "", .Region)
        txtPostalCode = IIf(IsNull(.PostalCode), "", .PostalCode)
        txtCountry = IIf(IsNull(.Country), "", .Country)
        txtPhone = IIf(IsNull(.Phone), "", .Phone)
        txtFax = IIf(IsNull(.Fax), "", .Fax)
        txtCustomerID = .CustomerID
        
    End With

End Sub

Private Sub cmdDelete_Click()
 Dim lngRetVal As Long
 
    lngRetVal = objCustomer.Delete
        
    If lngRetVal = 0 Then
        MsgBox "Operation Succeeded", vbInformation, "Results"
    Else
        MsgBox "Operation Failed", vbCritical, "Results"
        Exit Sub
    End If
    
    Call Clear_Controls
    Call requery_list
End Sub

Private Sub cmdNew_Click()
    Set objCustomer = Nothing           ' destroy old object
    Set objCustomer = New clsCustomer   ' create new object
    Call Clear_Controls
    cboCompanyName.ListIndex = -1
End Sub

Private Sub cmdSave_Click()
Dim lngRetVal As Long

    '   load up the object's properties with the user supplied data
    With objCustomer
    
        .CompanyName = txtCompanyName
        .ContactName = txtContactName
        .ContactTitle = txtContactTitle
        .Address = txtAddress
        .City = txtCity
        .Region = txtRegion
        .PostalCode = txtPostalCode
        .Country = txtCountry
        .Phone = txtPhone
        .Fax = txtFax
        .CustomerID = txtCustomerID
        
    End With

    lngRetVal = objCustomer.Update

    If lngRetVal = 0 Then
        MsgBox "Operation Succeeded", vbInformation, "Results"
    Else
        MsgBox "Operation Failed", vbCritical, "Results"
    End If
    
    Call requery_list
End Sub

Public Sub Form_Load()
    'MsgBox "Please Note:  This project requires the SQL Stored Procedures in the attached Procs.sql " & Chr(13) & "file to be added to the Northwind database 1st", vbInformation
    Call requery_list
End Sub

Private Function Clear_Controls()
 Dim ctl As Control
 
    For Each ctl In Me.Controls
        If TypeOf ctl Is TextBox Then ctl = ""
    Next ctl
End Function

Private Function requery_list()
 Dim rs As New ADODB.Recordset
 Dim strSQL As String
 
    cboCompanyName.Clear
    
    strSQL = "SELECT CustomerID FROM customers where customerID <> ''"
    rs.Open strSQL, g_objCn
    Do While Not rs.EOF
        cboCompanyName.AddItem rs(0)
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
End Function


Public Function Exec_prc_del_Customers(ByVal strCustomerID) As Long
 Dim strSQL As String
 Dim objCmd As ADODB.Command
 Dim objCn  As ADODB.Connection

    On Error GoTo PROC_ERR
    Set objCmd = New ADODB.Command
    Set objCn = New ADODB.Connection
    strSQL = "prc_del_Customers"

    objCn.Open g_strConnectionString
    With objCmd
        .CommandText = strSQL
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = objCn

        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, 0)
        .Parameters.Append .CreateParameter("CustomerID", adVarChar, adParamInput, 10, IIf(strCustomerID = vbNullString, Null, strCustomerID))
        .Execute Options:=adExecuteNoRecords
    
        Exec_prc_del_Customers = .Parameters("RETURN_VALUE")
    End With
    objCn.Close
    Set objCn = Nothing
    
    '   Added for test project
    Debug.Print Cmd2SQL(objCmd)

    Set objCmd = Nothing
    Exit Function
PROC_ERR:
    Exec_prc_del_Customers = Err.Number
End Function

Public Function Exec_prc_ins_Customers(ByVal strCustomerID, ByVal strCompanyName, _
            ByVal strContactName, ByVal strContactTitle, _
            ByVal strAddress, ByVal strCity, ByVal strRegion, _
            ByVal strPostalCode, ByVal strCountry, ByVal strPhone, _
            ByVal strFax) As Long
            
 Dim strSQL As String
 Dim objCmd As ADODB.Command
 Dim objCn  As ADODB.Connection

    On Error GoTo PROC_ERR
    Set objCmd = New ADODB.Command
    Set objCn = New ADODB.Connection
    strSQL = "prc_ins_Customers"

    objCn.Open g_strConnectionString
    With objCmd
        .CommandText = strSQL
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = objCn

        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, 0)
        .Parameters.Append .CreateParameter("CustomerID", adVarChar, adParamInput, 10, IIf(strCustomerID = vbNullString, Null, strCustomerID))
        .Parameters.Append .CreateParameter("CompanyName", adWChar, adParamInput, 40, IIf(strCompanyName = vbNullString, Null, strCompanyName))
        .Parameters.Append .CreateParameter("ContactName", adWChar, adParamInput, 30, IIf(strContactName = vbNullString, Null, strContactName))
        .Parameters.Append .CreateParameter("ContactTitle", adWChar, adParamInput, 30, IIf(strContactTitle = vbNullString, Null, strContactTitle))
        .Parameters.Append .CreateParameter("Address", adWChar, adParamInput, 60, IIf(strAddress = vbNullString, Null, strAddress))
        .Parameters.Append .CreateParameter("City", adWChar, adParamInput, 15, IIf(strCity = vbNullString, Null, strCity))
        .Parameters.Append .CreateParameter("Region", adWChar, adParamInput, 15, IIf(strRegion = vbNullString, Null, strRegion))
        .Parameters.Append .CreateParameter("PostalCode", adWChar, adParamInput, 10, IIf(strPostalCode = vbNullString, Null, strPostalCode))
        .Parameters.Append .CreateParameter("Country", adWChar, adParamInput, 15, IIf(strCountry = vbNullString, Null, strCountry))
        .Parameters.Append .CreateParameter("Phone", adWChar, adParamInput, 24, IIf(strPhone = vbNullString, Null, strPhone))
        .Parameters.Append .CreateParameter("Fax", adWChar, adParamInput, 24, IIf(strFax = vbNullString, Null, strFax))
        .Execute Options:=adExecuteNoRecords
    
        Exec_prc_ins_Customers = .Parameters("RETURN_VALUE")
    End With
    objCn.Close
    Set objCn = Nothing
    
    '   Added for test project
    Debug.Print Cmd2SQL(objCmd)

    Set objCmd = Nothing
    Exit Function
PROC_ERR:
    Exec_prc_ins_Customers = Err.Number
End Function

Public Function Exec_prc_upd_Customers(ByVal strCustomerID, ByVal strCompanyName, _
        ByVal strContactName, ByVal strContactTitle, _
        ByVal strAddress, ByVal strCity, ByVal strRegion, _
        ByVal strPostalCode, ByVal strCountry, ByVal strPhone, _
        ByVal strFax) As Long
        
 Dim strSQL As String
 Dim objCmd As ADODB.Command
 Dim objCn  As ADODB.Connection

    On Error GoTo PROC_ERR
    Set objCmd = New ADODB.Command
    Set objCn = New ADODB.Connection
    strSQL = "prc_upd_Customers"

    objCn.Open g_strConnectionString
    With objCmd
        .CommandText = strSQL
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = objCn

        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, 0)
        .Parameters.Append .CreateParameter("CustomerID", adVarChar, adParamInput, 10, IIf(strCustomerID = vbNullString, Null, strCustomerID))
        .Parameters.Append .CreateParameter("CompanyName", adWChar, adParamInput, 40, IIf(strCompanyName = vbNullString, Null, strCompanyName))
        .Parameters.Append .CreateParameter("ContactName", adWChar, adParamInput, 30, IIf(strContactName = vbNullString, Null, strContactName))
        .Parameters.Append .CreateParameter("ContactTitle", adWChar, adParamInput, 30, IIf(strContactTitle = vbNullString, Null, strContactTitle))
        .Parameters.Append .CreateParameter("Address", adWChar, adParamInput, 60, IIf(strAddress = vbNullString, Null, strAddress))
        .Parameters.Append .CreateParameter("City", adWChar, adParamInput, 15, IIf(strCity = vbNullString, Null, strCity))
        .Parameters.Append .CreateParameter("Region", adWChar, adParamInput, 15, IIf(strRegion = vbNullString, Null, strRegion))
        .Parameters.Append .CreateParameter("PostalCode", adWChar, adParamInput, 10, IIf(strPostalCode = vbNullString, Null, strPostalCode))
        .Parameters.Append .CreateParameter("Country", adWChar, adParamInput, 15, IIf(strCountry = vbNullString, Null, strCountry))
        .Parameters.Append .CreateParameter("Phone", adWChar, adParamInput, 24, IIf(strPhone = vbNullString, Null, strPhone))
        .Parameters.Append .CreateParameter("Fax", adWChar, adParamInput, 24, IIf(strFax = vbNullString, Null, strFax))
        .Execute Options:=adExecuteNoRecords
    
        Exec_prc_upd_Customers = .Parameters("RETURN_VALUE")
    End With
    objCn.Close
    Set objCn = Nothing
    
    '   Added for test project
    Debug.Print Cmd2SQL(objCmd)
    
    Set objCmd = Nothing
    Exit Function
PROC_ERR:
    Exec_prc_upd_Customers = Err.Number
End Function

Public Function Exec_prc_sel_Customers_Output(ByVal strCustomerID As String, ByRef strCompanyName As String, ByRef strContactName As String, ByRef strContactTitle As String, ByRef strAddress As String, ByRef strCity As String, ByRef strRegion As String, ByRef strPostalCode As String, ByRef strCountry As String, ByRef strPhone As String, ByRef strFax As String) As Long
 Dim strSQL As String
 Dim objCmd As ADODB.Command
 Dim objCn  As ADODB.Connection

    On Error GoTo PROC_ERR
    Set objCmd = New ADODB.Command
    Set objCn = New ADODB.Connection
    strSQL = "prc_sel_Customers_Output"

    objCn.Open g_strConnectionString
    With objCmd
        .CommandText = strSQL
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = objCn

        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, 0)
        .Parameters.Append .CreateParameter("CustomerID", adVarChar, adParamInputOutput, 10, IIf(strCustomerID = vbNullString, Null, strCustomerID))
        .Parameters.Append .CreateParameter("CompanyName", adWChar, adParamInputOutput, 40, Null)
        .Parameters.Append .CreateParameter("ContactName", adWChar, adParamInputOutput, 30, Null)
        .Parameters.Append .CreateParameter("ContactTitle", adWChar, adParamInputOutput, 30, Null)
        .Parameters.Append .CreateParameter("Address", adWChar, adParamInputOutput, 60, Null)
        .Parameters.Append .CreateParameter("City", adWChar, adParamInputOutput, 15, Null)
        .Parameters.Append .CreateParameter("Region", adWChar, adParamInputOutput, 15, Null)
        .Parameters.Append .CreateParameter("PostalCode", adWChar, adParamInputOutput, 10, Null)
        .Parameters.Append .CreateParameter("Country", adWChar, adParamInputOutput, 15, Null)
        .Parameters.Append .CreateParameter("Phone", adWChar, adParamInputOutput, 24, Null)
        .Parameters.Append .CreateParameter("Fax", adWChar, adParamInputOutput, 24, Null)
    
        .Execute Options:=adExecuteNoRecords
    
        strCustomerID = RTrim(.Parameters("CustomerID"))
        strCompanyName = RTrim(.Parameters("CompanyName"))
        strContactName = RTrim(.Parameters("ContactName"))
        strContactTitle = RTrim(.Parameters("ContactTitle"))
        strAddress = RTrim(.Parameters("Address"))
        strCity = RTrim(.Parameters("City"))
        strRegion = RTrim(IIf(IsNull(.Parameters("Region")), "", .Parameters("Region")))
        strPostalCode = RTrim(.Parameters("PostalCode"))
        strCountry = RTrim(.Parameters("Country"))
        strPhone = RTrim(.Parameters("Phone"))
        strFax = RTrim(IIf(IsNull(.Parameters("Fax")), "", .Parameters("Fax")))
        Exec_prc_sel_Customers_Output = .Parameters("RETURN_VALUE")
    End With
    objCn.Close
    Set objCn = Nothing
    
    '   Added for test project
    Debug.Print Cmd2SQL(objCmd)
    
    Set objCmd = Nothing
    Exit Function
PROC_ERR:
    Exec_prc_sel_Customers_Output = Err.Number
End Function

Public Function Exec_prc_sel_Customers(ByVal strCustomerID As String, ByRef objRs As Recordset) As Long
 Dim strSQL As String
 Dim objCmd As ADODB.Command
 Dim objCn  As ADODB.Connection

    On Error GoTo PROC_ERR
    Set objCmd = New ADODB.Command
    Set objCn = New ADODB.Connection
    strSQL = "prc_sel_Customers"

    objCn.Open g_strConnectionString
    With objCmd
        .CommandText = strSQL
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = objCn

        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, 0)
        .Parameters.Append .CreateParameter("CustomerID", adVarChar, adParamInput, 10, IIf(strCustomerID = vbNullString, Null, strCustomerID))
    End With
    With objRs
        .CursorLocation = adUseClient
        .Open objCmd, , adOpenStatic, adLockReadOnly
        Set .ActiveConnection = Nothing
    End With
    Exec_prc_sel_Customers = objCmd.Parameters("RETURN_VALUE")
    objCn.Close
    Set objCn = Nothing
    
    '   Added for test project
    Debug.Print Cmd2SQL(objCmd)

    Set objCmd = Nothing
    Exit Function
PROC_ERR:
    Exec_prc_sel_Customers = Err.Number
End Function



Private Function NullIt(ctl As Control) As Variant
    If TypeOf ctl Is ListBox Or TypeOf ctl Is ComboBox Then
        If ctl.ListIndex = -1 Then
            NullIt = Null
        Else
            NullIt = ctl.ItemData(ctl.ListIndex)
        End If
    ElseIf TypeOf ctl Is TextBox Then
        If ctl = "" Then
            NullIt = Null
        Else
            NullIt = ctl
        End If
    'Elseif ADD OTHER CONTROLS AS NECESSARY
    End If
End Function




