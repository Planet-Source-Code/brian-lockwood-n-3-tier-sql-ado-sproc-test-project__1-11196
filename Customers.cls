VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCustomer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

' ************************************************************************************

' Object    : Visual Basic 6.0 Class
' Name      : Customers
' Purpose   : Biz interface to Customers Entity
' Author    : Brian Lockwood
' Created   : 9/2/00 3:32:21 PM

' ************************************************************************************;

Option Explicit


'  Private constants;
Private Const ltUPDATE = 0
Private Const ltINSERT = 1
Private Const ltDELETE = 2

'  Private Read/Write Variables;
Private m_strCustomerID As String
Private m_strCompanyName As String
Private m_varContactName As Variant
Private m_varContactTitle As Variant
Private m_varAddress As Variant
Private m_varCity As Variant
Private m_varRegion As Variant
Private m_varPostalCode As Variant
Private m_varCountry As Variant
Private m_varPhone As Variant
Private m_varFax As Variant

'   Primary Key;
Private mudtPrimaryKey  As PrimaryKey

Private Type PrimaryKey
    CustomerID As String
End Type

'   Common Private Property Variables;
Private m_intStatus As String           '   insert, Update or Delete;
Private m_bolIsDirty As Boolean                 '   has object's data changed?;
Private m_strErrDesc As String          '   string value of last error;


Private Const m_strCLASS_ID = "Customers" '   Subjective name of the class;

Public Property Get PK(ByVal vintIndex As Integer) As Variant
    Select Case vintIndex
        Case Is = 0
            PK = mudtPrimaryKey.CustomerID
        Case Else
            Err.Raise vbObjectError + 3001, m_strCLASS_ID, "Case Else Error on PK retrieval"
    End Select
End Property

'  Public Property LETS/GETS;

'  CustomerID;
Public Property Let CustomerID(ByVal vData As String)
    m_bolIsDirty = IIf(HasVarChanged(m_strCustomerID, vData), True, m_bolIsDirty)
    m_strCustomerID = vData
End Property
Public Property Get CustomerID() As String
    CustomerID = m_strCustomerID
End Property

'  CompanyName;
Public Property Let CompanyName(ByVal vData As String)
    m_bolIsDirty = IIf(HasVarChanged(m_strCompanyName, vData), True, m_bolIsDirty)
    m_strCompanyName = vData
End Property
Public Property Get CompanyName() As String
    CompanyName = m_strCompanyName
End Property

'  ContactName;
Public Property Let ContactName(ByVal vData As Variant)
    m_bolIsDirty = IIf(HasVarChanged(m_varContactName, vData), True, m_bolIsDirty)
    m_varContactName = vData
End Property
Public Property Get ContactName() As Variant
    ContactName = m_varContactName
End Property

'  ContactTitle;
Public Property Let ContactTitle(ByVal vData As Variant)
    m_bolIsDirty = IIf(HasVarChanged(m_varContactTitle, vData), True, m_bolIsDirty)
    m_varContactTitle = vData
End Property
Public Property Get ContactTitle() As Variant
    ContactTitle = m_varContactTitle
End Property

'  Address;
Public Property Let Address(ByVal vData As Variant)
    m_bolIsDirty = IIf(HasVarChanged(m_varAddress, vData), True, m_bolIsDirty)
    m_varAddress = vData
End Property
Public Property Get Address() As Variant
    Address = m_varAddress
End Property

'  City;
Public Property Let City(ByVal vData As Variant)
    m_bolIsDirty = IIf(HasVarChanged(m_varCity, vData), True, m_bolIsDirty)
    m_varCity = vData
End Property
Public Property Get City() As Variant
    City = m_varCity
End Property

'  Region;
Public Property Let Region(ByVal vData As Variant)
    m_bolIsDirty = IIf(HasVarChanged(m_varRegion, vData), True, m_bolIsDirty)
    m_varRegion = vData
End Property
Public Property Get Region() As Variant
    Region = m_varRegion
End Property

'  PostalCode;
Public Property Let PostalCode(ByVal vData As Variant)
    m_bolIsDirty = IIf(HasVarChanged(m_varPostalCode, vData), True, m_bolIsDirty)
    m_varPostalCode = vData
End Property
Public Property Get PostalCode() As Variant
    PostalCode = m_varPostalCode
End Property

'  Country;
Public Property Let Country(ByVal vData As Variant)
    m_bolIsDirty = IIf(HasVarChanged(m_varCountry, vData), True, m_bolIsDirty)
    m_varCountry = vData
End Property
Public Property Get Country() As Variant
    Country = m_varCountry
End Property

'  Phone;
Public Property Let Phone(ByVal vData As Variant)
    m_bolIsDirty = IIf(HasVarChanged(m_varPhone, vData), True, m_bolIsDirty)
    m_varPhone = vData
End Property
Public Property Get Phone() As Variant
    Phone = m_varPhone
End Property

'  Fax;
Public Property Let Fax(ByVal vData As Variant)
    m_bolIsDirty = IIf(HasVarChanged(m_varFax, vData), True, m_bolIsDirty)
    m_varFax = vData
End Property
Public Property Get Fax() As Variant
    Fax = m_varFax
End Property

Public Function Find(ByVal vstrCustomerID As String) As Long
 Dim rs             As New ADODB.Recordset
 Dim lngRetVal      As Long

    On Error GoTo PROC_ERR

    lngRetVal = Exec_prc_sel_Customers(vstrCustomerID, rs)

    If lngRetVal <> 0 Then
        GoTo PROC_EXIT
    ElseIf IsEmpty(rs) Then
        Find = vbObjectError + 3002
        m_strErrDesc = "Empty Recordset"
        GoTo PROC_EXIT
    ElseIf rs.BOF And rs.EOF Then   '   no records returned;
        lngRetVal = vbObjectError + 3003
        m_strErrDesc = "Record not found"
        GoTo PROC_EXIT
    Else
                m_strCustomerID = rs("CustomerID")
                m_strCompanyName = rs("CompanyName")
                m_varContactName = rs("ContactName")
                m_varContactTitle = rs("ContactTitle")
                m_varAddress = rs("Address")
                m_varCity = rs("City")
                m_varRegion = rs("Region")
                m_varPostalCode = rs("PostalCode")
                m_varCountry = rs("Country")
                m_varPhone = rs("Phone")
                m_varFax = rs("Fax")
    End If

    rs.Close
    Set rs = Nothing

    '   Load Primary Key value(s);
        mudtPrimaryKey.CustomerID = m_strCustomerID

    m_bolIsDirty = False '   Set this flag to False because a New object is always Clean;

    m_intStatus = ltUPDATE

PROC_EXIT:
    Find = lngRetVal '  Set the return code to the return code from the txn object method and exit the function;
    Exit Function
PROC_ERR:
    m_strErrDesc = "Procedure: clsCustomers.Find Number = " & Err.Number & " Description = " & Err.Description & " Line = " & Erl
    lngRetVal = Err.Number
End Function

Public Function Update() As Long
 Dim lngRetVal      As Long

    On Error GoTo PROC_ERR

    '   Skip if this Object is being updated but it is unchanged since last update;
    If (m_bolIsDirty = False And m_intStatus = ltUPDATE) Then
        GoTo PROC_EXIT
    End If

   If m_intStatus = ltINSERT Then
       lngRetVal = Exec_prc_ins_Customers()
   ElseIf m_intStatus = ltUPDATE Then
       lngRetVal = Exec_prc_upd_Customers()
   Else
       lngRetVal = Exec_prc_del_Customers()
   End If

    If lngRetVal <> 0 Then
        '   An error occurred;
        GoTo PROC_EXIT
    End If

    '   A new record that has been Inserted becomes and existing record
    '   that will be Updtated next time;
    If m_intStatus = ltINSERT Then m_intStatus = ltUPDATE

PROC_EXIT:
    Update = lngRetVal '  Set the return code to the return code from the txn object method and exit the function;
    Exit Function
PROC_ERR:
    m_strErrDesc = "Procedure: clsCustomers.Find Number = " & Err.Number & " Description = " & Err.Description & " Line = " & Erl
    lngRetVal = Err.Number
End Function

Public Function Delete() As Long
 Dim lngRetVal As Long

    m_intStatus = ltDELETE
    lngRetVal = Update

    Delete = lngRetVal
End Function

'   GENERIC PROPERTY VARIABLES;

Public Property Get ClassID() As String
    ClassID = m_strCLASS_ID
End Property

'   IsDirty;
Public Property Get IsDirty() As Boolean
    IsDirty = m_bolIsDirty
End Property

'   Status;
Public Property Get Status() As Integer
    Status = m_intStatus
End Property

'   Error Description;
Public Property Get ErrorDesc() As String
    ErrorDesc = m_strErrDesc
End Property

'   Defaults;
Private Sub Class_Initialize()
    m_intStatus = ltINSERT
End Sub

Private Function HasVarChanged(val1 As Variant, val2 As Variant) As Boolean
    HasVarChanged = True
    '   If either of the values is a NULL check to make sure that they are NOT
    '   both NULLs  In this case the comparison would result in NULL and not =;
    If IsNull(val1) Or IsNull(val2) Then
        If IsNull(val1) And IsNull(val2) Then
            HasVarChanged = False
            Exit Function
        End If
    End If

    If val1 = val2 Then
        HasVarChanged = False
    End If
End Function

Private Function Exec_prc_del_Customers() As Long
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
                .Parameters.Append .CreateParameter("CustomerID", adVarChar, adParamInput, 10, IIf(m_strCustomerID = vbNullString, Null, m_strCustomerID))
                .Execute Options:=adExecuteNoRecords
        
                Exec_prc_del_Customers = .Parameters("RETURN_VALUE")
        End With
        objCn.Close
        Set objCn = Nothing

        Set objCmd = Nothing
        Exit Function
PROC_ERR:
        Exec_prc_del_Customers = Err.Number
        m_strErrDesc = Err.Description
End Function

Private Function Exec_prc_ins_Customers() As Long
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
                .Parameters.Append .CreateParameter("CustomerID", adVarChar, adParamInput, 10, IIf(m_strCustomerID = vbNullString, Null, m_strCustomerID))
                .Parameters.Append .CreateParameter("CompanyName", adWChar, adParamInput, 40, IIf(m_strCompanyName = vbNullString, Null, m_strCompanyName))
                .Parameters.Append .CreateParameter("ContactName", adWChar, adParamInput, 30, m_varContactName)
                .Parameters.Append .CreateParameter("ContactTitle", adWChar, adParamInput, 30, m_varContactTitle)
                .Parameters.Append .CreateParameter("Address", adWChar, adParamInput, 60, m_varAddress)
                .Parameters.Append .CreateParameter("City", adWChar, adParamInput, 15, m_varCity)
                .Parameters.Append .CreateParameter("Region", adWChar, adParamInput, 15, m_varRegion)
                .Parameters.Append .CreateParameter("PostalCode", adWChar, adParamInput, 10, m_varPostalCode)
                .Parameters.Append .CreateParameter("Country", adWChar, adParamInput, 15, m_varCountry)
                .Parameters.Append .CreateParameter("Phone", adWChar, adParamInput, 24, m_varPhone)
                .Parameters.Append .CreateParameter("Fax", adWChar, adParamInput, 24, m_varFax)
                .Execute Options:=adExecuteNoRecords
        
                Exec_prc_ins_Customers = .Parameters("RETURN_VALUE")
        End With
        objCn.Close
        Set objCn = Nothing

        Set objCmd = Nothing
        Exit Function
PROC_ERR:
        Exec_prc_ins_Customers = Err.Number
        m_strErrDesc = Err.Description
End Function

Private Function Exec_prc_upd_Customers() As Long
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
                .Parameters.Append .CreateParameter("CustomerID", adVarChar, adParamInput, 10, IIf(m_strCustomerID = vbNullString, Null, m_strCustomerID))
                .Parameters.Append .CreateParameter("CompanyName", adWChar, adParamInput, 40, IIf(m_strCompanyName = vbNullString, Null, m_strCompanyName))
                .Parameters.Append .CreateParameter("ContactName", adWChar, adParamInput, 30, m_varContactName)
                .Parameters.Append .CreateParameter("ContactTitle", adWChar, adParamInput, 30, m_varContactTitle)
                .Parameters.Append .CreateParameter("Address", adWChar, adParamInput, 60, m_varAddress)
                .Parameters.Append .CreateParameter("City", adWChar, adParamInput, 15, m_varCity)
                .Parameters.Append .CreateParameter("Region", adWChar, adParamInput, 15, m_varRegion)
                .Parameters.Append .CreateParameter("PostalCode", adWChar, adParamInput, 10, m_varPostalCode)
                .Parameters.Append .CreateParameter("Country", adWChar, adParamInput, 15, m_varCountry)
                .Parameters.Append .CreateParameter("Phone", adWChar, adParamInput, 24, m_varPhone)
                .Parameters.Append .CreateParameter("Fax", adWChar, adParamInput, 24, m_varFax)
                .Execute Options:=adExecuteNoRecords
        
                Exec_prc_upd_Customers = .Parameters("RETURN_VALUE")
        End With
        objCn.Close
        Set objCn = Nothing

        Set objCmd = Nothing
        Exit Function
PROC_ERR:
        Exec_prc_upd_Customers = Err.Number
        m_strErrDesc = Err.Description
End Function

Private Function Exec_prc_sel_Customers(ByVal m_strCustomerID As String, ByRef objRs As Recordset) As Long
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
                .Parameters.Append .CreateParameter("CustomerID", adVarChar, adParamInput, 10, IIf(m_strCustomerID = vbNullString, Null, m_strCustomerID))
        End With
        With objRs
                .CursorLocation = adUseClient
                .Open objCmd, , adOpenDynamic, adLockReadOnly
                Set .ActiveConnection = Nothing
        End With
        Exec_prc_sel_Customers = objCmd.Parameters("RETURN_VALUE")
        objCn.Close
        Set objCn = Nothing

        Set objCmd = Nothing
        Exit Function
PROC_ERR:
        Exec_prc_sel_Customers = Err.Number
        m_strErrDesc = Err.Description
End Function

