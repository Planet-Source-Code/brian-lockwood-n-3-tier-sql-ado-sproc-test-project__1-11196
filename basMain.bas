Attribute VB_Name = "basMain"
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

Public g_objCn As New ADODB.Connection
Public Const g_strConnectionString = "Northwind"

Sub main()
    Screen.MousePointer = vbHourglass
    g_objCn.Open g_strConnectionString
    Screen.MousePointer = vbNormal
    frmTest.Show
End Sub



