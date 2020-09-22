Attribute VB_Name = "basFunctions"
Option Explicit

Public Function getTypeDesc(xNo As Integer, xCon As ADODB.Connection) As String
    Dim strType As String
    Dim rsGetType As ADODB.Recordset
    Set rsGetType = New ADODB.Recordset
    strType = "Select * from tblType where typeno = " & xNo & " order by typedesc"
    rsGetType.Open strType, xCon, adOpenStatic, adLockOptimistic
    getTypeDesc = rsGetType.Fields("typedesc")
End Function

Public Function getTypeNo(xTypeDesc As String, xCon As ADODB.Connection) As Integer
    Dim strType As String
    Dim rsGetType As ADODB.Recordset
    Set rsGetType = New ADODB.Recordset
    strType = "Select * from tblType where typedesc = '" & xTypeDesc & "' order by typedesc"
    rsGetType.Open strType, xCon, adOpenStatic, adLockOptimistic
    getTypeNo = rsGetType.Fields("typeno")
End Function
