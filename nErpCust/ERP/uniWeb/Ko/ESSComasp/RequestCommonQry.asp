<% Option Explicit%>
<!-- #Include file="../inc/CommResponse.inc" -->
<!-- #Include file="../inc/adovbs.inc" -->
<%
Dim iADODBConnString
Dim iStrSQL
Dim iVarArray
Dim iLngErrCode
Dim iStrErrDesc
Dim iReturnCode

iADODBConnString = Trim(Request("ADODBConnString"))
iStrSQL = Request("StrSQL")

iReturnCode = CommonQueryRs(iADODBConnString,iStrSQL,iVarArray,iLngErrCode,iStrErrDesc)

If iReturnCode = True Then
    Response.Write iVarArray
End If

Function CommonQueryRs(ByVal gADODBConnString , ByVal pvStrSQL, prVarArray, prLngErrCode, prStrErrDesc)
    Dim adoRs
    
    On Error Resume Next
    Err.Clear
    
    CommonQueryRs = True
    prLngErrCode = 0
    prStrErrDesc = ""

    Set adoRs = Server.CreateObject("ADODB.Recordset")

    adoRs.Open pvStrSQL, gADODBConnString, adOpenForwardOnly, adLockReadOnly, adCmdText
    If Not (adoRs.EOF And adoRs.BOF) Then
       prVarArray = adoRs.GetString(, , Chr(11), Chr(12))
    Else
       CommonQueryRs = False
    End If
    Call CloseAdoObject(adoRs)
    If Err.number <> 0 Then
        CommonQueryRs = False
        prLngErrCode = Err.Number
        prStrErrDesc = Err.Description
    End If
End Function

Sub CloseAdoObject(pObject)

    If Not (pObject Is Nothing) Then
       If pObject.State = adStateOpen Then
          pObject.Close
       End If
       Set pObject = Nothing
    End If
End Sub
%>