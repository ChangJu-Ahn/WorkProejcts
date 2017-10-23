<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgentVariables.inc" -->
<!-- #Include file="../../ComAsp/LoadinfTB19029.asp" -->
<%
'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p1205mb2.asp
'*  4. Program Name         : Component Allocation (Query)
'*  5. Program Desc         :
'*  6. Component List       : DB Agent (p1205mb2)
'*  7. Modified date(First) : 2002/06/24
'*  8. Modified date(Last)  : 2002/11/21
'*  9. Modifier (First)     : Park, BumSoo
'* 10. Modifier (Last)      : Hong Chang Ho
'* 11. Comment              :
'**********************************************************************************************

On Error Resume Next

Call HideStatusWnd															'☜: 모든 작업 완료후 작업진행중 표시창을 Hide
Err.Clear

Call LoadBasisGlobalInf
Call LoadinfTB19029B("I", "*", "NOCOOKIE", "MB")

Dim ADF
Dim strRetMsg
Dim UNISqlId, UNIValue, UNILock, UNIFlag
Dim rs0
Dim iIntCnt, iLngMaxRows, strQryMode, iStrPrevKey2
Dim strData
Dim TmpBuffer
Dim iTotalStr

strQryMode = Request("lgIntFlgMode")
iStrPrevKey2 = Request("lgStrPrevKey2")
iLngMaxRows = Request("txtMaxRows")

Redim UNISqlId(0)
Redim UNIValue(0, 4)

UNISqlId(0) = "p1205mb2a"	'// 기존 p1205mb2를 p1205mb2a로 ID변경 

UNIValue(0, 0) = "^"
UNIValue(0, 1) = FilterVar(Request("txtPlantCd"), "''", "S")
UNIValue(0, 2) = FilterVar(Request("txtItemCd"), "''", "S")
UNIValue(0, 3) = FilterVar(Request("txtRoutNo"), "''", "S")
UNIValue(0, 4) = FilterVar(Request("txtOprNo"), "''", "S")
	

UNILock = DISCONNREAD :	UNIFlag = "1"
	
Set ADF = Server.CreateObject("prjPublic.cCtlTake")
strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)

If (rs0.EOF And rs0.BOF) Then
	Call DisplayMsgBox("181500", vbOKOnly, "", "", I_MKSCRIPT)
	rs0.Close
	Set rs0 = Nothing
	Set ADF = Nothing
	Response.Write "<Script Language = VBScript>" & vbCrLf
	Response.Write "		With Parent " & vbCrLf
	Response.Write "			Call .SetFocusToDocument(""M"") " & vbCrLf
	Response.Write "			.frm1.vspdData1.focus " & vbCrLf
	Response.Write "		End With " & vbCrLf
	Response.Write "</Script>" & vbCrLf
	Response.End
End If

Response.Write "<Script Language = VBScript>" & vbCrLf
Response.Write "With parent" & vbCrLf
	
If Not(rs0.EOF And rs0.BOF) Then
	ReDim TmpBuffer(rs0.RecordCount - 1)
	For iIntCnt = 0 To rs0.RecordCount - 1
		strData = ""
	    strData = strData & Chr(11) & UniConvNumberDBToCompany(rs0("RANK"), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
		strData = strData & Chr(11) & ConvSPChars(rs0("RESOURCE_CD"))						'Resource Cd
		strData = strData & Chr(11) & ""													'Resource Popup
		strData = strData & Chr(11) & ConvSPChars(rs0("RESOURCE_NM"))						'Resource Description
		strData = strData & Chr(11) & ConvSPChars(rs0("RESOURCE_TYPE_NM"))					'
		strData = strData & Chr(11) & UniConvNumberDBToCompany(rs0("BOR_EFFICIENCY"), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
	    strData = strData & Chr(11) & (iLngMaxRows + iIntCnt)
		strData = strData & Chr(11) & Chr(12)
		TmpBuffer(iIntCnt) = strData
		rs0.MoveNext
	Next
	
	iTotalStr = Join(TmpBuffer, "")

	Response.Write ".ggoSpread.Source = .frm1.vspdData2" & vbCrLf
	Response.Write ".ggoSpread.SSShowDataByClip """ & iTotalStr & """" & vbCrLf
		
	If rs0("OPR_NO") = Null Then
		Response.Write ".lgStrPrevKey2 = """"" & vbCrLf
	Else
		Response.Write ".lgStrPrevKey2 = """ & Trim(rs0("OPR_NO")) & """" & vbCrLf
	End If
End If

rs0.Close
Set rs0 = Nothing

Response.Write "If .frm1.vspdData2.MaxRows < .VisibleRowCnt(.frm1.vspdData2, 0) And .lgStrPrevKey2 <> """" Then" & vbCrLf
Response.Write "	.initData(" & iLngMaxRows & " + 1)" & vbCrLf
Response.Write "	.DbDtlQuery" & vbCrLf
Response.Write "Else" & vbCrLf
Response.Write "	.frm1.hPlantCd.value = """ & ConvSPChars(Request("txtPlantCd")) & """" & vbCrLf
Response.Write "	.frm1.hItemCd.value = """ & ConvSPChars(Request("txtItemCd")) & """" & vbCrLf
Response.Write "	.frm1.hRoutNo.value = """ & ConvSPChars(Request("txtRoutNo")) & """" & vbCrLf

Response.Write "	.DbDtlQueryOk(" & iLngMaxRows & " + 1)" & vbCrLf
Response.Write "End If" & vbCrLf

Response.Write "End With" & vbCrLf

Response.Write "</Script>" & vbCrLf
Set ADF = Nothing												'☜: ActiveX Data Factory Object Nothing
%>

<Script Language = VBScript RUNAT = Server>
Function ConvToTimeFormat(ByVal iVal)
	Dim iTime
	Dim iMin
	Dim iSec
				
	If IVal = 0 Then
		ConvToTimeFormat = "00:00:00"
	Else
		iMin = Fix(IVal Mod 3600)
		iTime = Fix(IVal /3600)
		
		iSec = Fix(iMin Mod 60)
		iMin = Fix(iMin / 60)
		
		ConvToTimeFormat = Right("0" & CStr(iTime),2) & ":" & Right("0" & CStr(iMin),2) & ":" & Right("0" & CStr(iSec),2)
		 
	End If
End Function
</Script>
