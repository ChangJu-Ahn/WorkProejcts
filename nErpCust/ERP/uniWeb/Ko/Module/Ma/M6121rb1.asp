<%@ LANGUAGE="VBSCRIPT" %>
<%Option Explicit    %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Prucurement
'*  2. Function Name        : 
'*  3. Program ID           : m6121rb1
'*  4. Program Name         : 배부내역 
'*  5. Program Desc         : 배부내역 
'*  6. Component List       : 
'*  7. Modified date(First) : 2004/11/15
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Byun Jee Hyun
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change" 
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************
-->
<!-- #Include file="../../inc/incSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDBAgent.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgentVariables.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%
Call LoadBasisGlobalInf()
Call LoadInfTB19029B("I", "*", "NOCOOKIE", "RB")
Call LoadBNumericFormatB("I", "*", "NOCOOKIE", "RB")

On Error Resume Next

Dim ADF										'ActiveX Data Factory 지정 변수선언 
Dim strRetMsg								'Record Set Return Message 변수선언 
Dim UNISqlId, UNIValue, UNILock, UNIFlag	'DBAgent Parameter 선언 
Dim rs0

Dim strPlantCd
Dim strBatchJobDt
Dim strDisbDt
Dim strDocumentNo

Dim LngMaxRow
Dim TmpBuffer
Dim iTotalStr
Dim LoopCnt
Dim i
Dim strData
Dim strQryMode
Dim strNextKey1, strNextKey2, strNextKey3

Const C_SHEETMAXROWS_D = 100

Err.Clear 

Call HideStatusWnd 

strQryMode = Request("lgIntFlgMode")

strPlantCd = FilterVar(Ucase(Trim(Request("txtPlantCd"))), "''","S")
strBatchJobDt = FilterVar(Ucase(Trim(Request("txtBatchJobDt"))), "''","S")
strDisbDt = FilterVar(Ucase(Trim(Request("txtDisbDt"))), "''","S")
strDocumentNo = FilterVar(Ucase(Trim(Request("txtDocumentNo"))), "''","S")

strNextKey1 = FilterVar(UCase(Trim(Request("lgStrPrevKey1"))), "''", "S")
strNextKey2 = FilterVar(UCase(Trim(Request("lgStrPrevKey2"))), "''", "S")
strNextKey3 = FilterVar(UCase(Trim(Request("lgStrPrevKey3"))), "''", "S")

LngMaxRow = Clng(Request("txtMaxRows"))

Redim UNISqlId(0)
Redim UNIValue(0, 4)

UNISqlId(0) = "M6121RA101"

UNIValue(0, 0) = strPlantCd
UNIValue(0, 1) = strBatchJobDt
UNIValue(0, 2) = strDisbDt
UNIValue(0, 3) = strDocumentNo

Select Case strQryMode
		Case CStr(OPMD_CMODE)
			
			UNIValue(0, 4) = "|"
			
		Case CStr(OPMD_UMODE) 
			UNIValue(0, 4) = " (c.charge_no > " & strNextKey1  
			UNIValue(0, 4) = UNIValue(0, 4) & " or (c.charge_no = " & strNextKey1 
			UNIValue(0, 4) = UNIValue(0, 4) & " and c.seq > " & strNextKey2 & ") "
			UNIValue(0, 4) = UNIValue(0, 4) & " or (c.charge_no = " & strNextKey1 
			UNIValue(0, 4) = UNIValue(0, 4) & " and c.seq = " & strNextKey2 
			UNIValue(0, 4) = UNIValue(0, 4) & " and c.mvmt_no >= " & strNextKey3 & ")) "
	End Select

UNILock = DISCONNREAD :	UNIFlag = "1"

Set ADF = Server.CreateObject("prjPublic.cCtlTake")

strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)

If (rs0.EOF And rs0.BOF) Then
	Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
	rs0.Close
	Set rs0 = Nothing
	Set ADF = Nothing	
else
	If C_SHEETMAXROWS_D < rs0.RecordCount Then 
		LoopCnt = C_SHEETMAXROWS_D - 1
	Else
		LoopCnt = rs0.RecordCount - 1
	End If
	
	ReDim TmpBuffer(LoopCnt)
		
	For i=0 to LoopCnt
		strData = ""
		strData = strData & Chr(11) & ConvSPChars(rs0("charge_no"))
		strData = strData & Chr(11) & ConvSPChars(rs0("bas_no"))
		strData = strData & Chr(11) & ConvSPChars(rs0("seq"))
		strData = strData & chr(11) & ConvSPChars(rs0("disb_type"))
		strData = strData & Chr(11) & UNINumClientFormat(rs0("item_qty"),ggQty.DecPoint,0)
		strData = strData & Chr(11) & UniConvNumDBToCompanyWithOutChange(rs0("base_charge_amt"),0)
		strData = strData & Chr(11) & UNINumClientFormat(rs0("mvmt_qty"),ggQty.DecPoint,0)
		strData = strData & Chr(11) & UniConvNumDBToCompanyWithOutChange(rs0("disb_amt"),0)
		strData = strData & Chr(11) & ConvSPChars(rs0("plant_cd"))
		strData = strData & Chr(11) & ConvSPChars(rs0("plant_nm"))
		strData = strData & Chr(11) & ConvSPChars(rs0("item_cd"))
		strData = strData & Chr(11) & ConvSPChars(rs0("item_nm"))
		strData = strData & Chr(11) & ConvSPChars(rs0("spec"))
		strData = strData & Chr(11) & ConvSPChars(rs0("mvmt_rcpt_no"))
		strData = strData & Chr(11) & ConvSPChars(rs0("po_no"))
		strData = strData & Chr(11) & ConvSPChars(rs0("po_seq_no"))
		
		strData = strData & Chr(11) & LngMaxRow + i
		strData = strData & Chr(11) & Chr(12)
		
		'Call ServerMesgBox(strData, vbCritical, I_MKSCRIPT) 
			
		TmpBuffer(i) = strData
		rs0.MoveNext
	Next
	
	If Not (rs0.eof) then
		strNextKey1 = ConvSPChars(rs0("charge_no"))
		strNextKey2 = ConvSPChars(rs0("seq"))
		strNextKey3 = ConvSPChars(rs0("mvmt_no"))
		
	Else
		strNextKey1 = ""	
		strNextKey2 = ""
		strNextKey3 = ""
	End if
		
	iTotalStr = Join(TmpBuffer, "")
		
	rs0.Close
	Set rs0 = Nothing
	Set ADF = Nothing
end if

Response.Write "<Script Language=vbscript>" & vbCr
Response.Write "With parent" & vbCr												'☜: 화면 처리 ASP 를 지칭함 
Response.Write "	.ggoSpread.Source          =  .frm1.vspdData " & vbCr
Response.Write "	.ggoSpread.SSShowData        """ & iTotalStr & """" & vbCr	
Response.Write "	.frm1.vspdData.Redraw = false " & vbCr
Response.Write "	.frm1.vspdData.Redraw = True " & vbCr
Response.Write "	.lgStrPrevKey1              = """ & StrNextKey1 & """" & vbCr
Response.Write "	.lgStrPrevKey2              = """ & StrNextKey2 & """" & vbCr  
Response.Write "	.lgStrPrevKey3              = """ & StrNextKey3 & """" & vbCr    
Response.Write " .DbQueryOk "		    	  & vbCr 
Response.Write " .frm1.vspdData.focus "		  & vbCr 
Response.Write "End With" & vbCr
Response.Write "</Script>" & vbCr

%>