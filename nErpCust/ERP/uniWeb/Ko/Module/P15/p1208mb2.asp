<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgentVariables.inc" -->
<%
'======================================================================================================
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p1208mb2.asp
'*  4. Program Name         : List Manufacturing Instruction (Lower Grid)
'*  5. Program Desc         : 
'*  6. Modified date(First) : 2002. 3. 24
'*  7. Modified date(Last)  : 2002/11/20
'*  8. Modifier (First)     : JaeHyun Chen
'*  9. Modifier (Last)      : Hong Chang Ho
'* 10. Comment              : 
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================

On Error Resume Next

Call HideStatusWnd															'☜: 모든 작업 완료후 작업진행중 표시창을 Hide
Err.Clear

Call LoadBasisGlobalInf

Dim ADF										'ActiveX Data Factory 지정 변수선언 
Dim strRetMsg								'Record Set Return Message 변수선언 
Dim UNISqlId, UNIValue, UNILock, UNIFlag	'DBAgent Parameter 선언 
Dim rs0										'DBAgent Parameter 선언 
Dim iIntCnt, iLngMaxRows3
Dim strData2, strData3
Dim TmpBuffer2, TmpBuffer3
Dim iTotalStr2, iTotalStr3

iLngMaxRows3 = Request("txtMaxRows3")

Redim UNISqlId(0)
Redim UNIValue(0, 6)

UNISqlId(0) = "P1208MB2"	

UNIValue(0, 0) = "^"
UNIValue(0, 1) = FilterVar(Request("txtPlantCd"), "''", "S")
UNIValue(0, 2) = FilterVar(Request("txtItemCd"), "''", "S")
UNIValue(0, 3) = FilterVar(Request("txtRoutNo"), "''", "S")
UNIValue(0, 4) = FilterVar(Request("txtOprNo"), "''", "S")
UNIValue(0, 5) = " " & FilterVar(UniConvDate(Request("txtStdDt")), "''", "S") & ""
UNIValue(0, 6) = " " & FilterVar(UniConvDate(Request("txtStdDt")), "''", "S") & ""
		
UNILock = DISCONNREAD :	UNIFlag = "1"
	
Set ADF = Server.CreateObject("prjPublic.cCtlTake")
strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
      
If (rs0.EOF And rs0.BOF) Then
	Call DisplayMsgBox("181421", vbOKOnly, "", "", I_MKSCRIPT)
	rs0.Close
	Set rs0 = Nothing
	Response.End													'☜: 비지니스 로직 처리를 종료함 
End If
	
Response.Write "<Script Language = VBScript>" & vbCrLf
Response.Write "With parent" & vbCrLf

	If Not(rs0.EOF And rs0.BOF) Then
	
		ReDim TmpBuffer2(rs0.RecordCount - 1)
		ReDim TmpBuffer3(rs0.RecordCount - 1)
		
		For iIntCnt = 0 To rs0.RecordCount - 1 
			strData2 = ""
			strData2 = strData2 & Chr(11) & ConvSPChars(rs0("SEQ"))
			strData2 = strData2 & Chr(11) & ConvSPChars(rs0("MFG_INSTRUCTION_DTL_CD"))
			strData2 = strData2 & Chr(11) & ""	' Work Center Popup
			strData2 = strData2 & Chr(11) & ConvSPChars(rs0("MFG_INSTRUCTION_DTL_DESC"))
			strData2 = strData2 & Chr(11) & UNIDateClientFormat(rs0("VALID_START_DT"))
			strData2 = strData2 & Chr(11) & UNIDateClientFormat(rs0("VALID_END_DT"))
			strData2 = strData2 & Chr(11) & ConvSPChars(rs0("PLANT_CD"))
			strData2 = strData2 & Chr(11) & ConvSPChars(rs0("ITEM_CD"))
			strData2 = strData2 & Chr(11) & ConvSPChars(rs0("ROUT_NO"))
			strData2 = strData2 & Chr(11) & ConvSPChars(rs0("OPR_NO"))
			strData2 = strData2 & Chr(11) & ConvSPChars(rs0("SEQ"))
			strData2 = strData2 & Chr(11) & iIntCnt
			strData2 = strData2 & Chr(11) & Chr(12)
			TmpBuffer2(iIntCnt) = strData2
			
			strData3 = ""
			strData3 = strData3 & Chr(11) & ConvSPChars(rs0("SEQ"))
			strData3 = strData3 & Chr(11) & ConvSPChars(rs0("MFG_INSTRUCTION_DTL_CD"))
			strData3 = strData3 & Chr(11) & ""	' Work Center Popup
			strData3 = strData3 & Chr(11) & ConvSPChars(rs0("MFG_INSTRUCTION_DTL_DESC"))
			strData3 = strData3 & Chr(11) & UNIDateClientFormat(rs0("VALID_START_DT"))
			strData3 = strData3 & Chr(11) & UNIDateClientFormat(rs0("VALID_END_DT"))
			strData3 = strData3 & Chr(11) & ConvSPChars(rs0("PLANT_CD"))
			strData3 = strData3 & Chr(11) & ConvSPChars(rs0("ITEM_CD"))
			strData3 = strData3 & Chr(11) & ConvSPChars(rs0("ROUT_NO"))
			strData3 = strData3 & Chr(11) & ConvSPChars(rs0("OPR_NO"))
			strData3 = strData3 & Chr(11) & ConvSPChars(rs0("SEQ"))
			strData3 = strData3 & Chr(11) & (iLngMaxRows3 + iIntCnt)
			strData3 = strData3 & Chr(11) & Chr(12)
			TmpBuffer3(iIntCnt) = strData3
			rs0.MoveNext
		Next
		iTotalStr2 = Join(TmpBuffer2, "")
		iTotalStr3 = Join(TmpBuffer3, "")
		Response.Write ".ggoSpread.Source = .frm1.vspdData2" & vbCrLf
		Response.Write ".ggoSpread.SSShowDataByClip """ & iTotalStr2 & """" & vbCrLf
		Response.Write ".ggoSpread.Source = .frm1.vspdData3" & vbCrLf
		Response.Write ".ggoSpread.SSShowDataByClip """ & iTotalStr3 & """" & vbCrLf
		
	End If

	rs0.Close
	Set rs0 = Nothing

	Response.Write ".DbDtlQueryOk(" & iLngMaxRows3 & " + 1)" & vbCrLf

Response.Write "End With" & vbCrLf

Response.Write "</Script>" & vbCrLf	
Set ADF = Nothing												'☜: ActiveX Data Factory Object Nothing
%>
