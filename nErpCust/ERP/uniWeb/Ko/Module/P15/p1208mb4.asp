<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgentVariables.inc" -->
<%'======================================================================================================
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p1208mb4.asp
'*  4. Program Name         : List Standard Manufacturing Instruction Copy 
'*  5. Program Desc         : 
'*  6. Modified date(First) : 2002/03/26
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
Dim iIntCnt, iLngMaxRows

iLngMaxRows = Request("txtMaxRows")

Redim UNISqlId(0)
Redim UNIValue(0, 3)

UNISqlId(0) = "P1208MB4"	

UNIValue(0, 0) = "^"
UNIValue(0, 1) = FilterVar(UCase(Request("txtStdWISet")), "''", "S")
UNIValue(0, 2) = " " & FilterVar(UniConvDate(Request("txtValidDt")), "''", "S") & ""
UNIValue(0, 3) = " " & FilterVar(UniConvDate(Request("txtValidDt")), "''", "S") & ""
	
		
UNILock = DISCONNREAD :	UNIFlag = "1"
	
Set ADF = Server.CreateObject("prjPublic.cCtlTake")
strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
      
If (rs0.EOF And rs0.BOF) Then
	Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
	rs0.Close
	Set rs0 = Nothing
	Response.End													'☜: 비지니스 로직 처리를 종료함 
End If
	
Response.Write "<Script Language = VBScript>" & vbCrLf
Response.Write "With parent" & vbCrLf

	If Not(rs0.EOF And rs0.BOF) Then
		For iIntCnt = 0 To rs0.RecordCount - 1

			Response.Write ".frm1.vspdData2.focus" & vbCrLf
			Response.Write "Set .gActiveElement = .document.activeElement " & vbCrLf
			Response.Write ".ggoSpread.Source = .frm1.vspdData2" & vbCrLf
			Response.Write ".ggoSpread.InsertRow" & vbCrLf
			Response.Write ".frm1.vspdData2.Row = .frm1.vspdData2.ActiveRow" & vbCrLf
		    Response.Write ".frm1.vspdData2.Col = .C_WICd2" & vbCrLf
			Response.Write ".frm1.vspdData2.Text = """ & ConvSPChars(rs0("MFG_INSTRUCTION_DTL_CD")) & """" & vbCrLf
			Response.Write ".frm1.vspdData2.Col = .C_WIDesc2" & vbCrLf
			Response.Write ".frm1.vspdData2.Text = """ & ConvSPChars(rs0("MFG_INSTRUCTION_DTL_DESC")) & """" & vbCrLf
			Response.Write ".frm1.vspdData2.Col = .C_ValidStartDt2" & vbCrLf
			Response.Write ".frm1.vspdData2.Text = """ & UNIDateClientFormat(rs0("VALID_START_DT")) & """" & vbCrLf
			Response.Write ".frm1.vspdData2.Col = .C_ValidEndDt2" & vbCrLf
			Response.Write ".frm1.vspdData2.Text = """ & UNIDateClientFormat(rs0("VALID_END_DT")) & """" & vbCrLf
			Response.Write ".frm1.vspdData2.Col = .C_PlantCd2" & vbCrLf
			Response.Write ".frm1.vspdData2.Text = """ & Trim(UCase(Request("txtPlantCd"))) & """" & vbCrLf
			Response.Write ".frm1.vspdData2.Col = .C_ItemCd2" & vbCrLf
			Response.Write ".frm1.vspdData2.Text = """ & Trim(UCase(Request("txtItemCd"))) & """" & vbCrLf
			Response.Write ".frm1.vspdData2.Col = .C_RoutingNo2" & vbCrLf
			Response.Write ".frm1.vspdData2.Text = """ & Trim(UCase(Request("txtRoutNo"))) & """" & vbCrLf
			Response.Write ".frm1.vspdData2.Col = .C_OprNo2" & vbCrLf
			Response.Write ".frm1.vspdData2.Text = """ & Trim(UCase(Request("txtOprNo"))) & """" & vbCrLf	
			
			Response.Write ".frm1.vspdData2.Col = 0" & vbCrLf
			Response.Write ".frm1.vspdData2.Text = .ggoSpread.InsertFlag " & vbCrLf
			
			Response.Write ".SetSpreadColor .frm1.vspdData2.ActiveRow, .frm1.vspdData2.ActiveRow" & vbCrLf

			rs0.MoveNext
		Next
	End If

	rs0.Close
	Set rs0 = Nothing
		
	Response.Write ".SetStdWISetOk(" & iLngMaxRows & " + 1)" & vbCrLf
Response.Write "End With" & vbCrLf
Response.Write "</Script>" & vbCrLf

Set ADF = Nothing												'☜: ActiveX Data Factory Object Nothing
%>    
