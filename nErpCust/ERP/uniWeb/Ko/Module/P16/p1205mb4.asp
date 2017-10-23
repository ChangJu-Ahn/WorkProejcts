<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgentVariables.inc" -->
<!-- #Include file="../../ComAsp/LoadinfTB19029.asp" -->
<!-- #Include file="../../inc/adoVbs.inc" -->
<!-- #Include file="../../inc/incServerAdoDb.asp" -->
<%
'======================================================================================================
'*  1. Module Name          : Basic Architect
'*  2. Function Name        : ADO Template
'*  3. Program ID           : p1205mb4.asp
'*  4. Program Name         :
'*  5. Program Desc         :
'*  6. Modified date(First) : 2000/11/27
'*  7. Modified date(Last)  : 2002/11/21
'*  8. Modifier (First)     : Jung Yu Kyung
'*  9. Modifier (Last)      : Hong Chang Ho
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================

Call HideStatusWnd															'☜: 모든 작업 완료후 작업진행중 표시창을 Hide

On Error Resume Next
Err.Clear

Call LoadBasisGlobalInf
Call LoadinfTB19029B("I", "*", "NOCOOKIE", "MB")

Dim ADF
Dim strRetMsg
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1
Dim iLngCurRows

iLngCurRows = Request("lgLngCurRows")
	

Redim UNISqlId(0)
Redim UNIValue(0, 1)
	
UNISqlId(0) = "180000san"
	
UNIValue(0, 0) = FilterVar(UCase(Request("txtPlantCd")), "''", "S")
UNIValue(0, 1) = FilterVar(UCase(Request("txtResourceCd")), "''", "S")

UNILock = DISCONNREAD :	UNIFlag = "1"
	
Set ADF = Server.CreateObject("prjPublic.cCtlTake")
strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
      
If (rs0.EOF And rs0.BOF) Then
	Call DisplayMsgBox("181600", vbOKOnly, "", "", I_MKSCRIPT)
	Response.Write "<Script Language=VBScript>" & vbCrLf
	Response.Write "With parent" & vbCrLf
		Response.Write ".frm1.vspdData2.Row = " & iLngCurRows & vbCrLf	    	
		Response.Write ".frm1.vspdData2.Col = .C_ResourceNm" & vbCrLf
		Response.Write ".frm1.vspdData2.Text =  """"" & vbCrLf							
		Response.Write ".frm1.vspdData2.Col = .C_ResourceType" & vbCrLf
		Response.Write ".frm1.vspdData2.Text =  """"" & vbCrLf
	Response.Write "End With" & vbCrLf
	Response.Write "</Script>" & vbCrLf

	rs0.Close
	Set rs0 = Nothing
	Set ADF = Nothing
	Response.End
End If

Response.Write "<Script Language=VBScript>" & vbCrLf
Response.Write "With parent" & vbCrLf
Response.Write "	.frm1.vspdData2.Row = " & iLngCurRows & vbCrLf
		
Response.Write "	.frm1.vspdData2.Col = .C_ResourceCd" & vbCrLf
Response.Write "	.frm1.vspdData2.Text = """ & rs0("RESOURCE_CD") & """" & vbCrLf
    	
Response.Write "	.frm1.vspdData2.Col = .C_ResourceNm" & vbCrLf
Response.Write "	.frm1.vspdData2.Text = """ & rs0("DESCRIPTION") & """" & vbCrLf
				
Response.Write "	.frm1.vspdData2.Col = .C_ResourceType" & vbCrLf
Response.Write "	.frm1.vspdData2.Text = """ & FuncCodeName(1,"P1502",rs0("resource_type")) & """" & vbCrLf
Response.Write "End With" & vbCrLf
Response.Write "</Script>" &vbCrLf

rs0.Close				
Set rs0 = Nothing				
Set ADF = Nothing
%>
