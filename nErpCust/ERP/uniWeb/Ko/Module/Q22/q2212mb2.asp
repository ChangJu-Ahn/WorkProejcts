<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029B("I", "*", "NOCOOKIE","MB") %>
<%Call LoadBasisGlobalInf%>
<%
'**********************************************************************************************
'*  1. Module Name          : Quality Management
'*  2. Function Name        : 
'*  3. Program ID           : Q2212MB2
'*  4. Program Name         : ������� 
'*  5. Program Desc         : Quality Configuration
'*  6. Component List       : 
'*  7. Modified date(First) : 2002/05/14
'*  8. Modified date(Last)  : 2003/05/15
'*  9. Modifier (First)     : Koh Jae Woo
'* 10. Modifier (Last)      : Park Hyun Soo
'* 11. Comment
'* 12. Common Coding Guide  : this mark(��) means that "Do not change" 
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'**********************************************************************************************

On Error Resume Next
Call HideStatusWnd

Dim strinsp_class_cd
strinsp_class_cd = "P"	'@@@���� 
Dim PQIG060																	'�� : ��ȸ�� ComProxy Dll ��� ���� 
Dim lgIntFlgMode
Dim LngMaxRow
Dim arrRowVal								'��: Spread Sheet �� ���� ���� Array ���� 
Dim arrColVal								'��: Spread Sheet �� ���� ���� Array ���� 
Dim strStatus								'��: Sheet �� ���� Row�� ���� (Create/Update/Delete)
Dim lGrpCnt								'��: Group Count
Dim strInspReqNo
Dim strPlantCd
Dim txtSpread
Dim iErrorPosition

Dim i			'2003-03-01 Release �߰� 
Dim SpdCount	'2003-03-01 Release �߰� 
	
strInspReqNo = UCase(Trim(Request("txtInspReqNo")))
strPlantCd = UCase(Trim(Request("txtPlantCd")))

'****** START
LngMaxRow = CInt(Request("txtMaxRows"))					'��: �ִ� ������Ʈ�� ���� 
'****** END
	
Dim I2_q_inspection_result
ReDim I2_q_inspection_result(2)
Const Q221_I2_insp_result_no = 0    '[CONVERSION INFORMATION]  View Name : import q_inspection_result
Const Q221_I2_plant_cd = 1
Const Q221_I2_insp_class_cd = 2

I2_q_inspection_result(Q221_I2_insp_result_no) = 1	
I2_q_inspection_result(Q221_I2_plant_cd) = strPlantCd
I2_q_inspection_result(Q221_I2_insp_class_cd) = strinsp_class_cd	'@@@���� 
	
Dim itxtSpread
Dim itxtSpreadArr
Dim itxtSpreadArrCount

Dim iCUCount
Dim iDCount

Dim ii

itxtSpread = ""
             
iCUCount = Request.Form("txtCUSpread").Count
iDCount  = Request.Form("txtDSpread").Count
             
itxtSpreadArrCount = -1
             
ReDim itxtSpreadArr(iCUCount + iDCount)
             
For ii = 1 To iDCount
    itxtSpreadArrCount = itxtSpreadArrCount + 1
    itxtSpreadArr(itxtSpreadArrCount) = Request.Form("txtDSpread")(ii)
Next
For ii = 1 To iCUCount
    itxtSpreadArrCount = itxtSpreadArrCount + 1
    itxtSpreadArr(itxtSpreadArrCount) = Request.Form("txtCUSpread")(ii)
Next

itxtSpread = Join(itxtSpreadArr,"")
	
Set PQIG060 = Server.CreateObject("PQIG060.cQMtInspMeaValSvr")

If CheckSYSTEMError(Err,True) = True Then
	Response.Write "<Script Language=VBScript>" & vbCrLF
	Response.Write "Call Parent.RemovedivTextArea" & vbCrLF
	Response.Write "</Script>" & vbCrLF
	Response.End
End If
	
'/* ��ü ���� ���� - START */
Call PQIG060.Q_MAINT_INSP_MEAS_VALUE_SVR(gStrGlobalCollection, _
										 strInspReqNo, _
										 I2_q_inspection_result, _
										 "N", _
										 itxtSpread, _
										 iErrorPosition)
'/* ��ü ���� ���� - END */
	
If CheckSYSTEMError2(Err,True,iErrorPosition & "��","","","","") = True Then
	If iErrorPosition <> "" Then			
%>
<Script Language=vbscript>
		Call Parent.RemovedivTextArea
		Call parent.SheetFocus(<%=iErrorPosition%>)
</Script>
<%	
		Set PQIG060 = Nothing
		Response.End
	End If		
	Response.End
End If	
	
Set PQIG060 = Nothing	
%>
<Script Language=vbscript>
With parent	
	.DbSaveOk															
End With
</Script>