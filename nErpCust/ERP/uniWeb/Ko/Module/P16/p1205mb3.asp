<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<%
'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p1201mb3.asp
'*  4. Program Name         : Entry Routing Component Allocation(Create, Update, Delete)
'*  5. Program Desc         :
'*  6. Component List       : PP1S608.cPMngBillOfRsrc
'*  7. Modified date(First) : 2000/03/27
'*  8. Modified date(Last)  : 2002/11/21
'*  9. Modifier (First)     : Im Hyun Soo
'* 10. Modifier (Last)      : Hong Chang Ho
'* 11. Comment              :
'**********************************************************************************************

On Error Resume Next

Call HideStatusWnd															'��: ��� �۾� �Ϸ��� �۾������� ǥ��â�� Hide
Err.Clear

Call LoadBasisGlobalInf

Dim pPP1S608																	'�� : �Է�/������ Component Dll ��� ���� 
Dim I1_plant_cd, I2_item_cd, I3_rout_no, I4_opr_no, l5_rank, l6_efficiency, iErrorPosition, strSpread

If Request("hPlantCd") = "" Then												'��: ��ȸ�� ���� ���� ���Դ��� üũ 
	Call DisplayMsgBox("900010", vbOKOnly, "", "", I_MKSCRIPT)              
	Response.End 
ElseIf Request("hRoutNo") = "" Then
	Call DisplayMsgBox("900010", vbOKOnly, "", "", I_MKSCRIPT)              
	Response.End 
ElseIf Request("hOprNo") = "" Then
	Call DisplayMsgBox("900010", vbOKOnly, "", "", I_MKSCRIPT)           
	Response.End 
ElseIf Request("hItemCd") = "" Then
	Call DisplayMsgBox("900010", vbOKOnly, "", "", I_MKSCRIPT)           
	Response.End 
End If
	
'-----------------------
'Data manipulate area
'-----------------------
strSpread = Request("txtSpread")
I1_plant_cd = UCase(Trim(Request("hPlantCd")))
I2_item_cd = UCase(Trim(Request("hItemCd")))
I3_rout_no = UCase(Trim(Request("hRoutNo")))
I4_opr_no = UCase(Trim(Request("hOprNo")))

Set pPP1S608 = Server.CreateObject("PP1S608.cPMngBillOfRsrc")

If CheckSYSTEMError(Err,True) = True Then
	Response.End
End If

Call pPP1S608.P_MANAGE_BILL_OF_RESOURCE2(gStrGlobalCollection, strSpread, I1_plant_cd, I2_item_cd, _
										I3_rout_no, I4_opr_no, iErrorPosition)

If CheckSYSTEMError2(Err, True, iErrorPosition & "��", "", "", "", "") = True Then
	Set pPP1S608 = Nothing														'��: Unload Component
	Response.End
End If

Set pPP1S608 = Nothing															'��: Unload Component

Response.Write "<Script Language=VBScript>" & vbCrLf
	Response.Write "parent.DbSaveOk" & vbCrLf
Response.Write "</Script>" & vbCrLf
Response.End
%>