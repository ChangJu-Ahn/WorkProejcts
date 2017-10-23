<%@ LANGUAGE = VBSCript%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<%
'**********************************************************************************************
'*  1. Module Name          : ����BOM���� 
'*  2. Function Name        : 
'*  3. Program ID           : p1713mb14.asp
'*  4. Program Name         : EBOM_TO_PBOM_MASTER & DETAIL Delete
'*  5. Program Desc         :
'*  6. Component List       : PP1S407.cYTransBomHdrMulti
'*  7. Modified date(First) : 2001/10/30
'*  8. Modified date(Last)  : 2002/11/19
'*  9. Modifier (First)     : Jung Yu Kyung
'* 10. Modifier (Last)      : Hong Chang Ho
'* 11. Comment              :
'**********************************************************************************************

On Error Resume Next                                                             '��: Protect system from crashing
Err.Clear                                                                        '��: Clear Error status

Call HideStatusWnd															'��: ��� �۾� �Ϸ��� �۾������� ǥ��â�� Hide

Call LoadBasisGlobalInf

Dim pPY3S113																	'�� : �Է�/������ ComProxy Dll ��� ���� 
Dim iCommandSent, iErrorPosition
Dim I1_select_char, I2_y_trans_bom_header, I3_plant_cd, I4_item_cd

Const Y311_I2_bom_no			= 0 
Const Y311_I2_req_trans_no		= 1 

If Request("txtDestPlantCd") = "" Then														'��: ��ȸ�� ���� ���� ���Դ��� üũ 
	Call DisplayMsgBox("700114", vbOKOnly, "", "", I_MKSCRIPT)            
	Response.End 
End If

If Request("txtItemCd") = "" Then													'��: ��ȸ�� ���� ���� ���Դ��� üũ 
	Call DisplayMsgBox("700114", vbOKOnly, "", "", I_MKSCRIPT)            
	Response.End
End If
	
'-----------------------
'Data manipulate area
'-----------------------												'��: Single ����Ÿ ���� 
Redim I2_y_trans_bom_header(Y311_I2_req_trans_no)

'I2_y_trans_bom_header(Y311_I2_bom_no)	= UCase(Trim(Request("txtBomType")))
I2_y_trans_bom_header(Y311_I2_bom_no)	= "1"
I2_y_trans_bom_header(Y311_I2_req_trans_no)	= UCase(Trim(Request("txtReqTransNo")))
I3_plant_cd		= UCase(Trim(Request("txtDestPlantCd")))
I4_item_cd		= UCase(Trim(Request("txtItemCd")))

iCommandSent = "DELETE"
	
'-----------------------
'Com action result check area(OS,internal)
'-----------------------
Set pPY3S113 = Server.CreateObject("PY3S113.cYTransBomHdrMulti")    

If CheckSYSTEMError(Err,True) = True Then
	Response.End
End If

Call pPY3S113.Y_TRANS_BOM_HEADER_MULTI(gStrGlobalCollection, iCommandSent, "", _
				 I1_select_char, I2_y_trans_bom_header, I3_plant_cd, I4_item_cd, iErrorPosition)

If CheckSYSTEMError2(Err, True, iErrorPosition & "��", "", "", "", "") = True Then
	Set pPY3S113 = Nothing															'��: Unload Component
	Response.End
End If

Set pPY3S113 = Nothing													'��: Unload Comproxy
	
Response.Write "<Script Language = VBScript>" & vbCrLf
	Response.Write "parent.DbDeleteOk" & vbCrLf
Response.Write "</Script>" & vbCrLf
Response.End
%>