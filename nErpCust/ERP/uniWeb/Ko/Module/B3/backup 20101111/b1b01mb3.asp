<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<%Call LoadBasisGlobalInf%>
<%
'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : b1b01mb3.asp
'*  4. Program Name         : Delete Item
'*  5. Program Desc         :
'*  6. Component List       : PB3S105.cBMngItem
'*  7. Modified date(First) : 2000/03/27
'*  8. Modified date(Last)  : 2002/11/14
'*  9. Modifier (First)     : Im Hyun Soo
'* 10. Modifier (Last)      : Hong Chang Ho
'* 11. Comment              :
'**********************************************************************************************

Call HideStatusWnd															'��: ��� �۾� �Ϸ��� �۾������� ǥ��â�� Hide

On Error Resume Next
Err.Clear

Dim pPB3S105																	'�� : ����� Component Dll ��� ���� 
Dim I2_b_item, iCommandSent

Const P025_I2_item_cd = 0
Redim I2_b_item(0)

If Request("txtItemCd1") = "" Then												'��: ��ȸ�� ���� ���� ���Դ��� üũ 
	Call DisplayMsgBox("700114", vbOKOnly, "", "", I_MKSCRIPT)           
	Response.End 
End If
    
'-----------------------
'Data manipulate area
'-----------------------												'��: Single ����Ÿ ���� 
I2_b_item(P025_I2_item_cd)  = Request("txtItemCd1")

iCommandSent = "DELETE"

Set pPB3S105 = Server.CreateObject("PB3S105.cBMngItem")
If CheckSYSTEMError(Err,True) = True Then
	Response.End
End If

Call pPB3S105.B_MANAGE_ITEM(gStrGlobalCollection, iCommandSent, , I2_b_item)
If CheckSYSTEMError(Err,True) = True Then
	Set pPB3S105 = Nothing                                                   '��: Unload Component
	Response.End
End If
	
Set pPB3S105 = Nothing                                                   '��: Unload Component

Response.Write "<Script Language=vbscript>" & vbCrLf
	Response.Write "parent.DbDeleteOk" & vbCrLf
Response.Write "</Script>" & vbCrLf
Response.End
%>
