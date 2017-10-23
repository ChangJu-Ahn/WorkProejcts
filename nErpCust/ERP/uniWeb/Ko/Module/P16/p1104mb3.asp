<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<%
'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p1104mb3.asp
'*  4. Program Name         : Entry Shift(Delete)
'*  5. Program Desc         :
'*  6. Component List       : PP1G602.cPMngShift
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

Dim pPP1G602																	'�� : �Է�/������ Component Dll ��� ���� 
Dim I1_plant_cd, I2_p_shift_header, iCommandSent, iErrorPosition
Dim strPlantCd

Const P159_I2_shift_cd = 0

Redim I2_p_shift_header(P159_I2_shift_cd)
	
strPlantCd = Request("txtPlantCd")
    
If strPlantCd = "" Then												'��: ��ȸ�� ���� ���� ���Դ��� üũ 
	Call DisplayMsgBox("700114", vbOKOnly, "", "", I_MKSCRIPT)
	Response.End 
ElseIf Request("txtShiftCd") = "" Then												'��: ��ȸ�� ���� ���� ���Դ��� üũ 
	Call DisplayMsgBox("700114", vbOKOnly, "", "", I_MKSCRIPT)
	Response.End 
End If

'Data manipulate area
I1_plant_cd	= strPlantCd
I2_p_shift_header(P159_I2_shift_cd)	= Request("txtShiftCd")
	
iCommandSent = "DELETE"

'Com action result check area(OS,internal)
Set pPP1G602 = Server.CreateObject("PP1G602.cPMngShift")

If CheckSYSTEMError(Err,True) = True Then
	Response.End
End If

Call pPP1G602.P_MANAGE_SHIFT(gStrGlobalCollection, iCommandSent, , I1_plant_cd, I2_p_shift_header, iErrorPosition)

If CheckSYSTEMError2(Err, True, iErrorPosition & "��", "", "", "", "") = True Then
	Set pPP1G602 = Nothing															'��: Unload Component
	Response.End
End If

Set pPP1G602 = Nothing																'��: Unload Component

Response.Write "<Script Language=vbscript>" & vbCrLf
	Response.Write "parent.DbDeleteOk" & vbCrLf
Response.Write "</Script>" & vbCrLf
Response.End
%>