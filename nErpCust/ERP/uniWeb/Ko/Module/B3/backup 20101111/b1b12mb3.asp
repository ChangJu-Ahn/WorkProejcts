<% Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp"  -->

<%
'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : b1b12mb3.asp
'*  4. Program Name         : Lot Control Delete
'*  5. Program Desc         :
'*  6. Component List       : +PB3S113.B_MANAGE_LOT_CONTROL.B_MANAGE_LOT_CONTROL
'*  7. Modified date(First) : 2002/09/03
'*  8. Modified date(Last)  : 2002/09/03
'*  9. Modifier (First)     : 
'* 10. Modifier (Last)      : Lee Hwa Jung
'* 11. Comment              :
'**********************************************************************************************
											'�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 
On Error Resume Next														'��: 
    Err.Clear																		'��: Protect system from crashing
    '[CONVERSION INFORMATION]  View Name : import b_lot_control

	Call LoadBasisGlobalInf() 
    Call HideStatusWnd                                                     '��: Hide Processing message

	Dim pPB3S113
	Dim iCommandSent																	'�� : �Է�/������ ComProxy Dll ��� ���� 
    Dim I3_b_lot_control
    Dim I1_b_item_cd
    Dim I2_b_plant_cd

	
	
    If Request("txtPlantCd") = "" Then												'��: ������ ���� ���� ���Դ��� üũ 
    
		Call DisplayMsgBox("900010", vbOKOnly, "", "", I_MKSCRIPT)					 '��: �����޼����� DBȭ �Ѵ�.           
		Response.End 
	End If
    
    If Request("txtItemCd") = "" Then												'��: ������ ���� ���� ���Դ��� üũ 
		Call DisplayMsgBox("900010", vbOKOnly, "", "", I_MKSCRIPT)						   '��: �����޼����� DBȭ �Ѵ�.           		
		Response.End 
	End If
	
	
    '-----------------------
    'Data manipulate area
    '-----------------------
	I2_b_plant_cd									= UCase(Request("txtPlantCd"))
	I1_b_item_cd									= UCase(Request("txtItemCd"))

	iCommandSent = "DELETE"
	
	Set pPB3S113 = Server.CreateObject("PB3S113.cBMngLotCtl")
	
	If CheckSYSTEMError(Err, True) = True Then
		Response.End 
	End if
	
	Call pPB3S113.B_MANAGE_LOT_CONTROL (gStrGlobalCollection, I1_b_item_cd, I2_b_plant_cd, , iCommandSent)

	If CheckSYSTEMError(Err, True) = True Then
		Set pPB3S113 = Nothing														'��: Unload Component
		Response.End
	End If
	
	Set pPB3S113 = Nothing																'��: Unload Component
	
%>

<Script Language=vbscript>
	With parent																			'��: ȭ�� ó�� ASP �� ��Ī�� 
		.DbDeleteOk
	End With
</Script>
