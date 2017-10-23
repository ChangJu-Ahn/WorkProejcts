<%@LANGUAGE = VBScript%>
<%'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p1502mb4.asp
'*  4. Program Name         : plant cd ��ȸ 
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 
'*  8. Modified date(Last)  : 2002/11/15
'*  9. Modifier (First)     : 
'* 10. Modifier (Last)      : RYU SUNG WON
'* 11. Comment              :
'**********************************************************************************************%>

<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<%
Call LoadBasisGlobalInf

Dim oPB6S101							'�� : ��ȸ�� ComProxy Dll ��� ���� 
Dim strCode								'�� : Lookup �� �ڵ� ���� ���� 
Dim strPlantCd
Dim strMode								'��: ���� MyBiz.asp �� ������¸� ��Ÿ�� 

Dim E1_B_Plant
Const C_E1_Plant_Cd = 0
Const C_E1_Plant_Nm = 1
Const C_E1_Plant_Cur_Cd = 2

Call HideStatusWnd

On Error Resume Next					
Err.Clear                               
	
	strMode = Request("txtMode")									'�� : ���� ���¸� ���� 
	strPlantCd = Request("txtPlantCd")
	
	If Request("txtPlantCd") = "" Then								'��: ��ȸ�� ���� ���� ���Դ��� üũ 
		Call DisplayMsgBox("700112", vbOKOnly, "", "", I_MKSCRIPT)        
		Response.End 
	End If
	
    
    Set oPB6S101 = Server.CreateObject("PB6S101.cBLkUpPlt")    

	If CheckSYSTEMError(Err,True) = True Then
	   Response.End 
	End If
	    
    call oPB6S101.B_LOOK_UP_PLANT(gStrGlobalCollection, _
									strPlantCd, _
									, _
									, _
									E1_B_Plant)
	
	If CheckSYSTEMError(Err,True) = True Then
		Set oPB6S101 = Nothing
       
		Response.Write "<Script Language=vbscript> " & vbCr
		Response.Write "	Call parent.CurCdLookUpNotOk " & vbCr
		Response.Write "</Script> " & vbCr
		
		Response.End		
    End If
    
    Set oPB6S101 = Nothing
    
	
	'-----------------------
	'Result data display area
	'----------------------- 
	' ����Ű�� ����Ű�� �������� ���� ��� Blank�� ������ ������ ������.

	Response.Write "<Script Language=vbscript> " & vbCr
	Response.Write "	With parent.frm1 " & vbCr

	Response.Write "		.txtPlantCd.value		= """ & E1_B_Plant(C_E1_Plant_Cd) & """" & vbCr		'��: Plant Code
	Response.Write "		.txtPlantNm.value		= """ & E1_B_Plant(C_E1_Plant_Nm) & """" & vbCr		'��: Plant Name		
	Response.Write "		.txtCurCd.value			= """ & E1_B_Plant(C_E1_Plant_Cur_Cd) & """" & vbCr	'��: Currency Code
			
	Response.Write "		Call parent.CurCdLookUpOk " & vbCr											'��: ��ȭ�� ���� 
	Response.Write "	End With " & vbCr
	Response.Write "</Script> " & vbCr

	Response.End																						'��: Process End

	'==============================================================================
	' ����� ���� ���� �Լ� 
	'==============================================================================
	
	Response.Write "<Script Language=vbscript> " & vbCr
	Response.Write "</Script> " & vbCr
%>