<%@LANGUAGE = VBScript%>
<%'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p1502mb4.asp
'*  4. Program Name         : plant cd 조회 
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

Dim oPB6S101							'☆ : 조회용 ComProxy Dll 사용 변수 
Dim strCode								'☆ : Lookup 용 코드 저장 변수 
Dim strPlantCd
Dim strMode								'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 

Dim E1_B_Plant
Const C_E1_Plant_Cd = 0
Const C_E1_Plant_Nm = 1
Const C_E1_Plant_Cur_Cd = 2

Call HideStatusWnd

On Error Resume Next					
Err.Clear                               
	
	strMode = Request("txtMode")									'☜ : 현재 상태를 받음 
	strPlantCd = Request("txtPlantCd")
	
	If Request("txtPlantCd") = "" Then								'⊙: 조회를 위한 값이 들어왔는지 체크 
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
	' 다음키와 이전키는 존재하지 않을 경우 Blank로 보내는 로직을 수행함.

	Response.Write "<Script Language=vbscript> " & vbCr
	Response.Write "	With parent.frm1 " & vbCr

	Response.Write "		.txtPlantCd.value		= """ & E1_B_Plant(C_E1_Plant_Cd) & """" & vbCr		'☆: Plant Code
	Response.Write "		.txtPlantNm.value		= """ & E1_B_Plant(C_E1_Plant_Nm) & """" & vbCr		'☆: Plant Name		
	Response.Write "		.txtCurCd.value			= """ & E1_B_Plant(C_E1_Plant_Cur_Cd) & """" & vbCr	'☆: Currency Code
			
	Response.Write "		Call parent.CurCdLookUpOk " & vbCr											'☜: 조화가 성공 
	Response.Write "	End With " & vbCr
	Response.Write "</Script> " & vbCr

	Response.End																						'☜: Process End

	'==============================================================================
	' 사용자 정의 서버 함수 
	'==============================================================================
	
	Response.Write "<Script Language=vbscript> " & vbCr
	Response.Write "</Script> " & vbCr
%>