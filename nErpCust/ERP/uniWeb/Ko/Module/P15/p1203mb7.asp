<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<%
'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p1203mb7.asp	
'*  4. Program Name         : Look Up Work Center
'*  5. Program Desc         :
'*  6. Component List       : PP1C201.cPLkUpWkCtr
'*  7. Modified date(First) : 2000/03/31
'*  8. Modified date(Last)  : 2002/11/20
'*  9. Modifier (First)     : Mr  Kim
'* 10. Modifier (Last)      : Hong Chang Ho
'* 11. Comment              :
'**********************************************************************************************

On Error Resume Next

Call HideStatusWnd															'☜: 모든 작업 완료후 작업진행중 표시창을 Hide
Err.Clear

Call LoadBasisGlobalInf

Dim pPP1C201																'☆ : 조회용 Component Dll 사용 변수 
Dim strCode																	'☆ : Lookup 용 코드 저장 변수 
Dim strWcNm
Dim strInsideFlg
Dim Row

Dim I1_plant_cd, I2_wc_cd
Dim E4_p_work_center

' E4_p_work_center
Const P191_E4_wc_cd = 0
Const P191_E4_wc_nm = 1
Const P191_E4_inside_flg = 2

If Request("txtPlantCd") = "" Then										'⊙: 조회를 위한 값이 들어왔는지 체크 
	Call DisplayMsgBox("700112", vbOKOnly, "", "", I_MKSCRIPT)           
	Response.End 
End If
If Request("txtWcCd") = "" Then										'⊙: 조회를 위한 값이 들어왔는지 체크 
	Call DisplayMsgBox("700112", vbOKOnly, "", "", I_MKSCRIPT)             
	Response.End 	
End If
	
'-----------------------
'Data manipulate  area(import view match)
'-----------------------
I1_plant_cd	= Trim(Request("txtPlantCd"))
I2_wc_cd = Trim(Request("txtWcCd"))
    
'-----------------------
'Com action area
'-----------------------
Set pPP1C201 = Server.CreateObject("PP1C201.cPLkUpWkCtr")    

If CheckSYSTEMError(Err,True) = True Then
	Response.End
End If

Call pPP1C201.P_LOOK_UP_WORK_CENTER(gStrGlobalCollection, I1_plant_cd, I2_wc_cd, , , , E4_p_work_center)

If CheckSYSTEMError(Err, True) = True Then
	Set pPP1C201 = Nothing															'☜: Unload Component
	Response.Write "<Script Language=vbscript>" & VBCrLf
	Response.Write "Call parent.LookUpWcNotOk(""" & Request("Row") & """)" & VBCrLf
	Response.Write "</Script>"
	Response.End
End If

Set pPP1C201 = Nothing															'☜: Unload Component
    
strWcNm = E4_p_work_center(P191_E4_wc_nm)
strInsideFlg = E4_p_work_center(P191_E4_inside_flg)
Row = Request("Row")

Response.Write "<Script Language = VBScript>" & VBCrLf
Response.Write "Call parent.LookUpWcOk(""" & ConvSPChars(strWcNm) & """,""" & strInsideFlg & """,""" & Row & """)" & VBCrLf
Response.Write "</Script>"
Response.End																	'☜: Process End
%>