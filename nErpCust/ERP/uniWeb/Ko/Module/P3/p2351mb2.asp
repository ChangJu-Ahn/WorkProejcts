<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgentVariables.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!--
'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p2346mb2.asp
'*  4. Program Name         : MRP Conversion Partial
'*  5. Program Desc         :
'*  6. Comproxy List        : PP2G102.cPCnfmMrpSvr
'*  7. Modified date(First) : 
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Lee Hyun Jae
'* 10. Modifier (Last)      : Jung Yu Kyung
'* 11. Comment              :
'**********************************************************************************************-->
<% 

Call LoadBasisGlobalInf
Call HideStatusWnd
Call LoadInfTB19029B("I", "*", "NOCOOKIE", "MB")  

On Error Resume Next

Dim pPP3G104
Dim I2_mrp_parameter
Dim I1_plant_cd, I3_select_char
Dim strSpread
Dim strSafeFlg

Redim I2_mrp_parameter(3)

Const P307_I1_plant_cd = 0
Const P307_I1_safe_flg = 1
Const P307_I1_idep_flg = 2
Const P307_I1_entdt = 3

	If Request("rdoSafeInvFlg") = "Y" Then
         strSafeFlg  = "Y"
    Else
    	 strSafeFlg  = "N"
    End If

 	Err.Clear
    
	I1_plant_cd			= UCase(Request("txtPlantCd"))
	I2_mrp_parameter(P307_I1_plant_cd)	= UCase(Request("txtPlantCd"))
	I2_mrp_parameter(P307_I1_safe_flg)	= strSafeFlg
	I2_mrp_parameter(P307_I1_idep_flg)	= "S"
	I2_mrp_parameter(P307_I1_entdt)	= UniConvDateToYYYYMMDD(GetSvrDate,gServerDateFormat,"")

	I3_select_char = "S"
	strSpread = Request("txtSpread")

	Set pPP3G104 = Server.CreateObject("PP3G104.cPCnfmPlanMrpSvr")
		    
	If CheckSYSTEMError(Err,True) = True Then
		Set pPP3G104 = Nothing		
		Response.End
	End If
	
	Call pPP3G104.P_CONFIRM_PLAN_MRP_SRV(gStrGlobalCollection, I1_plant_cd, I2_mrp_parameter, I3_select_char, strSpread)

	If CheckSYSTEMError(Err, True) = True Then
		Set pPP3G104 = Nothing
		Response.End
	End If
	
	Set pPP3G104 = Nothing      	
            
	Response.Write "<Script Language=VBScript>" & vbCrLf
		Response.Write "parent.DBSaveOk" & vbCrLf
	Response.Write "</Script>" & vbCrLf	
%>
<Script Language=vbscript RUNAT=server>

'==============================================================================
' Function : SheetFocus
' Description : 에러발생시 Spread Sheet에 포커스줌 
'==============================================================================
Function SheetFocus(Byval lRow, Byval lCol, Byval iLoc)
	If iLoc = I_INSCRIPT Then
		strHTML = "parent.frm1.vspdData.focus" & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.Row = " & lRow & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.Col = " & lCol & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.Action = 0" & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.SelStart = 0 " & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.SelLength = len(parent.frm1.vspdData.Text) " & vbCrLf
		Response.Write strHTML
	ElseIf iLoc = I_MKSCRIPT Then
		strHTML = "<" & "Script LANGUAGE=VBScript" & ">" & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.focus" & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.Row = " & lRow & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.Col = " & lCol & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.Action = 0" & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.SelStart = 0 " & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.SelLength = len(parent.frm1.vspdData.Text) " & vbCrLf
		strHTML = strHTML & "</" & "Script" & ">" & vbCrLf
		Response.Write strHTML
	End If
End Function
</Script>