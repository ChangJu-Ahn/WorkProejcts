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

Dim pPP2G102
Dim pPP2G150
Dim I2_mrp_parameter
Dim I1_plant_cd, I3_select_char
Dim strSpread
Dim strSpread2
    
Const P206_I2_plant_cd = 0    
Const P206_I2_safe_flg = 1
Const P206_I2_inv_flg = 2
Const P206_I2_idep_flg = 3
Const P206_I2_forward = 4
Const P206_I2_mpsscope = 5

 	Err.Clear
    
    ReDim I2_mrp_parameter(P206_I2_mpsscope)
    
	I1_plant_cd			= UCase(Request("txtPlantCd"))
	
	I2_mrp_parameter(P206_I2_plant_cd)	= UCase(Request("txtPlantCd"))
	I2_mrp_parameter(P206_I2_safe_flg)	= "Y"
	I2_mrp_parameter(P206_I2_inv_flg)	= "M"
	I2_mrp_parameter(P206_I2_idep_flg)	= "S" 
	I2_mrp_parameter(P206_I2_forward)	= UCase(Request("hMrpNo"))
	I2_mrp_parameter(P206_I2_mpsscope)	= "" 

	I3_select_char = "S"
	strSpread = Request("txtSpread")
	strSpread2 = Request("txtSpread2")
	
	Set pPP2G102 = Server.CreateObject("PP2G102.cPCnfmMrpSvr")
		    
	If CheckSYSTEMError(Err,True) = True Then
		Set pPP2G102 = Nothing		
		Response.End
	End If
	
	Call pPP2G102.P_CONFIRM_PLAN_PART_MRP(gStrGlobalCollection, _
								I1_plant_cd, _
								I2_mrp_parameter, _
								I3_select_char, _
								strSpread, _
								strSpread2)

	If CheckSYSTEMError(Err, True) = True Then
		Set pPP2G102 = Nothing
		Response.End
	End If
	
	Set pPP2G102 = Nothing     	
            

Dim ADF
Dim strRetMsg
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0
	
	Redim UNISqlId(0)
	Redim UNIValue(0, 1)
	
	UNISqlId(0) = "185000saa"

	UNIValue(0, 0) = "'" & Ucase(Trim(Request("txtPlantCd"))) & "'"
	UNIValue(0, 1) = "'" & Ucase(Trim(Request("txtPlantCd"))) & "'"
		
	UNILock = DISCONNREAD :	UNIFlag = "1"

    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)

	If Not(rs0.EOF And rs0.BOF) Then
		If rs0("error_qty") > 0 And rs0("error_qty") <> "" Then
			Call DisplayMsgBox("184304", vbInformation, "", "", I_MKSCRIPT)
		End If		
	End If
	
	rs0.Close
	Set rs0 = Nothing	
	Set ADF = Nothing			
	
	Response.Write "<Script Language=VBScript>" & vbCrLf
		Response.Write "parent.MRPConvOk" & vbCrLf
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