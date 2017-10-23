<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<%
'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p1201mb3_ko441.asp
'*  4. Program Name         : Routing Component Allocation
'*  5. Program Desc         :
'*  6. Component List       : PP1S506.cPMngCmpReqByRtng
'*  7. Modified date(First) : 2000/03/27
'*  8. Modified date(Last)  : 2008/01/31
'*  9. Modifier (First)     : Im Hyun Soo
'* 10. Modifier (Last)      : HAN cheol
'* 11. Comment              :
'**********************************************************************************************

On Error Resume Next

Call HideStatusWnd
Err.Clear

Call LoadBasisGlobalInf

Dim pPP1S506
Dim I1_plant_cd, I2_item_cd, I3_rout_no, I4_opr_no, iErrorPosition, strSpread

strSpread   = Request("txtSpread")
I1_plant_cd	= Trim(UCase(Request("txtPlantCd")))
I2_item_cd	= Trim(UCase(Request("txtItemCd")))
I3_rout_no	= Trim(UCase(Request("txtRoutNo")))
I4_opr_no	= Trim(UCase(Request("hOprNo")))



If I1_plant_cd = "" Then
	Call DisplayMsgBox("700112", vbOKOnly, "", "", I_MKSCRIPT)            
	Response.End 
ElseIf I3_rout_no = "" Then
	Call DisplayMsgBox("700112", vbOKOnly, "", "", I_MKSCRIPT)             
	Response.End 
ElseIf I4_opr_no = "" Then
	Call DisplayMsgBox("700112", vbOKOnly, "", "", I_MKSCRIPT)       
	Response.End 
ElseIf I2_item_cd = "" Then
	Call DisplayMsgBox("700112", vbOKOnly, "", "", I_MKSCRIPT)              
	Response.End 
End If
	
Set pPP1S506 = Server.CreateObject("PP1S506.cPMngCmpReqByRtng")

If CheckSYSTEMError(Err,True) = True Then
	Response.End
End If

Call pPP1S506.P_MANAGE_COMP_REQ_BY_ROUT(gStrGlobalCollection, strSpread, I1_plant_cd, I2_item_cd, _
										I3_rout_no, I4_opr_no, iErrorPosition)

If CheckSYSTEMError2(Err, True, iErrorPosition & "За", "", "", "", "") = True Then
	Set pPP1S506 = Nothing
	Response.End
End If

Set pPP1S506 = Nothing

Response.Write "<Script Language=VBScript>" & vbCrLf
Response.Write "	parent.DbSaveOk2" & vbCrLf
Response.Write "</Script>" & vbCrLf
Response.End 
%>