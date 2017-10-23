<%@ LANGUAGE = VBSCript%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<%
'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p1711mb6.asp
'*  4. Program Name         : 설계BOM Header Delete
'*  5. Program Desc         :
'*  6. Component List       : PP1S401.cPMngBomHdr
'*  7. Modified date(First) : 2000/05/2
'*  8. Modified date(Last)  : 2002/11/19
'*  9. Modifier (First)     : Im Hyun Soo
'* 10. Modifier (Last)      : Hong Chang Ho
'**********************************************************************************************

Call HideStatusWnd															

On Error Resume Next
Err.Clear																

Call LoadBasisGlobalInf

Dim pPP1S401																	'☆ : 입력/수정용 ComProxy Dll 사용 변수 
Dim iCommandSent, I1_plant_cd, I2_item_cd, I3_p_bom_header		

Const P131_I3_bom_no = 0															
	
    
'-----------------------
'Key Value Check---------------------------
'-----------------------
    
If Request("txtPlantCd") = "" Then										
	Call DisplayMsgBox("700114", vbOKOnly, "", "", I_MKSCRIPT)
	Response.End 
End If
	
If Request("txtItemCd") = "" Then										
	Call DisplayMsgBox("700114", vbOKOnly, "", "", I_MKSCRIPT)
	Response.End 
End If
	
If Request("txtBomNo") = "" Then										
	Call DisplayMsgBox("700114", vbOKOnly, "", "", I_MKSCRIPT)
	Response.End 
End If
	
'-----------------------
'Data manipulate area
'-----------------------												
Redim I3_p_bom_header(P131_I3_bom_no)

iCommandSent				= "DELETE"
I1_plant_cd		= Trim(Request("txtPlantCd"))
I2_item_cd		= Trim(Request("txtItemCd"))
I3_p_bom_header(P131_I3_bom_no)	= Trim(Request("txtBomNo"))
	
'-----------------------
'Server Create Object
'-----------------------
Set pPP1S401 = Server.CreateObject("PP1S401.cPMngBomHdr")
	    
If CheckSYSTEMError(Err,True) = True Then
	Set pPP1S401 = Nothing
	Response.End
End If

Call pPP1S401.P_MANAGE_BOM_HEADER(gStrGlobalCollection, iCommandSent, I1_plant_cd, I2_item_cd, I3_p_bom_header)
	
If CheckSYSTEMError(Err, True) = True Then
	Set pPP1S401 = Nothing															'☜: Unload Component
	Response.End
End If

Set pPP1S401 = Nothing   

Response.Write "<Script Language = VBScript>" & vbCrLf
		Response.Write "parent.DbDeleteOk" & vbCrLf
Response.Write "</Script>" & vbCrLf
Response.End
%>
