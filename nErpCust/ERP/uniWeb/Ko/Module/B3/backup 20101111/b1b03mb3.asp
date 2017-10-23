<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<%Call LoadBasisGlobalInf%>
<%
'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : b1b03mb3.asp
'*  4. Program Name         : Entry Item Group (Delete)	
'*  5. Program Desc         :
'*  6. Component List       : PB3G103.cMngItemGrp
'*  7. Modified date(First) : 2000/09/15
'*  8. Modified date(Last)  : 2002/11/12
'*  9. Modifier (First)     : Hook Eun Sook
'* 10. Modifier (Last)      : Hong Chang Ho
'* 11. Comment              :
'**********************************************************************************************

Call HideStatusWnd															'☜: 모든 작업 완료후 작업진행중 표시창을 Hide

On Error Resume Next
Err.Clear

Dim pPB3G103																	'☆ : 저장용 Component Dll 사용 변수 
Dim I1_b_item_group, iCommandSent

Const P019_I1_item_group_cd = 0

Redim I1_b_item_group(0)

If Request("txtItemGroupCd2") = "" Then												'⊙: 조회를 위한 값이 들어왔는지 체크 
	Call DisplayMsgBox("700114", vbOKOnly, "", "", I_MKSCRIPT)        
	Response.End 
End If
	
'-----------------------
'Data manipulate area
'-----------------------	
												
I1_b_item_group(P019_I1_item_group_cd) = Trim(Request("txtItemGroupCd2"))
iCommandSent = "DELETE"
	
Set pPB3G103 = Server.CreateObject("PB3G103.cBMngItemGrp")
If CheckSYSTEMError(Err,True) = True Then
	Response.End
End If

Call pPB3G103.B_MANAGE_ITEM_GROUP(gStrGlobalCollection, iCommandSent, I1_b_item_group)
If CheckSYSTEMError(Err,True) = True Then
	Set pPB3G103 = Nothing                                                   '☜: Unload Component
	Response.End
End If
	
Set pPB3G103 = Nothing                                                   '☜: Unload Component

Response.Write "<Script Language=VBScript>" & vbCrLf
Response.Write "parent.DbDeleteOk" & vbCrLf
Response.Write "</Script>" & vbCrLf
Response.End																				'☜: Process End
%>