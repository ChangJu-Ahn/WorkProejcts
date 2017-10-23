<% Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<%
'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           :  b1b03mb1.asp
'*  4. Program Name         :  품목그룹 조회 
'*  5. Program Desc         :
'*  6. Component List       : PB3S101.B_LOOK_UP_ITEM_GROUP.B_LOOK_UP_ITEM_GROUP_SVR
'*  7. Modified date(First) : 2000/09/15
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Hong Eun Sook
'* 10. Modifier (Last)      : Lee Hwa Jung
'* 11. Comment              :
'**********************************************************************************************
   
' E1_b_item_group
Const P044_E1_item_group_cd = 0
Const P044_E1_item_group_nm = 1
Const P044_E1_leaf_flg = 2
Const P044_E1_valid_from_dt = 3
Const P044_E1_valid_to_dt = 4
Const P044_E1_level = 5

' E2_b_item_group
Const P044_E2_item_group_cd = 0
Const P044_E2_item_group_nm = 1
Const P044_E2_leaf_flg = 2
Const P044_E2_valid_from_dt = 3
Const P044_E2_valid_to_dt = 4
Const P044_E2_level = 5

On Error Resume Next														'☜: 

Call HideStatusWnd															'☜: 모든 작업 완료후 작업진행중 표시창을 Hide
Call LoadBasisGlobalInf() 

Dim pPB3S101																	'☆ : 조회용 ComProxy Dll 사용 변수 
Dim strCode																	'☆ : Lookup 용 코드 저장 변수 
Dim intLevel
Dim iLevelCnt

Dim i
Dim strTemp

Dim strMode																	'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 

Dim I1_item_group_cd
Dim E1_b_item_group
Dim E2_b_item_group 

Redim I1_item_group(1)
strMode = Request("txtMode")												'☜ : 현재 상태를 받음 

Err.Clear                                                               '☜: Protect system from crashing
    
If Request("txtItemGroupCd") = "" Then										'⊙: 조회를 위한 값이 들어왔는지 체크 
	Call DisplayMsgBox("700112", vbOKOnly, "", "", I_MKSCRIPT)       
	Response.End 
End If
    
I1_item_group_cd = Trim(Request("txtItemGroupCd"))
Set pPB3S101 = Server.CreateObject("PB3S101.cBLkUpItemGrp")

If CheckSYSTEMError(Err, True) = True Then
	Response.Write "<Script Language = VBScript>" & vbCrLf
	Response.Write "parent.frm1.txtItemGroupNm.value = """"" & vbCrLf
	Response.Write "</Script>" & vbCrLf
	Response.End 
End if
    
'-----------------------
'Data manipulate  area(import view match)
'-----------------------
	
Call pPB3S101.B_LOOK_UP_ITEM_GROUP (gStrGlobalCollection, I1_item_group_cd, E1_b_item_group, E2_b_item_group )
'-----------------------
'Com action result check area(OS,internal)
'-----------------------
If CheckSYSTEMError(Err, True) = True Then
	Set pPB3S101 = Nothing								
	Response.End
End If

'------------------------------
' Level Setting
'------------------------------
If Trim(E1_b_item_group(P044_E1_level)) <> "0" Then
	For i = 1 To CInt(E1_b_item_group(P044_E1_level))
		strTemp = strTemp & "."			
	Next

	intLevel = strTemp & E1_b_item_group(P044_E1_level)
Else
	intLevel = E1_b_item_group(P044_E1_level)
End If 

'-----------------------
'Result data display area
'----------------------- 

Response.Write "<Script Language = VBScript>" & vbCrLf
	Response.Write "With parent.frm1" & vbCrLf
		
		Response.Write ".txtLevel1.value = """ & intLevel & """" & vbCrLf
		Response.Write ".txtItemGroupCd1.value = """ & ConvSPChars(UCase(E1_b_item_group(P044_E1_item_group_cd))) & """" & vbCrLf
		Response.Write ".txtItemGroupNm1.value = """ & ConvSPChars(E1_b_item_group(P044_E1_item_group_nm)) & """" & vbCrLf
		Response.Write ".txtUpperItemGroupCd.value = """ & ConvSPChars(E2_b_item_group(P044_E2_item_group_cd)) & """" & vbCrLf
		Response.Write ".txtUpperItemGroupNm.value = """ & ConvSPChars(E2_b_item_group(P044_E2_item_group_nm)) & """" & vbCrLf
				
		If UCase(Trim(E1_b_item_group(P044_E1_leaf_flg))) = "Y" Then
			Response.Write ".rdoLowItemGroupFlg1.checked = True" & vbCrLf
		Else
			Response.Write ".rdoLowItemGroupFlg2.checked = True" & vbCrLf
		End If
		 		
		Response.Write ".txtValidFromDt1.value = """ & UniDateClientFormat(E1_b_item_group(P044_E1_valid_from_dt)) & """" & vbCrLf
		Response.Write ".txtValidToDt1.value = """ & UniDateClientFormat(E1_b_item_group(P044_E1_valid_to_dt)) & """" & vbCrLf
		
		Response.Write "parent.LookUpItemGroupOk" & vbCrLf						

	Response.Write "End With" & vbCrLf
Response.Write "</Script>" & vbCrLf

Response.End										
%>
