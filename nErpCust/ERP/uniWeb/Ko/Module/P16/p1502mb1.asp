<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<%'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           :  p1502mb1.asp
'*  4. Program Name         :  ManageResourceGroup 조회 
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2000/09/07
'*  8. Modified date(Last)  : 2002/11/15
'*  9. Modifier (First)     : Hong Eun Sook
'* 10. Modifier (Last)      : RYU SUNG WON
'* 11. Comment              :
'**********************************************************************************************
%>

<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->

<Script Language=vbscript> 
	Dim strVar1
	Dim strVar2
	Dim strVar3
	Dim strVar4
	Dim strVar5
	Dim strVar6
	Dim strVar7

	Dim	TempstrResourceGroupCd
	
	TempstrResourceGroupCd	= "<%=Request("txtResourceGroupCd")%>"
	
	'자원그룹명 불러오기 
	Call parent.CommonQueryRs("RESOURCE_GROUP_CD,DESCRIPTION","P_RESOURCE_GROUP","RESOURCE_GROUP_CD =  " & parent.FilterVar(TempstrResourceGroupCd , "''", "S") & "",strVar1,strVar2,strVar3,strVar4,strVar5,strVar6,strVar7)
	strVar1 = Replace(strVar1,Chr(11), "")
	strVar2 = Replace(strVar2,Chr(11), "")
	Parent.frm1.txtResourceGroupCd.Value = strVar1
	Parent.frm1.txtResourceGroupNm.Value = strVar2

</Script>

<%
Call LoadBasisGlobalInf
Call LoadInfTB19029B("I", "*", "NOCOOKIE", "MB")

On Error Resume Next

Call HideStatusWnd
	
Const C_SHEETMAXROWS_D = 100

Dim oPP1G603


Dim StrNextKey
Dim LngMaxRow
Dim LngRow
Dim lgStrPrevKey
Dim strData

Dim TmpBuffer
Dim iTotalStr

Dim strPlantCd
Dim strResourceGroupCd
							
Dim E1_B_Plant									
Dim E2_P_Resource_Group

Const C_E1_plant_cd = 0
Const C_E1_plant_nm = 1
ReDim E1_B_Plant(C_E1_plant_nm)
	
Const C_E2_resource_group_cd = 0
Const C_E2_description = 1
ReDim E2_P_Resource_Group(C_E2_description)
		
strPlantCd			= Request("txtPlantCd")
strResourceGroupCd	= Request("txtResourceGroupCd")
lgStrPrevKey		= Request("lgStrPrevKey")
LngMaxRow			= Request("txtMaxRows")
	
If lgStrPrevKey <> "" then
	strResourceGroupCd = lgStrPrevKey
End If

Set oPP1G603 = Server.CreateObject("PP1G603.cPLkUpRsrcGrpSvr")    

If CheckSYSTEMError(Err,True) = True Then
	Response.End 
End If

Call oPP1G603.P_LIST_RESOURCE_GROUP_SVR(gStrGlobalCollection, _
										strPlantCd, _
										strResourceGroupCd, _
										E1_B_Plant, _
										E2_P_Resource_Group)
									
If CheckSYSTEMError(Err,True) = True Then
    Set oPP1G603 = Nothing
    Response.End 
End If    

If C_SHEETMAXROWS_D < UBound(E2_P_Resource_Group) Then 
	ReDim TmpBuffer(C_SHEETMAXROWS_D - 1)
Else
	ReDim TmpBuffer(UBound(E2_P_Resource_Group))
End If
    
For LngRow = 0 to UBound(E2_P_Resource_Group)
	If LngRow < C_SHEETMAXROWS_D Then
		strData = ""
		strData = strData & chr(11) & ConvSPChars(E2_P_Resource_Group(LngRow,C_E2_resource_group_cd))
		strData = strData & chr(11) & ConvSPChars(E2_P_Resource_Group(LngRow,C_E2_description))
		strData = strData & chr(11) & LngMaxRow + LngRow + 1
		strData = strData & Chr(11) & Chr(12)
		TmpBuffer(LngRow) = strData
	Else
		StrNextKey = E2_P_Resource_Group(LngRow,C_E2_resource_group_cd)
	End if
Next

iTotalStr = Join(TmpBuffer, "")

%>
<Script Language=vbscript>
With Parent
	.ggoSpread.Source = .frm1.vspdData 
	.ggoSpread.SSShowDataByClip "<%=iTotalStr%>"
	.lgStrPrevKey = "<%=ConvSPChars(StrNextKey)%>"
	
	' Request값을 hidden input으로 넘겨줌 
	If .frm1.vspdData.MaxRows < .parent.VisibleRowCnt(.frm1.vspdData, 0) And .lgStrPrevKey <> "" Then
		.DbQuery
	Else
		.frm1.txtPlantNm.Value = "<%=ConvSPChars(E1_B_Plant(C_E1_plant_nm))%>"
		.frm1.htxtPlantCd.Value = "<%=ConvSPChars(strPlantCd)%>"
		.frm1.htxtResourceGroupCd.Value = "<%=ConvSPChars(strNextKey)%>"
		.DbQueryOk
    End if
End with
</Script>
<%
Set oPP1G603 = Nothing
%>
