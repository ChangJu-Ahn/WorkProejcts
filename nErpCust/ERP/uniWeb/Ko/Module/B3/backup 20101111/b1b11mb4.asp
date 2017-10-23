<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<%Call LoadBasisGlobalInf%>
<%
'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : b1b11mb4.asp
'*  4. Program Name         : Entry Item By Plant (Look Up Item)
'*  5. Program Desc         :
'*  6. Component List       : PB3C104.cBLkUpItem
'*  7. Modified date(First) : 2000/03/25
'*  8. Modified date(Last)  : 2002/11/14
'*  9. Modifier (First)     : Im Hyun Soo
'* 10. Modifier (Last)      : Hong Chang Ho
'* 11. Comment              :
'**********************************************************************************************

On Error Resume Next														'☜: 
Err.Clear

Call HideStatusWnd															'☜: 모든 작업 완료후 작업진행중 표시창을 Hide

Dim pPB3C104																'☆ : 조회용 Component Dll 사용 변수 

Dim I2_item_cd
Dim E5_b_item

' E5_b_item
Const P020_E5_item_cd = 0
Const P020_E5_item_nm = 1
Const P020_E5_basic_unit = 4
Const P020_E5_item_acct = 5
Const P020_E5_phantom_flg = 7

I2_Item_cd  = UCase(Trim(Request("txtItemCd1")))

'-----------------------
'Com action area
'-----------------------
Set pPB3C104 = Server.CreateObject("PB3C104.cBLkUpItem")    

If CheckSYSTEMError(Err,True) = True Then
	Response.End
End If

Call pPB3C104.B_LOOK_UP_ITEM(gStrGlobalCollection, I2_item_cd, , , , , E5_b_item)

If CheckSYSTEMError(Err, True) = True Then
	Set pPB3C104 = Nothing															'☜: Unload Component
	Response.Write "<Script Language=vbscript>" & vbCrLf
		Response.Write "Call parent.LookUpItemNotOk()" & vbCrLf
	Response.Write "</Script>" & vbCrLf
	Response.End																				'☜: Process End
End If

Set pPB3C104 = Nothing															'☜: Unload Component

'-----------------------
'Result data display area
'----------------------- 
' 다음키와 이전키는 존재하지 않을 경우 Blank로 보내는 로직을 수행함.

Response.Write "<Script Language=VBScript>" & vbCrLf
	Response.Write "With parent.frm1" & vbCrLf
		Response.Write ".txtItemCd1.value = """ & ConvSPChars(E5_b_item(P020_E5_item_cd)) & """" & vbCrLf
		Response.Write ".txtItemNm1.value = """ & ConvSPChars(E5_b_item(P020_E5_item_nm)) & """" & vbCrLf
		Response.Write ".cboAccount.value = """ & E5_b_item(P020_E5_item_acct) & """" & vbCrLf
		Response.Write ".txtPhantomFlg.value = """ & E5_b_item(P020_E5_phantom_flg) & """" & vbCrLf
		Response.Write ".txtBasicUnit.value = """ & ConvSPChars(E5_b_item(P020_E5_basic_unit)) & """" & vbCrLf
		Response.Write "Call parent.LookUpItemOk()" & vbCrLf
	Response.Write "End With" & vbCrLf
Response.Write "</Script>" & vbCrLf
Response.End																	'☜: Process End
%>
