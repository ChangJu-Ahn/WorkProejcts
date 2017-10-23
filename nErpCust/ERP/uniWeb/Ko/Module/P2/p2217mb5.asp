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
'*  1. Module Name			: Production
'*  2. Function Name		: 
'*  3. Program ID			: p2217mb5.asp
'*  4. Program Name			: 
'*  5. Program Desc			: Item by Plant Á¶È¸ 
'*  6. Comproxy List		: PB3S106.cBLkUpItemByPlt
'*  7. Modified date(First)	: 2000/09/28
'*  8. Modified date(Last)	: 2002/12/10
'*  9. Modifier (First)		: Lee Hyun Jae
'* 10. Modifier (Last)		: Jung Yu Kyung
'* 11. Comment				:
'**********************************************************************************************-->
<% Call LoadBasisGlobalInf
   Call LoadInfTB19029B("Q", "P", "NOCOOKIE", "MB")

Call HideStatusWnd

On Error Resume Next

Dim pPB3S106
Dim I3_item_cd, I2_plant_cd
Dim E5_i_material_valuation, E6_b_plant, E7_b_item, E8_b_item_by_plant, iStatusCodeOfPrevNext

' export b_plant
Const P027_E6_plant_cd = 0
Const P027_E6_plant_nm = 1

' export b_item
Const P027_E7_item_cd = 0
Const P027_E7_item_nm = 1
Const P027_E7_spec = 3
Const P027_E7_basic_unit = 10

' export b_item_by_plant

Const P027_E8_tracking_flg = 60

I3_item_cd  = Request("txtItemCd")
I2_plant_cd = Request("txtPlantCd")

Set pPB3S106 = Server.CreateObject("PB3S106.cBLkUpItemByPlt")
If CheckSYSTEMError(Err,True) = True Then
	Response.End
End If

Call pPB3S106.B_LOOK_UP_ITEM_BY_PLANT_SVR(gStrGlobalCollection, "", I2_plant_cd, I3_item_cd,  , , , , _
                     E5_i_material_valuation, E6_b_plant, E7_b_item, E8_b_item_by_plant, iStatusCodeOfPrevNext)
           
If CheckSYSTEMError(Err, True) = True Then
	Set pPB3S106 = Nothing
	Response.End
End If

Set pPB3S106 = Nothing

If Trim(iStatusCodeOfPrevNext) = "900011" Or Trim(iStatusCodeOfPrevNext) = "900012" Then
	Call DisplayMsgBox(iStatusCodeOfPrevNext, VbOKOnly, "", "", I_MKSCRIPT)
End If

%>
<Script Language=vbscript>

	With parent.frm1.vspdData

		.Row = "<%=Request("CurRow")%>"

		If "<%=UCase(E8_b_item_by_plant(P027_E8_tracking_flg))%>" = "N" Then 'TRACKING_FLG
			parent.ggoSpread.SSSetProtected parent.C_TrackingNo, .Row, .Row
			parent.ggoSpread.SSSetProtected parent.C_TrackingNoPopup, .Row, .Row

			Call .SetText(parent.C_TrackingNo,<%=Request("CurRow")%>,"*")
		Else
		    parent.ggoSpread.SpreadUnLock parent.C_TrackingNo, .Row, parent.C_TrackingNoPopup, .Row
			parent.ggoSpread.SSSetRequired parent.C_TrackingNo, .Row, .Row

			Call .SetText(parent.C_TrackingNo,<%=Request("CurRow")%>,"")
		End If

		Call .SetText(parent.C_ItemName,<%=Request("CurRow")%>,"<%=ConvSPChars(E7_b_item(P027_E7_item_nm))%>")
		Call .SetText(parent.C_ItemSpec,<%=Request("CurRow")%>,"<%=ConvSPChars(E7_b_item(P027_E7_spec))%>")
		Call .SetText(parent.C_Unit,<%=Request("CurRow")%>,"<%=ConvSPChars(E7_b_item(P027_E7_basic_unit))%>")

	End With
</Script>