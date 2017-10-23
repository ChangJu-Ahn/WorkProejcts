<%@LANGUAGE = VBScript%>
<%'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p1601mb1.asp
'*  4. Program Name         : Copy Item By Plant
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 
'*  8. Modified date(Last)  : 2002/11/15
'*  9. Modifier (First)     : 
'* 10. Modifier (Last)      : RYU SUNG WON
'* 11. Comment              :
'**********************************************************************************************%>

<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%
Call LoadBasisGlobalInf
Call LoadInfTB19029B("I", "*", "NOCOOKIE", "MB")
													'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 
Dim ADF														'ActiveX Data Factory 지정 변수선언 
Dim strRetMsg												'Record Set Return Message 변수선언 
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3 ,rs4 				'DBAgent Parameter 선언 
Dim strQryMode

Const C_SHEETMAXROWS_D = 100

Dim strMode											'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 

'=======================================================================================================
'	아래 선언되어 있는 변수들은 COOL:Gen 의 Record Return Count 의 제한에 따른 것이다.
'	따라서, ADO를 사용할 경우 그와같은 문제성이 없기 때문에 아래의 변수들은 사용하지 않지만 추후 
'	uniERP2000 에서 한번에 조회되는 Record Count 의 수를 30으로 제한하고 있는 만큼 그에 따른 
'	표준은 동시에 추가될 예정이므로 변수삭제는 하지 않고 그대로 놔둔다.
'=======================================================================================================
Dim lgStrPrevKey
Dim i

Call HideStatusWnd

Dim strYear, strMonth, strDay, StartDate

Call ExtractDateFrom(GetSvrDate, gServerDateFormat, gServerDateType, strYear, strMonth, strDay)
StartDate = UniConvYYYYMMDDToDate(gDateFormat, strYear, strMonth, strDay)

strMode = Request("txtMode")						'☜ : 현재 상태를 받음 
strQryMode = Request("lgIntFlgMode")

On Error Resume Next
Err.Clear

Dim strPlantCd
Dim strItemCd
Dim strItemCd1
Dim strSumItemClass
Dim strItemAccount
Dim strItemGroup
Dim strStartDt
Dim strEndDt
Dim strPhantomFlg
Dim strItemGroupCd
Dim oPB3S101

	strItemGroupCd = Trim(Request("txtItemGroupCd"))
	'======================================================================================================
	'	품목그룹 체크 하는 부분 
	'======================================================================================================
	If strItemGroupCd <> "" Then
		Set oPB3S101 = Server.CreateObject("PB3S101.cBLkUpItemGrp")

		If CheckSYSTEMError(Err,True) = True Then
			Response.End
		End If

		Call oPB3S101.B_LOOK_UP_ITEM_GROUP(gStrGlobalCollection, strItemGroupCd)

		If CheckSYSTEMError(Err, True) = True Then
			Set oPB3S101 = Nothing
			Response.Write "<script languague = vbscript>" & vbCr
			Response.Write "	parent.frm1.txtHighItemGroupNm.value = """"" & vbCr '
			Response.Write "	parent.frm1.txtHighItemGroupCd.Focus()" & vbCr '
			Response.Write "</script>" & vbCr
			Response.End
		End If

		Set oPB3S101 = Nothing															'☜: Unload Component
	End If


	lgStrPrevKey = UCase(Trim(Request("lgStrPrevKey")))	

	If lgStrPrevKey = "" Then
	
		'======================================================================================================
		'	품목이름 처리해주는 부분 
		'======================================================================================================
		Redim UNISqlId(4)
		Redim UNIValue(4, 1)
		
		UNISqlId(0) = "122700sab"	'plant_nm
		UNISqlId(1) = "122600sac"	'item_nm
		UNISqlId(2) = "127400saa"	'item_group_nm
		UNISqlId(3) = "122600sac"	'item_nm
		UNISqlId(4) = "122600SAG"	'ProcType

		strItemCd = FilterVar(Request("txtItemCd"),"''","S")
		strItemGroup = FilterVar(strItemGroupCd ,"''","S")
		strPlantCd = FilterVar(Request("txtPlantCd"),"''","S")
		strItemCd1 = FilterVar(Request("txtItemCd1"),"''","S")

		
		UNIValue(0, 0) = strPlantCd
		UNIValue(1, 0) = strItemCd
		UNIValue(2, 0) = strItemGroup
		UNIValue(3, 0) = strItemCd1
		UNIValue(4, 0) = strPlantCd
		UNIValue(4, 1) = strItemCd1

		UNILock = DISCONNREAD :	UNIFlag = "1"

	    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
	    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3, rs4)
		
		'-----------------------
	    'Com action result check area(OS,internal)
	    '-----------------------
	    If Err.Number <> 0 Then
			Set rs0 = Nothing
			Set rs1 = Nothing
			Set rs2 = Nothing
			Set rs3 = Nothing
			Set rs4 = Nothing
			Set ADF = Nothing												'☜: ComProxy Unload
			Call ServerMesgBox(Err.description, vbCritical, I_MKSCRIPT)						'⊙:		
			Response.End														'☜: 비지니스 로직 처리를 종료함 
		End If
		
		If rs0.EOF And rs0.BOF Then												'☜: 공장코드 체크 
			Call DisplayMsgBox("125000", vbOKOnly, "", "", I_MKSCRIPT)
			Response.Write "<script languague = vbscript>" & vbCr
			Response.Write "	parent.frm1.txtPlantNm.value = """"" & vbCrLf
			Response.Write "	parent.frm1.txtPlantNm1.value = """"" & vbCrLf
			Response.write "	parent.frm1.txtPlantCd.Focus()" & vbCr
			Response.Write "</script>" & vbCr
			Response.End	
		Else
			Response.Write "<script languague = vbscript>" & vbCr
			Response.Write "	parent.frm1.txtPlantNm.value = """ & ConvSPChars(rs0(0)) & """" & vbCrLf
			Response.Write "</script>" & vbCr
		End If
		
		Response.Write "<Script Language = VBScript>" & vbCrLf
		
		If rs1.EOF And rs1.BOF Then
			Response.Write "parent.frm1.txtItemNm.value = """"" & vbCrLf
		Else
			Response.Write "parent.frm1.txtItemNm.value = """ & ConvSPChars(rs1(0)) & """" & vbCrLf
		End If
	
		If rs2.EOF And rs2.BOF Then
			Response.Write "parent.frm1.txtHighItemGroupNm.value = """"" & vbCrLf
		Else
			Response.Write "parent.frm1.txtHighItemGroupNm.value = """ & ConvSPChars(rs2(0)) & """" & vbCrLf
		End If
	
		If rs3.EOF And rs3.BOF Then
			Response.Write "parent.frm1.txtItemNm1.value = """"" & vbCrLf
			Response.Write "parent.frm1.txtItemSpec1.value = """"" & vbCrLf
			Response.Write "parent.frm1.txtItemProcType1.value = """"" & vbCrLf
		Else
			Response.Write "parent.frm1.txtItemNm1.value = """ & ConvSPChars(rs3(0)) & """" & vbCrLf
			Response.Write "parent.frm1.txtItemSpec1.value = """ & ConvSPChars(rs3(1)) & """" & vbCrLf
			Response.Write "parent.frm1.txtItemProcType1.value = """ & ConvSPChars(rs4(1)) & """" & vbCrLf
			Response.Write "parent.frm1.htxtItemProcType1.value = """ & ConvSPChars(rs4(0)) & """" & vbCrLf
		End If
						
		If rs4.EOF And rs4.BOF Then
			Response.Write "parent.frm1.txtItemProcType1.value = """"" & vbCrLf
		Else
			Response.Write "parent.frm1.txtItemProcType1.value = """ & ConvSPChars(rs4(1)) & """" & vbCrLf
			Response.Write "parent.frm1.htxtItemProcType1.value = """ & ConvSPChars(rs4(0)) & """" & vbCrLf
		End If
		Response.Write "</Script>" & vbCrLf					

		rs0.Close
		rs1.Close
		rs2.Close
		rs3.Close
		rs4.Close
			
		Set rs0 = Nothing
		Set rs1 = Nothing
		Set rs2 = Nothing
		Set rs3 = Nothing
		Set rs4 = Nothing
		Set ADF = Nothing				'☜: ActiveX Data Factory Object Nothing
	End If
	
	
	Err.Clear
	'=======================================================================================================
	'	만약, 선언한 변수가 배열이라면 아래와같은 Fix 된 배열로 Redim 을 해서 넘겨줘야 한다.
	'=======================================================================================================

	Redim UNISqlId(0)
	Redim UNIValue(0, 8)
	
	UNISqlId(0) = "P1601MB1"
	
	IF Request("txtPlantCd") = "" Then
	   strPlantCd = "|"
	ELSE
	   strPlantCd = FilterVar(Request("txtPlantCd"),"''","S")
	END IF
	
	IF Request("txtItemCd") = "" Then
	   strItemCd = "|"
	ELSE
	   strItemCd = FilterVar(Request("txtItemCd"),"''","S")
	END IF
		
	IF Request("cboItemAccount") = "" Then
	   strItemAccount = "|"
	ELSE
	   strItemAccount = FilterVar(Request("cboItemAccount") ,"''","S")
	END IF
	
	IF Request("txtItemGroupCd") = "" Then
	   strItemGroup = "|"
	ELSE
	   strItemGroup = FilterVar(Request("txtItemGroupCd"),"''","S")
	END IF
	
	IF Request("cboItemClass") = "" Then
	   strSumItemClass = "|"
	ELSE
	   strSumItemClass = FilterVar(Request("cboItemClass"),"''","S") 
	END IF
	
	If Request("rdoPhantomFlg") = "A" Then
		strPhantomFlg = "|"	
	Else
		strPhantomFlg = FilterVar(Request("rdoPhantomFlg"),"''","S")
	End If

	IF strItemGroupCd = "" Then
		strItemGroup = "|"
	Else
		strItemGroup = "A.ITEM_GROUP_CD in (select item_group_cd from ufn_P_ListItemGrp( " & FilterVar(strItemGroupCd, "''", "S") & ")) "
	End IF
		
	strStartDt = "|"
	strEndDt = "|"
	
	UNIValue(0, 0) = "^"
	UNIValue(0, 1) = strPlantCd
	
	Select Case strQryMode
		Case CStr(OPMD_CMODE)
			UNIValue(0, 2) = strItemCd
		Case CStr(OPMD_UMODE) 
			UNIValue(0, 2) = FilterVar(lgStrPrevKey,"''","S")
	End Select

	UNIValue(0, 3) = strSumItemClass
	UNIValue(0, 4) = strItemAccount
	UNIValue(0, 5) = strStartDt
	UNIValue(0, 6) = strEndDt
	UNIValue(0, 7) = strPhantomFlg
	UNIValue(0, 8) = strItemGroup
	
	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")

    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)

    If Err.Number <> 0 Then
		Set rs0 = Nothing					
		Set ADF = Nothing
		Call ServerMesgBox(Err.description, vbCritical, I_MKSCRIPT)						'⊙:		
		Response.End														'☜: 비지니스 로직 처리를 종료함 
	End If
  
    If rs0.EOF And rs0.BOF Then
		Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)	'⊙: DB 에러코드, 메세지타입, %처리, 스크립트유형 
		Set rs0 = Nothing					
		Set ADF = Nothing
		Response.End													'☜: 비지니스 로직 처리를 종료함 
	End If
%>

<Script Language=vbscript>
Dim LngMaxRow
Dim strData
Dim TmpBuffer
Dim iTotalStr

With parent																	'☜: 화면 처리 ASP 를 지칭함 
	LngMaxRow = .frm1.vspdData.MaxRows										'Save previous Maxrow
	
<%
		If C_SHEETMAXROWS_D < rs0.RecordCount Then 
%>
			ReDim TmpBuffer(<%=C_SHEETMAXROWS_D - 1%>)
<%
		Else
%>			
			ReDim TmpBuffer(<%=rs0.RecordCount - 1%>)
<%
		End If
		
		For i=0 to rs0.RecordCount-1 
		
			If i < C_SHEETMAXROWS_D Then
%>				
				strData = ""
				strData = strData & Chr(11) & ""													'1:C_Select
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("item_cd"))%>"					'2:C_Item
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("item_nm"))%>"					'3:C_ItmNm
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("spec"))%>"						'19:C_ItmSpec
				strData = strData & Chr(11) & "S"													'5:C_PrcCtrlInd
				strData = strData & Chr(11) & ""													'6:C_PrcCtrlIndNm
				strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany("",ggUnitCost.DecPoint,ggUnitCost.RndPolicy, ggUnitCost.RndUnit,0)%>"	'7:C_UnitPrice
				strData = strData & Chr(11) & "<%=startdate%>"					'29:C_ValidFromDt
				strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0("VALID_TO_DT"))%>"		'30:C_ValidToDt
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("formal_nm"))%>"					'4:C_ItmFormalNm
				strData = strData & Chr(11) & ""													'9:C_ItmAcc
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("basic_unit"))%>"					'6:C_Unit
				strData = strData & Chr(11) & ""													'7:C_UnitPopup
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("item_group_cd"))%>"				'8:C_ItmGroupCd
				strData = strData & Chr(11) & ""													'9:C_ItmGroupPopup
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("item_group_nm"))%>"				'10:C_ItmGroupNm				
				strData = strData & Chr(11) & "<%=rs0("phantom_flg")%>"								'11:C_Phantom
				strData = strData & Chr(11) & "<%=rs0("BLANKET_PUR_FLG")%>"							'12:C_BlanketPur
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("base_item_cd"))%>"				'13:C_BaseItm
				strData = strData & Chr(11) & ""													'14:C_BaseItmPopup
				strData = strData & Chr(11) & ""													'15:C_BaseItmNm
				strData = strData & Chr(11) & ""													'16:C_SumItmClass
				strData = strData & Chr(11) & "<%=rs0("VALID_FLG")%>"								'17:C_DefaultFlg
				strData = strData & Chr(11) & "<%=rs0("item_image_flg")%>"							'18:C_PicFlg
				strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("unit_weight"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit,0)%>" 	'20: C_UnitWeight
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("unit_of_weight"))%>"				'21:C_UnitOfWeight
				strData = strData & Chr(11) & ""													'22:C_WeightUnitPopup	
				
				strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("gross_weight"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit,0)%>" 	'20: C_UnitWeight
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("gross_unit"))%>"					'21:C_UnitOfWeight
				strData = strData & Chr(11) & ""													'22:C_WeightUnitPopup	
				strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("cbm"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit,0)%>" 	'20: C_UnitWeight
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("cbm_description"))%>"				'21:C_UnitOfWeight
				
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("draw_no"))%>"					'23:C_DrawNo
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("hs_cd"))%>"						'24:C_HsCd
				strData = strData & Chr(11) & ""													'25:C_HsCdPopup
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("hs_unit"))%>"					'26:C_HsUnit
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("vat_type"))%>"					'27:vat_type
				strData = strData & Chr(11) & ""													'25:C_HsCdPopup
				strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("vat_rate"),ggExchRate.DecPoint, ggExchRate.RndPolicy, ggExchRate.RndUnit,0)%>"		'28:vat_rate
				strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0("VALID_FROM_DT"))%>"		'29:C_ValidFromDt
				strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0("VALID_TO_DT"))%>"		'30:C_ValidToDt
				strData = strData & Chr(11) & "<%=rs0("item_class")%>"								'29:C_HdnSumItmClass
				strData = strData & Chr(11) & "<%=rs0("item_acct")%>"								'30:C_HdnItmAcc
				strData = strData & Chr(11) & LngMaxRow + "<%=i%>"									'31:
				strData = strData & Chr(11) & Chr(12)
				
				TmpBuffer(<%=i%>) = strData
<%			
				rs0.MoveNext
			End If
		Next
%>
		iTotalStr = Join(TmpBuffer, "")
		.ggoSpread.Source = .frm1.vspdData
		.ggoSpread.SSShowDataByClip iTotalStr
		
		.lgStrPrevKey = "<%=Trim(rs0("item_cd"))%>"
	
		.frm1.hItemCd.value			= "<%=ConvSPChars(Request("txtItemCd"))%>"
		.frm1.hPlantCd.value		= "<%=ConvSPChars(Request("txtPlantCd"))%>"
		.frm1.hItemAccount.value	= "<%=Request("cboItemAccount")%>"
		.frm1.hItemClass.value		= "<%=Request("CboItemClass")%>"	
		.frm1.hItemGroupCd.value	= "<%=ConvSPChars(Request("txtItemGroupCd"))%>"
		.frm1.hPhantomFlg.value		= "<%=Request("rdoPhantomFlg")%>"	
<%			
		rs0.Close
		Set rs0 = Nothing
%>
	.DbQueryOk(LngMaxRow + 1)
End With	
</Script>	
<%
Set ADF = Nothing												'☜: ActiveX Data Factory Object Nothing
'++++++++++++++++++++++++++++++++++++++++++++++++++++++
%>




