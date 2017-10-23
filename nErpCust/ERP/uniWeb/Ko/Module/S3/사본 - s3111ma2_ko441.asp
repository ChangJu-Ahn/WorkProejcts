<%@ LANGUAGE="VBSCRIPT" %>
<%
'**********************************************************************************************
'*  1. Module Name          : 영업 
'*  2. Function Name        : 
'*  3. Program ID           : s3111ma2.asp	
'*  4. Program Name         : 단가확정 
'*  5. Program Desc         : 단가확정 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2003/05/28
'*  8. Modified date(Last)  : 2005/05/26
'*  9. Modifier (First)     : Cho in kuk
'* 10. Modifier (Last)      : HJO
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************
%>
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMA_KO441.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit                                                             '☜: indicates that All variables must be declared in advance
<!-- #Include file="../../inc/lgvariables.inc" -->	

Dim IsOpenPop						' Popup
Dim lgBlnOpenedFlag
Dim	lgBlnSoldToPartyChg
Dim lgBlnSalesGrpChg
Dim	lgBlnPlantChg
Dim lgBlnFlgConChg

Dim lgBlnPriceRule         'price rule:T/N

Dim iDBSYSDate
Dim EndDate, StartDate

iDBSYSDate = "<%=GetSvrDate%>"
'------ ☆: 초기화면에 뿌려지는 마지막 날짜 ------
EndDate = UniConvDateAToB(iDBSYSDate, parent.gServerDateFormat, parent.gDateFormat)

'------ ☆: 초기화면에 뿌려지는 시작 날짜 ------
StartDate = UNIDateAdd("m", -1, EndDate, parent.gDateFormat)



Const BIZ_PGM_ID = "s3111mb2.asp"												'☆: Head Query 비지니스 로직 ASP명 
Const BIZ_POP_PGM_ID = "s3111mp2.asp"										'☆: PopUp Query 비지니스 로직 ASP명 

'☆: Spread Sheet의 Column별 상수 
Dim C_PriceFlag1				'재확정여부 
Dim C_PriceFlag				'단가여부 
Dim C_SoNo					'수주번호 
Dim C_SoSeq					'수주순번 
Dim C_ItemCode				'품목 
Dim C_ItemName				'품목명 
Dim C_PriceFlagY			'가단가 
Dim C_PriceFlagN			'진단가 
Dim C_SoldToParty			'주문처 
Dim C_SoldToPartyNm			'주문처명 
Dim C_SoDt					'수주일 
Dim C_DealType				'거래유형 
Dim C_DealTypeNm			'거래유형명 
Dim C_PayTerms				'결제방법 
Dim C_PayTermsNm			'결제방법명 
Dim C_BillToParty			'발행처 
Dim C_BillToPartyNm			'발행처명 
Dim C_NetAmt				'수주금액 
Dim C_Currency				'화폐 

'2002-12-22  추가 
Dim C_SalesGrp				'영업그룹	
Dim C_SalesGrpNm			'영업그룹명	
Dim C_Plant					'공장 
Dim C_PlantNm				'공장코드 
Dim C_ItemSpec				'규격 
Dim C_SoUnit				'수주단위 

'2006.02.28 추가 
Dim C_TrackingNo			'Tracking No

Const C_PopSoldToParty	= 1
Const C_PopSalesGrp		= 2
Const C_PopPlant		= 3

'============================================================================================================
Sub SetDefaultLGVal()
	lgBlnOpenedFlag = True
	lgBlnSoldToPartyChg = False
	lgBlnSalesGrpChg = False
	lgBlnPlantChg = False
	lgBlnFlgConChg = False
End Sub


'============================================================================================================
Sub initSpreadPosVariables()  

	C_PriceFlag1			= 1
	C_PriceFlag			= 2
	C_SoNo				= 3
	C_SoSeq				= 4
	C_ItemCode			= 5
	C_ItemName			= 6
	C_PriceFlagY		= 7		
	C_PriceFlagN		= 8		
	C_SoldToParty		= 9		
	C_SoldToPartyNm		= 10		
	C_SoDt				= 11
	C_DealType			= 12	
	C_DealTypeNm		= 13	
	C_PayTerms			= 14	
	C_PayTermsNm		= 15	
	C_BillToParty		= 16	
	C_BillToPartyNm		= 17	
	C_NetAmt			= 18	
	C_Currency			= 19
	C_SalesGrp			= 20
	C_SalesGrpNm		= 21
	C_Plant				= 22	
	C_PlantNm			= 23
	C_ItemSpec			= 24	'규격필드추가 
	C_SoUnit			= 25
	C_TrackingNo		= 26
	
End Sub

'============================================================================================================
Sub InitVariables()

    lgIntFlgMode = Parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
    lgStrPrevKey = ""
    lgLngCurRows = 0  
End Sub

'============================================================================================================
Sub SetDefaultVal()
	frm1.txtSoNo.focus
	frm1.txtFromDate.Text = StartDate
	frm1.txtToDate.Text = EndDate
	frm1.txtBaseDate.Text=EndDate
	If Not(ChkRulePrice) then	
		Exit Sub		
	End If

	If lgSGCd <> "" then 
		Call ggoOper.SetReqAttr(frm1.txtSalesGrp, "Q") 
        	frm1.txtSalesGrp.value = lgSGCd
	End If

	If lgPLCd <> "" then 
		Call ggoOper.SetReqAttr(frm1.txtPlant, "Q") 
        	frm1.txtPlant.value = lgPLCd
	End If
		
	frm1.btnAllselect.disabled	= True
	frm1.btnDeselect.disabled	= True
	frm1.btnOpenPrice.disabled	= True	
	lgBlnFlgChgValue = False
End Sub
'============================================================================================================
Function ChkRulePrice()
	DIM lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
	Dim arrRtn
	
	'lgBlnPriceRule  = False 
	
	ChkRulePrice=True	
	'두번째이후부터는 수정 가능하도록.
	If lgBlnPriceRule Then Exit Function
	
	Call CommonQueryRs(" MINOR_CD "," B_CONFIGURATION "," MAJOR_CD = " & FilterVar("S1000", "''", "S") & " And REFERENCE=" & FilterVar("Y", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	arrRtn=split(lgF0,chr(11))
	
    If ubound(arrRtn)=-1 then    
    	Call DisplayMsgBox("171214","X","X","X")
    	ChkRulePrice=False    
    	Exit Function 	
    ElseIf Trim(arrRtn(0))="T" then    		
		frm1.rdoPriceFlagT.checked=True
		frm1.txtPriceFlag.value = frm1.rdoPriceFlagT.value		
	Else	
		frm1.rdoPriceFlagN.checked=True
		frm1.txtPriceFlag.value = frm1.rdoPriceFlagN.value		
    End if

	
End Function
'============================================================================================================
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029A( "I", "*", "NOCOOKIE", "MA") %>
<% Call LoadBNumericFormatA("I", "*", "NOCOOKIE", "MA") %>
End Sub

'============================================================================================================
Sub InitSpreadSheet()
	Call initSpreadPosVariables()
	With frm1.vspdData
		ggoSpread.Source = frm1.vspdData	
		ggoSpread.Spreadinit "V20030710",, parent.gAllowDragDropSpread
		.ReDraw = false
	    .MaxCols = C_SoUnit + 1												'☜: 최대 Columns의 항상 1개 증가시킴 
	    .Col = .MaxCols														'☜: 공통콘트롤 사용 Hidden Column
	    .ColHidden = True

	    .MaxRows = 0
				
		Call GetSpreadColumnPos("A")
		
		ggoSpread.SSSetCheck C_PriceFlag1, "선택", 8,,,True		
		ggoSpread.SSSetEdit C_PriceFlag, "확정여부", 10
		ggoSpread.SSSetEdit C_SoNo, "수주번호", 18
		ggoSpread.SSSetEdit C_SoSeq, "수주순번", 10,1
	    ggoSpread.SSSetEdit C_ItemCode, "품목", 18,,,18,2
	    ggoSpread.SSSetEdit C_ItemName, "품목명", 25,,,40
		ggoSpread.SSSetFloat C_PriceFlagY,"가단가",15,"C",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
		ggoSpread.SSSetFloat C_PriceFlagN,"진단가",15,"C",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
	    ggoSpread.SSSetEdit C_SoldToParty, "주문처", 10,,,,2
	    ggoSpread.SSSetEdit C_SoldToPartyNm, "주문처명", 15
	    ggoSpread.SSSetDate C_SoDt, "수주일",10,2,Parent.gDateFormat
	    ggoSpread.SSSetEdit C_DealType, "거래유형", 10,,,,2
	    ggoSpread.SSSetEdit C_DealTypeNm, "거래유형", 12
	    ggoSpread.SSSetEdit C_PayTerms, "결제방법", 10,,,,2
	    ggoSpread.SSSetEdit C_PayTermsNm, "결제방법명", 12
	    ggoSpread.SSSetEdit C_BillToParty, "발행처", 10,,,,2
	    ggoSpread.SSSetEdit C_BillToPartyNm, "발행처명", 15
		ggoSpread.SSSetFloat C_NetAmt,"수주금액",15,"A",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
	    ggoSpread.SSSetEdit C_Currency, "화폐", 10	    
	    ggoSpread.SSSetEdit C_SalesGrp, "영업그룹",10,,,4,2
	    ggoSpread.SSSetEdit C_SalesGrpNm, "영업그룹명",20
	    ggoSpread.SSSetEdit C_Plant, "공장",10,,,4,2
	    ggoSpread.SSSetEdit C_PlantNm, "공장명",20
		ggoSpread.SSSetEdit C_ItemSpec, "규격", 20
		ggoSpread.SSSetEdit C_SoUnit, "수주단위", 10
		ggoSpread.SSSetEdit	C_TrackingNo,		"Tracking No",	15,		,					,	  25,	  2
		
       Call ggoSpread.SSSetColHidden(C_DealType,C_DealType,True)
       Call ggoSpread.SSSetColHidden(C_PayTerms, C_PayTerms, True)
       Call ggoSpread.SSSetColHidden(C_SoUnit, C_SoUnit, True)
    
        .ColsFrozen = 1
        
	   .ReDraw = true
   
    End With
    
End Sub

'============================================================================================================
Sub SetSpreadLock()

End Sub

'============================================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
    
    With frm1
    
	    .vspdData.ReDraw = False
		    ggoSpread.SSSetProtected C_SoNo, pvStartRow, pvEndRow
		    ggoSpread.SSSetProtected C_SoSeq, pvStartRow, pvEndRow
		    ggoSpread.SSSetProtected C_ItemCode, pvStartRow, pvEndRow
		    ggoSpread.SSSetProtected C_ItemName, pvStartRow, pvEndRow
		    ggoSpread.SSSetProtected C_SoldToParty, pvStartRow, pvEndRow
		    ggoSpread.SSSetProtected C_SoldToPartyNm, pvStartRow, pvEndRow
		    ggoSpread.SSSetProtected C_SoDt, pvStartRow, pvEndRow
		    ggoSpread.SSSetProtected C_DealType, pvStartRow, pvEndRow
		    ggoSpread.SSSetProtected C_DealTypeNm, pvStartRow, pvEndRow
		    ggoSpread.SSSetProtected C_PayTerms, pvStartRow, pvEndRow
		    ggoSpread.SSSetProtected C_PayTermsNm, pvStartRow, pvEndRow
		    ggoSpread.SSSetProtected C_BillToParty, pvStartRow, pvEndRow
		    ggoSpread.SSSetProtected C_BillToPartyNm, pvStartRow, pvEndRow
		    ggoSpread.SSSetProtected C_NetAmt, pvStartRow, pvEndRow
		    ggoSpread.SSSetProtected C_Currency, pvStartRow, pvEndRow
		    ggoSpread.SSSetProtected C_PriceFlagY, pvStartRow, pvEndRow
		    ggoSpread.SSSetRequired C_PriceFlagN, pvStartRow, pvEndRow
			ggoSpread.SSSetProtected C_SalesGrp, pvStartRow, pvEndRow
		    ggoSpread.SSSetProtected C_SalesGrpNm, pvStartRow, pvEndRow
		    ggoSpread.SSSetProtected C_Plant, pvStartRow, pvEndRow
		    ggoSpread.SSSetProtected C_PlantNm, pvStartRow, pvEndRow
			ggoSpread.SSSetProtected C_ItemSpec, pvStartRow, pvEndRow
			ggoSpread.SSSetProtected C_SoUnit, pvStartRow, pvEndRow
			ggoSpread.SSSetProtected C_TrackingNo, pvStartRow, pvEndRow
			
	    .vspdData.ReDraw = True

    
    End With

End Sub

'============================================================================================================
Function OpenSoNo()

	Dim strRet
	Dim iCalledAspName
	Dim IntRetCD				

	If IsOpenPop = True Then Exit Function			
	IsOpenPop = True
	
	iCalledAspName = AskPRAspName("S3111PA1_ko441")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "S3111PA1_ko441", "X")			
		IsOpenPop = False
		Exit Function
	End If
		
		
	strRet = window.showModalDialog(iCalledAspName, Array(window.Parent, ""), _
		"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If strRet = "" Then
		Exit Function
	Else
		frm1.txtSoNo.value = strRet
		frm1.txtSoNo.focus
	End If	

End Function

'============================================================================================================
Function OpenConSoldToParty()

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "주문처"								
	arrParam(1) = "b_biz_partner"							
	arrParam(2) = Trim(frm1.txtSoldToParty.Value)			
	
	arrParam(4) = "bp_type like " & FilterVar("C%", "''", "S") & ""							
	arrParam(5) = "주문처"								
	
	arrField(0) = "bp_cd"								
	arrField(1) = "bp_nm"									
    
	arrHeader(0) = "주문처"								
	arrHeader(1) = "주문처명"							

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetConPopup(arrRet, C_PopSoldToParty)
	End If	
	
End Function

'============================================================================================================
Function OpenConSalesGrp()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	
	If frm1.txtSalesGrp.className = "protected" Then Exit Function
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "영업그룹"					
	arrParam(1) = "B_SALES_GRP"						
	arrParam(2) = Trim(frm1.txtSalesGrp.value)		
	arrParam(3) = ""
	arrParam(4) = "USAGE_FLAG=" & FilterVar("Y", "''", "S") & " "					
	arrParam(5) = "영업그룹"					
		
	arrField(0) = "SALES_GRP"						
	arrField(1) = "SALES_GRP_NM"					
	    
	arrHeader(0) = "영업그룹"					
	arrHeader(1) = "영업그룹명"					

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetConPopup(arrRet, C_PopSalesGrp)
	End If	

End Function


'============================================================================================================
Function OpenConPlant()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If frm1.txtPlant.className = "protected" Then Exit Function
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "공장"					
	arrParam(1) = "B_PLANT"						
	arrParam(2) = Trim(frm1.txtPlant.value)		
	arrParam(3) = ""
	arrParam(4) = ""					
	arrParam(5) = "공장"					
		
	arrField(0) = "Plant_cd"						
	arrField(1) = "Plant_NM"					
	    
	arrHeader(0) = "공장"					
	arrHeader(1) = "공장명"					

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetConPopup(arrRet, C_PopPlant)
	End If	

End Function

Function OpenTrackingNo()

	Dim iCalledAspName
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	
	'2002-10-07 s3135pa1.asp 추가 
	Dim strRet
		
	Dim arrTNParam(5), i

	For i = 0 to UBound(arrTNParam)
		arrTNParam(i) = ""
	Next	

	'20021227 kangjungu dynamic popup
	iCalledAspName = AskPRAspName("s3135pa1")	
	if Trim(iCalledAspName) = "" then
		IntRetCD = DisplayMsgBox("900040",parent.VB_INFORMATION, "s3135pa1", "x")
		IsOpenPop = False
		exit Function
	end if

	strRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrTNParam), _
		"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If strRet = "" Then
		Exit Function
	Else
		frm1.txtTrackingNo.value = strRet 
	End If		
		
	frm1.txtTrackingNo.focus
		
End Function
'============================================================================================================
'단가불러오기 처리 
Function OpenPrice()

 'Err.Clear																
 
    Dim iLngRow    , iLngCol    
	Dim iStrVal
	Dim tmpRet
	Dim strFlag
	
    OpenPrice = False                                                      
    
	If LayerShowHide(1) = False Then
		Exit Function
	End If
	
	With frm1

	'선택여부 
	If frm1.rdoPriceFlagT.checked = True Then
		frm1.txtPriceFlag.value = frm1.rdoPriceFlagT.value 
	Else
		frm1.txtPriceFlag.value = frm1.rdoPriceFlagN.value 
	End If
		iStrVal = ""
		'-----------------------
		'Data manipulate area
		'-----------------------
		Dim iCurColumnPos
		
		For iLngRow = 1 To .vspdData.MaxRows    
		    .vspdData.Row = iLngRow  :		    .vspdData.Col =C_PriceFlag1		    
		    strFlag = .vspdData.Text
		    
		    If strFlag="1" then
				'박 정 순 수정 컬럼 위치 변경시 단가 못 불러옴.
   
		            	ggoSpread.Source = frm1.vspdData
            			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

				iStrVal =  iStrVal & iLngRow & Parent.gColSep
				
				For iLngCol =1 To .vspdData.MaxCols -1								
					.vspdData.Col = iCurColumnPos(iLngCol)					
					 iStrVal = iStrVal & Trim(.vspdData.Text) & Parent.gColSep			
				Next
				.vspdData.Col= .vspdData.MaxCols
				iStrVal = iStrVal & Trim(.vspdData.Text) & Parent.gRowSep


			End If
		Next		
		
		'msgbox istrval
		.txtSpread.value = iStrVal

		If TRim(iStrVal) <>"" then
			Call ExecMyBizASP(frm1, BIZ_POP_PGM_ID)										'☜: 비지니스 ASP 를 가동	
		Else
			Call LayerShowHide(0)                                                               '☜: Hide Processing message
			Exit function 
		End If

	End With
	
    OpenPrice = True                                                           '⊙: Processing is NG 
		

End Function
'============================================================================================================
Function SetConPopup(Byval pvArrRet,ByVal pvIntWhere)

	SetConPopup = False

	Select Case pvIntWhere
	Case C_PopSoldToParty
		frm1.txtSoldToParty.value = pvArrRet(0) 
		frm1.txtSoldToPartyNm.value = pvArrRet(1)   
	Case C_PopSalesGrp
		frm1.txtSalesGrp.value = pvArrRet(0) 
		frm1.txtSalesGrpNm.value = pvArrRet(1)   
	Case C_PopPlant
		frm1.txtPlant.value = pvArrRet(0) 
		frm1.txtPlantNm.value = pvArrRet(1)   
	End Select

	SetConPopup = True

End Function

'============================================================================================================
Sub SetQuerySpreadColor(ByVal lRow)
	
    With frm1

    .vspdData.ReDraw = False


'		If .txtPostFlag.value = .rdoPostFlagY.value Then
'			ggoSpread.SpreadLock C_PriceFlag1,lRow,C_PriceFlag,.vspdData.MaxRows
'		Else
			ggoSpread.SpreadUnLock C_PriceFlag1,lRow,C_PriceFlag,.vspdData.MaxRows
'		End If
				
		ggoSpread.SpreadLock C_PriceFlag,lRow,C_PriceFlag,.vspdData.MaxRows
		ggoSpread.SSSetProtected C_SoNo, lRow, .vspdData.MaxRows
		ggoSpread.SSSetProtected C_SoSeq, lRow, .vspdData.MaxRows
		ggoSpread.SSSetProtected C_ItemCode, lRow, .vspdData.MaxRows
		ggoSpread.SSSetProtected C_ItemName, lRow, .vspdData.MaxRows
		ggoSpread.SSSetProtected C_SoldToParty, lRow, .vspdData.MaxRows
		ggoSpread.SSSetProtected C_SoldToPartyNm, lRow, .vspdData.MaxRows
		ggoSpread.SSSetProtected C_SoDt, lRow, .vspdData.MaxRows
		ggoSpread.SSSetProtected C_DealType, lRow, .vspdData.MaxRows
		ggoSpread.SSSetProtected C_DealTypeNm, lRow, .vspdData.MaxRows
		ggoSpread.SSSetProtected C_PayTerms, lRow, .vspdData.MaxRows
		ggoSpread.SSSetProtected C_PayTermsNm, lRow, .vspdData.MaxRows
		ggoSpread.SSSetProtected C_BillToParty, lRow, .vspdData.MaxRows
		ggoSpread.SSSetProtected C_BillToPartyNm, lRow, .vspdData.MaxRows
		ggoSpread.SSSetProtected C_NetAmt, lRow, .vspdData.MaxRows
		ggoSpread.SSSetProtected C_Currency, lRow, .vspdData.MaxRows
		ggoSpread.SSSetProtected C_PriceFlagY, lRow, .vspdData.MaxRows		
		ggoSpread.SSSetProtected C_SalesGrp, lRow, .vspdData.MaxRows
		ggoSpread.SSSetProtected C_SalesGrpNm, lRow, .vspdData.MaxRows
		ggoSpread.SSSetProtected C_Plant, lRow, .vspdData.MaxRows
		ggoSpread.SSSetProtected C_PlantNm, lRow, .vspdData.MaxRows				
		ggoSpread.SSSetProtected C_ItemSpec, lRow, .vspdData.MaxRows
		ggoSpread.SSSetProtected C_SoUnit, lRow, .vspdData.MaxRows
		ggoSpread.SSSetProtected C_TrackingNo, lRow, .vspdData.MaxRows
		
    .vspdData.ReDraw = True
    
    End With

End Sub

'============================================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			
			C_PriceFlag1			= iCurColumnPos(1)
			C_PriceFlag			= iCurColumnPos(2)
			C_SoNo				= iCurColumnPos(3)
			C_SoSeq				= iCurColumnPos(4)
			C_ItemCode			= iCurColumnPos(5)
			C_ItemName			= iCurColumnPos(6)
			C_PriceFlagY		= iCurColumnPos(7)
			C_PriceFlagN		= iCurColumnPos(8)
			C_SoldToParty		= iCurColumnPos(9)
			C_SoldToPartyNm		= iCurColumnPos(10)
			C_SoDt				= iCurColumnPos(11)
			C_DealType			= iCurColumnPos(12)
			C_DealTypeNm		= iCurColumnPos(13)
			C_PayTerms			= iCurColumnPos(14)
			C_PayTermsNm		= iCurColumnPos(15)
			C_BillToParty		= iCurColumnPos(16)
			C_BillToPartyNm		= iCurColumnPos(17)
			C_NetAmt			= iCurColumnPos(18)
			C_Currency			= iCurColumnPos(19)
			C_SalesGrp			= iCurColumnPos(20)
			C_SalesGrpNm		= iCurColumnPos(21)
			C_Plant				= iCurColumnPos(22)
			C_PlantNm			= iCurColumnPos(23)
			C_ItemSpec			= iCurColumnPos(24)
			C_SoUnit			= iCurColumnPos(25)	
			C_TrackingNo		= iCurColumnPos(26)	
    End Select    
End Sub

'============================================================================================================
Sub Form_Load()

	Call InitVariables	
	
	
    Call LoadInfTB19029    
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)		
    Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field
    '----------  Coding part  -------------------------------------------------------------
    Call GetValue_ko441()    
	Call InitSpreadSheet
	Call SetDefaultVal	
	Call SetDefaultLGVal	
    Call SetToolbar("11000000000011")										'⊙: 버튼 툴바 제어 

End Sub

'============================================================================================================
Function txtSoldToParty_OnKeyDown()
	lgBlnSoldToPartyChg = True
	lgBlnFlgConChg = True
End Function

'============================================================================================================
Function txtSalesGrp_OnKeyDown()
	lgBlnSalesGrpChg = True
	lgBlnFlgConChg = True
End Function

'============================================================================================================
Function txtPlant_OnKeyDown()
	lgBlnPlantChg = True
	lgBlnFlgConChg = True
End Function


'============================================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)

	Exit Sub

	If lgIntFlgMode = Parent.OPMD_CMODE Then Exit Sub

	If Col = C_PriceFlag1 And Row > 0 Then
	    Select Case ButtonDown
	    Case 0				
			ggoSpread.Source = frm1.vspdData
			ggoSpread.EditUndo
			lgBlnFlgChgValue = False
	    Case 1
			ggoSpread.Source = frm1.vspdData
			ggoSpread.UpdateRow Row
			lgBlnFlgChgValue = True
	    End Select
    End If

End Sub


'============================================================================================================
Sub vspdData_Click(ByVal Col , ByVal Row )
    
    Call SetPopupMenuItemInf("0000111111")
    gMouseClickStatus = "SPC"
    
    Set gActiveSpdSheet = frm1.vspdData

    If frm1.vspdData.MaxRows = 0 Then                                                    
       Exit Sub
   	End If
   	    
	If Row <= 0 Then
	    ggoSpread.Source = frm1.vspdData
	    If lgSortKey = 1 Then
	        ggoSpread.SSSort Col				'Sort in Ascending
	        lgSortKey = 2
	    Else
	        ggoSpread.SSSort Col, lgSortKey		'Sort in Descending
	        lgSortKey = 1
	    End If
		 Exit Sub     
	End If

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    	frm1.vspdData.Row = Row
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub

'============================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub

'============================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)				
    If Row <= 0 Then
		Exit Sub
    End If
    
    If frm1.vspdData.MaxRows = 0 Then
    	Exit Sub
    End If

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub


'============================================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If

End Sub


'============================================================================================================
Sub vspdData_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    If Col <= C_PriceFlag Or NewCol <= C_PriceFlag Then
        Cancel = True
        Exit Sub
    End If
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    
    Call GetSpreadColumnPos("A")
End Sub


'============================================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )

	Frm1.vspdData.Row = Row
    Frm1.vspdData.Col = Col
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	
    Select Case Col
	     Case  C_PriceFlagN
             Call FixDecimalPlaceByCurrency(frm1.vspdData,Row,C_Currency,C_PriceFlagN,"C" ,"X","X")
             Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData,Row,Row,C_Currency,C_PriceFlagN, "C" ,"I","X","X")            
    End Select    
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
             
    Call CheckMinNumSpread(frm1.vspdData, Col, Row)		
    
	ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row
        
	lgBlnFlgChgValue = True

	Select Case Col
	Case C_PriceFlagN
		frm1.vspdData.Row = Row
		frm1.vspdData.Col = C_PriceFlag1
		frm1.vspdData.Text = "1"
	End Select
    
End Sub

'============================================================================================================
Sub vspdData_EditMode(ByVal Col, ByVal Row, ByVal Mode, ByVal ChangeMade)
    Select Case Col
        Case C_PriceFlagN
            Call EditModeCheck(frm1.vspdData, Row, C_Currency, C_PriceFlagN, "C" ,"I", Mode, "X", "X")
    End Select
End Sub

'============================================================================================================
Sub vspdData_LeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)    
End Sub

'============================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )    
    If OldLeft <> NewLeft Then Exit Sub
	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	    
    	
    	If lgStrPrevKey <> "" Then				    		
			Call DisableToolBar(parent.TBC_QUERY)
			Call DBQuery
    	End If
    End If    
End Sub

'============================================================================================================
Sub txtFromDate_DblClick(Button)
	If Button = 1 Then
		frm1.txtFromDate.Action = 7
		Call SetFocusToDocument("M")
        frm1.txtFromDate.Focus
	End If
End Sub

'============================================================================================================
Sub txtToDate_DblClick(Button)
	If Button = 1 Then
		frm1.txtToDate.Action = 7
		Call SetFocusToDocument("M")
        frm1.txtToDate.Focus
	End If
End Sub

'============================================================================================================
Sub txtBaseDate_DblClick(Button)
	If Button = 1 Then
		frm1.txtBaseDate.Action = 7
		Call SetFocusToDocument("M")
        frm1.txtBaseDate.Focus
	End If
End Sub
'============================================================================================================
Sub txtFromDate_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
End Sub

Sub txtToDate_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
End Sub


Sub txtBaseDate_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
End Sub
'============================================================================================================
Sub btnAllSelect_OnClick()
	Dim i
	
	with frm1.vspdData
		for i= 1 to .MaxRows
			.Row= i :			.Col = 1			
			.value=	"1"
			ggoSpread.Source = frm1.vspdData
			ggoSpread.UpdateRow .Row
			lgBlnFlgChgValue = True
		next
	End With

'	frm1.btnAllselect.disabled=True
'	frm1.btnDeselect.disabled=False
End Sub
'============================================================================================================
Sub btnDeSelect_OnClick()
	Dim i
	
	with frm1.vspdData
		for i= 1 to .MaxRows
			.Row= i :			.Col = 1
			.value="0"
			ggoSpread.Source = frm1.vspdData
			ggoSpread.EditUndo .Row
			lgBlnFlgChgValue = False			
		next		
	End With
'	frm1.btnAllselect.disabled=False
'	frm1.btnDeselect.disabled=True
End Sub
'============================================================================================================
Function FncQuery() 
    Dim IntRetCD 
    
    FncQuery = False                                                     
    
    Err.Clear                                                              

	'** ValidDateCheck(pObjFromDt, pObjToDt) : 'pObjToDt'이 'pObjFromDt'보다 크거나 같아야 할때 **
	If ValidDateCheck(frm1.txtFromDate, frm1.txtToDate) = False Then Exit Function

	' 조회조건 유효값 check
	If 	lgBlnFlgConChg Then
		If Not ChkValidityQueryCon Then	Exit Function
	End If
		'단가적용규칙 
	If Not(ChkRulePrice) then	
		Exit Function	
	End If
	
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO, "X", "X")		
		If IntRetCD = vbNo Then
		    Exit Function
		End If
    End If
    
    Call ggoOper.ClearField(Document, "2")										
    Call InitVariables															

    If Not chkField(Document, "1") Then								
       Exit Function
    End If

	If frm1.rdoPostFlagY.checked = True Then
		frm1.txtPostFlag.value = frm1.rdoPostFlagY.value 
	Else
		frm1.txtPostFlag.value = frm1.rdoPostFlagN.value 
	End If
	'선택여부 
	If frm1.rdoPriceFlagT.checked = True Then
		frm1.txtPriceFlag.value = frm1.rdoPriceFlagT.value 
	Else
		frm1.txtPriceFlag.value = frm1.rdoPriceFlagN.value 
	End If
	
    Call DbQuery															

    FncQuery = True																
        
End Function
'============================================================================================================
Function FncQuery2() 
    Dim IntRetCD 
    
    FncQuery2 = False                                                     
    
    Err.Clear                                                              

	'** ValidDateCheck(pObjFromDt, pObjToDt) : 'pObjToDt'이 'pObjFromDt'보다 크거나 같아야 할때 **
	If ValidDateCheck(frm1.txtFromDate, frm1.txtToDate) = False Then Exit Function

	' 조회조건 유효값 check
	If 	lgBlnFlgConChg Then
		If Not ChkValidityQueryCon Then	Exit Function
	End If
		'단가적용규칙 
	If Not(ChkRulePrice) then	
		Exit Function	
	End If
	
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO, "X", "X")		
		If IntRetCD = vbNo Then
		    Exit Function
		End If
    End If
    
    Call ggoOper.ClearField(Document, "2")										
    Call InitVariables															

    If Not chkField(Document, "1") Then								
       Exit Function
    End If

	If frm1.rdoPostFlagY.checked = True Then
		frm1.txtPostFlag.value = frm1.rdoPostFlagY.value 
	Else
		frm1.rdoPostFlagY.checked= True
		frm1.txtPostFlag.value = frm1.rdoPostFlagY.value 
	End If
	'선택여부 
	If frm1.rdoPriceFlagT.checked = True Then
		frm1.txtPriceFlag.value = frm1.rdoPriceFlagT.value 
	Else
		frm1.txtPriceFlag.value = frm1.rdoPriceFlagN.value 
	End If
	
    Call DbQuery2															

    FncQuery2 = True																
        
End Function
'============================================================================================================
Function FncNew() 
    Dim IntRetCD 
    
    FncNew = False                                                        
    
    If lgBlnFlgChgValue = True Then
        IntRetCD = DisplayMsgBox("900015", Parent.VB_YES_NO, "X", "X") 	
		If IntRetCD = vbNo Then
		    Exit Function
		End If
    End If
    
    Call ggoOper.ClearField(Document, "A")                                                           
    Call ggoOper.LockField(Document, "N")                                      
    Call SetDefaultVal
    Call InitVariables														

    Call SetToolbar("11000000000011")	    

    FncNew = True															

End Function


'============================================================================================================
Function FncSave() 
    Dim IntRetCD 
    
    FncSave = False                                                       
    
    Err.Clear                                                              

    ggoSpread.Source = frm1.vspdData
    If ggoSpread.Source = False Then
        IntRetCD = DisplayMsgBox("900001", "X", "X", "X")       
        Exit Function
    End If
    
	ggoSpread.Source = frm1.vspdData

    If Not chkField(Document, "2") Then    
       Exit Function
    End If

    If Not ggoSpread.SSDefaultCheck  Then    
       Exit Function
    End If
      
    CAll DbSave		 
    
    FncSave = True                                                         
    
End Function


'============================================================================================================
Function FncCancel() 
	If frm1.vspdData.MaxRows < 1 Then Exit Function
    ggoSpread.Source = frm1.vspdData	
    ggoSpread.EditUndo  
End Function

'============================================================================================================
Function FncPrint() 
    Call parent.FncPrint()
End Function

'============================================================================================================
Function FncExcel() 
	Call parent.FncExport(Parent.C_SINGLEMULTI)
End Function

'============================================================================================================
Function FncFind() 
	Call parent.FncFind(Parent.C_SINGLEMULTI, False)
End Function

'============================================================================================================
Sub FncSplitColumn()

    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   

End Sub

'============================================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub

'============================================================================================================
Sub PopRestoreSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()      
    
	Call ggoSpread.ReOrderingSpreadData()
	Call SetQuerySpreadColor(1)
End Sub

'============================================================================================================
Function FncExit()
	Dim IntRetCD
	FncExit = False

    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO, "X", "X")		
		If IntRetCD = vbNo Then
		    Exit Function
		End If
    End If

    FncExit = True
End Function

'============================================================================================================
Function DbDelete() 
    On Error Resume Next                                               
End Function

'============================================================================================================
Function DbDeleteOk()													
    On Error Resume Next                                                   
End Function

'============================================================================================================
Function DbQuery() 
on error resume next

    Err.Clear                                                              
    
    DbQuery = False                                                       

	If LayerShowHide(1) = False Then
		Exit Function
	End If
	    
    Dim iStrVal
    
    If lgIntFlgMode = Parent.OPMD_UMODE Then
		iStrVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001 & _								
							"&txtSoldToParty=" & Trim(frm1.txtHSoldToParty.value) & _			
							"&txtFromDate=" & Trim(frm1.txtHFromDate.value) & _	
							"&txtToDate=" & Trim(frm1.txtHToDate.value) & _									
							"&txtPostFlag=" & Trim(frm1.txtPostFlag.value) & _
							"&txtSoNo=" & Trim(frm1.txtHSoNo.value) & _
							"&txtSalesGrp=" & Trim(frm1.txtHSalesGrp.value) & _
							"&txtPlant=" & Trim(frm1.txtHPlant.value) & _		
							"&txtPriceFlag=" & Trim(frm1.txtPriceFlag.value) & _
							"&txtBaseDate=" & Trim(frm1.txtHBaseDate.value) & _
							"&txtTrackingNo= " & Trim(frm1.txtHTrackingNo.value) & _
							"&txtMaxRows=" & frm1.vspdData.MaxRows & _		
							"&lgStrPrevKey=" & lgStrPrevKey		
    Else
		iStrVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001 & _								
							"&txtSoldToParty=" & Trim(frm1.txtSoldToParty.value) & _			
							"&txtFromDate=" & Trim(frm1.txtFromDate.Text) & _	
							"&txtToDate=" & Trim(frm1.txtToDate.Text) & _									
							"&txtPostFlag=" & Trim(frm1.txtPostFlag.value) & _
							"&txtSoNo=" & Trim(frm1.txtSoNo.value) & _
							"&txtSalesGrp=" & Trim(frm1.txtSalesGrp.value) & _
							"&txtPlant=" & Trim(frm1.txtPlant.value) & _
							"&txtPriceFlag=" & Trim(frm1.txtPriceFlag.value) & _
							"&txtBaseDate=" & Trim(frm1.txtBaseDate.value) & _
							"&txtTrackingNo= " & Trim(frm1.txtTrackingNo.value) & _
							"&txtMaxRows=" & frm1.vspdData.MaxRows & _		
							"&lgStrPrevKey=" & lgStrPrevKey
    End If    

	Call RunMyBizASP(MyBizASP, iStrVal)											
	
    DbQuery = True																	

End Function
'============================================================================================================
Function DbQuery2() 
on error resume next

    Err.Clear                                                              
    
    DbQuery2 = False                                                       

	If LayerShowHide(1) = False Then
		Exit Function
	End If
	    
    Dim iStrVal
    
     If lgIntFlgMode = Parent.OPMD_UMODE Then
		iStrVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001 & _								
							"&txtSoldToParty=" & Trim(frm1.txtHSoldToParty.value) & _			
							"&txtFromDate=" & Trim(frm1.txtHFromDate.value) & _	
							"&txtToDate=" & Trim(frm1.txtHToDate.value) & _									
							"&txtPostFlag=" & Trim(frm1.txtPostFlag.value) & _
							"&txtSoNo=" & Trim(frm1.txtHSoNo.value) & _
							"&txtSalesGrp=" & Trim(frm1.txtHSalesGrp.value) & _
							"&txtPlant=" & Trim(frm1.txtHPlant.value) & _		
							"&txtPriceFlag=" & Trim(frm1.txtPriceFlag.value) & _
							"&txtBaseDate=" & Trim(frm1.txtHBaseDate.value) & _
							"&txtMaxRows=" & frm1.vspdData.MaxRows & _		
							"&lgStrPrevKey=" & lgStrPrevKey		
    Else
		iStrVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001 & _								
							"&txtSoldToParty=" & Trim(frm1.txtSoldToParty.value) & _			
							"&txtFromDate=" & Trim(frm1.txtFromDate.Text) & _	
							"&txtToDate=" & Trim(frm1.txtToDate.Text) & _									
							"&txtPostFlag=" & Trim(frm1.txtPostFlag.value) & _
							"&txtSoNo=" & Trim(frm1.txtSoNo.value) & _
							"&txtSalesGrp=" & Trim(frm1.txtSalesGrp.value) & _
							"&txtPlant=" & Trim(frm1.txtPlant.value) & _
							"&txtPriceFlag=" & Trim(frm1.txtPriceFlag.value) & _
							"&txtBaseDate=" & Trim(frm1.txtBaseDate.value) & _
							"&txtMaxRows=" & frm1.vspdData.MaxRows & _		
							"&lgStrPrevKey=" & lgStrPrevKey
    End If    

	Call RunMyBizASP(MyBizASP, iStrVal)											
	
    DbQuery2 = True																	

End Function



'============================================================================================================
Function DbQueryOk()														
	
    '-----------------------
    'Reset variables area
    '-----------------------
    lgIntFlgMode = Parent.OPMD_UMODE											
  
    Call SetToolbar("11101000000111")					   
	Call SetQuerySpreadColor(1)

	Call ButtonVisible(1)

	lgBlnPriceRule = true
	
	lgBlnFlgChgValue = False
End Function

'============================================================================================================
Function DbSave() 

    Err.Clear																
 
    Dim iLngRow        
	Dim iStrVal
	
    DbSave = False                                                      
    
	If LayerShowHide(1) = False Then
		Exit Function
	End If
	
	With frm1
	
		.txtMode.value = Parent.UID_M0002

		iStrVal = ""
		'-----------------------
		'Data manipulate area
		'-----------------------
		For iLngRow = 1 To .vspdData.MaxRows
    
		    .vspdData.Row = iLngRow
		    .vspdData.Col = C_PriceFlag1
		    		    
			If Trim(.vspdData.Text) = "1" Then				
				
				.vspdData.Col = 0				
				Select Case .vspdData.Text
				
					Case ggoSpread.UpdateFlag							'☜: 수정		
						iStrVal = iStrVal & iLngRow & Parent.gColSep

						'--- 수주번호 
				        .vspdData.Col = C_SoNo		            
				        iStrVal = iStrVal & Trim(.vspdData.Text) & Parent.gColSep
						'--- 수주순번 
				        .vspdData.Col = C_SoSeq	
				        iStrVal = iStrVal & Trim(UNIConvNum(.vspdData.Text, 0)) & Parent.gColSep
						'--- 진단가 
				        .vspdData.Col = C_PriceFlagN		            
				        iStrVal = iStrVal & Trim(UNIConvNum(.vspdData.Text, 0)) & Parent.gRowSep

				End Select
		        
			End if      
		Next
		
		.txtSpread.value = iStrVal

		Call ExecMyBizASP(frm1, BIZ_PGM_ID)										'☜: 비지니스 ASP 를 가동 
	
	End With
	
    DbSave = True                                                           '⊙: Processing is NG
    
End Function

'============================================================================================================
Function DbSaveOk()															

    Call InitVariables
    
    Call fncQuery2()

End Function

'============================================================================================================
Function ChkValidityQueryCon()
	Dim iStrCode

	ChkValidityQueryCon = True

	If lgBlnSoldToPartyChg Then
		iStrCode = Trim(frm1.txtSoldToParty.value)
		If iStrCode <> "" Then
			iStrCode = " " & FilterVar(iStrCode, "''", "S") & ""
			If Not GetCodeName(iStrCode, "" & FilterVar("C%", "''", "S") & "", "default", "default", "default", "" & FilterVar("BP", "''", "S") & "", C_PopSoldToParty) Then
				Call DisplayMsgBox("970000", "X", frm1.txtSoldtoparty.alt, "X")
				frm1.txtSoldtoparty.focus
				ChkValidityQueryCon = False
				Exit Function
			End If
		Else
			frm1.txtSoldToPartyNm.value = ""
		End If
		lgBlnSoldToPartyChg	= False
	End If

	If lgBlnSalesGrpChg Then
		iStrCode = Trim(frm1.txtSalesGrp.value)
		If iStrCode <> "" Then
			iStrCode = " " & FilterVar(iStrCode, "''", "S") & ""
			If Not GetCodeName(iStrCode, "default", "default", "default", "default", "" & FilterVar("SG", "''", "S") & "", C_PopSalesGrp) Then
				Call DisplayMsgBox("970000", "X", frm1.txtSalesGrp.alt, "X")
				frm1.txtSalesGrp.focus
				ChkValidityQueryCon = False
				Exit Function
			End If
		Else
			frm1.txtSalesGrpNm.value = ""
		End If
		lgBlnSalesGrpChg = False
	End If
			
	If lgBlnPlantChg Then
		iStrCode = Trim(frm1.txtPlant.value)
		If iStrCode <> "" Then
			iStrCode = " " & FilterVar(iStrCode, "''", "S") & ""
			If Not GetCodeName(iStrCode, "default", "default", "default", "default", "" & FilterVar("PT", "''", "S") & "", C_PopPlant) Then
				Call DisplayMsgBox("970000", "X", frm1.txtPlant.alt, "X")
				frm1.txtPlant.focus
				ChkValidityQueryCon = False
				Exit Function
			End If
		Else
			frm1.txtPlantNm.value = ""
		End If
		lgBlnPlantChg = False
	End If
			
End Function


'============================================================================================================
Function GetCodeName(ByVal pvStrArg1, ByVal pvStrArg2, ByVal pvStrArg3, ByVal pvStrArg4, ByVal pvIntArg5, ByVal pvStrFlag, ByVal pvIntWhere)

	Dim iStrSelectList, iStrFromList, iStrWhereList
	Dim iStrRs
	Dim iArrRs(2), iArrTemp
	
	GetCodeName = False

	iStrSelectList = " * "
	iStrFromList = " dbo.ufn_s_GetCodeName (" & pvStrArg1 & ", " & pvStrArg2 & ", " & pvStrArg3 & ", " & pvStrArg4 & ", " & pvIntArg5 & ", " & pvStrFlag & ") "
	iStrWhereList = ""
	
	Err.Clear
	
	If CommonQueryRs2by2(iStrSelectList,iStrFromList,iStrWhereList,iStrRs) Then
		
		If iStrRs = "" Then
			Exit Function
		End If
		
		iArrTemp = Split(iStrRs, Chr(11))
		iArrRs(0) = iArrTemp(1)
		iArrRs(1) = iArrTemp(2)
		iArrRs(2) = iArrTemp(3)
		GetCodeName = SetConPopup(iArrRs, pvIntWhere)
	Else
		
	End if
End Function
'========================================================================================================
Function ButtonVisible(ByVal BRow)

	ButtonVisible = False
		
	If frm1.vspdData.MaxRows <1 Then
		frm1.btnAllselect.disabled = True
		frm1.btndeselect.disabled = True
		frm1.btnOpenPrice.disabled = True
		Exit Function
	End IF
	
	If    BRow >= 1  Then	
				frm1.vspdData.Row = BRow
				frm1.btnAllselect.disabled = False
				frm1.btndeselect.disabled = False
				frm1.btnOpenPrice.disabled = False	
	Else
			frm1.btnAllselect.disabled = True
			frm1.btndeselect.disabled = True
			frm1.btnOpenPrice.disabled = False
	End If

	ButtonVisible = True

End Function
</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>

<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">

<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR >
		<TD <%=HEIGHT_TYPE_00%>></TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSLTAB">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSLTAB"><font color=white>단가확정</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=*>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=*>
		<TD WIDTH=100% CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS="TD5" NOWRAP>수주번호</TD>
									<TD CLASS="TD6"><INPUT NAME="txtSoNo" ALT="수주번호" TYPE="Text" MAXLENGTH=18 SiZE=34 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSoHDR" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenSoNo"></TD>
									<TD CLASS=TD5 NOWRAP>수주일</TD>
									<TD CLASS=TD6 NOWRAP>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 NAME="txtFromDate" CLASS=FPDTYYYYMMDD tag="12X1" ALT="수주시작일" Title="FPDATETIME"></OBJECT>');</SCRIPT>
										&nbsp;~&nbsp;
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime2 NAME="txtToDate" CLASS=FPDTYYYYMMDD tag="12X1" ALT="수주종료일" Title="FPDATETIME"></OBJECT>');</SCRIPT>
									</TD>
								</TR>
								<TR>	
									<TD CLASS=TD5 NOWRAP>주문처</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtSoldToParty" ALT="주문처" TYPE="Text" MAXLENGTH="10" SIZE=10 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDnHDR" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConSoldToParty()">&nbsp;<INPUT NAME="txtSoldToPartyNm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="14"></TD>
									<TD CLASS=TD5 NOWRAP>단가확정여부</TD>
									<TD CLASS=TD6 NOWRAP>
										<input type=radio CLASS="RADIO" name="rdoPostFlag" id="rdoPostFlagY" value="Y" tag = "11">
											<label for="rdoPostFlagY">확정</label>&nbsp;&nbsp;&nbsp;&nbsp;
										<input type=radio CLASS = "RADIO" name="rdoPostFlag" id="rdoPostFlagN" value="N" tag = "11" checked>
											<label for="rdoPostFlagN">미확정</label>
									</TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>영업그룹</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtSalesGrp" TYPE="Text" MAXLENGTH="4" SIZE=10 tag="11XXXU" ALT="영업그룹"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSalesGrp" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConSalesGrp()">&nbsp;<INPUT NAME="txtSalesGrpNm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="14"></TD>
									<TD CLASS="TD5" NOWRAP>공장</TD>
									<TD CLASS="TD6" NOWRAP><INPUT NAME="txtPlant" ALT="공장" TYPE="Text" MAXLENGTH=4 SiZE=10 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlant" align=top TYPE="BUTTON" ONCLICK ="vbscript:OpenConPlant()">&nbsp;<INPUT NAME="txtPlantNm" TYPE="Text" MAXLENGTH="40" SIZE=25 tag="14"></TD>
								</TR>	
								<TR>
									<TD CLASS="TD5" NOWRAP>Tracking No</TD>
									<TD CLASS="TD6"><INPUT NAME="txtTrackingNo" ALT="Tracking No" TYPE="Text" MAXLENGTH=25 SiZE=34 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSoHDR" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenTrackingNo()"></TD>
									<TD CLASS="TD5" NOWRAP></TD>
									<TD CLASS="TD6" NOWRAP></TD>
								</TR>								
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=* VALIGN=TOP>
						<TABLE <%=LR_SPACE_TYPE_60%>>						
							<TR>
									<TD CLASS=TD5 NOWRAP>단가적용기준일</TD>
									<TD CLASS=TD6 NOWRAP>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 NAME="txtBaseDate" CLASS=FPDTYYYYMMDD tag="12X1" ALT="단가적용기준일" Title="FPDATETIME"></OBJECT>');</SCRIPT></TD>
									<TD CLASS="TD5" NOWRAP>단가적용규칙</TD>
									<TD CLASS="TD6" NOWRAP><input type=radio CLASS="RADIO" name="rdoPriceFlag" id="rdoPriceFlagT" value="T" tag = "11" >
											<label for="rdoPriceFlagT">진단가</label>&nbsp;&nbsp;&nbsp;
										<input type=radio CLASS = "RADIO" name="rdoPriceFlag" id="rdoPriceFlagN" value="N" tag = "11">
											<label for="rdoPriceFlagN">최신단가</label>
									</TD>
								</TR>
							<TR>
								<TD HEIGHT="100%" WIDTH="100%" COLSPAN=4>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=vaSpread1> <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR >
		<TD <%=HEIGHT_TYPE_01%>></TD>
	</TR>
		<TR HEIGHT=20>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_30%>>
			   <TR>
				 <TD WIDTH=10>&nbsp;</TD>
				 <TD>
				  <BUTTON NAME="btnAllSelect" CLASS="CLSSBTN">일괄선택</BUTTON>&nbsp;
				  <BUTTON NAME="btnDeselect" CLASS="CLSSBTN">일괄선택취소</BUTTON>&nbsp;
				  <BUTTON NAME="btnOpenPrice" CLASS="CLSSBTN" onClick="OpenPrice()">단가불러오기</BUTTON>&nbsp;				  
				 </TD>				
			   </TR>
			</TABLE>
		</TD>
	</TR>
	<TR >
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=No  noresize framespacing=0  TABINDEX = -1></IFRAME>
		</TD>
	</TR>	
</TABLE>

<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"  TABINDEX = -1>
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24" TABINDEX = -1>
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" TABINDEX = -1>
<TEXTAREA CLASS="hidden" NAME="txtSpread" TAG="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtBatch" tag="24" TABINDEX = -1>
<INPUT TYPE=HIDDEN NAME="txtPostFlag" tag="24" TABINDEX = -1>
<INPUT TYPE=HIDDEN NAME="txtPriceFlag" tag="24" TABINDEX = -1>
<INPUT TYPE=HIDDEN NAME="txtHBaseDate" tag="24" TABINDEX = -1>
<INPUT TYPE=HIDDEN NAME="txtHSalesGrp" tag="24"TABINDEX = -1>
<INPUT TYPE=HIDDEN NAME="txtHPlant" tag="24"TABINDEX = -1>
<INPUT TYPE=HIDDEN NAME="txtHSoldToParty" tag="24" TABINDEX = -1>
<INPUT TYPE=HIDDEN NAME="txtHFromDate" tag="24" TABINDEX = -1>
<INPUT TYPE=HIDDEN NAME="txtHToDate" tag="24" TABINDEX = -1>
<INPUT TYPE=HIDDEN NAME="txtHSoNo" tag="24" TABINDEX = -1>
<INPUT TYPE=HIDDEN NAME="txtHTrackingNo" tag="24">

</FORM>
<DIV ID="MousePT" NAME="MousePT" STYLE="visible:true">
	<IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm"></IFRAME>
</DIV>	
</BODY>
</HTML>
