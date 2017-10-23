<%@ LANGUAGE="VBSCRIPT" %>
<%
'======================================================================================================
'*  1. Module Name          : Sales
'*  2. Function Name        : 
'*  3. Program ID           : S2216MA1
'*  4. Program Name         : 공장별일별품목판매계획조정 
'*  5. Program Desc         : 공장별일별품목판매계획조정 
'*  6. Comproxy List        : +
'*  7. Modified date(First) : 2003/01/16
'*  8. Modified date(Last)  :
'*  9. Modifier (First)     : Seongbae Hwang
'* 10. Modifier (Last)      :
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'======================================================================================================= 
%>
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/IncSvrCcm.Inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT> 
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs"> </SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit                             '☜: indicates that All variables must be declared in advance

Const BIZ_PGM_ID = "s2216mb1.asp"            '☆: Head Query 비지니스 로직 ASP명 
Const BIZ_JUMP_ID = "s2216ba1"				 '☆: JUMP시 비지니스 로직 ASP명 

Const C_PopFrSpPeriod	= 1
Const C_PopToSpPeriod	= 2
Const C_PopPlantCd		= 3
Const C_PopSalesGrp		= 4
Const C_PopItemCd		= 5
Const C_PopSoldToParty	= 6

<!-- #Include file="../../inc/lgvariables.inc" --> 

'========================================================================================================
'☆: Spread Sheet의 Column
'☜:--------Spreadsheet #1-----------------------------------------------------------------------------   
Dim C_SpDt
Dim C_PlantCd
Dim C_PlantNm
Dim C_SalesGrp
Dim C_SalesGrpNm
Dim C_LocExpFlag
Dim C_LocExpFlagNm
Dim C_SoldToParty
Dim C_SoldToPartyNm
Dim C_ItemCd
Dim C_ItemNm
Dim C_Qty
Dim C_Unit
Dim C_QtyOrderUnitMfg
Dim C_OrderUnitMfg
Dim C_CfmFlag
Dim C_SpPeriod
Dim C_SpPeriodDesc
Dim C_SpMonth
Dim C_SpWeek

Dim C_Pointer
Dim C_OldQty
Dim C_OldQtyOrderUnitMfg

'☜:--------Spreadsheet #2-----------------------------------------------------------------------------   
Dim C_SpPeriod2
Dim C_SpPeriodDesc2
Dim C_PlantCd2
Dim C_PlantNm2
Dim C_TotQty
Dim	C_Unit2
Dim C_TotQtyOrderUnitMfg
Dim C_OrderUnitMfg2

'========================================================================================================
Dim lgBlnOpenPop
Dim lgStrWhere					' Scrollbar를 조회조건 
Dim lgBlnExistsSpConfig
Dim	lgLngUseStep
Dim lgLngProcessByPlant

Dim lgLngStartRow		' Start row to be queryed

Dim iDBSYSDate
Dim iStrFromDt

iDBSYSDate = "<%=GetSvrDate%>"
iStrFromDt = UNIConvDateAtoB(iDBSYSDate, parent.gServerDateFormat, parent.gDateFormat)

'========================================================================================================
Sub initSpreadPosVariables()  
	'☜:--------Spreadsheet #1-----------------------------------------------------------------------------   
   	If gMouseClickStatus = "N" Or gMouseClickStatus = "SPCRP" Then
		C_SpDt				= 1
		C_PlantCd			= 2
		C_PlantNm			= 3
		C_SalesGrp			= 4
		C_SalesGrpNm		= 5
		C_LocExpFlag		= 6
		C_LocExpFlagNm		= 7
		C_SoldToParty		= 8
		C_SoldToPartyNm		= 9
		C_ItemCd			= 10
		C_ItemNm			= 11
		C_Qty				= 12
		C_Unit				= 13
		C_QtyOrderUnitMfg	= 14
		C_OrderUnitMfg		= 15
		C_CfmFlag			= 16
		C_SpPeriod			= 17
		C_SpPeriodDesc		= 18
		C_SpMonth			= 19
		C_SpWeek			= 20

		C_Pointer			= 21
		C_OldQty			= 22
		C_OldQtyOrderUnitMfg= 23
	End If
	
	'☜:--------Spreadsheet #2-----------------------------------------------------------------------------   
   	If gMouseClickStatus = "N" Or gMouseClickStatus = "SP2CRP" Then
		C_SpPeriod2			= 1
		C_SpPeriodDesc2		= 2
		C_PlantCd2			= 3
		C_PlantNm2			= 4
		C_TotQty			= 5
		C_Unit2				= 6
		C_TotQtyOrderUnitMfg= 7
		C_OrderUnitMfg2		= 8
	End If
	
End Sub

'========================================================================================================
Sub InitVariables()
    lgIntFlgMode = Parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
    lgStrPrevKey = ""
    lgLngCurRows = 0  
    lgBlnOpenPop = False
End Sub

'=========================================================================================================
Sub SetDefaultVal()
	'공장 Default값처리 
	If Parent.gPlant <> "" And Trim(frm1.txtConPlantCd.value) = "" Then
		frm1.txtConPlantCd.value = parent.gPlant
	End If

	'영업그룹 Default값처리 
	If Parent.gSalesGrp <> "" And Trim(frm1.txtConSalesGrp.value) = "" Then
		frm1.txtConSalesGrp.value = parent.gSalesGrp
	End If

	Call GetSpConfig()
	
	If (lgLngProcessByPlant And 8192) > 0 Then
		Call ggoOper.SetReqAttr(frm1.txtConPlantCd, "N")		
	Else
		Call ggoOper.SetReqAttr(frm1.txtConSalesGrp, "N")		
	End If
	' 계획일 
	If frm1.txtConFromDt.Text = "" Then
		frm1.txtConFromDt.Text = iStrFromDt
	End If		
	'Set initial focus
	frm1.txtConFromDt.focus
End Sub

'==========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"-->
	<% Call LoadInfTB19029A("I","*","NOCOOKIE", "MA") %>
	<% Call LoadBNumericFormatA("I", "*", "NOCOOKIE", "MA") %>
End Sub

'==========================================================================================================
Sub InitSpreadSheet()
	Call initSpreadPosVariables() 
	
   	'☜:--------Spreadsheet #1-----------------------------------------------------------------------------   
   	' gMouseClickStatus = "N" : when the form is loaded
   	If gMouseClickStatus = "N" Or gMouseClickStatus = "SPCRP" Then
		With frm1.vspdData		
			
		   	'☜:--------Spreadsheet #1-----------------------------------------------------------------------------   
			ggoSpread.Source = frm1.vspdData
			'patch version
		    ggoSpread.Spreadinit "V20021214",,parent.gAllowDragDropSpread    		
			.ReDraw = false
			
			.MaxRows = 0 : .MaxCols = 0
			.MaxCols = C_OldQtyOrderUnitMfg + 1            '☜: 최대 Columns의 항상 1개 증가시킴 
		    
		    Call GetSpreadColumnPos("A")
		    ' SSSetEdit(Col, Header, ColWidth, HAlign, Row, Length, CharCase)

			ggoSpread.SSSetDate		C_SpDt,			"계획일", 10, 2, parent.gDateFormat
			ggoSpread.SSSetEdit		C_PlantCd,		"공장", 10,,,4,2
			ggoSpread.SSSetEdit		C_PlantNm,		"공장명", 18
			ggoSpread.SSSetEdit		C_SalesGrp,		"영업그룹", 10,,,4,2
			ggoSpread.SSSetEdit		C_SalesGrpNm,	"영업그룹명", 18
            ggoSpread.SSSetCombo	C_LocExpFlag,	"거래구분", 1
            ggoSpread.SSSetCombo	C_LocExpFlagNm, "거래구분", 10
			ggoSpread.SSSetEdit		C_SoldToParty,	"거래처", 18,,,10,2
			ggoSpread.SSSetEdit		C_SoldToPartyNm, "거래처명", 18
			ggoSpread.SSSetEdit		C_ItemCd,		"품목", 18,,,18,2 
			ggoSpread.SSSetEdit		C_ItemNm,		"품목명", 18
			ggoSpread.SSSetFloat	C_Qty,			"수량" ,15,Parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
			ggoSpread.SSSetEdit		C_Unit,			"단위", 8,2,,3,2
			ggoSpread.SSSetFloat	C_QtyOrderUnitMfg,"생산단위수량" ,15,Parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
			ggoSpread.SSSetEdit		C_OrderUnitMfg,	"생산단위", 8,2,,3,2
			ggoSpread.SSSetEdit		C_CfmFlag,		"확정여부", 10,2,,1,2
			ggoSpread.SSSetEdit		C_SpPeriod,		"계획기간", 10,2,,8
			ggoSpread.SSSetEdit		C_SpPeriodDesc,	"계획기간설명", 18,,,30
			ggoSpread.SSSetEdit		C_SpMonth,		"월", 10,2,,2
			ggoSpread.SSSetEdit		C_SpWeek,		"주", 10,2,,2

			ggoSpread.SSSetEdit		C_Pointer,		"", 1
			ggoSpread.SSSetEdit		C_OldQty,		"", 1
			ggoSpread.SSSetEdit		C_OldQtyOrderUnitMfg,		"", 1

		    Call ggoSpread.SSSetColHidden(C_LocExpFlag,C_LocExpFlag,True)
		    Call ggoSpread.SSSetColHidden(C_Pointer,.MaxCols,True)   '☜: 공통콘트롤 사용 Hidden Column
		    
   		    Call SetSpreadLock()

			.ReDraw = True
		End With
	End If
    
   	'☜:--------Spreadsheet #2-----------------------------------------------------------------------------   
   	If gMouseClickStatus = "N" Or gMouseClickStatus = "SP2CRP" Then
		With frm1.vspdData2		
			
		   	'☜:--------Spreadsheet #1-----------------------------------------------------------------------------   
			ggoSpread.Source = frm1.vspdData2
			'patch version
		    ggoSpread.Spreadinit "V20021214",,parent.gAllowDragDropSpread    		
			.ReDraw = false
			
			.MaxRows = 0 : .MaxCols = 0
			.MaxCols = C_OrderUnitMfg2 + 1            '☜: 최대 Columns의 항상 1개 증가시킴 
		    
		    Call GetSpreadColumnPos("B")
		    
		    ' SSSetEdit(Col, Header, ColWidth, HAlign, Row, Length, CharCase)
			ggoSpread.SSSetEdit		C_SpPeriod2,		"계획기간", 18,,,8
			ggoSpread.SSSetEdit		C_SpPeriodDesc2,	"계획기간설명", 18,,,30
			ggoSpread.SSSetEdit		C_PlantCd2,			"공장", 10,,,4,2
			ggoSpread.SSSetEdit		C_PlantNm2,			"공장명", 18
			ggoSpread.SSSetFloat	C_TotQty,			"수량" ,15,Parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
			ggoSpread.SSSetEdit		C_Unit2,			"단위", 8,2,,3,2
			ggoSpread.SSSetFloat	C_TotQtyOrderUnitMfg,"생산단위수량" ,15,Parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
			ggoSpread.SSSetEdit		C_OrderUnitMfg2,	"생산단위", 8,2,,3,2

		    Call ggoSpread.SSSetColHidden(.MaxCols,.MaxCols,True)   '☜: 공통콘트롤 사용 Hidden Column

		    ' Lock the sheet
		    Call SetSpreadLock2()
		    .OperationMode = 3
			.ReDraw = True
		End With
	End If
End Sub

'==========================================================================================================
Sub SetSpreadLock()
	ggoSpread.SpreadLock C_SpDt, -1, C_ItemNm
	ggoSpread.SpreadLock C_Unit, -1
End Sub

Sub SetSpreadLock2()
	ggoSpread.SpreadLock 1, -1
End Sub

'==========================================================================================================
' After query
Sub SetQuerySpreadColor(ByVal pvStartRow)
	Dim iLngIndex
	
	ggoSpread.Source = frm1.vspdData
	With frm1.vspdData
		For iLngIndex = pvStartRow to .MaxRows
			.Row = iLngIndex
			.Col = C_CfmFlag
			If .Text = "N" Then
				ggoSpread.SSSetRequired  C_Qty , iLngIndex, iLngIndex
			Else
				ggoSpread.SSSetProtected  C_Qty , iLngIndex, iLngIndex
			End If
		Next
	End With
	
End Sub

'================================== 3.1.8 SubSetErrPos() ==================================================
' Desc : This method set focus to position of error
'      : This method is called in MB area
'==========================================================================================================
Sub SubSetErrPos(iPosArr)
    Dim iDx
    Dim iRow
    iPosArr = Split(iPosArr,Parent.gColSep)
    If IsNumeric(iPosArr(0)) Then
       iRow = CInt(iPosArr(0))
       For iDx = 1 To  frm1.vspdData.MaxCols - 1
           Frm1.vspdData.Col = iDx
           Frm1.vspdData.Row = iRow

           If Not Frm1.vspdData.ColHidden Then
              Frm1.vspdData.Col = iDx
              Frm1.vspdData.Row = iRow
              Frm1.vspdData.Action = 0 ' go to 
              Exit For
           End If
           
       Next
          
    End If   
End Sub

'==========================================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos

    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

			C_SpDt				= iCurColumnPos(1)
			C_PlantCd			= iCurColumnPos(2)
			C_PlantNm			= iCurColumnPos(3)
			C_SalesGrp			= iCurColumnPos(4)
			C_SalesGrpNm		= iCurColumnPos(5)
			C_LocExpFlag		= iCurColumnPos(6)
			C_LocExpFlagNm		= iCurColumnPos(7)
			C_SoldToParty		= iCurColumnPos(8)
			C_SoldToPartyNm		= iCurColumnPos(9)
			C_ItemCd			= iCurColumnPos(10)
			C_ItemNm			= iCurColumnPos(11)
			C_Qty				= iCurColumnPos(12)
			C_Unit				= iCurColumnPos(13)
			C_QtyOrderUnitMfg	= iCurColumnPos(14)
			C_OrderUnitMfg		= iCurColumnPos(15)
			C_CfmFlag			= iCurColumnPos(16)
			C_SpPeriod			= iCurColumnPos(17)
			C_SpPeriodDesc		= iCurColumnPos(18)
			C_SpMonth			= iCurColumnPos(19)
			C_SpWeek			= iCurColumnPos(20)

			C_Pointer			= iCurColumnPos(21)
			C_OldQty			= iCurColumnPos(22)
			C_OldQtyOrderUnitMfg= iCurColumnPos(23)
			
       Case "B"
            ggoSpread.Source = frm1.vspdData2
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
            
			C_SpPeriod2			= iCurColumnPos(1)
			C_SpPeriodDesc2		= iCurColumnPos(2)
			C_PlantCd2			= iCurColumnPos(3)
			C_PlantNm2			= iCurColumnPos(4)
			C_TotQty			= iCurColumnPos(5)
			C_Unit2				= iCurColumnPos(6)
			C_TotQtyOrderUnitMfg= iCurColumnPos(7)
			C_OrderUnitMfg2		= iCurColumnPos(8)
    End Select    
End Sub

'==========================================================================================================
Sub InitComboBox()
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
		'거래구분 
	Call CommonQueryRs(" MINOR_CD,MINOR_NM", " B_MINOR ", " MAJOR_CD=" & FilterVar("S4225", "''", "S") & " ", _
    	               lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)    
	Call SetCombo2(frm1.cboConLocExpFlag, lgF0,lgF1, parent.gColSep)

	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub

'	Name : InitSpreadComboBox()
Sub InitSpreadComboBox()
	Dim iStrCboData    ''lgF0
	Dim iStrCboData2    ''lgF1

	Call CommonQueryRs(" MINOR_CD,MINOR_NM ", " B_MINOR ", " MAJOR_CD = " & FilterVar("S4225", "''", "S") & " ", _	
    	               lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    
	iStrCboData = Replace(lgF0,chr(11),vbTab)
    iStrCboData2 = Replace(lgF1,chr(11),vbTab)
    iStrCboData = Left(iStrCboData,Len(iStrCboData) - 1)
    iStrCboData2 = Left(iStrCboData2,Len(iStrCboData2) - 1)
    
    ggoSpread.Source = frm1.vspdData
	ggoSpread.SetCombo iStrCboData,  C_LocExpFlag
	ggoSpread.SetCombo iStrCboData2, C_LocExpFlagNM

End Sub

'==========================================================================================================
Function CookiePage(Byval pvKubun)

	On Error Resume Next
	Const CookieSplit = 4877      <%'Cookie Split String : CookiePage Function Use%>
	Dim iStrTemp, iArrVal
	Dim iStrFromDt, iStrToDt

	With frm1
		If pvKubun = 1 Then
			WriteCookie CookieSplit , .txtConSalesGrp.value & Parent.gColSep & .txtConSalesGrpNm.value & Parent.gColSep & _
									  .txtConPlantCd.value & parent.gColSep & .txtConPlantNm.value
		ElseIf pvKubun = 0 Then
			iStrTemp = ReadCookie(CookieSplit)
			
			If Trim(Replace(iStrTemp, Parent.gColSep, "")) = "" then Exit Function
			
			iArrVal = Split(iStrTemp, Parent.gColSep)

			.txtConSalesGrp.value	= iArrVal(0)
			.txtConSalesGrpNm.value = iArrVal(1)
			.txtConPlantCd.value	= iArrVal(2)
			.txtConPlantNm.value	= iArrVal(3)
			If GetSpPeriodInfo(iArrVal(4), iArrVal(6), iStrFromDt, iStrToDt) Then
				.txtConFromDt.Text	= iStrFromDt
				If iArrVal(6) <> "" Then
					.txtConToDt.Text	= iStrToDt
				End If
			End If
			WriteCookie CookieSplit , ""
		End If
	End With
End Function

'==========================================================================================================
Function JumpChgCheck(byVal pvStrJumpPgmId)

	Dim IntRetCD

	'************ 멀티인 경우 **************
	ggoSpread.Source = frm1.vspdData 
	If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900017", Parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If

	Call CookiePage(1)
	Call PgmJump(pvStrJumpPgmId)

End Function

'==========================================================================================================
Sub Form_Load()
on error resume next
	Call LoadInfTB19029             '⊙: Load table , B_numeric_format
	Call ggoOper.FormatField(Document, "A",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec) '⊙: Format Contents  Field
	Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field
	'----------  Coding part  -------------------------------------------------------------
	Call InitComboBox()
	Call InitSpreadSheet
	Call InitSpreadComboBox()
	Call CookiePage(0)
	Call SetDefaultVal    
	Call InitVariables              '⊙: Initializes local global variables
	If lgBlnExistsSpConfig Then
		If (lgLngUseStep And 4096) = 0 Then
			Call DisplayMsgBox("202415", "X", "", "")
			Call SetToolbar("10000000000011")          '⊙: 버튼 툴바 제어 
		Else
			Call SetToolbar("11000000000011")          '⊙: 버튼 툴바 제어 
		End If
	Else
		Call SetToolbar("10000000000011")          '⊙: 버튼 툴바 제어 
	End If
End Sub

'==========================================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )
End Sub

'==========================================================================================================
Function FncQuery() 
    Dim IntRetCD 
    FncQuery = False                                  <%'⊙: Processing is NG%>
    
    Err.Clear             
                                                      <%'☜: Protect system from crashing%>
    '-----------------------
    'Check previous data area
    '-----------------------
	ggoSpread.Source = frm1.vspdData
	If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO, "X", "X")
		'IntRetCD = MsgBox("데이타가 변경되었습니다. 조회하시겠습니까?", vbYesNo)
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If

	'-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then         <%'⊙: This function check indispensable field%>
       Exit Function
    End If

	'-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")          <%'⊙: Clear Contents  Field%>
    Call ggoSpread.ClearSpreadData()
    Call InitVariables               <%'⊙: Initializes local global variables%>
	'-----------------------
    'Query function call area
    '-----------------------
    Call DbQuery                <%'☜: Query db data%>

    FncQuery = True                <%'⊙: Processing is OK%>
        
End Function

'========================================================================================================
Function FncSave() 
    Dim IntRetCD 
    
    FncSave = False                                                         <%'⊙: Processing is NG%>
    
    Err.Clear                                                               <%'☜: Protect system from crashing%>

	ggoSpread.Source = frm1.vspdData 
	If ggoSpread.SSCheckChange = False Then
		IntRetCD = DisplayMsgBox("900001", "X", "X", "X")
		'Call MsgBox("No data changed!!", vbInformation)
		Exit Function
	End If
    
<%  '-----------------------
    'Check content area
    '-----------------------%>
    If Not chkField(Document, "2") Then     <%'⊙: Check contents area%>
       Exit Function
    End If
    If ggoSpread.SSDefaultCheck = False Then     <%'⊙: Check contents area%>
       Exit Function
    End If

<%  '-----------------------
    'Save function call area
    '-----------------------%>
    CAll  DbSave                                                   <%'☜: Save db data%>
    
    FncSave = True                                                          <%'⊙: Processing is OK%>
    
End Function

'========================================================================================================
Function FncCancel() 
	On Error Resume Next
	Dim iDblNewQty, iDblOldQty
	Dim iDblNewQtyOrderUnitMfg, iDblOldQtyOrderUnitMfg
	
	If frm1.vspdData.MaxRows < 1 Then Exit Function

    ggoSpread.Source = frm1.vspdData 
	
	With frm1.vspdData
		.Row = .ActiveRow
		.Col = C_OldQty				: iDblOldQty = UNICDbl(.Text)
		.Col = C_OldQtyOrderUnitMfg	: iDblOldQtyOrderUnitMfg = UNICDbl(.Text)
		.Col = 0
	    
		Select Case	.Text
			Case ggoSpread.UpdateFlag
			    ggoSpread.EditUndo
				.Col = C_Qty				:	iDblNewQty = UNICDbl(.Text)
				.Col = C_QtyOrderUnitMfg	:	iDblNewQtyOrderUnitMfg = UNICDbl(.Text)
				
				Call ReCalcSpread2(.ActiveRow, iDblNewQty - iDblOldQty, iDblNewQtyOrderUnitMfg - iDblOldQtyOrderUnitMfg)
		End Select

	End With

End Function

'========================================================================================================
Function FncPrint() 
 Call parent.FncPrint()
End Function

'========================================================================================================
Function FncExcel() 
 On Error Resume Next                                                             '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncExcel = False                                                              '☜: Processing is NG

    'Call FncExport(parent.C_MULTI)            '☜: 화면 유형 

	Call parent.FncExport(Parent.C_SINGLEMULTI)	                     			  '☜: 화면 유형 

    If Err.number = 0 Then	 
       FncExcel = True                                                            '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement   
End Function

'========================================================================================================
Function FncFind()
	On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncFind = False                                                               '☜: Processing is NG

'    Call parent.FncFind(parent.C_MULTI, False)                                         '☜:화면 유형, Tab 유무 
     
    Call parent.FncFind(Parent.C_SINGLEMULTI, False)                              '☜:화면 유형, Tab 유무 
    
    If Err.number = 0 Then	 
       FncFind = True                                                             '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement   
End Function

'========================================================================================================
Sub FncSplitColumn()
	If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
	   Exit Sub
	End If

	ggoSpread.Source = gActiveSpdSheet
	ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   
    
End Sub

'========================================================================================================
Function FncExit()
 Dim IntRetCD
 FncExit = False

	ggoSpread.Source = frm1.vspdData 
	If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO, "X", "X")
		'IntRetCD = MsgBox("데이타가 변경되었습니다. 종료 하시겠습니까?", vbYesNo)
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If

	FncExit = True
End Function

'========================================================================================================
Function DbQuery() 
	Err.Clear                                                               <%'☜: Protect system from crashing%>
	    
	DbQuery = False                                                         <%'⊙: Processing is NG%>
	If  LayerShowHide(1) = False Then
		Exit Function 
	End If
	    
	Dim iStrVal
	With Frm1
		iStrVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001         <%'☜: 비지니스 처리 ASP의 상태 %>

	    If lgIntFlgMode = Parent.OPMD_UMODE Then    
			' Scroll
			iStrVal = iStrVal & lgStrWhere
		Else
			' Initial query
			lgStrWhere = "&txtWhere="
			lgStrWhere = lgStrWhere & UNIConvDate(.txtConFromDt.Text) & parent.gColSep' Sales Plan Date
			lgStrWhere = lgStrWhere & Trim(.txtConPlantCd.value) & parent.gColSep			' Plant
			lgStrWhere = lgStrWhere & Trim(.txtConSalesGrp.value) & parent.gColSep			' Sales Group
			lgStrWhere = lgStrWhere & Trim(.txtConItemCd.value) & parent.gColSep			' Item Code
			lgStrWhere = lgStrWhere & Trim(.txtConSoldToParty.value) & parent.gColSep		' Slod to party
			lgStrWhere = lgStrWhere & Trim(.cboConLocExpFlag.value) & parent.gColSep		' Local/Export Flag
			If .txtConToDt.Text <> "" Then
				lgStrWhere = lgStrWhere & UNIConvDate(.txtConToDt.Text) & parent.gColSep	' Sales Plan Date
			Else
				lgStrWhere = lgStrWhere & "" & parent.gColSep
			End If

			iStrVal = iStrVal & lgStrWhere
		End If 
		iStrVal = iStrVal & "&lgStrPrevKey=" & lgStrPrevKey
		iStrVal = iStrVal & "&txtLastRow=" & frm1.vspdData.MaxRows

		lgLngStartRow = frm1.vspdData.MaxRows + 1
	End With
	Call RunMyBizASP(MyBizASP, iStrVal)            <%'☜: 비지니스 ASP 를 가동 %>

	DbQuery = True   
                          <%'⊙: Processing is NG%>
End Function

'========================================================================================================
Function DbQueryOk()              <%'☆: 조회 성공후 실행로직 %>
    '-----------------------
    'Reset variables area
    '-----------------------
    If lgIntFlgMode <> Parent.OPMD_UMODE Then
		lgIntFlgMode = Parent.OPMD_UMODE            <%'⊙: Indicates that current mode is Update mode%>
	End If
    
	If Trim(lgStrPrevKey) = "" Then
		lgStrWhere = ""
    End If

	Call SetQuerySpreadColor(lgLngStartRow)
	
	Call SetToolbar("11001001000111")

	If frm1.vspdData.MaxRows > 0 Then	 
		frm1.vspdData.focus
	Else
		Call SetFocusToDocument("M")
		frm1.txtConFromDt.focus
	End If

End Function

'========================================================================================================
Function DbSave() 

	Err.Clear                <%'☜: Protect system from crashing%>
	 
	Dim iStrIns, iStrUpd, iStrDel, iStrKey
	Dim iLngRow
		 
	DbSave = False                                                          '⊙: Processing is NG
		    
	On Error Resume Next                                                   '☜: Protect system from crashing
		   
	If LayerShowHide(1) = False Then
		Exit Function 
	End If

  '-----------------------
  'Data manipulate area
  '-----------------------
  iStrInt = ""
  iStrUpd = ""
  iStrDel = ""
   
  '-----------------------
  'Data manipulate area
  '-----------------------
	With frm1.vspdData
		For iLngRow = 1 To .MaxRows
    
			.Row = iLngRow
			.Col = 0

			if .Text <> "" Then
				iStrKey = CStr(iLngRow) & Parent.gColSep		' Row No.
				.Col = C_SpDt			' Sales Plan Period(PK)
				iStrKey = iStrKey & .Text & Parent.gColSep
				.Col = C_PlantCd		' Plant Cd(PK)
				iStrKey = iStrKey & .Text & Parent.gColSep
				.Col = C_SalesGrp       ' Sales Group(PK)
				iStrKey = iStrKey & .Text & Parent.gColSep
				.Col = C_SoldToParty    ' Slod to party(PK)
				iStrKey = iStrKey & .Text & Parent.gColSep
				.Col = C_ItemCd			' Item Code(PK)
				iStrKey = iStrKey & .Text & Parent.gColSep
				.Col = C_LocExpFlag		' Local/Export Flag(PK)
				iStrKey = iStrKey & .Text & Parent.gColSep
				
				.Col = 0
				Select Case .Text
					Case ggoSpread.UpdateFlag       '☜: 수정 

						iStrUpd = iStrUpd & iStrKey
						
						.Col = C_Qty				' Quantity
						iStrUpd = iStrUpd & UNIConvNum(.Text,0) & Parent.gColSep
						
						.Col = C_QtyOrderUnitMfg	' Quantity
						iStrUpd = iStrUpd & UNIConvNum(.Text,0) & Parent.gColSep
						
						iStrUpd = iStrUpd & Parent.gUsrID & Parent.gColSep & Parent.gRowSep
				End Select
			End If
		Next
	End With
 
	With frm1
	  .txtMode.value = Parent.UID_M0002
	  .txtSpreadIns.value = iStrIns
	  .txtSpreadUpd.value = iStrUpd
	  .txtSpreadDel.value = iStrDel
	End With

 	Call ExecMyBizASP(frm1, BIZ_PGM_ID)         '☜: 비지니스 ASP 를 가동 
 
    DbSave = True                                                           '⊙: Processing is NG
    
End Function

'========================================================================================================
Function DbSaveOk()               <%'☆: 저장 성공후 실행 로직 %>
	Call ggoOper.ClearField(Document, "2")
    Call MainQuery()
End Function

'========================================================================================================
Function OpenConPopup(ByVal pvIntWhere)
	Dim iArrRet
	Dim iArrParam(5), iArrField(6), iArrHeader(6)

	OpenConPopup = False
	
	If lgBlnOpenPop Then Exit Function

	lgBlnOpenPop = True
	
	Select Case pvIntWhere
	Case C_PopPlantCd
		iArrParam(1) = "B_PLANT"							<%' TABLE 명칭 %>
		iArrParam(2) = Trim(frm1.txtConPlantCd.value)		<%' Code Condition%>
		iArrParam(3) = ""									<%' Name Cindition%>
		iArrParam(4) = ""									<%' Where Condition%>
		iArrParam(5) = "공장"							<%' TextBox 명칭 %>
		
		iArrField(0) = "ED15" & Parent.gColSep & "PLANT_CD"	<%' Field명(0)%>
		iArrField(1) = "ED30" & Parent.gColSep & "PLANT_NM"	<%' Field명(1)%>
    
	    iArrHeader(0) = "공장"							<%' Header명(0)%>
	    iArrHeader(1) = "공장명"						<%' Header명(1)%>

		frm1.txtConPlantCd.focus 

	Case C_PopSalesGrp												
		iArrParam(1) = "B_SALES_GRP"
		iArrParam(2) = Trim(frm1.txtConSalesGrp.value)
		iArrParam(3) = ""
		iArrParam(4) = "USAGE_FLAG=" & FilterVar("Y", "''", "S") & " "
		iArrParam(5) = "영업그룹"
		
		iArrField(0) = "ED15" & Parent.gColSep & "SALES_GRP"
		iArrField(1) = "ED30" & Parent.gColSep & "SALES_GRP_NM"
    
	    iArrHeader(0) = "영업그룹"
	    iArrHeader(1) = "영업그룹명"

		frm1.txtConSalesGrp.focus 

	Case C_PopItemCd
		OpenConPopup = OpenConItemPopup(C_PopItemCd, frm1.txtConItemCd.value)
		frm1.txtConItemCd.focus
		Exit Function

	Case C_PopSoldToParty												
		iArrParam(1) = "B_BIZ_PARTNER BP"
		iArrParam(2) = Trim(frm1.txtConSoldToParty.value)
		iArrParam(3) = ""
		iArrParam(4) = "BP.USAGE_FLAG=" & FilterVar("Y", "''", "S") & "  AND BP.BP_TYPE LIKE " & FilterVar("C%", "''", "S") & " "
		iArrParam(5) = "거래처"
			
		iArrField(0) = "ED15" & Parent.gColSep & "BP.BP_CD"
		iArrField(1) = "ED30" & Parent.gColSep & "BP.BP_NM"
		iArrField(2) = "ED8" & Parent.gColSep & "BP.CURRENCY"
		    
		iArrHeader(0) = "거래처"
		iArrHeader(1) = "거래처명"
		iArrHeader(2) = "화폐"

		frm1.txtConSoldToParty.focus

	End Select
 
	iArrParam(0) = iArrParam(5)							<%' 팝업 명칭 %> 

	iArrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(iArrParam, iArrField, iArrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgBlnOpenPop = False
	
	If iArrRet(0) = "" Then
		Exit Function
	Else
		OpenConPopup = SetConPopup(iArrRet,pvIntWhere)
	End If	

End Function

' Sales planning period Popup
Function OpenConSpPeriodPopup(ByVal pvIntWhere, ByVal pvStrData)
	Dim iArrRet
	Dim iArrParam(2)
	Dim iCalledAspName

	OpenConSpPeriodPopup = False

	iCalledAspName = AskPRAspName("s2211pa3")
	
	If Trim(iCalledAspName) = "" Then
		Call DisplayMsgBox("900040", parent.VB_INFORMATION, "s2211pa3", "X")
		lgBlnOpenPop = False
		Exit Function
	End If

	iArrParam(0) = pvStrData
	
	iArrRet = window.showModalDialog(iCalledAspName & "?txtDisplayFlag=N", Array(window.parent,iArrParam), _
	 "dialogWidth=690px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgBlnOpenPop = False

	If iArrRet(0) = "" Then
		Exit Function
	Else
		OpenConSpPeriodPopup = SetConPopup(iArrRet,pvIntWhere)
	End If	
End Function

' Item Popup
Function OpenConItemPopup(ByVal pvIntWhere, ByVal pvStrData)
	Dim iArrRet
	Dim iArrParam(3)
	Dim iCalledAspName

	OpenConItemPopup = False

	iCalledAspName = AskPRAspName("s2210pa1")
	
	If Trim(iCalledAspName) = "" Then
		Call DisplayMsgBox("900040", parent.VB_INFORMATION, "s2210pa1", "X")
		lgBlnOpenPop = False
		Exit Function
	End If

	iArrParam(0) = pvStrData
	iArrParam(3) = Trim(frm1.txtConPlantCd.value)
	
	iArrRet = window.showModalDialog(iCalledAspName, Array(window.parent,iArrParam), _
	 "dialogWidth=850px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgBlnOpenPop = False

	If iArrRet(0) = "" Then
		Exit Function
	Else
		OpenConItemPopup = SetConPopup(iArrRet,pvIntWhere)
	End If	
End Function

'===========================================================================================================
Function SetConPopup(Byval pvArrRet,ByVal pvIntWhere)
	SetConPopup = False

	With frm1
		Select Case pvIntWhere
		Case C_PopPlantCd
			.txtConPlantCd.value = pvArrRet(0) 
			.txtConPlantNm.value = pvArrRet(1)   
			
		Case C_PopSalesGrp
			.txtConSalesGrp.value = pvArrRet(0) 
			.txtConSalesGrpNm.value = pvArrRet(1)   
			
		Case C_PopItemCd
			frm1.txtConItemCd.value = pvArrRet(0) 
			frm1.txtConItemNm.value = pvArrRet(1)   

		Case C_PopSoldToParty
			frm1.txtConSoldToParty.value = pvArrRet(0) 
			frm1.txtConSoldToPartyNm.value = pvArrRet(1)   

		End Select
	End With

	SetConPopup = True
End Function

'=============================================  SetRowStatus() =============================================
'   Event Desc : Update the Row Status
'===========================================================================================================
Sub SetRowStatus(intRow)
    ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow intRow

 lgBlnFlgChgValue = True
End Sub


'========================================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub

'========================================================================================================
Sub PopRestoreSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()      

	Call ggoSpread.ReOrderingSpreadData()

	 '------ Developer Coding part (Start ) --------------------------------------------------------------
	If gMouseClickStatus = "SPCRP" Then	SetQuerySpreadColor(1)

    '------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub

'========================================================================================================
Sub vspdData_Click(ByVal Col , ByVal Row )
	' Context 메뉴의 입력, 삭제, 데이터 입력, 취소 
	Call SetPopupMenuItemInf(Mid(gToolBarBit, 6, 2) + "0" + Mid(gToolBarBit, 8, 1) & "111111")
    gMouseClickStatus = "SPC"
    Set gActiveSpdSheet = frm1.vspdData
    
    If frm1.vspdData.MaxRows = 0 Then 
		Exit Sub
	End If  
	   
	If Row <= 0 Then
		ggoSpread.Source = frm1.vspdData
		
		If lgSortKey = 1 Then
			ggoSpread.SSSort Col				'Sort in Ascending
			lgSortkey = 2
		Else
			ggoSpread.SSSort Col, lgSortKey		'Sort in Descending
			lgSortkey = 1
		End If
		
		Exit Sub
	End If    	
End Sub

Sub vspdData2_Click(ByVal Col, ByVal Row)

    'Call SetPopupMenuItemInf("0000111111")
	gMouseClickStatus = "SP2C"	'Split 상태코드 
	   
    Set gActiveSpdSheet = frm1.vspdData2
    
    ' spread1에서 spread2의 Pointer 갖고 있어 spread2의 정렬은 disalbe 시킴 
    Exit Sub

    If frm1.vspdData2.MaxRows = 0 Then 
		Exit Sub
	End If  
	   
	If Row <= 0 Then
		ggoSpread.Source = frm1.vspdData2
		
		If lgSortKey = 1 Then
			ggoSpread.SSSort Col				'Sort in Ascending
			lgSortkey = 2
		Else
			ggoSpread.SSSort Col, lgSortKey	'Sort in Descending
			lgSortkey = 1
		End If
		
		Exit Sub
	End If    	
End Sub

'========================================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )
	Dim iStrData
	
	ggoSpread.Source = frm1.vspdData

	With frm1.vspdData
		.Row = Row
		.Col = 0
		If .Text = ggoSpread.DeleteFlag Then Exit Sub
		
		CALL SetRowStatus(Row)

		.Col = Col	: iStrData = .Text
		
		If iStrData = "" Then Exit Sub
		
		Select Case Col
			Case C_Qty
				Call CalcQtyOrderUnitMfg(Row, iStrData)

		End Select
	End With

End Sub

'==========================================================================================
'   Event Desc : Combo 변경 
'==========================================================================================
Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
	Dim iIntIndex

	With frm1.vspdData
		.Row = Row
		
		Select Case Col
			Case C_LocExpFlagNm
				.Col = Col			:	iIntIndex = .Value
				.Col = C_LocExpFlag	:	.Value = iIntIndex
		End Select
	End With
End Sub

'========================================================================================================
Sub vspdData_LeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
End Sub

'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		

    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

End Sub

Sub vspdData2_ColWidthChange(ByVal Col1, ByVal Col2)
   ggoSpread.Source = frm1.vspdData2
  Call ggoSpread.SSSetColWidth(Col1,Col2)

End Sub

'========================================================================================================
Sub vspdData_GotFocus()
    ggoSpread.Source = Frm1.vspdData

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub

'========================================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)
    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If

End Sub

Sub vspdData2_MouseDown(Button , Shift , x , y)

 If Button = 2 And gMouseClickStatus = "SP2C" Then
  gMouseClickStatus = "SP2CR"
 End If

End Sub

'========================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
End Sub

Sub vspdData2_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    
    Call GetSpreadColumnPos("B")
End Sub

'========================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
	If OldLeft <> NewLeft Then	Exit Sub
	    
	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) And lgStrPrevKey <> "" Then
		If CheckRunningBizProcess = True Then Exit Sub
	    
		Call DisableToolBar(Parent.TBC_QUERY)
		If DBQuery = False Then
			Call RestoreToolBar()
			Exit Sub
		End If
	End if    

End Sub

'========================================================================================================
Sub txtConFromDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtConFromDt.Action = 7		
		Call SetFocusToDocument("M")   
		Frm1.txtConFromDt.Focus
	End If
End Sub

Sub txtConToDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtConToDt.Action = 7
		Call SetFocusToDocument("M")   
		Frm1.txtConToDt.Focus
	End If
End Sub

Sub txtConFromDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13 Then Call MainQuery()
End Sub

Sub txtConToDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13 Then Call MainQuery()
End Sub

<%'======================================   GetSpConfig()  =====================================
'	Description : 판매계획환경정보를 Fetch한다.
'==================================================================================================== %>
Sub GetSpConfig()

	Dim iStrSelectList, iStrFromList, iStrWhereList
	Dim iStrRs
	Dim iArrRs
	
	iStrSelectList = " USE_STEP, PROCESS_BY_PLANT "
	iStrFromList = " dbo.S_SP_CONFIG "
	iStrWhereList = "SP_TYPE = " & FilterVar("E", "''", "S") & " "
	
	Err.Clear
	
	If CommonQueryRs2by2(iStrSelectList,iStrFromList,iStrWhereList,iStrRs) Then
		iArrRs = Split(iStrRs, parent.gColSep)
		lgLngUseStep = CLng(iArrRs(1))
		lgLngProcessByPlant = CLng(iArrRs(2))
		lgBlnExistsSpConfig = True
	Else
		'판매계획환경설정 정보가 없습니다.
		Call DisplayMsgBox("202403", "X", "", "")
		lgBlnExistsSpConfig = False
	End if
End Sub

'=============================================== GetSpPeriodInfo() =============================================
' Description : 판매계획기간 정보에 대한 날짜 정보를 Fetch한다.
'===========================================================================================================
Function GetSpPeriodInfo(ByVal pvStrFromSpPeriod, ByVal pvStrToSpPeriod, ByRef prStrFromDt, ByRef prStrToDt)
	Dim iStrSelectList, iStrFromList, iStrWhereList
	Dim iStrRs
	Dim iArrSpPeriodInfo
	
	GetSpPeriodInfo = False
	
	If pvStrFromSpPeriod = "" Then Exit function
	iStrSelectList = " MIN(FROM_DT), MAX(TO_DT) "
	iStrFromList   = " dbo.S_SP_PERIOD_INFO "
	
	If pvStrToSpPeriod <> "" Then
		iStrWhereList  = " SP_TYPE = " & FilterVar("E", "''", "S") & "  AND (SP_PERIOD =  " & FilterVar(pvStrFromSpPeriod , "''", "S") & " OR SP_PERIOD =  " & FilterVar(pvStrToSpPeriod , "''", "S") & ")" 
	Else
		iStrWhereList  = " SP_TYPE = " & FilterVar("E", "''", "S") & "  AND SP_PERIOD =  " & FilterVar(pvStrFromSpPeriod , "''", "S") & "" 
	End If

	Err.Clear
	    
	If CommonQueryRs2by2(iStrSelectList,iStrFromList,iStrWhereList,iStrRs) Then
		iArrSpPeriodInfo = Split(iStrRs, Chr(11))
		prStrFromDt = UNIDateClientFormat(Trim(iArrSpPeriodInfo(1)))
		prStrToDt = UNIDateClientFormat(Trim(iArrSpPeriodInfo(2)))
		GetSpPeriodInfo = True
	Else
		If Err.number <> 0 Then
			MsgBox Err.description 
			Err.Clear
		End If
	End If

End Function

'=========================================== CalcQtyOrderUnitMfg() =========================================
' Description : 수량변경시 재고단위 수량을 재계산한다.
'===========================================================================================================
Sub CalcQtyOrderUnitMfg(ByVal pvLngRow, ByVal pvStrData)
	Dim iDblNewQty, iDblOldQty, iDblNewQtyOrderUnitMfg, iDblOldQtyOrderUnitMfg
	Dim iStrNewQtyOrderUnitMfg, iStrOldQty
	Dim iStrUnit, iStrOrderUnitMfg
	Dim iStrSelectList, iStrFromList, iStrWhereList
	Dim iStrRs
	Dim iArrQty
	
	With frm1.vspdData
		.Row = pvLngRow
		.Col = C_Unit				:	iStrUnit = .Text
		.Col = C_OldQty				:	iStrOldQty = .Text		:	iDblOldQty = UNICDbl(.Text)
		.Col = C_OrderUnitMfg		:	iStrOrderUnitMfg = .Text
		.Col = C_OldQtyOrderUnitMfg	:	iDblOldQtyOrderUnitMfg = UNICDbl(.Text)
		
		iDblNewQty = UNICDbl(pvStrData)
		If iStrUnit = iStrOrderUnitMfg Then
			iDblNewQtyOrderUnitMfg = iDblNewQty
			iStrNewQtyOrderUnitMfg = pvStrData
		Else
			.Col = C_ItemCd
			
			iStrSelectList = " ISNULL(ROUND(" & CStr(iDblNewQty) & " * CONV_FACTOR + ISNULL(UR.ROUNDING_UNIT, 0.1) * (-5), ISNULL(UR.DECIMALS, 0)), 0) "
			iStrFromList   = " (SELECT dbo.ufn_s_GetUnitConversionFactor( " & FilterVar(.Text, "''", "S") & ",  " & FilterVar(iStrUnit, "''", "S") & ", " & FilterVar(iStrOrderUnitMfg, "''", "S") & ") AS CONV_FACTOR,  " & FilterVar(iStrOrderUnitMfg, "''", "S") & " AS ORDER_UNIT_MFG) T "
			iStrFromList   = iStrFromList & " LEFT OUTER JOIN dbo.S_SP_UNIT_ROUNDING_POLICY UR ON (UR.UNIT = T.ORDER_UNIT_MFG) "
			iStrWhereList  = ""

			If CommonQueryRs2by2(iStrSelectList,iStrFromList,iStrWhereList,iStrRs) Then
				iArrQty = Split(iStrRs, parent.gColSep)
				iDblNewQtyOrderUnitMfg = CDbl(iArrQty(1))
				If iDblNewQty <> 0 And iDblNewQtyOrderUnitMfg = 0 Then
					.Col = C_Qty	: .Text = iStrOldQty
					Call DisplayMsgBox("123800", "X", "X", "X")
					Exit Sub
				Else
					iStrNewQtyOrderUnitMfg = UNIFormatNumber(iDblNewQtyOrderUnitMfg,ggQty.DecPoint,-2,0,ggQty.RndPolicy,ggQty.RndUnit)	
				End If
			Else
				If Err.number <> 0 Then
					.Col = C_Qty	: .Text = iStrOldQty
					MsgBox Err.description 
					Err.Clear
					Exit Sub
				End If
			End If
		End If
		
		If iDblOldQty <> iDblNewQty Then
			Call ReCalcSpread2(pvLngRow, iDblNewQty - iDblOldQty, iDblNewQtyOrderUnitMfg - iDblOldQtyOrderUnitMfg)
			.Col = C_OldQty				: .Text = pvStrData
			.Col = C_QtyOrderUnitMfg	: .Text = iStrNewQtyOrderUnitMfg
			.Col = C_OldQtyOrderUnitMfg	: .Text = iStrNewQtyOrderUnitMfg
		End If
	End With

End Sub

'=============================================== ReCalcSpread2() =============================================
' Description : 집계 Spread 수량 재계산 
'===========================================================================================================
Sub ReCalcSpread2(ByVal pvLngRow, ByVal pvDblQty, ByVal pvDblQtyOrderUnitMfg)
	Dim iStrSpPeriod, iStrPlantCd, iStrUnit, iStrOrderUnitMfg, iStrPointer
	Dim iStrSpPeriod2, iStrPlantCd2, iStrUnit2, iStrOrderUnitMfg2
	Dim iLngRow
	With frm1.vspdData
		.Row = pvLngRow
		.Col = C_Pointer	: iStrPointer = Trim(.Text)
		
		If iStrPointer <> "" Then
			With frm1.vspdData2
				.Row = CLng(iStrPointer)
				
				.Col = C_TotQty
				.Text = UNIFormatNumber(UNICDbl(.Text) + pvDblQty,ggQty.DecPoint,-2,0,ggQty.RndPolicy,ggQty.RndUnit)
				.Col = C_TotQtyOrderUnitMfg
				.Text = UNIFormatNumber(UNICDbl(.Text) + pvDblQtyOrderUnitMfg,ggQty.DecPoint,-2,0,ggQty.RndPolicy,ggQty.RndUnit)
			End With
		Else
			.Col = C_SpPeriod		:	iStrSpPeriod = Trim(.Text)
			.Col = C_PlantCd		:	iStrPlantCd = Trim(.Text)
			.Col = C_Unit			:	iStrUnit = Trim(.Text)
			.Col = C_OrderUnitMfg	:	iStrOrderUnitMfg = Trim(.Text)
			With frm1.vspdData2
				For iLngRow = 1 To .MaxRows
					.Row = iLngRow
					.Col = C_SpPeriod2		: 	iStrSpPeriod2 = Trim(.Text)
					.Col = C_PlantCd2		:	iStrPlantCd2 = Trim(.Text)
					.Col = C_Unit2			:	iStrUnit2 = Trim(.Text)
					.Col = C_OrderUnitMfg2	:	iStrOrderUnitMfg2 = Trim(.Text)

					If iStrSpPeriod = iStrSpPeriod2 And iStrPlantCd = iStrPlantCd2 And iStrUnit = iStrUnit2 And iStrOrderUnitMfg = iStrOrderUnitMfg2 Then
						.Col = C_TotQty
						.Text = UNIFormatNumber(UNICDbl(.Text) + pvDblQty,ggQty.DecPoint,-2,0,ggQty.RndPolicy,ggQty.RndUnit)
						.Col = C_TotQtyOrderUnitMfg
						.Text = UNIFormatNumber(UNICDbl(.Text) + pvDblQtyOrderUnitMfg,ggQty.DecPoint,-2,0,ggQty.RndPolicy,ggQty.RndUnit)
						
						iStrPointer = CStr(iLngRow)
						Exit For
					End If
				Next
			End With
			
			' Set the Pointer
			.Col = C_Pointer
			.Text = iStrPointer
		End If
	End With
End Sub

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->
</HEAD>

<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="post">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR>
		<TD <%=HEIGHT_TYPE_00%>></TD>
	</TR>
	
	<TR HEIGHT=23>
		<TD WIDTH="100%">
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><IMG src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23" ></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>공장별일별품목판매계획조정</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><IMG src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23" ></td>
							</TR>
						</TABLE>
					</TD>
					<TD WIDTH=*>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	
	<TR HEIGHT=*>
		<TD WIDTH="100%" CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH="100%">
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS="TD5" NOWRAP>계획일</TD>
									<TD CLASS="TD6" NOWRAP>
										<TABLE CELLSPACING=0 CELLPADDING=0>
											<TR>
												<TD>
													<script language =javascript src='./js/s2216ma1_fpDateTime1_txtConFromDt.js'></script>
												</TD>
												<TD>
													&nbsp;~&nbsp;
												</TD>
												<TD>
													<script language =javascript src='./js/s2216ma1_fpDateTime2_txtConToDt.js'></script>
												</TD>
											</TR>
										</TABLE>
									</TD>
									<TD CLASS="TD5" NOWRAP>영업그룹</TD>
									<TD CLASS="TD6" NOWRAP><INPUT NAME="txtConSalesGrp" ALT="영업그룹" TYPE="Text" MAXLENGTH=4 SiZE=10 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnConSalesGrp" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenConPopUp(C_PopSalesGrp)">&nbsp;<INPUT NAME="txtConSalesGrpNm" TYPE="Text" MAXLENGTH="50" SIZE=25 tag="14"></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>공장</TD>
									<TD CLASS="TD6"><INPUT NAME="txtConPlantCd" ALT="공장" TYPE="Text" MAXLENGTH=4 SiZE=10 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnConPlantCd" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenConPopUp(C_PopPlantCd)">&nbsp;<INPUT NAME="txtConPlantNm" TYPE="Text" MAXLENGTH="50" SIZE=25 tag="14"></TD>
									<TD CLASS="TD5" NOWRAP>품목</TD>
									<TD CLASS="TD6"><INPUT NAME="txtConItemCd" ALT="품목" TYPE="Text" MAXLENGTH=18 SiZE=10 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnConItemCd" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenConPopUp(C_PopItemCd)">&nbsp;<INPUT NAME="txtConItemNm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="14"></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>거래처</TD>
									<TD CLASS="TD6"><INPUT NAME="txtConSoldToParty" ALT="거래처" TYPE="Text" MAXLENGTH=10 SiZE=10 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnConSoldToParty" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenConPopUp(C_PopSoldToParty)">&nbsp;<INPUT NAME="txtConSoldToPartyNm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="14"></TD>
									<TD CLASS="TD5" NOWRAP>거래구분</TD>
									<TD CLASS="TD6"><SELECT NAME="cboConLocExpFlag" tag="11X" STYLE="WIDTH: 150px;"><OPTION Value=""></SELECT></TD>									
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
								<TD HEIGHT="68%" WIDTH="100%" COLSPAN=4>
									<script language =javascript src='./js/s2216ma1_OBJECT3_vspdData.js'></script>
								</TD>
							</TR>
							<TR>
								<TD HEIGHT="100%" WIDTH="100%" COLSPAN=4>
									<script language =javascript src='./js/s2216ma1_OBJECT1_vspdData2.js'></script>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>    
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_01%>></TD>
	</TR>
	
	<TR HEIGHT=20>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_30%>>
				<TR>
					<TD WIDTH=10></TD>
					<TD WIDTH=10>&nbsp;</TD>
					<TD WIDTH=* Align=RIGHT><a href = "VBSCRIPT:JumpChgCheck(BIZ_JUMP_ID)">판매계획확정</a></TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH="100%" HEIGHT=<%=BizSize%>>
			<IFRAME NAME="MyBizASP" WIDTH="100%" src="../../blank.htm"  HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=no noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>

<TEXTAREA class=hidden name=txtSpreadIns tag="24" TABINDEX="-1"></TEXTAREA>
<TEXTAREA class=hidden name=txtSpreadUpd tag="24" TABINDEX="-1"></TEXTAREA>
<TEXTAREA class=hidden name=txtSpreadDel tag="24" TABINDEX="-1"></TEXTAREA>

<INPUT TYPE=hidden NAME="txtMode" tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="txtFlgMode" tag="24" TABINDEX="-1">

<INPUT TYPE=hidden NAME="txtMaxRows" tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="txtMaxRows2" tag="24" TABINDEX="-1">

</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
