<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Accounting
'*  2. Function Name        : 
'*  3. Program ID           : a5101ma1
'*  4. Program Name         : 본지점대차대조표조회 
'*  5. Program Desc         : 본지점대차대조표조회 
'*  6. Component List       : 
'*  7. Modified date(First) : 2001/04/18
'*  8. Modified date(Last)  : 2003/06/05
'*  9. Modifier (First)     : 
'* 10. Modifier (Last)      : Lim YOung Woon
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change" 
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************
-->

<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/IncSvrCcm.inc"  -->
<!-- #Include file="../../inc/IncSvrHTML.inc"  -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs">          </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs">         </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs">       </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs">     </SCRIPT>
<SCRIPT LANGUAGE="VBScript"	  SRC="../../inc/incCliRdsQuery.vbs">        </SCRIPT>
<SCRIPT LANGUAGE="VBScript"	  SRC="../../inc/incCliDBAgentA.vbs">        </SCRIPT>
<SCRIPT LANGUAGE="VBScript"	  SRC="../../inc/incCliDBAgentVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incEB.vbs">                 </SCRIPT>

<Script Language="VBScript">
Option Explicit 
'########################################################################################################
'#                       4.  Data Declaration Part

'=                       4.1 External ASP File
Const BIZ_PGM_ID 		= "a8109mb1.asp"                              '☆: Biz Logic ASP Name

'========================================================================================================
'=                       4.2 Constant variables 
Const C_MaxKey          = 0					                          '☆: SpreadSheet의 키의 갯수 

'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim lgIsOpenPop
Dim lgFiscStart
Dim lgFromGlDt
Dim lgToGlDt
Dim lgPreFromGlDt
Dim lgPreToGlDt

'########################################################################################################
'#                       5.Method Declaration Part
'========================================================================================================
'                        5.1 Common Method-1
'========================================================================================================	
Sub InitVariables()

    lgStrPrevKey     = ""
    lgPageNo         = ""
    lgIntFlgMode     = Parent.OPMD_CMODE                          'Indicates that current mode is Create mode
	lgBlnFlgChgValue = False
    lgSortKey        = 1

End Sub

'========================================================================================================
Sub SetDefaultVal()

	Dim PreStartDate
	Dim PreEndDate	
    PreStartDate = UNIDateAdd("m", -12, Parent.gFiscStart,Parent.gServerDateFormat)
	PreEndDate   = UNIDateAdd("m", -12, Parent.gFiscEnd,  Parent.gServerDateFormat)

	frm1.txtFromGlDt.Text		= UniConvDateAToB(Parent.gFiscStart ,Parent.gServerDateFormat,Parent.gDateFormat)
	frm1.txtToGlDt.Text			= UniConvDateAToB(Parent.gFiscStart ,Parent.gServerDateFormat,Parent.gDateFormat) 
	frm1.txtPreFromGlDt.Text	= UniConvDateAToB(PreStartDate ,Parent.gServerDateFormat,Parent.gDateFormat)
	frm1.txtPreToGlDt.Text		= UniConvDateAToB(PreEnddate ,Parent.gServerDateFormat,Parent.gDateFormat) 

End Sub

'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "A", "COOKIE", "QA") %>                                '☆: 
    <% Call LoadBNumericFormatA("Q", "*", "COOKIE", "QA") %>
End Sub


'========================================================================================================
Sub InitSpreadSheet()
    Call SetZAdoSpreadSheet("A5108MA1_GRD01", "S", "A", "V20021211", parent.C_SORT_DBAGENT, frm1.vspdData, C_MaxKey, "X", "X")
	Call SetSpreadLock()
End Sub

'========================================================================================================
Sub SetSpreadLock()
    With frm1.vspdData    
		.ReDraw = False
		ggoSpread.Source = frm1.vspdData
		ggoSpread.SpreadLockWithOddEvenRowColor()
		.ReDraw = True
    End With		
End Sub

'========================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub

'========================================================================================
Sub PopRestoreSpreadColumnInf()
 
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()      
	Call ggoSpread.ReOrderingSpreadData()
	Call InitData()

End Sub

'========================================================================================================
Sub Form_Load()

    Call LoadInfTB19029()
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
    Call ggoOper.LockField(Document, "N")

	Call InitVariables()
	Call SetDefaultVal()
	Call InitSpreadSheet()
    Call SetToolBar("1100000000001111")
    frm1.txtFromGlDt.focus

End Sub
'========================================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )
 
End Sub

'========================================================================================================
Function FncQuery() 

    on Error Resume Next
    Err.Clear

    FncQuery = False

    If Not chkField(Document, "1") Then							
       Exit Function
    End If

    If CompareDateByFormat(frm1.txtFromGlDt.Text, frm1.txtToGlDt.Text, frm1.txtFromGlDt.Alt, frm1.txtToGlDt.Alt, _
	 "970024", frm1.txtFromGlDt.UserDefinedFormat, Parent.gComDateType, True) = False then
		frm1.txtToGlDt.Focus
		Exit Function
	End If

	If CompareDateByFormat(frm1.txtPreFromGlDt.Text, frm1.txtPreToGlDt.Text, frm1.txtPreFromGlDt.Alt, frm1.txtPreToGlDt.Alt, _
	 "970024", frm1.txtPreFromGlDt.UserDefinedFormat, Parent.gComDateType, True) = False then
		frm1.txtPreToGlDt.Focus
		Exit Function
	End If

    If CompareDateByFormat(frm1.txtPreToGlDt.Text, frm1.txtFromGlDt.Text, frm1.txtPreToGlDt.Alt, frm1.txtFromGlDt.Alt, _
	 "970024", frm1.txtPreToGlDt.UserDefinedFormat,Parent.gComDateType, True) = False then
		frm1.txtFromGlDt.Focus
		Exit Function
	End If

	Call ggoOper.ClearField(Document, "2")
    ggoSpread.Source = frm1.vspdData
	Call ggoSpread.ClearSpreadData()

    If frm1.txtBizAreaCd.value = "" Then
		frm1.txtBizAreaNm.value = ""
    End If

	Call InitVariables()    

    If DbQuery = False Then 
		Exit Function
	End If

	If Err.number = 0 Then
		FncQuery = True
	End If

	Set gActiveElement = document.ActiveElement

End Function


'========================================================================================================
Function FncPrint()

    on Error Resume Next
    Err.Clear
    
    FncPrint = False                                                             '☜: Processing is NG
	Call Parent.FncPrint()                                                       '☜: Protect system from crashing

	If Err.number = 0 Then
		FncPrint = True                                        
	End If

    Set gActiveElement = document.ActiveElement

End Function

'========================================================================================================
Function FncExcel() 

    on Error Resume Next
    Err.Clear

    FncExcel = False               

	Call Parent.FncExport(Parent.C_MULTI)
	
	If Err.number = 0 Then	
		FncExcel = True                                        
	End If

    Set gActiveElement = document.ActiveElement

End Function

'========================================================================================================
Function FncFind() 

    on Error Resume Next
    Err.Clear

    FncFind = False                

	Call Parent.FncFind(Parent.C_MULTI, True)

	If Err.number = 0 Then
		FncFind = True                                         
	End If

    Set gActiveElement = document.ActiveElement

End Function

'========================================================================================
Function FncSplitColumn()
    
    Dim ACol
    Dim ARow
    Dim iRet
    Dim iColumnLimit
    
    iColumnLimit  = Frm1.vspdData.MaxCols
    
    If gMouseClickStatus = "SPCRP" Then
		ACol = Frm1.vspdData.ActiveCol
		ARow = Frm1.vspdData.ActiveRow

		If ACol > iColumnLimit Then
			Frm1.vspdData.Col = iColumnLimit	:	Frm1.vspdData.Row = 0
			iRet = DisplayMsgBox("900030", "X", Trim(frm1.vspdData.Text), "X")
		   Exit Function
		End If   
    
		Frm1.vspdData.ScrollBars = Parent.SS_SCROLLBAR_NONE
		ggoSpread.Source = Frm1.vspdData
		ggoSpread.SSSetSplit(ACol)
		Frm1.vspdData.Col = ACol
		Frm1.vspdData.Row = ARow
		Frm1.vspdData.Action = 0
		Frm1.vspdData.ScrollBars = Parent.SS_SCROLLBAR_BOTH
	End If
    
End Function


'========================================================================================================
Function FncExit()

	Dim IntRetCD

    On Error Resume Next   
    Err.Clear          

	FncExit = False

    ggoSpread.Source = frm1.vspdData
    If lgBlnFlgChgValue = True OR ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO, "X", "X")			'⊙: "Will you destory previous data"
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End if

    If Err.number = 0 Then
       FncExit = True                                           
    End If

    Set gActiveElement = document.ActiveElement
    
End Function

'========================================================================================================
Function DbQuery()

	Dim strVal, strZeroFg
    Dim strYYYY, strMM, strDD    
    
    On Error Resume Next
    Err.Clear       
                                                  
    DbQuery = False
    
    Call DisableToolBar(parent.TBC_QUERY)
	Call LayerShowHide(1)

	Call GetQueryDate()		    

    With frm1
		If .ZeroFg1.checked = True Then
			strZeroFg = "Y"
		Else
			strZeroFg = "N"
		End If

        strVal = BIZ_PGM_ID
        If lgIntFlgMode  <> Parent.OPMD_UMODE Then
			strVal = strVal & "?txtFromGlDt="    & lgFromGlDt
			strVal = strVal & "&txtToGlDt="      & lgToGlDt
			strVal = strVal & "&txtPreFromGlDt=" & lgPreFromGlDt
			strVal = strVal & "&txtPreToGlDt="   & lgPreToGlDt
			strVal = strval & "&txtClassType="   & .txtClassType.value
			strVal = strVal & "&txtBizAreaCd="	 & .txtBizAreaCd.value
			strVal = strVal & "&strHqBrchFg="	 & "Y"
			strVal = strVal & "&strZeroFg="		 & strZeroFg
        	strVal = strVal & "&strUserId="		 & Parent.gUsrID
        Else
			strVal = strVal & "?txtFromGlDt="    & lgFromGlDt
        End If
        strVal = strVal & "&lgPageNo="       & lgPageNo
        strVal = strVal & "&lgMaxCount="     & C_SHEETMAXROWS_D
        strVal = strVal & "&lgSelectListDT=" & GetSQLSelectLIstDataType("A")
        strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
		strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))

        
    End With

    Call RunMyBizASP(MyBizASP, strVal)

	If Err.number = 0 Then
       DbQuery = True														  '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement

End Function

'========================================================================================
Sub DbQueryOk()												

    lgIntFlgMode     = Parent.OPMD_UMODE									 '⊙: Indicates that current mode is Update mode
    Call SetToolBar("1100000000011111")	

End Sub

'========================================================================================================
Function OpenBizAreaCd()

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	arrParam(0) = "사업장 팝업"				' 팝업 명칭 
	arrParam(1) = "B_BIZ_AREA"					' TABLE 명칭 
	arrParam(2) = Trim(frm1.txtBizAreaCd.Value)	' Code Condition
	arrParam(3) = ""							' Name Cindition
	arrParam(4) = ""							' Where Condition
	arrParam(5) = "사업장 코드"			
	
    arrField(0) = "BIZ_AREA_CD"					' Field명(0)
    arrField(1) = "BIZ_AREA_NM"					' Field명(1)
    
    arrHeader(0) = "사업장코드"				' Header명(0)
	arrHeader(1) = "사업장명"				' Header명(1)
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	lgIsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetReturnVal(arrRet,1)
	End If	
	
	Call EscPopup(1)

End Function


'========================================================================================================
Function OpenClassTypeCd()

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	arrParam(0) = "재무제표유형 팝업"		' 팝업 명칭 
	arrParam(1) = "A_ACCT_CLASS_TYPE"			' TABLE 명칭 
	arrParam(2) = Trim(frm1.txtClassType.Value)		' Code Condition
	arrParam(3) = ""							' Name Cindition
	arrParam(4) = "CLASS_TYPE LIKE " & FilterVar("BS%", "''", "S") & " "							' Where Condition
	arrParam(5) = "재무제표코드"			
	
    arrField(0) = "CLASS_TYPE"					' Field명(0)
    arrField(1) = "CLASS_TYPE_NM"				' Field명(1)
    
    arrHeader(0) = "재무제표코드"			' Header명(0)
	arrHeader(1) = "재무제표명"				' Header명(1)
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	lgIsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetReturnVal(arrRet,2)
	End If	
	Call EscPopup( 2)
End Function

'========================================================================================================
Function SetReturnVal(byref arrRet,byval field_fg)

	With frm1	
		Select case field_fg
			case 1
				.txtBizAreaCd.Value	  = arrRet(0)
				.txtBizAreaNm.Value	  = arrRet(1)
			case 2
				.txtClassType.Value	  = arrRet(0)
				.txtClassTypeNm.Value = arrRet(1)
		End select	
	End With

End Function
'=======================================================================================================
Function EscPopup(Byval iWhere)
	With frm1
		Select Case iWhere
			Case 1  
				.txtBizAreaCd.focus
			Case 2 
				.txtClassType.focus
		End Select    
	End With

End Function
'========================================================================================================
Sub SetPrintCond(StrEbrFile,varFromDt, varTodt, varPreFromDt, varPreToDt,varBizArea)
	
	StrEbrFile = "A5108MA1"

	If Trim(frm1.txtBizArea.value) = "" Then
		varBizArea = "*"
	Else
		varBizArea = Trim(frm1.txtBizArea.value)
	End If	

	varFromDt		 = lgFromGlDt	
	varToDt			 = lgToGlDt		
	varPreFromDt	 = lgPreFromGlDt	
	varPreToDt		 = lgPreToGlDt		
	
	

End Sub    

'========================================================================================================
Function BtnPreview()

	Dim StrEbrFile,varFromDt, varTodt, varPreFromDt, varPreToDt,varBizArea, VarClassType
	Dim StrUrl
	Dim lngPos
	Dim intCnt,IntRetCD
	Dim arrParam, arrField, arrHeader
	
	If lgIntFlgMode <> Parent.OPMD_UMODE Then	
		IntRetCD = DisplayMsgBox("900002","x","x","x")   ' 조회를 먼저 하십시오.	
		Exit Function
	End If
	
    If Not chkField(Document, "1") Then					 '⊙: This function check indispensable field
       Exit Function
    End If	
    
	If frm1.vspddata.MaxRows < 1 Then
		IntRetCD = DisplayMsgBox("900014","x","x","x")   '☜ 바뀐부분 
		Exit Function
	End If
	
    Call SetPrintCond(StrEbrFile,varFromDt, varTodt, varPreFromDt, varPreToDt,varBizArea)
    
    ObjName = AskEBDocumentName(StrEbrFile, "ebr")
    
    lngPos = 0
        		
	For intCnt = 1 to 3
	    lngPos = InStr(lngPos + 1, GetUserPath, "/")
	Next
	
	StrUrl = StrUrl & "varFromDt|"			& varFromDt
	StrUrl = StrUrl & "|varToDt|"			& varToDt
	StrUrl = StrUrl & "|varPreFromDt|"      & varPreFromDt
	StrUrl = StrUrl & "|varPreToDt|"		& varPreToDt
	StrUrl = StrUrl & "|BizAreaCd|"			& varBizArea

	Call FncEBRPreview(ObjName, StrUrl)
	
End Function


'========================================================================================================
Function BtnPrint()

	Dim IntRetCD,intCnt	
	Dim StrEbrFile,varFromDt, varTodt, varPreFromDt, varPreToDt,varBizArea, VarClassType
	Dim StrUrl
	Dim lngPos

	If lgIntFlgMode = Parent.OPMD_UMODE Then
		IntRetCD = DisplayMsgBox("900019", Parent.VB_YES_NO,"x","x") '☜ 바뀐부분 
		If IntRetCD = vbNo Then	Exit Function		
	Else
		IntRetCD = DisplayMsgBox("900002","x","x","x") '☜ 바뀐부분 
		
		 Exit Function
	End If
		
    If Not chkField(Document, "1") Then     			'⊙: This function check indispensable field
       Exit Function
    End If	
  
	If frm1.vspddata.MaxRows < 1 Then
		IntRetCD = DisplayMsgBox("900014","x","x","x") '☜ 바뀐부분 
		Exit Function
	End If
	  
    Call SetPrintCond(StrEbrFile,varFromDt, varTodt, varPreFromDt, varPreToDt,varBizArea)
    
    ObjName = AskEBDocumentName(StrEbrFile, "ebr")
    
    lngPos = 0
        		
	For intCnt = 1 to 3
	    lngPos = InStr(lngPos + 1, GetUserPath, "/")
	Next

	StrUrl = StrUrl & "varFromDt|" & varFromDt
	StrUrl = StrUrl & "|varToDt|"			& varToDt
	StrUrl = StrUrl & "|varPreFromDt|"      & varPreFromDt
	StrUrl = StrUrl & "|varPreToDt|"		& varPreToDt
	StrUrl = StrUrl & "|BizAreaCd|"			& varBizArea
	
	Call FncEBRPrint(EBAction, ObjName, StrUrl)
	
End Function


'========================================================================================================
Function PopZAdoConfigGrid()

	Dim arrRet
	
	If lgIsOpenPop = True Then Exit Function
	lgIsOpenPop = True

	arrRet = window.showModalDialog("../../ComAsp/ZADOSortPopup.asp",Array(ggoSpread.GetXMLData("A"),gMethodText),"dialogWidth=" & parent.SORTW_WIDTH & "px; dialogHeight=" & parent.SORTW_HEIGHT & "px;; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False
	
	If arrRet(0) = "X" Then
	   Exit Function
	Else
	   Call ggoSpread.SaveXMLData("A",arrRet(0),arrRet(1))
       Call InitVariables
       Call InitSpreadSheet()       
   End If

End Function



'========================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)

	Call SetPopupMenuItemInf("00000000001") 
    gMouseClickStatus = "SPC"	
	
	Set gActiveSpdSheet = frm1.vspdData   
    
	If frm1.vspdData.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If

    If Row <= 0 Then
        ggoSpread.Source = frm1.vspdData
        If lgSortKey = 1 Then
            ggoSpread.SSSort, lgSortKey
            lgSortKey = 2
        Else
            ggoSpread.SSSort, lgSortKey
            lgSortKey = 1
        End If    
    End If

    Call SetSpreadColumnValue("A", frm1.vspdData, Col, Row)

End Sub

'========================================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)

   If Button = 2 And gMouseClickStatus = "SPC" Then
      gMouseClickStatus = "SPCR"
   End If

End Sub    


'========================================================================================================
Sub vspdData_GotFocus()

    ggoSpread.Source = frm1.vspdData

End Sub


'========================================================================================================
sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)

	ggoSpread.source = frm1.vspdData
	Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

End Sub

	
'========================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If

	If CheckRunningBizProcess = True Then
	   Exit Sub
	End If
    
	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	    
    	If lgPageNo <> "" Then								
           Call DisableToolBar(Parent.TBC_QUERY)
           If DbQuery = False Then
              Call RestoreToolBar()
              Exit Sub
           End if
    	End If
    End If
    
End Sub

'=======================================================================================================
Sub txtFromGlDt_KeyDown(KeyCode, Shift)

	If KeyCode = 13 Then Call MainQuery()

End Sub

'=======================================================================================================
Sub txtToGlDt_KeyDown(KeyCode, Shift)

	If KeyCode = 13 Then Call MainQuery()

End Sub

'=======================================================================================================
Sub txtPreFromGlDt_KeyDown(KeyCode, Shift)

	If KeyCode = 13 Then Call MainQuery()

End Sub

'=======================================================================================================
Sub txtPreToGlDt_KeyDown(KeyCode, Shift)

	If KeyCode = 13 Then Call MainQuery()

End Sub

'=======================================================================================================
Sub txtFromGlDt_DblClick(Button)

    If Button = 1 Then
        frm1.txtFromGlDt.Action = 7
        Call SetFocusToDocument("M")	
        frm1.txtFromGlDt.Focus
    End If

End Sub

'=======================================================================================================
Sub txtToGlDt_DblClick(Button)

    If Button = 1 Then
        frm1.txtToGlDt.Action = 7
        Call SetFocusToDocument("M")	
        frm1.txtToGlDt.Focus
    End If

End Sub
'=======================================================================================================
Sub txtPreFromGlDt_DblClick(Button)

    If Button = 1 Then
        frm1.txtPreFromGlDt.Action = 7
        Call SetFocusToDocument("M")
        frm1.txtPreFromGlDt.Focus
    End If

End Sub

'=======================================================================================================
Sub txtPreToGlDt_DblClick(Button)

    If Button = 1 Then
        frm1.txtPreToGlDt.Action = 7
        Call SetFocusToDocument("M")	
        frm1.txtPreToGlDt.Focus
    End If

End Sub

'====================================================================================================
Sub CurFormatNumSprSheet()

	With frm1
		ggoSpread.Source = frm1.vspdData		
		ggoSpread.SSSetFloatByCellOfCur C_ItemAmt,-1, .txtDocCur.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gComNum1000, Parent.gComNumDec		
	End With

End Sub

'========================================================================================================
Sub GetQueryDate()
		
	Dim strFromYYYY, strFromMM, strFromDD
	Dim strToYYYY, strToMM, strToDD
	Dim strPreFromYYYY, strPreFromMM, strPreFromDD
	Dim strPreToYYYY, strPreToMM, strPreToDD
	
	Call ExtractDateFrom(frm1.txtFromGlDt.text,		Parent.gDateFormat,	Parent.gComDateType,	strFromYYYY,	strFromMM,		strFromDD)
	Call ExtractDateFrom(frm1.txtToGlDt.text,		Parent.gDateFormat,	Parent.gComDateType,	strToYYYY,		strToMM,		strToDD)
	Call ExtractDateFrom(frm1.txtPreFromGlDt.text,	Parent.gDateFormat,	Parent.gComDateType,	strPreFromYYYY,	strPreFromMM,	strPreFromDD)
	Call ExtractDateFrom(frm1.txtPreToGlDt.text,	Parent.gDateFormat,	Parent.gComDateType,	strPreToYYYY,	strPreToMM,		strPreToDD)
	
	lgFiscStart		= GetFiscDate(frm1.txtFromGlDt.Text)
	lgFromGlDt		= strFromYYYY		& strFromMM			& strFromDD
	lgToGlDt		= strToYYYY			& strToMM			& strToDD
	lgPreFromGlDt	= strPreFromYYYY	& strPreFromMM		& strPreFromDD
	lgPreToGlDt		= strPreToYYYY		& strPreToMM		& strPreToDD

End Sub

'========================================================================================================
Function GetFiscDate( ByVal strFromDate)

	Dim strFiscYYYY, strFiscMM, strFiscDD	
	Dim strFromYYYY, strFromMM, strFromDD
	
	GetFiscDate				="19000101"	
	
	Call ExtractDateFrom(Parent.gFiscStart,	Parent.gServerDateFormat,	Parent.gServerDateType,	strFiscYYYY,	strFiscMM,	strFiscDD)
	Call ExtractDateFrom(strFromDate,	Parent.gDateFormat,		Parent.gComDateType,		strFromYYYY,	strFromMM,	strFromDD)
	
	strFiscYYYY =  strFromYYYY
	
	If isnumeric(strFromYYYY) And isnumeric(strFromMM) And isnumeric(strFiscMM) Then
	
		If Cint(strFiscMM) > Cint(strFromMM)  then                         
		   GetFiscDate	= Cstr(Cint(strFromYYYY) - 1) & strFiscMM & strFiscDD		   
		Else
		   GetFiscDate	= strFromYYYY & strFiscMM & strFiscDD	 	    
		End If
	
	End If

End Function
</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
</HEAD>

<BODY TABINDEX="-1" SCROLL=NO>
<FORM NAME=frm1 TARGET=MyBizASP METHOD=POST>
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR>
		<TD <%=HEIGHT_TYPE_00%>><% ' 상위 여백 %></TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH="100%">
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD>
						<TABLE ID=MyTab CELLSPACING=0 CELLPADDING=0>
							<TR>
								<TD BACKGROUND="../../../CShared/image/table/seltab_up_bg.gif"><IMG SRC="../../../CShared/image/table/seltab_up_left.gif" WIDTH=9 HEIGHT=23></TD>
								<TD BACKGROUND="../../../CShared/image/table/seltab_up_bg.gif" ALIGN=CENTER CLASS=CLSMTAB><FONT COLOR=WHITE>본지점대차대조표조회(출력)</FONT></TD>
								<TD BACKGROUND="../../../CShared/image/table/seltab_up_bg.gif" ALIGN=RIGHT><IMG SRC="../../../CShared/image/table/seltab_up_right.gif" WIDTH=10 HEIGHT=23></TD>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=*>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT="100%">
		<TD WIDTH="100%" CLASS=Tab11>
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD <%=HEIGHT_TYPE_02%> WIDTH="100%"></TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH="100%">
						<FIELDSET>
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS=TD5 NOWRAP>회계일(당기)</TD>
									<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/a8109ma1_fpDateTime1_txtFromGlDt.js'></script>&nbsp;~&nbsp;
												         <script language =javascript src='./js/a8109ma1_fpDateTime2_txtToGlDt.js'></script></TD>
									</TD>
									<TD CLASS=TD5>사업장</TD>
									<TD CLASS=TD6><INPUT TYPE=TEXT ID=txtBizArea   NAME=txtBizAreaCd SIZE=10 MAXLENGTH=10 STYLE="TEXT-ALIGN: left" tag="1XNXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME=btn ALIGN=TOP TYPE=BUTTON ONCLICK="vbscript:Call OpenBizAreaCd()">&nbsp;
												  <INPUT TYPE=TEXT ID=txtBizAreaNm NAME=txtBizAreaNm SIZE=20 MAXLENGTH=20 STYLE="TEXT-ALIGN: left" tag="14X"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>회계일(전기)</TD>
									<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/a8109ma1_fpDateTime3_txtPreFromGlDt.js'></script>&nbsp;~&nbsp;
												         <script language =javascript src='./js/a8109ma1_fpDateTime4_txtPreToGlDt.js'></script></TD>
									</TD>
									<TD CLASS=TD5>재무제표코드</TD>
									<TD CLASS=TD6><INPUT TYPE=TEXT ID=txtClassType   NAME=txtClassType   SIZE=10 MAXLENGTH=10 STYLE="TEXT-ALIGN: left" tag="12XXXU" ALT="재무제표코드"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME=btn ALIGN=TOP TYPE=BUTTON ONCLICK="vbscript:Call OpenClassTypeCd()">&nbsp;
												  <INPUT TYPE=TEXT ID=txtClassTypeNm NAME=txtClassTypeNm SIZE=20 MAXLENGTH=20 STYLE="TEXT-ALIGN: left" tag="14X"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>조회구분</TD>
									<TD CLASS=TD6 NOWRAP>
										<SPAN STYLE="WIDTH: 80px"><INPUT TYPE=RADIO CLASS=RADIO NAME=ZeroFg CHECKED ID=ZeroFg1 VALUE=Y tag="25"><LABEL FOR=ZeroFg1>전체</LABEL></SPAN>
										<SPAN STYLE="WIDTH: 80px"><INPUT TYPE=RADIO CLASS=RADIO NAME=ZeroFg ID=ZeroFg2 VALUE=N tag="25"><LABEL FOR=ZeroFg2>발생금액</LABEL></SPAN></TD>
									<TD CLASS=TD5 NOWRAP></TD>
									<TD CLASS=TD6 NOWRAP></TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH="100%"></TD>
				</TR>
				<TR>
					<TD WIDTH="100%" HEIGHT=* VALIGN=TOP>
						<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD HEIGHT="100%">
								<script language =javascript src='./js/a8109ma1_vaSpread1_vspdData.js'></script></TD>
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
		<TD WIDTH="100%">
			<TABLE <%=LR_SPACE_TYPE_30%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD><BUTTON NAME=bttnPreview CLASS=CLSSBTN ONCLICK="vbscript:BtnPreview()" Flag=1>미리보기</BUTTON>&nbsp;
						<BUTTON NAME=bttnPrint	 CLASS=CLSSBTN ONCLICK="vbscript:BtnPrint()"   Flag=1>인쇄</BUTTON></TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH="100%" HEIGHT=<%=BizSize%>><IFRAME NAME=MyBizASP SRC="../../blank.htm" WIDTH="100%" HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO NORESIZE FRAMESPACING=0 TABINDEX="-1"></IFRAME></TD>
	</TR>
</TABLE>
<TEXTAREA CLASS=HIDDEN NAME=txtSpread	tag="24" TABINDEX="-1"></TEXTAREA>
<INPUT    TYPE=HIDDEN  NAME=txtMode		tag="24" TABINDEX="-1">
</FORM>
<DIV ID=MousePT NAME=MousePT>
	<IFRAME NAME=MouseWindow FRAMEBORDER=0 SCROLLING=NO NORESIZE FRAMESPACING=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm"></IFRAME>
</DIV>
<FORM NAME=EBAction TARGET=MyBizASP METHOD=POST>
	<INPUT TYPE=HIDDEN NAME=uname    TABINDEX="-1">
	<INPUT TYPE=HIDDEN NAME=dbname   TABINDEX="-1">
	<INPUT TYPE=HIDDEN NAME=filename TABINDEX="-1">
	<INPUT TYPE=HIDDEN NAME=condvar  TABINDEX="-1">
	<INPUT TYPE=HIDDEN NAME=date     TABINDEX="-1">	
</FORM>
</BODY>
</HTML>
