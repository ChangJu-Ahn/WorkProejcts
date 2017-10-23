<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Accounting
'*  2. Function Name        : 
'*  3. Program ID           : a5101ma1
'*  4. Program Name         : 본지점시산표조회 
'*  5. Program Desc         : 본지점시산표조회 
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

<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAMain.vbs">			</SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAEvent.vbs">			</SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAOperation.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliVariables.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"	SRC="../../inc/incCliDBAgentA.vbs">			</SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliDBAgentVariables.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"	SRC="../../inc/incCliRdsQuery.vbs">			</SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incEB.vbs">					</SCRIPT>

<Script Language="VBScript">
Option Explicit
	

'########################################################################################################
'#                       4.  Data Declaration Part
'========================================================================================================
'=                       4.1 External ASP File
Const BIZ_PGM_ID 		= "a8108MB1.asp"
Const BIZ_PGM_ID_SP 	= "a8108MB2.asp"
'========================================================================================================
'=                       4.2 Constant variables 
Const C_MaxKey          = 0					                          '☆: SpreadSheet의 키의 갯수 
'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       4.4 User-defind Variables
Dim C_ListSeq
Dim C_LBal   
Dim C_LSum   
Dim C_LThis  
Dim C_Result 
Dim C_RThis  
Dim C_RSum   
Dim C_RBal   

Dim lgIsOpenPop
Dim lgMaxFieldCount
Dim lgCookValue
Dim lgFiscStart
Dim lgStartDt
Dim lgEndDt

'########################################################################################################
'#                       5.Method Declaration Part
'========================================================================================================
'                        5.1 Common Method-1
'========================================================================================================	
Sub InitVariables()

    lgStrPrevKey     = ""
    lgPageNo         = ""
    lgIntFlgMode     = Parent.OPMD_CMODE  
    lgSortKey        = 1

	frm1.hSum.value = "합계"
	frm1.hUnBalance.value = "대차착오"

End Sub

'========================================================================================================
Sub SetDefaultVal()

	frm1.txtStartDT.Text	= UniConvDateAToB(Parent.gFiscStart ,Parent.gServerDateFormat,Parent.gDateFormat)
	frm1.txtEndDT.Text		= UniConvDateAToB("<%=GetSvrDate%>" ,Parent.gServerDateFormat,Parent.gDateFormat)
End Sub

'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "A", "COOKIE", "QA") %>                                '☆: 
End Sub


'========================================================================================================
Sub initSpreadPosVariables()  

	 C_ListSeq= 1
	 C_LBal   = 2
	 C_LSum   = 3
	 C_LThis  = 4
	 C_Result = 5
	 C_RThis  = 6
	 C_RSum   = 7
	 C_RBal   = 8
End Sub

'========================================================================================================
Sub InitSpreadSheet()

	Call SetZAdoSpreadSheet("A5112MA1_GRD01","S","A","V20021220",parent.C_SORT_DBAGENT,frm1.vspdData, C_MaxKey, "X","X")
	
	Call initSpreadPosVariables()    
    
    ggoSpread.Source = frm1.vspdData2
    ggoSpread.Spreadinit "V20021220",,parent.gAllowDragDropSpread
    
    With frm1.vspdData2
    
        .ReDraw = False
		
        .MaxCols = C_RBal + 1
        .RowHeaderDisplay = 0
        .Row = 0
        .RowHidden = True
        
        Call ggoSpread.ClearSpreadData()
		Call GetSpreadColumnPos("B")

         ggoSpread.SSSetEdit C_ListSeq, "", 15
         ggoSpread.SSSetFloat C_LBal, "", 15, Parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
         ggoSpread.SSSetFloat C_LSum, "", 15, Parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
         ggoSpread.SSSetFloat C_LThis, "", 15, Parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
         ggoSpread.SSSetEdit C_Result, "", 25, 2, , 40
         ggoSpread.SSSetFloat C_RThis, "", 15, Parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
         ggoSpread.SSSetFloat C_RSum, "", 15, Parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
         ggoSpread.SSSetFloat C_RBal, "", 15, Parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec

         .ScrollBars = 0

		Call ggoSpread.SSSetColHidden(C_ListSeq, C_ListSeq, True)
		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
			
        .ReDraw = True

    End With

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
    
    With frm1.vspdData2
		.ReDraw = False		
		ggoSpread.Source = frm1.vspdData2		
		ggoSpread.SpreadLockWithOddEvenRowColor()
		.ReDraw = True
    End With

End Sub

'========================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "B"
            ggoSpread.Source = frm1.vspdData2
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			C_ListSeq = iCurColumnPos(1)
			C_LBal    = iCurColumnPos(2)
			C_LSum    = iCurColumnPos(3)    
			C_LThis   = iCurColumnPos(4)
			C_Result  = iCurColumnPos(5)
			C_RThis   = iCurColumnPos(6)
			C_RSum    = iCurColumnPos(7)
			C_RBal    = iCurColumnPos(8)

    End Select    
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

    Call LoadInfTB19029														
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
    Call ggoOper.LockField(Document, "N")                                   

	Call InitVariables()
	Call SetDefaultVal()

	Call InitSpreadSheet()

    Call SetToolBar("1100000000001111")										
    frm1.txtStartDT.focus
   
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
    	
	If CompareDateByFormat(frm1.txtStartDt.Text, frm1.txtEndDt.Text, frm1.txtStartDt.Alt, frm1.txtEndDt.Alt, "970025", _
		frm1.txtStartDt.UserDefinedFormat, Parent.gComDateType, True) = False Then
		Exit Function
	End If
	
    Call ggoOper.ClearField(Document, "2")
    ggoSpread.Source = frm1.vspdData
	Call ggoSpread.ClearSpreadData()
	ggoSpread.Source = frm1.vspdData2
	Call ggoSpread.ClearSpreadData()
    Call InitVariables()

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
    
    FncPrint = False   
	Call Parent.FncPrint()   
    
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

	ggoSpread.Source = Frm1.vspdData
	If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO, "X", "X")         '⊙: "Will you destory previous data"
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If

    If Err.number = 0 Then	 
       FncExit = True    
    End If

    Set gActiveElement = document.ActiveElement

End Function

'========================================================================================================
Function DbQuery() 

	Dim strValSp, strZeroFg

    On Error Resume Next   
    Err.Clear   

    DbQuery = False

    Call DisableToolBar(parent.TBC_QUERY)
	Call LayerShowHide(1)

	Call GetQueryDate()

	With frm1
		
		If .PrintOpt1.checked = True Then
			.txtPrintOpt.value = "1"
		Else
			.txtPrintOpt.value = "2"
		End If
		
		If .ZeroFg1.checked = True Then
			strZeroFg = "Y"
		Else
			strZeroFg = "N"
		End If
		
		strValSp	= BIZ_PGM_ID_SP
        If lgIntFlgMode  <> Parent.OPMD_UMODE Then   ' This means that it is first search	
			'sp를 호출한다.        				
			strValSp = strValSp & "?lgFiscStart="	& Trim(lgFiscStart)
			strValSp = strValSp & "&lgStartDt="     & Trim(lgStartDt)
			strValSp = strValSp & "&lgEndDt="       & Trim(lgEndDt)
        	strValSp = strValSp & "&txtClassType=" & Trim(.txtClassType.value)
        	strValSp = strValSp & "&txtBizArea="	& Trim(.txtBizArea.value)
        	strValSp = strValSp & "&strHqBrchFg="   & "Y"
        	strValSp = strValSp & "&strZeroFg="		& strZeroFg
        	strValSp = strValSp & "&txtPrintOpt="   & Trim(.txtPrintOpt.value)
        	strValSp = strValSp & "&strUserId="		& Parent.gUsrID
        End If
    End With

    Call RunMyBizASP(MyBizASP, strValSp)

	If Err.number = 0 Then
       DbQuery = True														  '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement

End Function

'========================================================================================================
Function DbQuery2()
	Dim strVal

    On Error Resume Next   
    Err.Clear   

    DbQuery2 = False
    
    Call DisableToolBar(parent.TBC_QUERY)
	Call LayerShowHide(1)
    
	With frm1
 	
        strVal		= BIZ_PGM_ID
        If lgIntFlgMode  <> Parent.OPMD_UMODE Then   ' This means that it is first search
			strVal = strVal & "?txtClassType="   & .txtClassType.value
           	strVal = strVal & "&txtBizArea="	 & .txtBizArea.value
        Else
			strVal = strVal & "?txtClassType="   & .htxtClassType.value
           	strVal = strVal & "&txtBizArea="	 & .htxtBizArea.value
        End If

		strVal = strVal & "&lgPageNo="       & lgPageNo
		strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")
		strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
		strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))
    End With
	
	Call RunMyBizASP(MyBizASP, strVal)
    
	If Err.number = 0 Then
       DbQuery2 = True														  '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement

End Function

'========================================================================================================
Function DbQueryOk()												

	If DbQuery2 = False Then 
		Exit Function
	End If
	
End Function

'========================================================================================================
Function DbQuery2Ok()												

    lgIntFlgMode     = Parent.OPMD_UMODE										'⊙: Indicates that current mode is Update mode
    Call SetToolBar("1100000000011111")	
    
End Function

'========================================================================================================
Function OpenPopUp(Byval strCode, Byval iWhere)

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim strgChangeOrgId

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	Select Case iWhere
		Case 0
			arrParam(0) = "사업장 팝업"				' 팝업 명칭 
			arrParam(1) = "B_BIZ_AREA" 					' TABLE 명칭 
			arrParam(2) = strCode						' Code Condition
			arrParam(3) = ""							' Name Cindition
			arrParam(4) = ""							' Where Condition
			arrParam(5) = "사업장코드"				' 조건필드의 라벨 명칭 

			arrField(0) = "BIZ_AREA_CD"					' Field명(0)
			arrField(1) = "BIZ_AREA_NM"					' Field명(1)
    
			arrHeader(0) = "사업장코드"				' Header명(0)
			arrHeader(1) = "사업장명"				' Header명(1)

		Case 1
			arrParam(0) = "재무제표코드팝업"		' 팝업 명칭 
			arrParam(1) = "A_ACCT_CLASS_TYPE" 			' TABLE 명칭 
			arrParam(2) = strCode						' Code Condition
			arrParam(3) = ""							' Name Cindition
			arrParam(4) = "CLASS_TYPE LIKE " & FilterVar("TB%", "''", "S") & " "		' Where Condition
			arrParam(5) = "재무제표코드"			' 조건필드의 라벨 명칭 

			arrField(0) = "CLASS_TYPE"					' Field명(0)
			arrField(1) = "CLASS_TYPE_NM"				' Field명(1)
    
			arrHeader(0) = "재무제표코드"			' Header명(0)
			arrHeader(1) = "재무제표명"			    ' Header명(1)
			
	End Select
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	lgIsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetPopUp(arrRet, iWhere)
	End If	
	
	Call EscPopup( iWhere)
End Function

'========================================================================================================
'	Name : SetPopUp()
'	Description : CtrlItem Popup에서 Return되는 값 setting
'========================================================================================================
Function SetPopUp(ByRef arrRet, Byval iWhere)
	With frm1
		Select Case iWhere
			Case 0		'BIZ_AREA
				.txtBizArea.value      = arrRet(0)
				.txtBizAreaNm.value    = arrRet(1)
			Case 1	
				.txtClassType.value    = arrRet(0)
				.txtClassTypeNm.value  = arrRet(1)			
		End Select

	End With

End Function
'=======================================================================================================
Function EscPopup(Byval iWhere)
	With frm1
		Select Case iWhere
			Case 0  
				.txtBizArea.focus
			Case 1 
				.txtClassType.focus
		End Select    
	End With

End Function
'========================================================================================================
Sub SetPrintCond(StrEbrFile,VarBizArea,varString2,varString3,varFromDt,varToDt,varFiscStartDt,ClassType)

	StrEbrFile = "A5112MA1"

	If Trim(frm1.txtBizArea.value) = "" Then
		varBizArea = "*"
	Else
		varBizArea = Trim(frm1.txtBizArea.value)
	End If	

	ClassType	= frm1.txtClassType.value
	varString2	= frm1.hSum.value
	varString3	= frm1.hUnBalance.value
	
'	당기시작일은 DB(AP)Server Format의 날짜이다.
	varFiscStartDt	= lgFiscStart	
	varFromDt		= lgStartDt	
	varToDt			= lgEndDt	

End Sub    

'========================================================================================================
'   Event Name : BtnPreview
'   Event Desc : Print Button
'========================================================================================================
Function BtnPreview()

	Dim StrEbrFile,VarBizArea,varString2,varString3,varFromDt,varToDt,varFiscStartDt,ClassType
	Dim StrUrl,IntRetCD
	
	if lgIntFlgMode <> Parent.OPMD_UMODE then	
		IntRetCD = DisplayMsgBox("900002","x","x","x")
		Exit Function
	end if		
	
    If Not chkField(Document, "1") Then
       Exit Function
    End If	
    
    Call SetPrintCond(StrEbrFile,VarBizArea,varString2,varString3,varFromDt,varToDt,varFiscStartDt,ClassType)
    ObjName = AskEBDocumentName(StrEBrFile, "ebr")
    
	StrUrl = StrUrl & "varFromDt|"		& varFromDt
	StrUrl = StrUrl & "|varToDt|"		& varToDt
	StrUrl = StrUrl & "|varFiscStartDt|" & varFiscStartDt
	StrUrl = StrUrl & "|ClassType|"		& ClassType
	StrUrl = StrUrl & "|VarBizArea|"	& varBizArea
	StrUrl = StrUrl & "|varString2|"	& varString2
	StrUrl = StrUrl & "|varString3|"	& varString3
	'@@
	
	With frm1.vspdData2
	
		.Row = 1
		.Col  = 2				
		StrUrl = StrUrl & "|bal_lamt|"		& .Text
		.Col  = 3				
		StrUrl = StrUrl & "|tot_lamt|"		& .Text
		.Col  = 4
		StrUrl = StrUrl & "|this_lamt|"		& .Text
		.Col  = 6
		StrUrl = StrUrl & "|this_ramt|"		& .Text
		.Col  = 7			
		StrUrl = StrUrl & "|tot_ramt|"		& .Text
		.Col  = 8				
		StrUrl = StrUrl & "|bal_ramt|"		& .Text
			
	End With
	
	Call FncEBRPreview(ObjName,StrUrl)
			
End Function    

'========================================================================================================
Function BtnPrint()

	Dim StrEbrFile,VarBizArea,varString2,varString3,varFromDt,varToDt,varFiscStartDt,ClassType
	Dim StrUrl,IntRetCD
	
	if lgIntFlgMode <> Parent.OPMD_UMODE then	
		IntRetCD = DisplayMsgBox("900002","x","x","x")
		Exit Function
	end if		
	
    If Not chkField(Document, "1") Then
       Exit Function
    End If	
    
    Call SetPrintCond(StrEbrFile,VarBizArea,varString2,varString3,varFromDt,varToDt,varFiscStartDt,ClassType)
    ObjName = AskEBDocumentName(StrEBrFile, "ebr")
    
	StrUrl = StrUrl & "varFromDt|"		& varFromDt
	StrUrl = StrUrl & "|varToDt|"		& varToDt
	StrUrl = StrUrl & "|varFiscStartDt|" & varFiscStartDt
	StrUrl = StrUrl & "|ClassType|"		& ClassType
	StrUrl = StrUrl & "|VarBizArea|"	& varBizArea
	StrUrl = StrUrl & "|varString2|"	& varString2
	StrUrl = StrUrl & "|varString3|"	& varString3

	With frm1.vspdData2

		.Row = 1
		.Col  = 2				
		StrUrl = StrUrl & "|bal_lamt|"		& .Text
		.Col  = 3				
		StrUrl = StrUrl & "|tot_lamt|"		& .Text
		.Col  = 4
		StrUrl = StrUrl & "|this_lamt|"		& .Text
		.Col  = 6
		StrUrl = StrUrl & "|this_ramt|"		& .Text
		.Col  = 7			
		StrUrl = StrUrl & "|tot_ramt|"		& .Text
		.Col  = 8				
		StrUrl = StrUrl & "|bal_ramt|"		& .Text
			
	End With

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
       Call InitVariables()
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
Function vspdData_DblClick(ByVal Col, ByVal Row)

	If frm1.vspdData.MaxRows > 0 Then
		If frm1.vspdData.ActiveRow = Row Or frm1.vspdData.ActiveRow > 0 Then
		End If
	End If
	
End Function

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
           
           If DbQuery2 = False Then
              Call RestoreToolBar()
              Exit Sub
           End if
    	End If
    End If
    
End Sub

'========================================================================================================
Sub txtStartDT_DblClick(Button)
	If Button = 1 Then
       frm1.txtStartDT.Action = 7
       Call SetFocusToDocument("M")	
       frm1.txtStartDT.Focus
	End If
End Sub

'========================================================================================================
Sub txtStartDT_Change()
	
End Sub
'========================================================================================================
Sub txtEndDT_DblClick(Button)
	If Button = 1 Then
       frm1.txtEndDT.Action = 7
       Call SetFocusToDocument("M")	
       frm1.txtEndDT.Focus
	End If
End Sub

'========================================================================================================
Sub txtStartDT_KeyPress(KeyAscii)

	If KeyAscii = 13 Then 
	   Call MainQuery()
	End If   

End Sub

'========================================================================================================
Sub txtEndDT_KeyPress(KeyAscii)

	If KeyAscii = 13 Then 
	   Call MainQuery()
	End If   

End Sub

'========================================================================================================
Function PrintOpt1_OnClick() 

	If frm1.PrintOpt1.checked = True then
		bs_pl_fg.innerHTML = "재무제표코드"
		Call ElementVisible(frm1.txtClassType, 1)
		Call ElementVisible(frm1.txtClassTypeNm, 1)
		Call ElementVisible(frm1.btnClassType, 1)

		frm1.txtClassType.value		= ""
		frm1.txtClassTypeNm.value	= ""
	End if

End Function

'========================================================================================================
Function PrintOpt2_OnClick() 

	If frm1.PrintOpt2.checked = True then
		bs_pl_fg.innerHTML = ""
		Call ElementVisible(frm1.txtClassType, 0)
		Call ElementVisible(frm1.txtClassTypeNm, 0)
		Call ElementVisible(frm1.btnClassType, 0)

		frm1.txtClassType.value		= "*"
	End if

End Function

'========================================================================================================
Sub GetQueryDate()
		
	Dim strFromYYYY, strFromMM, strFromDD
	Dim strToYYYY, strToMM, strToDD
	
	Call ExtractDateFrom(frm1.txtStartDT.text,	Parent.gDateFormat,	Parent.gComDateType,	strFromYYYY,	strFromMM,	strFromDD)
	Call ExtractDateFrom(frm1.txtEndDT.text,	Parent.gDateFormat,	Parent.gComDateType,	strToYYYY,		strToMM,	strToDD)
	
	
	lgFiscStart		= GetFiscDate(frm1.txtStartDT.Text)
	lgStartDt		= strFromYYYY	& strFromMM		& strFromDD
	lgEndDt			= strToYYYY		& strToMM		& strToDD

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
		<TD <%=HEIGHT_TYPE_00%>></TD>
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
								<TD BACKGROUND="../../../CShared/image/table/seltab_up_bg.gif" ALIGN=CENTER CLASS=CLSMTAB><FONT COLOR=WHITE>본지점시산표조회(출력)</FONT></TD>
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
									<TD CLASS=TD5 NOWRAP>회계일자</TD>
									<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/a8108ma1_fpDateTime1_txtStartDT.js'></script>&nbsp;~&nbsp;<script language =javascript src='./js/a8108ma1_fpDateTime2_txtEndDT.js'></script></TD>
									<TD CLASS=TD5 NOWRAP>조회유형</TD>
									<TD CLASS=TD6 NOWRAP>
										<SPAN STYLE="WIDTH: 80px"><INPUT TYPE=RADIO CLASS=RADIO NAME=PrintOpt CHECKED ID=PrintOpt1 VALUE="Y" tag="15"><LABEL FOR=PrintOpt1>재무제표구분</LABEL></SPAN>
										<SPAN STYLE="WIDTH: 80px"><INPUT TYPE=RADIO CLASS=RADIO NAME=PrintOpt ID=PrintOpt2 VALUE="N" tag="15"><LABEL FOR=PrintOpt2>계정그룹</LABEL></SPAN></TD>
								</TR>	
								<TR>
									<TD CLASS=TD5 ID=bs_pl_fg NOWRAP>재무제표코드</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME=txtClassType   SIZE=10 MAXLENGTH=4 STYLE="TEXT-ALIGN: left" tag="12XXXU" ALT="재무제표코드"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME=btnClassType ALIGN=TOP TYPE=BUTTON ONCLICK="vbscript:Call OpenPopUp(frm1.txtClassType.Value, 1)">&nbsp;<INPUT TYPE=TEXT ID=txtClassTypeNm NAME=txtClassTypeNm SIZE=20 MAXLENGTH=20 STYLE="TEXT-ALIGN: left" tag="14X"></TD>
									<TD CLASS=TD5 NOWRAP>조회구분</TD>
									<TD CLASS=TD6 NOWRAP>
										<SPAN STYLE="WIDTH: 80px"><INPUT TYPE=RADIO CLASS=RADIO NAME=ZeroFg CHECKED ID=ZeroFg1 VALUE=Y tag="15"><LABEL FOR=ZeroFg1>전체</LABEL></SPAN>
										<SPAN STYLE="WIDTH: 80px"><INPUT TYPE=RADIO CLASS=RADIO NAME=ZeroFg ID=ZeroFg2 VALUE=N tag="15"><LABEL FOR=ZeroFg2>발생금액</LABEL></SPAN></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>사업장</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT ID=txtBizArea NAME=txtBizArea SIZE=10 MAXLENGTH=10 STYLE="TEXT-ALIGN: left" tag="1XNXXU" ALT="사업장코드"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME=btnBizArea ALIGN=TOP TYPE=BUTTON ONCLICK="vbscript:Call OpenPopUp(frm1.txtBizArea.Value, 0)">&nbsp;<INPUT TYPE=TEXT ID=txtBizAreaNm NAME=txtBizAreaNm SIZE=20 MAXLENGTH=20 STYLE="TEXT-ALIGN: left" tag="14X"></TD>
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
								<TD HEIGHT="94%"><!--94%-->
									<script language =javascript src='./js/a8108ma1_I841141289_vspdData.js'></script>
								</TD>
							</TR>
							<TR>
								<TD HEIGHT="6%"><!--6%-->
									<script language =javascript src='./js/a8108ma1_I845229780_vspdData2.js'></script>
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
		<TD WIDTH="100%">
			<TABLE <%=LR_SPACE_TYPE_30%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD><BUTTON NAME=bttnPreview CLASS=CLSSBTN ONCLICK="vbscript:BtnPreview()"  Flag=1>미리보기</BUTTON>&nbsp;
						<BUTTON NAME=bttnPrint   CLASS=CLSSBTN ONCLICK="vbscript:BtnPrint()"    Flag=1>인쇄</BUTTON></TD>
				</TR>
			</TABLE>	
		</TD>
	</TR>			
	<TR>
		<TD WIDTH="100%" HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO NORESIZE FRAMESPACING=0 TABINDEX="-1"></IFRAME></TD>
	</TR>
</TABLE>
<TEXTAREA CLASS=HIDDEN NAME=txtSpread	    tag="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME=txtMode			    tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME=txtMaxRows		    tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME=hBizArea			tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME=hClassType		    tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME=hFiscStart		    tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME=hStartDT			tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME=hEndDT			    tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME=hSum				tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME=hUnBalance		    tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME=txtPrintOpt		    tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME=htxtClassType		tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME=htxtBizArea		    tag="24" TABINDEX="-1">
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
