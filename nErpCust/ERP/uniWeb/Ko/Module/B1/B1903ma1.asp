<%@ LANGUAGE="VBSCRIPT" %>
<%
'**********************************************************************************************
'*  1. Module Name          : Basis Architect
'*  2. Function Name        : Master Data(Numeric Format)
'*  3. Program ID           : B1903ma1.asp
'*  4. Program Name         : B1903ma1.asp
'*  5. Program Desc         :
'*  6. Modified date(First) : 2000/09/18
'*  7. Modified date(Last)  : 2002/12/10
'*  8. Modifier (First)     : Hwang Jeong-won
'*  9. Modifier (Last)      : Sim Hae Young
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================
%>
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="javascript" SRC="../../inc/TabScript.js"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit																	'☜: indicates that All variables must be declared in advance

Const BIZ_PGM_ID = "B1903mb1.asp"												'☆: 비지니스 로직 ASP명 
Const BIZ_PGM_COMMON_FORMAT = "b1901ma1"
Const BIZ_PGM_COUNT_FORMAT = "b1902ma1"

Const TAB1 = 1										            <%'Tab의 위치 %>
Const TAB2 = 2

<!-- #Include file="../../inc/lgvariables.inc" -->	

Dim C_DataTypeNm  
Dim C_DataType    
Dim C_Currency    
Dim C_CurrencyNm  
Dim C_ModuleNm    
Dim C_ModuleCD    
Dim C_FormTypeNm  
Dim C_FormType    
Dim C_Decimals    
Dim C_RndUnit     
Dim C_RndPolicyNm 
Dim C_RndPolicy   
Dim C_DataFormat  

Dim C_DataTypeNm2  
Dim C_DataType2    
Dim C_Currency2    
Dim C_CurrencyNm2  
Dim C_ModuleNm2    
Dim C_ModuleCD2    
Dim C_FormTypeNm2  
Dim C_FormType2    
Dim C_Decimals2    
Dim C_RndUnit2     
Dim C_RndPolicyNm2 
Dim C_RndPolicy2   
Dim C_DataFormat2  


Dim IsOpenPop
Dim gSelframeFlg                                            <%'Current Tab Page%>

Sub InitSpreadPosVariables(ByVal pvSpdNo)  
    If pvSpdNo = "A" Then
        C_DataTypeNm  = 1
        C_DataType    = 2
        C_Currency    = 3
        C_CurrencyNm  = 4  
        C_ModuleNm    = 5  
        C_ModuleCD    = 6  
        C_FormTypeNm  = 7  
        C_FormType    = 8  
        C_Decimals    = 9  
        C_RndUnit     = 10 
        C_RndPolicyNm = 11 
        C_RndPolicy   = 12 
        C_DataFormat  = 13 
    ElseIf pvSpdNo = "B" Then
        C_DataTypeNm2  = 1
        C_DataType2    = 2
        C_Currency2    = 3
        C_CurrencyNm2  = 4  
        C_ModuleNm2    = 5  
        C_ModuleCD2    = 6  
        C_FormTypeNm2  = 7  
        C_FormType2    = 8  
        C_Decimals2    = 9  
        C_RndUnit2     = 10 
        C_RndPolicyNm2 = 11 
        C_RndPolicy2   = 12 
        C_DataFormat2  = 13 
    End If
End Sub

Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size

    lgLngCurRows = 0                            'initializes Deleted Rows Count
    lgSortKey = 1
End Sub

Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029A("I", "B","NOCOOKIE","MA") %>
End Sub

Sub InitSpreadSheet(ByVal pvSpdNo)
    If pvSpdNo = "" OR pvSpdNo = "A" Then

	    Call initSpreadPosVariables("A")

	    With frm1.vspdData
            ggoSpread.Source = frm1.vspdData
            ggoSpread.Spreadinit "V20021202",,parent.gAllowDragDropSpread    
	
	        .ReDraw = false
	
	        .MaxCols = C_DataFormat + 1												'☜: 최대 Columns의 항상 1개 증가시킴 
	        .Col = .MaxCols														'☆: 사용자 별 Hidden Column
            .ColHidden = True
    
            .MaxRows = 0
	        ggoSpread.ClearSpreadData

            Call GetSpreadColumnPos("A") 

            ggoSpread.SSSetCombo C_DataTypeNm, "데이터종류", 17	'1
            ggoSpread.SSSetCombo C_DataType, "", 4			'2
            ggoSpread.SSSetEdit C_Currency, "통화",17,,,3,2	'3
            ggoSpread.SSSetButton C_CurrencyNm       		'4
            ggoSpread.SSSetCombo C_ModuleNm, "업무", 18			'5
            ggoSpread.SSSetCombo C_ModuleCD, "", 4			'6
            ggoSpread.SSSetCombo C_FormTypeNm, "화면종류", 12		'7
            ggoSpread.SSSetCombo C_FormType, "", 4			'8
    
            ggoSpread.SSSetFloat C_Decimals,"소수점자리수" ,16,"6",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,,"0","6"
            ggoSpread.SSSetEdit C_RndUnit, "올림처리단위", 18, 1	'10
            ggoSpread.SSSetCombo C_RndPolicyNm, "올림구분", 20		'11
            ggoSpread.SSSetCombo C_RndPolicy, "", 4			'12
            ggoSpread.SSSetEdit C_DataFormat, "포맷", 26 , 1 '13                                  
            
            Call ggoSpread.MakePairsColumn(C_Currency,C_CurrencyNm)    
            
            Call ggoSpread.SSSetColHidden(C_DataType,C_DataType,True)
            Call ggoSpread.SSSetColHidden(C_ModuleNm,C_FormType,True)
            Call ggoSpread.SSSetColHidden(C_RndPolicy,C_RndPolicy,True)

	        .ReDraw = true
	
            Call SetSpreadLock("A") 
        End With

    End If
    
    If pvSpdNo = "" OR pvSpdNo = "B" Then
	    
	    Call initSpreadPosVariables("B")

	    With frm1.vspdData2
	        ggoSpread.Source = frm1.vspdData2
            ggoSpread.Spreadinit "V20021203",,parent.gAllowDragDropSpread    

	        .ReDraw = false
	
	        .MaxCols = C_DataFormat2 + 1												'☜: 최대 Columns의 항상 1개 증가시킴 
	        .Col = .MaxCols														'☆: 사용자 별 Hidden Column
            .ColHidden = True
    
            .MaxRows = 0
	        ggoSpread.ClearSpreadData

            Call GetSpreadColumnPos("B") 
            

            ggoSpread.SSSetCombo C_DataTypeNm2, "데이터종류", 14	'1
            ggoSpread.SSSetCombo C_DataType2, "", 4			'2
            ggoSpread.SSSetEdit C_Currency2, "통화",17,,,3,2	'3
            ggoSpread.SSSetButton C_CurrencyNm2       		'4
            ggoSpread.SSSetCombo C_ModuleNm2, "업무", 20			'5
            ggoSpread.SSSetCombo C_ModuleCD2, "", 4			'6
            ggoSpread.SSSetCombo C_FormTypeNm2, "화면종류", 12		'7
            ggoSpread.SSSetCombo C_FormType2, "", 4			'8
    
            ggoSpread.SSSetFloat C_Decimals2,"소수점자리수" ,15,"6",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,,"0","6"
    
            ggoSpread.SSSetEdit C_RndUnit2, "올림처리단위", 15, 1	'10
            ggoSpread.SSSetCombo C_RndPolicyNm2, "올림구분", 14		'11
            ggoSpread.SSSetCombo C_RndPolicy2, "", 4			'12
            ggoSpread.SSSetEdit C_DataFormat2, "포맷", 23, 1   '13                                  

            Call ggoSpread.MakePairsColumn(C_Currency2,C_CurrencyNm2)    

            Call ggoSpread.SSSetColHidden(C_DataType2,C_DataType2,True)
            Call ggoSpread.SSSetColHidden(C_ModuleCD2,C_FormType2,True)
            Call ggoSpread.SSSetColHidden(C_RndPolicy2,C_RndPolicy2,True)

	        .ReDraw = true
	
            Call SetSpreadLock("B") 
    
        End With
    End If
End Sub

Sub SetSpreadLock(ByVal pvSpdNo)
    If pvSpdNo = "A" Then
        ggoSpread.Source = Frm1.vspdData

        With frm1
    
        .vspdData.ReDraw = False
    
        ggoSpread.SpreadLock C_DataTypeNm, -1, C_FormType
        ggoSpread.SpreadLock C_Currency, -1, C_Currency
        ggoSpread.SSSetRequired	C_Decimals, -1, -1
        ggoSpread.SSSetRequired	C_RndUnit, -1, -1
        ggoSpread.SSSetRequired	C_RndPolicyNm, -1, -1
        ggoSpread.SpreadLock C_DataFormat, -1, C_DataFormat
        ggoSpread.SpreadLock C_RndUnit, -1, C_RndUnit
		ggoSpread.SSSetProtected .vspdData.MaxCols, -1, -1
        .vspdData.ReDraw = True

        End With
    
    ElseIf pvSpdNo = "B" Then
        ggoSpread.Source = Frm1.vspdData2
        With frm1
    
        .vspdData2.ReDraw = False
    
        ggoSpread.SpreadLock C_DataTypeNm2, -1, C_FormType2
        ggoSpread.SpreadLock C_Currency2, -1, C_Currency2
        ggoSpread.SSSetRequired	C_Decimals2, -1, -1
        ggoSpread.SSSetRequired	C_RndPolicyNm2, -1, -1
        ggoSpread.SpreadLock C_DataFormat2, -1, C_DataFormat2
        ggoSpread.SpreadLock C_RndUnit2, -1, C_RndUnit2
		ggoSpread.SSSetProtected .vspdData2.MaxCols, -1, -1
        .vspdData2.ReDraw = True

        End With
    End If
End Sub

Sub SetSpreadColor(ByVal pvStartRow,ByVal pvEndRow)

	If gSelframeFlg = TAB1 Then
		With frm1
    
		.vspdData.ReDraw = False
		ggoSpread.SSSetRequired C_DataTypeNm,  pvStartRow, pvEndRow
		ggoSpread.SSSetRequired C_Currency,    pvStartRow, pvEndRow
		ggoSpread.SSSetRequired C_Decimals,    pvStartRow, pvEndRow
		ggoSpread.SSSetRequired C_RndPolicyNm, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_RndUnit,    pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_DataFormat, pvStartRow, pvEndRow
		.vspdData.ReDraw = True
    
		End With
    Else    
    
		With frm1
    
		.vspdData2.ReDraw = False
		ggoSpread.SSSetRequired C_DataTypeNm2, pvStartRow, pvEndRow
		ggoSpread.SSSetRequired C_Currency2,   pvStartRow, pvEndRow
		ggoSpread.SSSetRequired C_ModuleNm2,   pvStartRow, pvEndRow
		ggoSpread.SSSetRequired C_Decimals2,   pvStartRow, pvEndRow
		ggoSpread.SSSetRequired C_RndPolicyNm2,pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_RndUnit2,   pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_DataFormat2,pvStartRow, pvEndRow
		.vspdData2.ReDraw = True
    
		End With
	End If

End Sub

Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
            C_DataTypeNm  = iCurColumnPos(1)
            C_DataType    = iCurColumnPos(2)
            C_Currency    = iCurColumnPos(3)
            C_CurrencyNm  = iCurColumnPos(4)  
            C_ModuleNm    = iCurColumnPos(5)  
            C_ModuleCD    = iCurColumnPos(6)  
            C_FormTypeNm  = iCurColumnPos(7)  
            C_FormType    = iCurColumnPos(8)  
            C_Decimals    = iCurColumnPos(9)  
            C_RndUnit     = iCurColumnPos(10) 
            C_RndPolicyNm = iCurColumnPos(11) 
            C_RndPolicy   = iCurColumnPos(12) 
            C_DataFormat  = iCurColumnPos(13) 
    
       Case "B"
            ggoSpread.Source = frm1.vspdData2
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

            C_DataTypeNm2  = iCurColumnPos(1)
            C_DataType2    = iCurColumnPos(2)
            C_Currency2    = iCurColumnPos(3)
            C_CurrencyNm2  = iCurColumnPos(4)  
            C_ModuleNm2    = iCurColumnPos(5)  
            C_ModuleCD2    = iCurColumnPos(6)  
            C_FormTypeNm2  = iCurColumnPos(7)  
            C_FormType2    = iCurColumnPos(8)  
            C_Decimals2    = iCurColumnPos(9)  
            C_RndUnit2     = iCurColumnPos(10) 
            C_RndPolicyNm2 = iCurColumnPos(11) 
            C_RndPolicy2   = iCurColumnPos(12) 
            C_DataFormat2  = iCurColumnPos(13) 
    End Select    
End Sub

Function ClickTab1()
	Dim IntRetCD
	
	If gSelframeFlg = TAB1 Then Exit Function
	
	ggoSpread.Source = frm1.vspdData2
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "X", "X") '☜ 바뀐부분 
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
	
	Call changeTabs(TAB1)                                               <%'첫번째 Tab%>
	gSelframeFlg = TAB1
	frm1.cboModuleCd.value = "*"
	frm1.cboModuleCd.disabled = True
	frm1.cboFormType.value = "I"
	Call SetToolbar("1100110000111111")										'⊙: 버튼 툴바 제어 
	Call MainQuery()
	
End Function

Function ClickTab2()	
	Dim IntRetCD
	
	If gSelframeFlg = TAB2 Then Exit Function
	
	ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "X", "X") '☜ 바뀐부분 
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
		
	Call changeTabs(TAB2)
	gSelframeFlg = TAB2
	frm1.cboModuleCd.value = " "
	frm1.cboModuleCd.disabled = False
	frm1.cboFormType.value = "Q"
	Call SetToolbar("1100110000111111")										'⊙: 버튼 툴바 제어 
	Call MainQuery()    
	
End Function

Sub InitSpreadComboBox()
	Dim strCboData    ''lgF0
	Dim strCboData2   ''lgF1
	
	Call CommonQueryRs(" MINOR_CD,MINOR_NM ", " B_MINOR ", " MAJOR_CD = " & FilterVar("B0003", "''", "S") & "  ", _	
    	               lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    
	strCboData = Replace(lgF0,chr(11),vbTab)
    strCboData2 = Replace(lgF1,chr(11),vbTab)
    strCboData = Left(strCboData,Len(strCboData) - 1)
    strCboData2 = Left(strCboData2,Len(strCboData2) - 1)
    
    ggoSpread.Source = frm1.vspdData
	ggoSpread.SetCombo strCboData,  C_DataType
	ggoSpread.SetCombo strCboData2, C_DataTypeNm

	ggoSpread.Source = frm1.vspdData2
	ggoSpread.SetCombo strCboData,  C_DataType2
	ggoSpread.SetCombo strCboData2, C_DataTypeNm2
	
	''MODULE
	Call CommonQueryRs(" MINOR_CD,MINOR_NM ", " B_MINOR ", " MAJOR_CD = " & FilterVar("B0001", "''", "S") & "  ", _
    	               lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)    
	strCboData = Replace(lgF0,chr(11),vbTab)
    strCboData2 = Replace(lgF1,chr(11),vbTab)
    strCboData = Left(strCboData,Len(strCboData) - 1)
    strCboData2 = Left(strCboData2,Len(strCboData2) - 1)
    
    ggoSpread.Source = frm1.vspdData
	ggoSpread.SetCombo strCboData, C_ModuleCD
	ggoSpread.SetCombo strCboData2, C_ModuleNm
	
	ggoSpread.Source = frm1.vspdData2
	ggoSpread.SetCombo strCboData, C_ModuleCD2
	ggoSpread.SetCombo strCboData2, C_ModuleNm2
	
	''FORM TYPE
	Call CommonQueryRs(" MINOR_CD,MINOR_NM ", " B_MINOR ", " MAJOR_CD = " & FilterVar("B0002", "''", "S") & "  ", _
    	               lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)    
	strCboData = Replace(lgF0,chr(11),vbTab)
    strCboData2 = Replace(lgF1,chr(11),vbTab)
    strCboData = Left(strCboData,Len(strCboData) - 1)
    strCboData2 = Left(strCboData2,Len(strCboData2) - 1)
    
    ggoSpread.Source = frm1.vspdData
	ggoSpread.SetCombo strCboData, C_FormType
	ggoSpread.SetCombo strCboData2, C_FormTypeNm
	
	ggoSpread.Source = frm1.vspdData2
	ggoSpread.SetCombo strCboData, C_FormType2
	ggoSpread.SetCombo strCboData2, C_FormTypeNm2
	
	
	''FLAG(올림/반올림)
	Call CommonQueryRs(" MINOR_CD,MINOR_NM ", " B_MINOR ", " MAJOR_CD = " & FilterVar("B0004", "''", "S") & "  ", _	
    	               lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)    
	strCboData = Replace(lgF0,chr(11),vbTab)
    strCboData2 = Replace(lgF1,chr(11),vbTab)
    strCboData = Left(strCboData,Len(strCboData) - 1)
    strCboData2 = Left(strCboData2,Len(strCboData2) - 1)
    
    ggoSpread.Source = frm1.vspdData
	ggoSpread.SetCombo strCboData, C_RndPolicy
	ggoSpread.SetCombo strCboData2, C_RndPolicyNm
	
	ggoSpread.Source = frm1.vspdData2
	ggoSpread.SetCombo strCboData, C_RndPolicy2
	ggoSpread.SetCombo strCboData2, C_RndPolicyNm2
End Sub

Sub InitComboBox()
	Dim strCboData    ''lgF0
	Dim strCboData2   ''lgF1
	
	Call CommonQueryRs(" MINOR_CD,MINOR_NM ", " B_MINOR ", " MAJOR_CD = " & FilterVar("B0003", "''", "S") & "  ", _	
    	               lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboDataType, "" & Chr(11) & lgF0, "" & Chr(11) & lgF1, Chr(11))
	
	''MODULE
	Call CommonQueryRs(" MINOR_CD,MINOR_NM ", " B_MINOR ", " MAJOR_CD = " & FilterVar("B0001", "''", "S") & "  ", _
    	               lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)    
	Call SetCombo2(frm1.cboModuleCd, lgF0, lgF1, Chr(11))	

	''FORM TYPE
	Call CommonQueryRs(" MINOR_CD,MINOR_NM ", " B_MINOR ", " MAJOR_CD = " & FilterVar("B0002", "''", "S") & "  ", _
    	               lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)    
	Call SetCombo2(frm1.cboFormType, lgF0, lgF1, Chr(11))
	
	''FLAG(올림/반올림)
	Call CommonQueryRs(" MINOR_CD,MINOR_NM ", " B_MINOR ", " MAJOR_CD = " & FilterVar("B0004", "''", "S") & "  ", _	
    	               lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)    
	''Call SetCombo2(frm1.cboDataType, lgF0, lgF1, Chr(11))
End Sub

Function OpenCurrency(Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "통화 팝업"					<%' 팝업 명칭 %>
	arrParam(1) = "b_currency"					<%' TABLE 명칭 %>
	
	If iWhere = 0 then
	    arrParam(2) = frm1.txtCurrency.value		<%' Code Condition%>
	ElseIf iWhere = 1 then
	    arrParam(2) = frm1.vspdData.Text		<%' Code Condition%>
	ElseIf iWhere = 2 then
	    arrParam(2) = frm1.vspdData2.Text		<%' Code Condition%>
	End If
	arrParam(3) = ""							<%' Name Cindition%>
	arrParam(4) = ""							<%' Where Condition%>
	arrParam(5) = "통 화"						<%' 조건필드의 라벨 명칭 %>
	
    arrField(0) = "currency"					<%' Field명(0)%>
    arrField(1) = "currency_desc"					<%' Field명(1)%>
    
    arrHeader(0) = "통화코드"					<%' Header명(0)%>
    arrHeader(1) = "통화명"						<%' Header명(1)%>
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	frm1.txtCurrency.focus
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetCurrency(arrRet,iWhere)
	End If	
	
End Function

Function SetCurrency(byval arrRet,Byval iWhere)
    With frm1
		If iWhere = 0 Then 'TextBox(Condition)
			.txtCurrency.Value    = arrRet(0)		
	        .txtCurrencyNm.Value  = arrRet(1)		
		ElseIf iWhere = 1 Then 'Spread1(Condition)
			.vspdData.Col = C_Currency
			.vspdData.Text = arrRet(0)
			
		ElseIf iWhere = 2 Then 'Spread2(Condition)
		    .vspdData2.Col = C_Currency2
			.vspdData2.Text = arrRet(0)
			'lgBlnFlgChgValue = True
		End If
	End With
End Function

Function LoadCommonFormat()
    
    PgmJump(BIZ_PGM_COMMON_FORMAT)

End Function

Function LoadCountFormat()
    
    PgmJump(BIZ_PGM_COUNT_FORMAT)

End Function

Sub Form_Load()
	
    Call LoadInfTB19029                                                     <%'Load table , B_numeric_format%>
    
    Call ggoOper.LockField(Document, "N")                                   <%'Lock  Suitable  Field%>                         
    			                                                            <%'Format Numeric Contents Field%>                                                                            
    			                                                            
    Call InitSpreadSheet("")    
    Call InitVariables                                                     '⊙: Initializes local global variables
    
    '----------  Coding part  -------------------------------------------------------------    
	Call InitSpreadComboBox
	Call InitComboBox	
       
    gIsTab     = "Y"
    gTabMaxCnt = 2
    
    Call ClickTab1
   
    frm1.cboDatatype.focus
End Sub

Sub Form_QueryUnload(Cancel , UnloadMode )

End Sub

Sub vspdData_Change(ByVal Col, ByVal Row )
Dim i, j, x
Dim intIndex
Dim lRound, lRoundP
Dim strCellText

    ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row

    lgBlnFlgChgValue = True
    
    With frm1.vspdData
'''''''''''''''''''''''''''''hmkwon modify (2002/08/23)  ISSUE 461              
		.Row = Row
    
		Select Case Col
			Case  C_DataTypeNm
				.Col = Col
				intIndex = .Value
				.Col = C_DataType
				.Value = intIndex
			
		    Case  C_ModuleNm
				.Col = Col
				intIndex = .Value
				.Col = C_ModuleCD
				.Value = intIndex
			
			Case  C_FormTypeNm
				.Col = Col
				intIndex = .Value
				.Col = C_FormType
				.Value = intIndex	
				
			Case  C_RndPolicyNm
				.Col = Col
				intIndex = .Value
				.Col = C_RndPolicy
				.Value = intIndex	
		End Select
'''''''''''''''''''''''''''''hmkwon modify end 
                  
		If Col = C_Decimals Or Col = C_DataTypeNm Then
           Col = C_Decimals  ''없음 안됨 
		  .Col = Col
		  .Row = Row
		  j = .value
    
		  lRound = 0.1
		  lRoundP = 1
		  		      
		    If j > 0 Then
		        For i = 1 To j
		            lRound = lRound * 0.1
		        Next
		        
		        .Col = C_RndUnit
		        .Row = Row
		        .value = lRound
		        
		    ElseIf j = 0 Then
		        .Col = C_RndUnit
		        .Row = Row
		        .value = lRound
		        
		    Else
		        For i = 1 To (j * -1)
		            lRoundP = lRoundP * 10
		        Next
		        
		        lRoundP = lRoundP / 10
		        .Col = C_RndUnit
		        .Row = Row
		        .value = lRoundP
		        
		    End If    		   
		  

		'***   frm1 안에 들어가야 하는데...vspdData이안에 들어가서 에러가 났었다...0416
		
		.value = replace (.value, ".", "@")
		.value = replace (.value, ",", "$")
		.value = replace (.value, "@", parent.gComNumDec)
		.value = replace (.value, "$", parent.gComNum1000)

''''''''''''''''''''''''''''''Srh추가/수정(2002/08/23)        
        .Col = C_DataTypeNm  : .Row = Row   ''데이타종류 
        strCellText = Trim(.text)
                
		.Col = 13 : .Row = Row			    
		Select Case CInt(j)
		    Case -1
		        Select Case strCellText
		            Case "금액"
			            .Text = "#,###,###,###,##0"
			        Case "단가"
			            .Text = "##,###,###,##0"
			        Case "환율"
			            .Text = "###,###,##0"
			        Case Else
			            .Text = "#,###,###,###,##0"
			    End Select 
			Case -2
				Select Case strCellText
		            Case "금액"
			            .Text = "#,###,###,###,#00"
			        Case "단가"
			            .Text = "##,###,###,#00"
			        Case "환율"
			            .Text = "###,###,#00"
			        Case Else
			            .Text = "#,###,###,###,#00"
			    End Select 
			Case -3
				Select Case strCellText
		            Case "금액"
			            .Text = "#,###,###,###,000"
			        Case "단가"
			            .Text = "##,###,###,000"
			        Case "환율"
			            .Text = "###,###,000"
			        Case Else
			            .Text = "#,###,###,###,000"
			    End Select 
			Case -4
				Select Case strCellText
		            Case "금액"
			            .Text = "#,###,###,##0,000"
			        Case "단가"
			            .Text = "##,###,##0,000"
			        Case "환율"
			            .Text = "###,##0,000"
			        Case Else
			            .Text = "#,###,###,##0,000"
			    End Select 
			Case 0
			    Select Case strCellText   ''금액/단가/환율=13/11/9(정수부 자릿수)
			        Case "금액"
			            .Text = "#,###,###,###,###"
			        Case "단가"
			            .Text = "##,###,###,###"
			        Case "환율"
			            .Text = "###,###,###"
			        Case Else
			            .Text = "#,###,###,###,###"
			    End Select 					
			Case 1
			    Select Case strCellText
		            Case "금액"
			            .Text = "#,###,###,###,###.0"
			        Case "단가"
			            .Text = "##,###,###,###.0"
			        Case "환율"
			            .Text = "###,###,###.0"
			        Case Else
			            .Text = "#,###,###,###,###.0"
			    End Select 					
			Case 2
			    Select Case strCellText
		            Case "금액"
			            .Text = "#,###,###,###,###.00"
			        Case "단가"
			            .Text = "##,###,###,###.00"
			        Case "환율"
			            .Text = "###,###,###.00"
			        Case Else
			            .Text = "#,###,###,###,###.00"
			    End Select
			Case 3
			    Select Case strCellText
		            Case "금액"
			            .Text = "###,###,###,###.000"
			        Case "단가"
			            .Text = "##,###,###,###.000"
			        Case "환율"
			            .Text = "###,###,###.000"
			        Case Else
			            .Text = "###,###,###,###.000"
			    End Select					
			Case 4
			    Select Case strCellText
		            Case "금액"
			            .Text = "##,###,###,###.0000"
			        Case "단가"
			            .Text = "##,###,###,###.0000"
			        Case "환율"
			            .Text = "###,###,###.0000"
			        Case Else
			            .Text = "##,###,###,###.0000"
			    End Select					
			Case 5
			    Select Case strCellText
		            Case "금액"
			            .Text = "#,###,###,###.00000"
			        Case "단가"
			            .Text = "#,###,###,###.00000"
			        Case "환율"
			            .Text = "###,###,###.00000"
			        Case Else
			            .Text = "#,###,###,###.00000"
			    End Select					
			Case 6
			    Select Case strCellText
		            Case "금액"
			            .Text = "###,###,###.000000"
			        Case "단가"
			            .Text = "###,###,###.000000"
			        Case "환율"
			            .Text = "###,###,###.000000"
			        Case Else
			            .Text = "###,###,###.000000"
			    End Select									
			Case Else
				.Text = "#,###,###,###,###"					
		End Select
''''''''''''''''''''''''''''''Srh추가/수정(2002/08/23)

		  .text = replace (.text, ".", "@")
		  .text = replace (.text, ",", "$")
		  .text = replace (.text, "@", parent.gComNumDec)			'parent.gComNumDec
		  .text = replace (.text, "$", parent.gComNum1000)			'parent.gComNum1000
			
		End If											
		
    End with
       
End Sub

Sub vspdData2_Change(ByVal Col, ByVal Row )
Dim i, j
Dim intIndex
Dim lRound, lRoundP
Dim strCellText

    ggoSpread.Source = frm1.vspdData2
    ggoSpread.UpdateRow Row

    lgBlnFlgChgValue = True
    
    With frm1.vspdData2
'''''''''''''''''''''''''''''hmkwon modify (2002/08/23)  ISSUE 461          
    	.Row = Row
    	Select Case Col
			Case  C_DataTypeNm2
				.Col = Col
				intIndex = .Value
				.Col = C_DataType2
				.Value = intIndex
			
		    Case  C_ModuleNm2
				.Col = Col
				intIndex = .Value
				.Col = C_ModuleCD2
				.Value = intIndex
			
			Case  C_FormTypeNm2
				.Col = Col
				intIndex = .Value
				.Col = C_FormType2
				.Value = intIndex	
				
			Case  C_RndPolicyNm2
				.Col = Col
				intIndex = .Value
				.Col = C_RndPolicy2
				.Value = intIndex	
		End Select
'''''''''''''''''''''''''''''hmkwon modify end 
              
		If Col = C_Decimals2 Or Col = C_DataTypeNm2 Then
           Col = C_Decimals2 ''없음 안됨    
		  .Col = Col
		  .Row = Row
		  j = .value
    
		  lRound = 0.1
		  lRoundP = 1
    
		    If j > 0 Then
		        For i = 1 To j
		            lRound = lRound * 0.1
		        Next
		        
		        .Col = C_RndUnit2
		        .Row = Row
		        .value = lRound
		        
		    ElseIf j = 0 Then
		        .Col = C_RndUnit2
		        .Row = Row
		        .value = lRound
		        
		    Else
		        For i = 1 To (j * -1)
		            lRoundP = lRoundP * 10
		        Next
		        
		        lRoundP = lRoundP / 10
		        .Col = C_RndUnit2
		        .Row = Row
		        .value = lRoundP
		        
		    End If    

			
		.value = replace (.value, ".", "@")
		.value = replace (.value, ",", "$")
		.value = replace (.value, "@", parent.gComNumDec)
		.value = replace (.value, "$", parent.gComNum1000)
		
		    
''''''''''''''''''''''''''''''Srh추가/수정(2002/08/23)        
        .Col = C_DataTypeNm2   : .Row = Row   ''데이타종류 
        strCellText = Trim(.text)
        
		.Col = C_DataFormat2 : .Row = Row			    
		Select Case CInt(j)
		    Case -1
		        Select Case strCellText
		            Case "금액"
			            .Text = "#,###,###,###,##0"
			        Case "단가"
			            .Text = "##,###,###,##0"
			        Case "환율"
			            .Text = "###,###,##0"
			        Case Else
			            .Text = "#,###,###,###,##0"
			    End Select 
			Case -2
				Select Case strCellText
		            Case "금액"
			            .Text = "#,###,###,###,#00"
			        Case "단가"
			            .Text = "##,###,###,#00"
			        Case "환율"
			            .Text = "###,###,#00"
			        Case Else
			            .Text = "#,###,###,###,#00"
			    End Select 
			Case -3
				Select Case strCellText
		            Case "금액"
			            .Text = "#,###,###,###,000"
			        Case "단가"
			            .Text = "##,###,###,000"
			        Case "환율"
			            .Text = "###,###,000"
			        Case Else
			            .Text = "#,###,###,###,000"
			    End Select 
			Case -4
				Select Case strCellText
		            Case "금액"
			            .Text = "#,###,###,##0,000"
			        Case "단가"
			            .Text = "##,###,##0,000"
			        Case "환율"
			            .Text = "###,##0,000"
			        Case Else
			            .Text = "#,###,###,##0,000"
			    End Select 
			Case 0
			    Select Case strCellText   ''금액/단가/환율=13/11/9(정수부 자릿수)
			        Case "금액"
			            .Text = "#,###,###,###,###"
			        Case "단가"
			            .Text = "##,###,###,###"
			        Case "환율"
			            .Text = "###,###,###"
			        Case Else
			            .Text = "#,###,###,###,###"
			    End Select 					
			Case 1
			    Select Case strCellText
		            Case "금액"
			            .Text = "#,###,###,###,###.0"
			        Case "단가"
			            .Text = "##,###,###,###.0"
			        Case "환율"
			            .Text = "###,###,###.0"
			        Case Else
			            .Text = "#,###,###,###,###.0"
			    End Select 					
			Case 2
			    Select Case strCellText
		            Case "금액"
			            .Text = "#,###,###,###,###.00"
			        Case "단가"
			            .Text = "##,###,###,###.00"
			        Case "환율"
			            .Text = "###,###,###.00"
			        Case Else
			            .Text = "#,###,###,###,###.00"
			    End Select
			Case 3
			    Select Case strCellText
		            Case "금액"
			            .Text = "###,###,###,###.000"
			        Case "단가"
			            .Text = "##,###,###,###.000"
			        Case "환율"
			            .Text = "###,###,###.000"
			        Case Else
			            .Text = "###,###,###,###.000"
			    End Select					
			Case 4
			    Select Case strCellText
		            Case "금액"
			            .Text = "##,###,###,###.0000"
			        Case "단가"
			            .Text = "##,###,###,###.0000"
			        Case "환율"
			            .Text = "###,###,###.0000"
			        Case Else
			            .Text = "##,###,###,###.0000"
			    End Select					
			Case 5
			    Select Case strCellText
		            Case "금액"
			            .Text = "#,###,###,###.00000"
			        Case "단가"
			            .Text = "#,###,###,###.00000"
			        Case "환율"
			            .Text = "###,###,###.00000"
			        Case Else
			            .Text = "#,###,###,###.00000"
			    End Select					
			Case 6
			    Select Case strCellText
		            Case "금액"
			            .Text = "###,###,###.000000"
			        Case "단가"
			            .Text = "###,###,###.000000"
			        Case "환율"
			            .Text = "###,###,###.000000"
			        Case Else
			            .Text = "###,###,###.000000"
			    End Select									
			Case Else
				.Text = "#,###,###,###,###"					
		End Select
''''''''''''''''''''''''''''''Srh추가/수정(2002/08/23)


		  .text = replace (.text, ".", "@")
		  .text = replace (.text, ",", "$")
		  .text = replace (.text, "@", parent.gComNumDec)			'parent.gComNumDec
		  .text = replace (.text, "$", parent.gComNum1000)			'parent.gComNum1000
			
			
		End If											
		
    End with
       
End Sub

function FormatChanging(ByVal cnt,ByVal strCellText )
	
	dim strFormat
			    
	''''''''''''''''''''''''''''''Srh추가/수정(2002/08/23)        
        Select Case CInt(cnt)
		    Case -1
		        Select Case strCellText
		            Case "금액"
			            strFormat = "#,###,###,###,##0"
			        Case "단가"
			            strFormat = "##,###,###,##0"
			        Case "환율"
			            strFormat = "###,###,##0"
			        Case Else
			            strFormat = "#,###,###,###,##0"
			    End Select 
			Case -2
				Select Case strCellText
		            Case "금액"
			            strFormat = "#,###,###,###,#00"
			        Case "단가"
			            strFormat = "##,###,###,#00"
			        Case "환율"
			            strFormat = "###,###,#00"
			        Case Else
			            strFormat = "#,###,###,###,#00"
			    End Select 
			Case -3
				Select Case strCellText
		            Case "금액"
			            strFormat = "#,###,###,###,000"
			        Case "단가"
			            strFormat = "##,###,###,000"
			        Case "환율"
			            strFormat = "###,###,000"
			        Case Else
			            strFormat = "#,###,###,###,000"
			    End Select 
			Case -4
				Select Case strCellText
		            Case "금액"
			            strFormat = "#,###,###,##0,000"
			        Case "단가"
			            strFormat = "##,###,##0,000"
			        Case "환율"
			            strFormat = "###,##0,000"
			        Case Else
			            strFormat = "#,###,###,##0,000"
			    End Select 
			Case 0
			    Select Case strCellText   ''금액/단가/환율=13/11/9(정수부 자릿수)
			        Case "금액"
			            strFormat = "#,###,###,###,###"
			        Case "단가"
			            strFormat = "##,###,###,###"
			        Case "환율"
			            strFormat = "###,###,###"
			        Case Else
			            strFormat = "#,###,###,###,###"
			    End Select 					
			Case 1
			    Select Case strCellText
		            Case "금액"
			            strFormat = "#,###,###,###,###.0"
			        Case "단가"
			            strFormat = "##,###,###,###.0"
			        Case "환율"
			            strFormat = "###,###,###.0"
			        Case Else
			            strFormat = "#,###,###,###,###.0"
			    End Select 					
			Case 2
			    Select Case strCellText
		            Case "금액"
			            strFormat = "#,###,###,###,###.00"
			        Case "단가"
			            strFormat = "##,###,###,###.00"
			        Case "환율"
			            strFormat = "###,###,###.00"
			        Case Else
			            strFormat = "#,###,###,###,###.00"
			    End Select
			Case 3
			    Select Case strCellText
		            Case "금액"
			            strFormat = "###,###,###,###.000"
			        Case "단가"
			            strFormat = "##,###,###,###.000"
			        Case "환율"
			            strFormat = "###,###,###.000"
			        Case Else
			            strFormat = "###,###,###,###.000"
			    End Select					
			Case 4
			    Select Case strCellText
		            Case "금액"
			            strFormat = "##,###,###,###.0000"
			        Case "단가"
			            strFormat = "##,###,###,###.0000"
			        Case "환율"
			            strFormat = "###,###,###.0000"
			        Case Else
			            strFormat = "##,###,###,###.0000"
			    End Select					
			Case 5
			    Select Case strCellText
		            Case "금액"
			            strFormat = "#,###,###,###.00000"
			        Case "단가"
			            strFormat = "#,###,###,###.00000"
			        Case "환율"
			            strFormat = "###,###,###.00000"
			        Case Else
			            strFormat = "#,###,###,###.00000"
			    End Select					
			Case 6
			    Select Case strCellText
		            Case "금액"
			            strFormat = "###,###,###.000000"
			        Case "단가"
			            strFormat = "###,###,###.000000"
			        Case "환율"
			            strFormat = "###,###,###.000000"
			        Case Else
			            strFormat = "###,###,###.000000"
			    End Select									
			Case Else
				strFormat = "#,###,###,###,###"					
		End Select
''''''''''''''''''''''''''''''Srh추가/수정(2002/08/23)
	
	strFormat = replace (strFormat, ".", "@")
	strFormat = replace (strFormat, ",", "$")
	strFormat = replace (strFormat, "@", parent.gComNumDec)			'parent.gComNumDec
	strFormat = replace (strFormat, "$", parent.gComNum1000)			'parent.gComNum1000
		
	FormatChanging = strFormat            

End function

Sub vspdData_Click(ByVal Col, ByVal Row)
    Call SetPopupMenuItemInf("1101011111") 
    
    gMouseClickStatus = "SPC"   

    Set gActiveSpdSheet = frm1.vspdData
   
    If frm1.vspdData.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If

    If Row = 0 Then
        ggoSpread.Source = frm1.vspdData
        If lgSortKey = 1 Then
            ggoSpread.SSSort Col
            lgSortKey = 2
        Else
            ggoSpread.SSSort Col,lgSortKey
            lgSortKey = 1
        End If    
 
        Exit Sub
    End If

End Sub

Sub vspdData2_Click(ByVal Col, ByVal Row)
    
    Call SetPopupMenuItemInf("1101011111") 

    gMouseClickStatus = "SP1C"   

    Set gActiveSpdSheet = frm1.vspdData2
   
    If frm1.vspdData2.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If

    If Row = 0 Then
        ggoSpread.Source = frm1.vspdData2
        If lgSortKey = 1 Then
            ggoSpread.SSSort Col
            lgSortKey = 2
        Else
            ggoSpread.SSSort Col,lgSortKey
            lgSortKey = 1
        End If 
        Exit Sub   
    End If

End Sub

Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub

Sub vspdData2_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub

Sub vspdData_DblClick(ByVal Col, ByVal Row)				
    Dim iColumnName
    
    If Row <= 0 Then
        Exit Sub
    End If
    
    If frm1.vspdData.MaxRows = 0 Then
        Exit Sub
    End If

End Sub

Sub vspdData2_DblClick(ByVal Col, ByVal Row)				
    Dim iColumnName
    
    If Row <= 0 Then
        Exit Sub
    End If
    
    If frm1.vspdData2.MaxRows = 0 Then
        Exit Sub
    End If
	
End Sub

Sub vspdData_GotFocus()
    ggoSpread.Source = Frm1.vspdData
End Sub

Sub vspdData2_GotFocus()
    ggoSpread.Source = Frm1.vspdData2

End Sub

Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If

End Sub    

Sub vspdData2_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SP1C" Then
       gMouseClickStatus = "SP1CR"
    End If

End Sub    

Sub vspdData_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    
    Call GetSpreadColumnPos("A")
End Sub

Sub vspdData2_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    
    Call GetSpreadColumnPos("B")
End Sub

Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If

	If CheckRunningBizProcess = True Then
	   Exit Sub
	End If
    
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	           
    	If lgStrPrevKeyIndex <> "" Then                         
      	   Call DisableToolBar(Parent.TBC_QUERY)
			If DbQuery = False Then
				Call RestoreTooBar()
			    Exit Sub
			End If  				
    	End If
    End if
End Sub

Sub vspdData2_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If

	If CheckRunningBizProcess = True Then
	   Exit Sub
	End If
    
    if frm1.vspdData2.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData2,NewTop) Then	           
    	If lgStrPrevKeyIndex <> "" Then                         
      	   Call DisableToolBar(Parent.TBC_QUERY)
			If DbQuery = False Then
				Call RestoreTooBar()
			    Exit Sub
			End If  				
    	End If
    End if
End Sub

Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub

Sub PopRestoreSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()

    Select Case gActiveSpdSheet.id
		Case "vaSpread"
			Call InitSpreadSheet("A")
		Case "vaSpread2"
			Call InitSpreadSheet("B")      		
	End Select 

    Call InitSpreadComboBox

    ggoSpread.Source = gActiveSpdSheet
	Call ggoSpread.ReOrderingSpreadData()
	
	Call InitData()
	
End Sub

Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
Dim strTemp
Dim intPos1
   
	With frm1.vspdData 
	
    ggoSpread.Source = frm1.vspdData
   
    If Row > 0 And Col = C_Currency + 1 Then
        .Col = Col - 1
        .Row = Row
        
        Call OpenCurrency(1)
        
    End If
    
    End With
      
End Sub

Sub vspdData2_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
Dim strTemp
Dim intPos1
   
	With frm1.vspdData2 
	
    ggoSpread.Source = frm1.vspdData2
   
    If Row > 0 And Col = C_Currency2 + 1 Then
        .Col = Col - 1
        .Row = Row
        
        Call OpenCurrency(2)
        
    End If
    
    End With
      
End Sub


Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If CheckRunningBizProcess = True Then					'☜: 조회중이면 다음 조회 안하도록 체크 
        Exit Sub
	End If
       
    If OldLeft <> NewLeft Then
        Exit Sub
    End If

End Sub

Sub vspdData2_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If CheckRunningBizProcess = True Then					'☜: 조회중이면 다음 조회 안하도록 체크 
        Exit Sub
	End If
	
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
 
End Sub

Function FncQuery() 
    Dim IntRetCD 
    
    FncQuery = False                                                        '⊙: Processing is NG
    
    Err.Clear                                                               '☜: Protect system from crashing
	
<%  '-----------------------
    'Check previous data area
    '----------------------- %>        
    If gSelframeFlg = TAB1 Then
    	ggoSpread.Source = frm1.vspdData
    	If ggoSpread.SSCheckChange = True Then
			IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "X", "X")			    <%'데이타가 변경되었습니다. 조회하시겠습니까?%>
    		If IntRetCD = vbNo Then
      			Exit Function
    		End If
    	End If
    Else
    	ggoSpread.Source = frm1.vspdData2
    	If ggoSpread.SSCheckChange = True Then
			IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "X", "X")			    <%'데이타가 변경되었습니다. 조회하시겠습니까?%>
    		If IntRetCD = vbNo Then
      			Exit Function
    		End If
    	End If
	End If            
	
    '-----------------------
    'Erase contents area
    '-----------------------    
    Call ggoOper.ClearField(Document, "2")										'⊙: Clear Contents  Field
    ggoSpread.ClearSpreadData

    Call InitVariables	
    
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then									'⊙: This function check indispensable field
       Exit Function
    End If
    
    '-----------------------
    'Query function call area
    '-----------------------
    If DbQuery = false Then														<%'☜: Query db data%>
		Exit Function 
	End If																'☜: Query db data
	       
    FncQuery = True																'⊙: Processing is OK
    
End Function

Function FncNew() 
    Dim IntRetCD 
    
    FncNew = False                                                          '⊙: Processing is NG
    
    Err.Clear                                                               '☜: Protect system from crashing
    On Error Resume Next                                                    '☜: Protect system from crashing
        
    '-----------------------
    'Check previous data area
    '-----------------------
    If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900015", parent.VB_YES_NO, "X", "X") '☜ 바뀐부분 
		If IntRetCD = vbNo Then
			Exit Function
		End If
       
    End If
    
    '-----------------------
    'Erase condition area
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "A")                                        
    

    ggoSpread.Source = frm1.vspdData
    ggoSpread.ClearSpreadData
    
    ggoSpread.Source = frm1.vspdData2
    ggoSpread.ClearSpreadData

    Call ggoOper.LockField(Document, "N")                                          '⊙: Lock  Suitable  Field
    Call InitVariables                                                      '⊙: Initializes local global variables
        
    FncNew = True                                                           '⊙: Processing is OK

End Function

Function FncDelete() 
    Dim IntRetCD 
    
    FncDelete = False                                                       '⊙: Processing is NG
    
    Err.Clear                                                               '☜: Protect system from crashing
    On Error Resume Next                                                    '☜: Protect system from crashing
    
    
    '-----------------------
    'Precheck area
    '-----------------------
    If lgIntFlgMode <> parent.OPMD_UMODE Then                                      'Check if there is retrived data
        Call DisplayMsgBox("900002", "X", "X", "X")                                  '☆:
        Exit Function
    End If
    
    IntRetCD = DisplayMsgBox("900003", parent.VB_YES_NO, "X", "X")  '☜ 바뀐부분 
    If IntRetCD = vbNo Then
        Exit Function
    End If
    
    '-----------------------
    'Delete function call area
    '-----------------------
    If DbDelete = False Then                                                '☜: Delete db data
       Exit Function                                                        '☜:
    End If
    
    '-----------------------
    'Erase condition area
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "A")    

    ggoSpread.Source = frm1.vspdData
    ggoSpread.ClearSpreadData
    
    ggoSpread.Source = frm1.vspdData2
    ggoSpread.ClearSpreadData

    FncDelete = True                                                        '⊙: Processing is OK
    
End Function

Function FncSave()     
    
    FncSave = False                                                         '⊙: Processing is NG
    
    Err.Clear                                                               '☜: Protect system from crashing
    On Error Resume Next                                                    '☜: Protect system from crashing
    
    frm1.txtLogInCnt.value = "0"	'현재로그인 user를 초기화 시킴 
    
    'Precheck area
    '-----------------------
    If gSelframeFlg = TAB1 Then
		ggoSpread.Source = frm1.vspdData
		If ggoSpread.SSCheckChange = False Then
			Call DisplayMsgBox("900001", "X", "X", "X")                            '⊙: No data changed!!        
			Exit Function
		End If
		
		If Not ggoSpread.SSDefaultCheck Then  		'Not chkField(Document, "2") Or
			Call changeTabs(TAB1)
			Exit Function
		End If
	Else
		ggoSpread.Source = frm1.vspdData2
		If ggoSpread.SSCheckChange = False Then
			Call DisplayMsgBox("900001", "X", "X", "X")                            '⊙: No data changed!!        
			Exit Function
		End If
		
		If Not ggoSpread.SSDefaultCheck Then  		'Not chkField(Document, "2") Or
			Call changeTabs(TAB2)
			Exit Function
		End If
	End If
    
    '-----------------------
    'Save function call area
    '-----------------------
    If DbSave = False Then				                                                  '☜: Save db data
		Exit Function
	End If
	
    FncSave = True                                                          '⊙: Processing is OK
    
End Function

Function FncCopy() 

    FncCopy = False                                                               '☜: Processing is NG

	If gSelframeFlg = TAB1 Then

        If Frm1.vspdData.MaxRows < 1 Then
           Exit Function
        End If

        With frm1.vspdData
		    If .ActiveRow > 0 Then
		    	.focus
		    	.ReDraw = false
		
		    	ggoSpread.Source = frm1.vspdData	
		    	ggoSpread.CopyRow

    		    SetSpreadColor .ActiveRow, .ActiveRow
    
		    	.Col=C_DataTypeNm
		    	.Text=""
		    	.Col=C_Currency
		    	.Text= ""
		    	.Col=C_ModuleNm
		    	.Text=""
		    	.Col=C_FormTypeNm
		    	.Text=""		
		        		    
		    	.ReDraw = true
		    End If
		End with
    Else
        If Frm1.vspdData2.MaxRows < 1 Then
           Exit Function
        End If

    	With frm1.vspdData2
    		If .ActiveRow > 0 Then
    			.focus
    			.ReDraw = false
    	    	
    			ggoSpread.Source = frm1.vspdData2
    			ggoSpread.CopyRow

    			SetSpreadColor .ActiveRow, .ActiveRow
         
    			.Col=C_DataTypeNm2
    			.Text=""
    			.Col=C_ModuleNm2
    			.Text=""
    			.Col=C_FormTypeNm2
    			.Text=""		
    	    	    		    
    			.ReDraw = true
    		End If
    	End with
    End If
   
    If Err.number = 0 Then	
       FncCopy = True                                                            '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement   
   
End Function

Function FncCancel()
Dim lRows  
	If gSelframeFlg = TAB1 Then
		ggoSpread.Source = frm1.vspdData
		lRows = frm1.vspdData.ActiveRow
		ggoSpread.EditUndo                                                  '☜: Protect system from crashing    
		Call InitData(lRows)
	Else
		ggoSpread.Source = frm1.vspdData2
		lRows = frm1.vspdData2.ActiveRow
		ggoSpread.EditUndo 	                                                '☜: Protect system from crashing
		Call InitData(lRows)
	End If

End Function

Function FncInsertRow(ByVal pvRowCnt)
    Dim imRow
    Dim iRow
    
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status
 
    FncInsertRow = False                                                         '☜: Processing is NG

    If IsNumeric(Trim(pvRowCnt)) Then
        imRow = CInt(pvRowCnt)
    Else
        imRow = AskSpdSheetAddRowCount()
        If imRow = "" Then
            Exit Function
        End If
    End If
    
	If gSelframeFlg = TAB1 Then
		With frm1
	
		.vspdData.ReDraw = False
    	.vspdData.focus
		ggoSpread.Source = .vspdData
	
	    ggoSpread.InsertRow .vspdData.ActiveRow, imRow
        SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1

		For iRow = .vspdData.ActiveRow To .vspdData.ActiveRow + imRow - 1
		    .vspdData.Row = iRow
		    
    		.vspdData.Col = C_Currency
    		.vspdData.Text = ""

    		.vspdData.Col = C_ModuleCd
    		.vspdData.Text = .cboModuleCd.value
    		
    		.vspdData.Col = C_FormType
    		.vspdData.Text = .cboFormType.value
    		
    		.vspdData.Col = C_RndUnit
            .vspdData.text = "0.1"
    		.vspdData.text = replace (.vspdData.text, ".", "@")		'***********************************020422수정 
    		.vspdData.text = replace (.vspdData.text, ",", "$")
    		.vspdData.text = replace (.vspdData.text, "@", parent.gComNumDec)
    		.vspdData.text = replace (.vspdData.text, "$", parent.gComNum1000)
        

    		.vspdData.Col = C_DataFormat    
    	    .vspdData.text = "#,###,###,###,###"

    		.vspdData.text = replace (.vspdData.text, ".", "@")		'***********************************020422수정 
    		.vspdData.text = replace (.vspdData.text, ",", "$")
    		.vspdData.text = replace (.vspdData.text, "@", parent.gComNumDec)			'parent.gComNumDec
    		.vspdData.text = replace (.vspdData.text, "$", parent.gComNum1000)			'parent.gComNum1000
        Next 
    		.vspdData.ReDraw = True
		End With
    Else
		With frm1
	
		.vspdData2.ReDraw = False
		.vspdData2.focus
		ggoSpread.Source = .vspdData2
        ggoSpread.InsertRow .vspdData2.ActiveRow, imRow
        SetSpreadColor .vspdData2.ActiveRow, .vspdData2.ActiveRow + imRow - 1

		For iRow = .vspdData2.ActiveRow To .vspdData2.ActiveRow + imRow - 1
            		    
            .vspdData2.Row = iRow		

		    .vspdData2.Col = C_Currency2
		    .vspdData2.Text = ""
		    	
		    .vspdData2.Col = C_FormType2
		    .vspdData2.Text = .cboFormType.value
		
		    .vspdData2.Col = C_RndUnit2
            .vspdData2.text = "0.1"
		    .vspdData2.text = replace (.vspdData2.text, ".", "@")		'***********************************020422수정 
		    .vspdData2.text = replace (.vspdData2.text, ",", "$")
		    .vspdData2.text = replace (.vspdData2.text, "@", parent.gComNumDec)
		    .vspdData2.text = replace (.vspdData2.text, "$", parent.gComNum1000)
    
            .vspdData2.Col = C_DataFormat2
	        .vspdData2.text = "#,###,###,###,###"

		    .vspdData2.text = replace (.vspdData2.text, ".", "@")		'***********************************020422수정 
		    .vspdData2.text = replace (.vspdData2.text, ",", "$")
		    .vspdData2.text = replace (.vspdData2.text, "@", parent.gComNumDec)			'parent.gComNumDec
		    .vspdData2.text = replace (.vspdData2.text, "$", parent.gComNum1000)			'parent.gComNum1000
		Next
	
		.vspdData2.ReDraw = True
    	End With
    End IF
    
    If Err.number = 0 Then
       FncInsertRow = True                                                          '☜: Processing is OK
    End If   
	
    Set gActiveElement = document.ActiveElement   

End Function

Function FncDeleteRow() 
    Dim lDelRows
    Dim iDelRowCnt, i
    
    If gSelframeFlg = TAB1 Then
		With frm1.vspdData 
    
		.focus
		ggoSpread.Source = frm1.vspdData 
    
		lDelRows = ggoSpread.DeleteRow
    
		End With
	Else
		With frm1.vspdData2
    
		.focus
		ggoSpread.Source = frm1.vspdData2
    
		lDelRows = ggoSpread.DeleteRow
    
		End With
	End If
End Function

Function FncPrint() 
    Call parent.FncPrint()                                                   '☜: Protect system from crashing
End Function

Function FncExcel() 
    Call parent.FncExport(parent.C_MULTI)                                                   <%'☜: Protect system from crashing%>
End Function

Function FncFind() 
    Call parent.FncFind(parent.C_MULTI, False)                                         <%'☜:화면 유형, Tab 유무 %>
End Function

Function FncExit()
	Dim IntRetCD
	FncExit = False
    
    If gSelframeFlg = TAB1 Then	
		ggoSpread.source = frm1.vspdData
		If ggoSpread.SSCheckChange = True Then
			IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO, "X", "X")			<%'⊙: "Will you destory previous data"%>
			If IntRetCD = vbNo Then
				Exit Function
			End If
		End If
	Else
		ggoSpread.source = frm1.vspdData2
		If ggoSpread.SSCheckChange = True Then
			IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO, "X", "X")			<%'⊙: "Will you destory previous data"%>
			If IntRetCD = vbNo Then
				Exit Function
			End If
		End If
	End If
    FncExit = True
End Function

Function DbQuery() 
    Dim LngLastRow      
    Dim LngMaxRow       
    Dim LngRow          
    Dim strTemp         
    Dim StrNextKey      
    
    
    DbQuery = False
    
    Call LayerShowHide(1)
    
    Err.Clear                                                               '☜: Protect system from crashing

	Dim strVal
    
    With frm1 
    
    If gSelframeFlg = Tab1 Then
    
		If lgIntFlgMode = parent.OPMD_UMODE Then
		strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001							'☜: 
		StrVal = strVal & "&FrmFlag="	  & "1"
		strVal = strVal & "&cboDataType=" & .hDataType.value 				'☆: 조회 조건 데이타 
		strVal = strVal & "&txtCurrency=" & .hCurrency.value 
		strVal = strVal & "&cboModuleCd=" & .hModuleCd.value 				'☆: 조회 조건 데이타 
		strVal = strVal & "&cboFormType=" & .hFormType.value 
		strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows   
		Else
		strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001							'☜: 
		StrVal = strVal & "&FrmFlag="	  & "1"
		strVal = strVal & "&cboDataType=" & Trim(.cboDataType.value)				'☆: 조회 조건 데이타 
		
		If "" = Trim(.txtCurrency.value) Then
			strVal = strVal & "&txtCurrency=" & ""'EP-1513라인 parent.gCurrency 020215창 
		Else
			strVal = strVal & "&txtCurrency=" & Trim(.txtCurrency.value)
		End If	
		
		strVal = strVal & "&cboModuleCd=" & Trim(.cboModuleCd.value)				'☆: 조회 조건 데이타 
		strVal = strVal & "&cboFormType=" & Trim(.cboFormType.value)
		strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
		End If    
    
    Else    
	    If lgIntFlgMode = parent.OPMD_UMODE Then
		strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001							'☜: 
		StrVal = strVal & "&FrmFlag="	  & "2"
		strVal = strVal & "&cboDataType=" & .hDataType.value 				'☆: 조회 조건 데이타 
		strVal = strVal & "&txtCurrency=" & .hCurrency.value 
		strVal = strVal & "&cboModuleCd=" & .hModuleCd.value 				'☆: 조회 조건 데이타 
		strVal = strVal & "&cboFormType=" & .hFormType.value 
		strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows 
		
		Else
		strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001							'☜: 
		StrVal = strVal & "&FrmFlag="	  & "2"
		strVal = strVal & "&cboDataType=" & Trim(.cboDataType.value)				'☆: 조회 조건 데이타 
		
		If "" = Trim(.txtCurrency.value) Then
			strVal = strVal & "&txtCurrency=" & ""	                                  'parent.gCurrency 020215창 
		Else
			strVal = strVal & "&txtCurrency=" & Trim(.txtCurrency.value)
		End If
		
		strVal = strVal & "&cboModuleCd=" & Trim(.cboModuleCd.value)				'☆: 조회 조건 데이타 
		strVal = strVal & "&cboFormType=" & Trim(.cboFormType.value)
		strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
		End If
    
    End If   
    
   
	Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 
        
    End With
    
    DbQuery = True    

End Function

Function DbQueryOk(lngMaxRow)														'☆: 조회 성공후 실행로직 
	
    '-----------------------
    'Reset variables area
    '-----------------------
    lgIntFlgMode = parent.OPMD_UMODE												'⊙: Indicates that current mode is Update mode
    
	Call SetToolbar("1100111100111111")									'⊙: 버튼 툴바 제어 
    InitData(lngMaxRow)
    
    If gSelframeFlg = TAB1 Then
		frm1.cboModuleCd.value = "*"
		frm1.cboModuleCd.disabled = True
    Else
		If frm1.hModuleCd.value = "" Then
			frm1.cboModuleCd.value = ""
		End If
		frm1.cboModuleCd.disabled = False
    End If

End Function

Sub InitData(ByVal lngStartRow)
Dim intRow, intCol
Dim intIndex 
	If gSelframeFlg = TAB1 Then	
		With frm1.vspdData
			For intRow = lngStartRow To .MaxRows
			
				.Row = intRow
			
				For intCol = 1 To .MaxCols
			
					Select Case intCol
						Case C_DataType
							.Col = C_DataType
							intIndex = .value
							.col = C_DataTypeNm
							.value = intindex
						Case C_ModuleCD
							.Col = C_ModuleCD
							intIndex = .value
							.col = C_ModuleNm
							.value = intindex
						Case C_FormType
							.Col = C_FormType
							intIndex = .value
							.col = C_FormTypeNm
							.value = intindex	
						Case C_RndPolicy
							.Col = C_RndPolicy
							intIndex = .value
							.col = C_RndPolicyNm
							.value = intindex		
					End Select
					
				Next
			
			Next	
		End With
	Else
		With frm1.vspdData2
			For intRow = lngStartRow To .MaxRows
			
				.Row = intRow
			
				For intCol = 1 To .MaxCols
			
					Select Case intCol
						Case C_DataType2
							.Col = C_DataType2
							intIndex = .value
							.col = C_DataTypeNm2
							.value = intindex
						Case C_ModuleCD2
							.Col = C_ModuleCD2
							intIndex = .value
							.col = C_ModuleNm2
							.value = intindex
						Case C_FormType2
							.Col = C_FormType2
							intIndex = .value
							.col = C_FormTypeNm2
							.value = intindex	
						Case C_RndPolicy2
							.Col = C_RndPolicy2
							intIndex = .value
							.col = C_RndPolicyNm2
							.value = intindex		
					End Select
					
				Next
			
			Next	
		End With
	End If
End Sub

Function DbSave() 
    Dim lRow        
    Dim lGrpCnt     
    Dim retVal      
    Dim lStartRow   
    Dim lEndRow     
    Dim lRestGrpCnt 
	Dim strVal, strDel
	Dim x, y
	Dim strTempDataFormat
	Dim strCellText 
	
    DbSave = False                                                          '⊙: Processing is NG
    
    Call LayerShowHide(1)
    
    On Error Resume Next                                                   '☜: Protect system from crashing

	With frm1
		.txtMode.value = parent.UID_M0002
		.txtInsrtUserId.value = parent.gUsrID
		
    
    '-----------------------
    'Data manipulate area
    '-----------------------
    lGrpCnt = 1    
    strVal = ""
    strDel = ""
    
    '-----------------------
    'Data manipulate area
    '-----------------------
  If gSelframeFlg = TAB1 Then
		ggoSpread.Source = .vspdData
	For lRow = 1 To .vspdData.MaxRows
    
        .vspdData.Row = lRow
        .vspdData.col = C_DataTypeNm  ''1
        strCellText = Trim(.vspdData.Text)
        
        .vspdData.Row = lRow
        .vspdData.col = 0
        Select Case .vspdData.Text

            Case ggoSpread.InsertFlag											'☜: 신규 
				
				strVal = strVal & "C" & parent.gColSep & lRow & parent.gColSep				'☜: C=Create
                
                .vspdData.col = C_DataType	'2
                strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                x = Trim(.vspdData.Text)
                
                .vspdData.Col = C_Currency	'3
                strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                
                .vspdData.Col = C_ModuleCd	'6
                strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                
                .vspdData.Col = C_FormType	'8
                strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                                                
                .vspdData.Col = C_Decimals	'9
                strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                y = CInt(Trim(.vspdData.Text))
                
                If x = "2" and y > 2 Then
					Call LayerShowHide(0)
					Call DisplayMsgBox("122104", "X", "X", "X")
					.vspdData.Action = 0
					Exit Function
				Elseif x = "4" and y > 6 Then
					Call LayerShowHide(0)
					Call DisplayMsgBox("122105", "X", "X", "X")
					.vspdData.Action = 0
					Exit Function
				End If
                
                
                .vspdData.Col = C_RndUnit	'10
                strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & parent.gColSep  'Round Unit     
                
                .vspdData.Col = C_RndPolicy	'12                
                strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                
                ''.vspdData.Col = C_DataFormat'13
                .vspdData.Col = C_Decimals'13
                strTempDataFormat = FormatChanging(Trim(.vspdData.Text),strCellText)
		        strVal = strVal & strTempDataFormat & parent.gRowSep      
                
                lGrpCnt = lGrpCnt + 1
                
            Case ggoSpread.UpdateFlag											'☜: 신규 

				strVal = strVal & "U" & parent.gColSep & lRow & parent.gColSep				'☜: U=Update

                .vspdData.Col = C_DataType	'2
                strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                x = Trim(.vspdData.Text)
                
                .vspdData.Col = C_Currency	'3
                strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                
                .vspdData.Col = C_ModuleCd	'6
                strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                
                .vspdData.Col = C_FormType	'8
                strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                                
                .vspdData.Col = C_Decimals	'9
                strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                y = CInt(Trim(.vspdData.Text))
                
                If x = "2" and y > 2 Then
					Call LayerShowHide(0)
					Call DisplayMsgBox("122104", "X", "X", "X")					
					.vspdData.Action = 0
					Exit Function
				Elseif x = "4" and y > 6 Then
					Call LayerShowHide(0)
					Call DisplayMsgBox("122105", "X", "X", "X")
					.vspdData.Action = 0
					Exit Function
				End If
                 
                .vspdData.Col = C_RndUnit	'10
                strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & parent.gColSep  'Round Unit     
                 
                .vspdData.Col = C_RndPolicy	'12                                
                strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
               
                ''.vspdData.Col = C_DataFormat'13
                .vspdData.Col = C_Decimals'13
                strTempDataFormat = FormatChanging(Trim(.vspdData.Text),strCellText)
		        strVal = strVal & strTempDataFormat & parent.gRowSep      
                
                lGrpCnt = lGrpCnt + 1
                
            Case ggoSpread.DeleteFlag													'☜: 삭제 

				
				strDel = strDel & "D" & parent.gColSep & lRow & parent.gColSep				'☜: D=Delete

                .vspdData.Col = C_DataType	'2
                strDel = strDel & Trim(.vspdData.Text) & parent.gColSep
                
                .vspdData.Col = C_Currency	'3
                strDel = strDel & Trim(.vspdData.Text) & parent.gColSep
                
                .vspdData.Col = C_ModuleCd	'6
                strDel = strDel & Trim(.vspdData.Text) & parent.gColSep
                
                .vspdData.Col = C_FormType	'8
                strDel = strDel & Trim(.vspdData.Text) & parent.gRowSep
                
                lGrpCnt = lGrpCnt + 1
                
        End Select
                
    Next
  Else  
	ggoSpread.Source = .vspdData2
    For lRow = 1 To .vspdData2.MaxRows
		
        .vspdData2.Row = lRow
        .vspdData2.col = 0
        
        Select Case .vspdData2.Text

            Case ggoSpread.InsertFlag											'☜: 신규 
				
				strVal = strVal & "C" & parent.gColSep & lRow & parent.gColSep				'☜: C=Create
                
                .vspdData2.col = C_DataType2	'2
                strVal = strVal & Trim(.vspdData2.Text) & parent.gColSep
                x = Trim(.vspdData2.Text)
                
                .vspdData2.Col = C_Currency2	'3
                strVal = strVal & Trim(.vspdData2.Text) & parent.gColSep
                
                .vspdData2.Col = C_ModuleCd2	'6
                strVal = strVal & Trim(.vspdData2.Text) & parent.gColSep
                
                .vspdData2.Col = C_FormType2	'8
                strVal = strVal & Trim(.vspdData2.Text) & parent.gColSep
                                                                
                .vspdData2.Col = C_Decimals2	'9
                strVal = strVal & Trim(.vspdData2.Text) & parent.gColSep                
                y = CInt(Trim(.vspdData2.Text))
                
                If x = "2" and y > 2 Then
					Call LayerShowHide(0)
					Call DisplayMsgBox("122104", "X", "X", "X")
					.vspdData2.Action = 0
					Exit Function
				Elseif x = "4" and y > 6 Then
					Call LayerShowHide(0)
					Call DisplayMsgBox("122105", "X", "X", "X")
					.vspdData2.Action = 0
					Exit Function
				End If
                
                .vspdData2.Col = C_RndUnit2	'10                
                strVal = strVal & UNIConvNum(Trim(.vspdData2.Text),0) & parent.gColSep  'Round Unit    
                
                .vspdData2.Col = C_RndPolicy2	'12
                strVal = strVal & Trim(.vspdData2.Text) & parent.gColSep
                
                ''.vspdData2.Col = C_DataFormat2'13
                .vspdData2.Col = C_Decimals2'13
                strTempDataFormat = FormatChanging(Trim(.vspdData2.Text),strCellText)
		        strVal = strVal & strTempDataFormat & parent.gRowSep      
                
                lGrpCnt = lGrpCnt + 1
                
            Case ggoSpread.UpdateFlag											'☜: 신규 

				strVal = strVal & "U" & parent.gColSep & lRow & parent.gColSep				'☜: U=Update

                .vspdData2.Col = C_DataType2	'2
                strVal = strVal & Trim(.vspdData2.Text) & parent.gColSep
                x = Trim(.vspdData2.Text)
                
                .vspdData2.Col = C_Currency2	'3
                strVal = strVal & Trim(.vspdData2.Text) & parent.gColSep
                
                .vspdData2.Col = C_ModuleCd2	'6
                strVal = strVal & Trim(.vspdData2.Text) & parent.gColSep
                
                .vspdData2.Col = C_FormType2	'8
                strVal = strVal & Trim(.vspdData2.Text) & parent.gColSep
                
                .vspdData2.Col = C_Decimals2	'9
                strVal = strVal & Trim(.vspdData2.Text) & parent.gColSep                
                
                y = CInt(Trim(.vspdData2.Text))
                
                If x = "2" and y > 2 Then
					Call LayerShowHide(0)
					Call DisplayMsgBox("122104", "X", "X", "X")
					.vspdData2.Action = 0
					Exit Function
				Elseif x = "4" and y > 6 Then
					Call LayerShowHide(0)
					Call DisplayMsgBox("122105", "X", "X", "X")
					.vspdData2.Action = 0
					Exit Function
				End If
				
                .vspdData2.Col = C_RndUnit2	'10
                strVal = strVal & UNIConvNum(Trim(.vspdData2.Text),0) & parent.gColSep  'Round Unit    
                
                .vspdData2.Col = C_RndPolicy2	'12
                strVal = strVal & Trim(.vspdData2.Text) & parent.gColSep
                
                ''.vspdData2.Col = C_DataFormat2'13
                .vspdData2.Col = C_Decimals2'13
                strTempDataFormat = FormatChanging(Trim(.vspdData2.Text),strCellText)
		        strVal = strVal & strTempDataFormat & parent.gRowSep      
                
                lGrpCnt = lGrpCnt + 1
                
            Case ggoSpread.DeleteFlag													'☜: 삭제 

				strDel = strDel & "D" & parent.gColSep & lRow & parent.gColSep				'☜: D=Delete

                .vspdData2.Col = C_DataType2	'2
                strDel = strDel & Trim(.vspdData2.Text) & parent.gColSep
                
                .vspdData2.Col = C_Currency2	'3
                strDel = strDel & Trim(.vspdData2.Text) & parent.gColSep
                
                .vspdData2.Col = C_ModuleCd2	'6
                strDel = strDel & Trim(.vspdData2.Text) & parent.gColSep
                
                .vspdData2.Col = C_FormType2	'8
                strDel = strDel & Trim(.vspdData2.Text) & parent.gRowSep
                
                lGrpCnt = lGrpCnt + 1
                
        End Select
                
    Next
  End If
  
	.txtMaxRows.value = lGrpCnt-1
	.txtSpread.value = strDel & strVal		
	
	Call ExecMyBizASP(frm1, BIZ_PGM_ID)										'☜: 비지니스 ASP 를 가동 
	
	End With
	
    DbSave = True                                                           '⊙: Processing is NG
    
End Function

Function CheckValidate(ByVal strData, ByVal intCol, ByVal lngRow)

	Dim sName

	CheckValidate = False

	If gSelframeFlg = TAB1 Then	
		With frm1.vspdData
		
			.Col = intCol
			.Row = 0
			sName = .text

			If len(strData) > 6 Then
				Call DisplayMsgBox("970025", "X", sName, "6")
				'Call DisplayMsgBox("970025", "X", "소수점 자리수", "6")
				.Col = intCol
				.Row = lngRow
				.Action = 0
			    .EditMode = True
				Exit Function
			ElseIf len(strData) < -4 Then
				Call DisplayMsgBox("970023", "X", sName, "-4")
				'Call DisplayMsgBox("970023", "X", "소수점 자리수", "-4")
				.Col = intCol
				.Row = lngRow
				.Action = 0
				.EditMode = True
				Exit Function
			End If
		
		End With
	Else
		With frm1.vspdData2
		
			.Col = intCol
			.Row = 0
			sName = .text

			If len(strData) > 6 Then
				Call DisplayMsgBox("970025", "X", sName, "6")
				'Call DisplayMsgBox("970025", , "소수점 자리수", "6")
				.Col = intCol
				.Row = lngRow
				.Action = 0
			    .EditMode = True
				Exit Function
			ElseIf len(strData) < -4 Then
				Call DisplayMsgBox("970023", "X", sName, "-4")
				'Call DisplayMsgBox("970023", , "소수점 자리수", "-4")
				.Col = intCol
				.Row = lngRow
				.Action = 0
				.EditMode = True
				Exit Function
			End If
		
		End With
	End If
	
	CheckValidate = True
	
End Function

Function DbSaveOk()													'☆: 저장 성공후 실행 로직 
   
	Call InitVariables
	frm1.vspdData.MaxRows = 0
	ggoSpread.Source = frm1.vspdData
	ggoSpread.ClearSpreadData

	frm1.vspdData2.MaxRows = 0
	ggoSpread.Source = frm1.vspdData2
	ggoSpread.ClearSpreadData
	
    Call MainQuery()

End Function

Function CheckLogInUser() 
    Dim IntRetCD 
	Dim strLogInCnt
	Dim arrRet
	Dim arrParam(5)
    Dim tempMsg
    Dim iCalledAspName
    
	arrParam(0) = ""
	arrParam(1) = ""
    
    Err.Clear			
    strLogInCnt = Cint(frm1.txtLogInCnt.value)
    
    tempMsg = "접속중인 사용자가 존재하므로 저장할 수 없습니다 " & vbCrLf
    tempMsg = tempMsg & "이 자료는 시스템관리자 1명만 접속했을 때 저장할 수 있습니다" & vbCrLf
    tempMsg = tempMsg & "접속중인 사용자 정보를 보시겠습니까?"
      
    intRetCD = MsgBox(tempMsg,vbExclamation + vbYesNo, gLogoName & "-[Warning]")
    
    If IntRetCD = vbNo Then
		Exit Function
	End If

	iCalledAspName = AskPRAspName("LoginUserList")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "LoginUserList", "X")
		lgIsOpenPop = False
		Exit Function
	End If


	arrRet = window.showModalDialog(iCalledAspName,Array(window.parent,arrParam),, "dialogWidth=400px; dialogHeight=600px; center: Yes; help: No; resizable: No; status: No;")


End Function

Function DbDelete() 
End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>

<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">

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
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="ClickTab1()">
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><IMG height=23 src="../../../CShared/image/table/seltab_up_left.gif" width=9></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>Numeric포맷(입력)</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><IMG height=23 src="../../../CShared/image/table/seltab_up_right.gif" width=10></td>
						    </TR>
						</TABLE>
					</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="ClickTab2()">
							<TR>
								<td background="../../../CShared/image/table/tab_up_bg.gif"><IMG height=23 src="../../../CShared/image/table/tab_up_left.gif" width=9></td>
								<td background="../../../CShared/image/table/tab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>Numeric포맷(조회)</font></td>
								<td background="../../../CShared/image/table/tab_up_bg.gif" align="right"><IMG height=23 src="../../../CShared/image/table/tab_up_right.gif" width=10></td>
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
					<FIELDSET CLASS="CLSFLD"><TABLE <%=LR_SPACE_TYPE_40%>>
					<TR>
            		 	<TD CLASS="TD5">데이터종류</TD>
            			<TD CLASS="TD6"><SELECT NAME="cboDataType" tag="11X" STYLE="WIDTH:160px;" ALT="데이터종류"></SELECT></TD>
            			<TD CLASS="TD5">통 화</TD>
            			<TD CLASS="TD6"><INPUT TYPE=TEXT NAME="txtCurrency" SIZE=10  MAXLENGTH=3 tag="11XXXU" ALT="통화코드"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCurrencyCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenCurrency(0)">
										<INPUT TYPE=TEXT NAME="txtCurrencyNm" tag="14X">
						</TD>
            		</TR>
            		<TR>
            			<TD CLASS="TD5">업무</TD>
            			<TD CLASS="TD6"><SELECT NAME="cboModuleCd"tag="11X" STYLE="WIDTH: 160px;"><OPTION value=""></OPTION></SELECT></TD>
            			<TD CLASS="TD5">화면종류</TD>
            			<TD CLASS="TD6"><SELECT NAME="cboFormType"tag="14X" STYLE="WIDTH: 160px;"><OPTION value=""></OPTION></SELECT></TD>
            		</TR>
						</TABLE></FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=* VALIGN=TOP>
	
				<!-- 첫번째 탭 내용 -->
					<DIV ID="TabDiv" STYLE="DISPLAY: none" SCROLL=no>
					<TABLE <%=LR_SPACE_TYPE_20%>>
						<TR>
							<TD HEIGHT="100%">
								<script language =javascript src='./js/b1903ma1_vaSpread_vspdData.js'></script>
							</TD>
						</TR>
					</TABLE>
					</DIV>	

				<!-- 두번째 탭 내용 -->
					<DIV ID="TabDiv" STYLE="DISPLAY: none" SCROLL=no>
					<TABLE <%=LR_SPACE_TYPE_20%>>
						<TR>
							<TD HEIGHT="100%">
								<script language =javascript src='./js/b1903ma1_vaSpread2_vspdData2.js'></script>
							</TD>
						</TR>
					</TABLE>
					</DIV>	
		</TABLE></TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_01%>></TD>
	</TR>
	<TR HEIGHT=20>
			<TD WIDTH=100%>
				<TABLE <%=LR_SPACE_TYPE_30%>>
					<TR>
					<TD WIDTH=* ALIGN=RIGHT><A HREF="vbscript:LoadCommonFormat">공통포맷</A>&nbsp;|&nbsp;<A HREF="vbscript:LoadCountFormat">수량포맷</A></TD>
					<TD WIDTH=10>&nbsp;</TD>
					</TR>
				</TABLE>
			</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="B1902mb1.asp" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="hDataType" tag="24">
<INPUT TYPE=HIDDEN NAME="hCurrency" tag="24">
<INPUT TYPE=HIDDEN NAME="hModuleCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hFormType" tag="24">
<INPUT TYPE=HIDDEN NAME="txtLogInCnt" tag="24">

</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
