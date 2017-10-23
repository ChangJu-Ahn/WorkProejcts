<%@ LANGUAGE="VBSCRIPT" %>
<%'======================================================================================================
'*  1. Module Name          : 법인세 
'*  2. Function Name        : 기준정보 
'*  3. Program ID           : WB117MA1
'*  4. Program Name         : WB117MA1.asp
'*  5. Program Desc         : 작업진행조회 및 마감 
'*  6. Modified date(First) : 2005/02/16
'*  7. Modified date(Last)  : 2005/02/16
'*  8. Modifier (First)     : 최영태 
'*  9. Modifier (Last)      : 최영태 
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../wcm/inc_SvrOperation.asp" -->
<%
'=========================  로긴중인 유저의 법인코드를 출력하기 위해  ======================
    Call LoadBasisGlobalInf()
%>
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript" SRC="../../inc/TabScript.js"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"	SRC="../WCM/inc_CliGrid.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"	SRC="../WCM/inc_CliOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit

'============================================  상수/변수 선언  ====================================

<!-- #Include file="../../inc/lgvariables.inc" -->

Const BIZ_MNU_ID = "wb113ma1"
Const BIZ_PGM_ID = "wb113mb1.asp"											 '☆: 비지니스 로직 ASP명 
Const BIZ_PGM_ID2 = "wb113mb2.asp"

Const TAB1 = 1																	'☜: Tab의 위치 
Const TAB2 = 2

Dim C_W1
Dim C_W1_NM
Dim C_W_CHK
Dim C_W2
Dim C_UPDT_USER
Dim C_UPDT_DT
Dim C_W3
Dim C_W4

Dim IsOpenPop    
Dim gSelframeFlg , lgCurrGrid      
Dim lgStrPrevKey2, IsRunEvents, lgblnConfig
Dim lgRefMode	' 레퍼런스 참조상태 : ERP 추출사용유무 

'============================================  초기화 함수  ====================================
Sub InitSpreadPosVariables()

	C_W1		= 1	
	C_W1_NM		= 2
	C_W_CHK		= 3
	C_W2		= 4
	C_UPDT_USER	= 5
	C_UPDT_DT	= 6
	C_W3		= 7
	C_W4		= 8

End Sub

Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
    
    lgStrPrevKey = ""                           'initializes Previous Key
    lgStrPrevKey2 = ""                          'initializes Previous Key
    lgLngCurRows = 0                            'initializes Deleted Rows Count
    lgSortKey = 1
    lgRefMode = False
    lgblnConfig = False
End Sub

Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029A("I", "*","NOCOOKIE","MA") %>
End Sub



'============================================  신고구분 콤보 박스 채우기  ====================================

Sub InitComboBox()
	Dim IntRetCD1
	call CommonQueryRs("MINOR_CD,MINOR_NM"," B_MINOR "," MAJOR_CD = 'W1018' ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboREP_TYPE ,lgF0  ,lgF1  ,Chr(11))

	call CommonQueryRs("MINOR_CD,MINOR_NM"," B_MINOR "," MAJOR_CD = 'W1077' ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboW1 ,lgF0  ,lgF1  ,Chr(11))
    
	call CommonQueryRs("MINOR_CD,MINOR_NM"," B_MINOR "," MAJOR_CD = 'W1078' ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboW2 ,lgF0  ,lgF1  ,Chr(11))
End Sub

Sub InitSpreadSheet()
	Dim ret
	
    Call initSpreadPosVariables()  

	With frm1.vspdData
	
	ggoSpread.Source = frm1.vspdData	
   'patch version
    ggoSpread.Spreadinit "V20041222",,parent.gAllowDragDropSpread    
    
	.ReDraw = false

    .MaxCols = C_W4 + 1							'☜: 최대 Columns의 항상 1개 증가시킴 
	.Col = .MaxCols														'☆: 사용자 별 Hidden Column
	.ColHidden = True    
	       
    .MaxRows = 0
    ggoSpread.ClearSpreadData
    
    ggoSpread.SSSetEdit		C_W1	,		"항목 코드", 10,,,10,1
    ggoSpread.SSSetEdit		C_W1_NM	,		"추출 항목", 30,,,100,1
    ggoSpread.SSSetCheck	C_W_CHK,		"추출", 7,,,True
	ggoSpread.SSSetFloat	C_W2	,		"추출갯수",		10,		"8",		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
    ggoSpread.SSSetEdit		C_UPDT_USER,	"최종작업자", 10,,,20,1
    ggoSpread.SSSetEdit		C_UPDT_DT,		"최종작업일", 20,,,25,2
	ggoSpread.SSSetEdit		C_W3,			"비고", 10,,,50,1
    ggoSpread.SSSetEdit		C_W4,			"로그", 50,,,4000
    	
	Call ggoSpread.SSSetColHidden(C_W1, C_W1, True)
	Call ggoSpread.SSSetColHidden(C_W4, C_W4, True)
	
	.ReDraw = true
	
    Call SetSpreadLock 
    
    End With

End Sub


'============================================  그리드 함수  ====================================

Sub InitSpreadComboBox()

End Sub


Sub SetSpreadLock()
    With frm1

    .vspdData.ReDraw = False
    
	ggoSpread.SpreadLock C_W1, -1, C_W1_NM
	ggoSpread.SpreadLock C_W2, -1, C_W4
	
    .vspdData.ReDraw = True

    End With
End Sub

Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)

End Sub

Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
 
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
            C_W1			= iCurColumnPos(1)
            C_W1_NM			= iCurColumnPos(2)
            C_W_CHK			= iCurColumnPos(3)
            C_W2			= iCurColumnPos(4)
            C_UPDT_USER		= iCurColumnPos(5)
            C_UPDT_DT		= iCurColumnPos(6)
            C_W3			= iCurColumnPos(7)
            C_W4			= iCurColumnPos(8)
    End Select    
End Sub

Sub InitData()
	frm1.txtFISC_YEAR.text 	= "<%=wgFISC_YEAR%>"
	frm1.txtCO_CD.value 	= "<%=wgCO_CD%>"
	frm1.txtCO_NM.value 	= "<%=wgCO_NM%>"
	frm1.cboREP_TYPE.value 	= "<%=wgREP_TYPE%>"

End Sub

'============================================  조회조건 함수  ====================================
Sub BtnAllChk()
	Dim iRow, iMaxRows 
	ggoSpread.Source = frm1.vspdData
    
	With frm1.vspdData
		iMaxRows = .MaxRows
		
		For iRow = 1 To iMaxRows
			.Row = iRow
			.Col = C_W_CHK
			If .value = "0" Then 
				.value = "1"
				lgBlnFlgChgValue = True
				ggoSpread.UpdateRow iRow
			End If
		Next
		
	End With
End Sub

Sub BtnChkCancel()
	Dim iRow, iMaxRows
	ggoSpread.Source = frm1.vspdData
	
	With frm1.vspdData
		iMaxRows = .MaxRows
		
		For iRow = 1 To iMaxRows
			.Row = iRow
			.Col = C_W_CHK
			If .value = "1" Then 
				.value = "0"
				lgBlnFlgChgValue = True
				ggoSpread.UpdateRow iRow
			End If
		Next
		
	End With
End Sub

Sub BtnERPGet()
    
	Call FncSave2()
End Sub


Sub BtnERPReset()
	Call FncDelete
End Sub

'====================================== 탭 함수 =========================================
Function ClickTab1()

	If gSelframeFlg = TAB1 Then Exit Function
	Call SetToolbar("1100000000000111")
	Call changeTabs(TAB1)
	gSelframeFlg = TAB1

End Function

Function ClickTab2()	
	Dim i, blnChange

	If gSelframeFlg = TAB2 Then Exit Function
	
	If wgConfirmFlg = "Y" Then
		Call SetToolbar("1100000000000111")
	Else
		Call SetToolbar("1100100000000111")
	End If
	Call changeTabs(TAB2)
	gSelframeFlg = TAB2
	
End Function

'============================================  폼 함수  ====================================
Sub Form_Load()

    Call LoadInfTB19029     
                                                    <%'Load table , B_numeric_format%>
    Call ggoOper.LockField(Document, "N")                                   <%'Lock  Suitable  Field%>                         
   
    Call InitSpreadSheet                                                    <%'Setup the Spread sheet%>
    Call InitVariables                                                      <%'Initializes local global variables%>
    
    Call SetToolbar("1100000000100111")										<%'버튼 툴바 제어 %>

	' 변경한곳 
    Call InitComboBox
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.FormatDate(frm1.txtFISC_YEAR, parent.gDateFormat,3)

	Call InitData()
	Call ClickTab1
	
	Call fncQuery()
     
    
End Sub


'============================================  이벤트 함수  ====================================
Sub txtFISC_YEAR_DblClick(Button)
    If Button = 1 Then
        frm1.txtFISC_YEAR.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtFISC_YEAR.Focus
    End If
End Sub


Sub cboW1_onChange()	' -- 연결유형 
	With frm1 
		If .cboW1.value = "4" Then
			Call ggoOper.SetReqAttr(frm1.cboW2, "Q")
			.cboW2.value = "1"
			divInfo(0).style.display = "none"
			divInfo(1).style.display = "none"
		Else
			Call ggoOper.SetReqAttr(frm1.cboW2, "N")
		End If
	End With
	lgBlnFlgChgValue = TRUE
End Sub

Sub cboW2_onChange()	' -- 연결방식 
	With frm1
		If .cboW2.value = "2" Then	
			divInfo(0).style.display = ""
			divInfo(1).style.display = "none"
			Call ggoOper.SetReqAttr(frm1.txtW3, "N")
			Call ggoOper.SetReqAttr(frm1.txtW6, "D")
			'.txtW3.setAttribute "tag", "23X"
		ElseIf .cboW2.value = "1" Then	
			divInfo(0).style.display = "none"
			divInfo(1).style.display = ""
			Call ggoOper.SetReqAttr(frm1.txtW3, "D")
			Call ggoOper.SetReqAttr(frm1.txtW6, "N")
		Else
			divInfo(0).style.display = "none"
			divInfo(1).style.display = "none"
			Call ggoOper.SetReqAttr(frm1.txtW3, "D")
			Call ggoOper.SetReqAttr(frm1.txtW6, "D")
		End If
	End With
	lgBlnFlgChgValue = TRUE
End Sub

Sub txtW3_onChange()
	lgBlnFlgChgValue = TRUE
End Sub

Sub txtW4_onChange()
	lgBlnFlgChgValue = TRUE
End Sub

Sub txtW5_onChange()
	lgBlnFlgChgValue = TRUE
End Sub


'==========================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )

End Sub

Sub vspdData_Click(ByVal Col, ByVal Row)

End Sub

Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		

End Sub

Sub vspdData_DblClick(ByVal Col, ByVal Row)				
 
End Sub

Sub vspdData_GotFocus()
    ggoSpread.Source = Frm1.vspdData

End Sub

Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If

End Sub    

Sub vspdData_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    
    Call GetSpreadColumnPos("A")
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

Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	If IsRunEvents = True Then Exit Sub	' 밑에 타 체크박스를 꺼는 행위시 같은 이벤트가 발생함 
	
	IsRunEvents = True
	
	With frm1.vspdData
	
		If Col = C_W_CHK Then
			lgBlnFlgChgValue= True ' 변경여부 
			.Row = Row
			.Col = Col
	
			If .Value = "1" Then
				ggoSpread.Source = frm1.vspdData
				ggoSpread.UpdateRow Row
			Else
				Call FncCancel
			End If
		End If
	End With
    
    IsRunEvents = False
End Sub

Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub

Sub PopRestoreSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()      
	Call ggoSpread.ReOrderingSpreadData()
End Sub


'============================================  툴바지원 함수  ====================================

Function FncNew() 
    Dim IntRetCD 

    FncNew = False

  '-----------------------
    'Check previous data area
    '-----------------------
    If lgBlnFlgChgValue = True Then
       IntRetCD = DisplayMsgBox("900015", parent.VB_YES_NO, "X", "X")           '⊙: "Will you destory previous data"
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

  '-----------------------
    'Erase condition area
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "A")
    Call ggoOper.LockField(Document, "N")
    Call InitVariables
    Call InitData

    Call SetToolbar("1100000000000111")

	frm1.txtCO_CD.focus

    FncNew = True

End Function

Function FncQuery() 
    Dim IntRetCD 
    
    FncQuery = False                                                        
    
    Err.Clear                                                               <%'Protect system from crashing%>

<%  '-----------------------
    'Check previous data area
    '----------------------- %>
    If lgBlnFlgChgValue Then
		ggoSpread.Source = frm1.vspdData
		If ggoSpread.SSCheckChange = True Then
			IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "X", "X")			    <%'데이타가 변경되었습니다. 조회하시겠습니까?%>
			If IntRetCD = vbNo Then
		  	Exit Function
			End If
		End If
    End If
<%  '-----------------------
    'Erase contents area
    '----------------------- %>
    Call ggoOper.ClearField(Document, "2")									<%'Clear Contents  Field%>
    ggoSpread.ClearSpreadData
    Call InitVariables                                                      <%'Initializes local global variables%>
    															
<%  '-----------------------
    'Check condition area
    '----------------------- %>
    If Not chkField(Document, "1") Then								<%'This function check indispensable field%>
       Exit Function
    End If    
	
     
    CALL DBQuery()
    
End Function

Function FncSave() 
    Dim blnChange, dblSum
    
    FncSave = False                                                         
    blnChange = False
    Err.Clear                                                               <%'☜: Protect system from crashing%>
    'On Error Resume Next                                                   <%'☜: Protect system from crashing%>    
    
<%  '-----------------------
    'Precheck area
    '----------------------- %>
    If lgBlnFlgChgValue = False Then
		Call DisplayMsgBox("900001", "X", "X", "X")                          <%'No data changed!!%>
		Exit Function
	End If

	If Not chkField(Document, "2") Then                             '⊙: Check contents area
	   Exit Function
	End If	    
<%  '-----------------------
    'Save function call area
    '----------------------- %>
    If DbSave = False Then Exit Function                                        '☜: Save db data
    
    FncSave = True                                                          
    
End Function

Function FncSave2() 
    Dim blnChange, dblSum
    Dim RetFlag
    FncSave2 = False                                                         

    Err.Clear                                                               <%'☜: Protect system from crashing%>
    'On Error Resume Next                                                   <%'☜: Protect system from crashing%>    
    
<%  '-----------------------
    'Precheck area
    '----------------------- %>
    If lgBlnFlgChgValue = False Then
		Call DisplayMsgBox("900001", "X", "X", "X")                          <%'No data changed!!%>
		Exit Function
	End If
	
	If lgblnConfig = False Then
		Call DisplayMsgBox("WB0004", "X", "X", "X")                          <%'No data changed!!%>
		Exit Function
	End If
	
	
	 RetFlag = DisplayMsgBox("900018", parent.VB_YES_NO,"X","X")   '☜ 바뀐부분 
	If RetFlag = VBNO Then
		Exit Function
	End If   
	
<%  '-----------------------
    'Save function call area
    '----------------------- %>
    If DbSave2 = False Then Exit Function                                        '☜: Save db data
    
    FncSave2 = True                                                         
    
End Function

Function FncCopy() 
 
End Function

Function FncCancel() 
    ggoSpread.Source = frm1.vspdData	
    ggoSpread.EditUndo                                                  '☜: Protect system from crashing
End Function

Function FncInsertRow(ByVal pvRowCnt) 

End Function

Function FncDeleteRow() 

End Function

Function FncPrint()
    Call parent.FncPrint()                                                   '☜: Protect system from crashing
End Function

Function FncExcel() 
    Call parent.FncExport(parent.C_MULTI)											 <%'☜: 화면 유형 %>
End Function

Function FncFind() 
    Call parent.FncFind(parent.C_MULTI, False)                                      <%'☜:화면 유형, Tab 유무 %>
End Function

Function FncExit()
Dim IntRetCD
	
	FncExit = False
	
    ggoSpread.Source = frm1.vspdData	
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO, "X", "X")                <%'데이타가 변경되었습니다. 종료 하시겠습니까?%>
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    FncExit = True
End Function

'========================================================================================
Function FncDelete() 
    Dim IntRetCD

    FncDelete = False
    If lgIntFlgMode <> parent.OPMD_UMODE Then
        Call DisplayMsgBox("900002", "X", "X", "X")
        Exit Function
    End If

    IntRetCD = DisplayMsgBox("900003",parent.VB_YES_NO, "X", "X")
    IF IntRetCD = vbNO Then
		Exit Function
	End IF

    Call DbDelete

    FncDelete = True
End Function
'============================================  DB 억세스 함수  ====================================
Function DbQuery() 

    DbQuery = False
    
    Err.Clear                                                               <%'☜: Protect system from crashing%>
	
	Call LayerShowHide(1)
	
	Dim strVal
    
    With Frm1
    
		strVal = BIZ_PGM_ID & "?txtMode="            & parent.UID_M0001						         
        strVal = strVal     & "&txtFISC_YEAR="       & Frm1.txtFISC_YEAR.Text      '☜: Query Key        
        strVal = strVal     & "&cboREP_TYPE="        & Frm1.cboREP_TYPE.Value      '☜: Query Key   
        'strVal = strVal     & "&txtCurrGrid="        & lgCurrGrid      
		
		Call RunMyBizASP(MyBizASP, strVal)   
    End With                                           '☜:  Run biz logic

    DbQuery = True  
  
End Function

Function DbQueryOk()													<%'조회 성공후 실행로직 %>
	
    Dim iNameArr
    Dim iDx    
    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
    
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	' 세무정보 조사 : 컨펌되면 락된다.
	Call CheckTaxDoc(frm1.txtCO_CD.value, frm1.txtFISC_YEAR.text, frm1.cboREP_TYPE.value, BIZ_MNU_ID)
	
	If wgConfirmFlg = "Y" Then

		Call SetToolbar("1100000000000111")	
		
		frm1.btn1.disabled = True
	Else
	
		'-----------------------
		'Reset variables area
		'-----------------------
		If frm1.vspdData.MaxRows > 0 Then
			'-----------------------
			'Reset variables area
			'-----------------------
			'Call SetGridSpan
			
			Call SetToolbar("1100000000000111")										<%'버튼 툴바 제어 %>

		End If
	
		If lgblnConfig Then	' -- wb113mb1.asp에서 변경됨 
			lgIntFlgMode = parent.OPMD_UMODE	
		End If

	End If
	'frm1.vspdData.focus			
End Function

'========================================================================================================
' Name : DbSave
' Desc : This function is data query and display
'========================================================================================================
Function DbSave() 
    DbSave = False                                                          
    
	if LayerShowHide(1) = false then
		exit Function
	end if

	Frm1.txtMode.value        =  Parent.UID_M0002
	Frm1.txtFlgMode.value     =  lgIntFlgMode
	'.txtUpdtUserId.value  =  Parent.gUsrID
	'.txtInsrtUserId.value =  Parent.gUsrID
				
	Call ExecMyBizASP(frm1, BIZ_PGM_ID2) 
		
End Function

'========================================================================================================
' Name : DbSave
' Desc : This function is data query and display
'========================================================================================================
Function DbSave2() 
    Dim lRow, lCol, lMaxRows, lMaxCols , i    
    Dim lGrpCnt     
    Dim retVal      
    Dim boolCheck   
    Dim lStartRow, lEndRow     
    Dim lRestGrpCnt 
    Dim strVal, strDel
 
    DbSave2 = False                                                          
    
	if LayerShowHide(1) = false then
		exit Function
	end if
   strVal = ""
    strDel = ""
    lGrpCnt = 1
    
	With frm1.vspdData
		' ----- 1번째 그리드 
		ggoSpread.Source = frm1.vspdData
		lMaxRows = .MaxRows : lMaxCols = .MaxCols
				
		For lRow = 1 To lMaxRows
		    
		   .Row = lRow : .Col = 0
		   
		   ' I/U/D 플래그 처리 
		   Select Case .Text
		       Case  ggoSpread.UpdateFlag                                      '☜: Update                                                  
		                                           strVal = strVal & "U"  &  Parent.gColSep                                                 
		            lGrpCnt = lGrpCnt + 1                                                 
		  End Select
		 .Col = 0
		  ' 모든 그리드 데이타 보냄     
		  If .Text = ggoSpread.InsertFlag Or .Text = ggoSpread.UpdateFlag Or .Text = ggoSpread.DeleteFlag Then
				For lCol = 1 To lMaxCols
					.Col = lCol : strVal = strVal & Trim(.Value) &  Parent.gColSep
				Next
				strVal = strVal & Trim(.Text) &  Parent.gRowSep
		  End If  
 
		Next
	End With

	Frm1.txtSpread.value      =  strVal
	Frm1.txtMode.value        =  Parent.UID_M0002
	'.txtUpdtUserId.value  =  Parent.gUsrID
	'.txtInsrtUserId.value =  Parent.gUsrID
				
	Call ExecMyBizASP(frm1, BIZ_PGM_ID) 
	
    DbSave2 = True                                                           
End Function


Function DbSaveOk()		
	Dim iRow											        <%' 저장 성공후 실행 로직 %>
	Call InitVariables
	
    Call MainQuery()
End Function

'========================================================================================
Function DbDelete() 
    Err.Clear
    DbDelete = False

    Dim strVal

    strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0003							'☜: 비지니스 처리 ASP의 상태 
    strVal = strVal     & "&txtFISC_YEAR="       & Frm1.txtFISC_YEAR.Text      '☜: Query Key        
    strVal = strVal     & "&cboREP_TYPE="        & Frm1.cboREP_TYPE.Value      '☜: Query Key            
	
	Call RunMyBizASP(MyBizASP, strVal)

    DbDelete = True
End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->	
</HEAD>
<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR>
		<TD <%=HEIGHT_TYPE_00%>></TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 width=200 onclick="ClickTab1()">
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" NOWRAP><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white><%=Request("strASPMnuMnuNm")%></font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 width=200 onclick="ClickTab2()">
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" NOWRAP><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>ERP연결 환경설정</font></td>
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
									<TD CLASS="TD5">사업연도</TD>
									<TD CLASS="TD6"><script language =javascript src='./js/wb113ma1_txtFISC_YEAR_txtFISC_YEAR.js'></script>
									<TD CLASS="TD5">법인명</TD>
									<TD CLASS="TD6"><INPUT TYPE=TEXT NAME="txtCO_CD" Size=10 tag="14">
										<INPUT TYPE=TEXT NAME="txtCO_NM" Size=20 tag="14">
									</TD>
									</TD>
								</TR>
								<TR>
									<TD CLASS="TD5">신고구분</TD>
									<TD CLASS="TD6"><SELECT NAME="cboREP_TYPE" ALT="신고구분" STYLE="WIDTH: 50%" tag="14X"></SELECT>
									</TD>
									<TD CLASS="TD5"></TD>
									<TD CLASS="TD6"></TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=* valign=top>
					<DIV ID="TabDiv" STYLE="FLOAT: left; HEIGHT: 100%; WIDTH: 100%;" SCROLL=no>
						<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR HEIGHT="100%">
								<TD>
									<script language =javascript src='./js/wb113ma1_vaSpread1_vspdData.js'></script>
								</TD>
							</TR>
						</TABLE>
					</DIV>
					<DIV ID="TabDiv" STYLE="FLOAT: left; HEIGHT: 100%; WIDTH: 100%;" SCROLL=no>
						<TABLE WIDTH=100%>
							<TR>
								<TD CLASS="TD5">ERP연결 유형</TD>
								<TD CLASS="TD6"><SELECT NAME="cboW1" ALT="ERP연결 유형" STYLE="WIDTH: 100%" tag="23X">
								<OPTION></OPTION>
								</SELECT></TD>
								<TD CLASS="TD5">&nbsp;</TD>
								<TD CLASS="TD6">&nbsp;</TD>
								</TD>
							</TR>
							<TR>
								<TD CLASS="TD5">ERP연결 방식</TD>
								<TD CLASS="TD6"><SELECT NAME="cboW2" ALT="ERP연결 방식" STYLE="WIDTH: 100%" tag="23X">
								<OPTION></OPTION>
								</SELECT></TD>
								<TD CLASS="TD5">&nbsp;</TD>
								<TD CLASS="TD6">&nbsp;</TD>
							</TR>
						</TABLE>
						<TABLE  WIDTH=100% style="display:'none'" ID="divInfo">
							<TR>
								<TD CLASS="TD5">원격지 URL</TD>
								<TD CLASS="TD6" COLSPAN=3><INPUT TYPE=TEXT NAME="txtW3" tag="25" STYLE="width: 100%" maxlength=1000></TD>
							</TR>
							<TR>
								<TD CLASS="TD5">원격지 아이디(ID)</TD>
								<TD CLASS="TD6"><INPUT TYPE=TEXT NAME="txtW4" tag="25" STYLE="width: 100%" maxlength=200></SELECT></TD>
								<TD CLASS="TD5">원격지 암호(PWD)</TD>
								<TD CLASS="TD6"><INPUT TYPE=TEXT NAME="txtW5" tag="25" STYLE="width: 100%" maxlength=100></TD>
							</TR>
						</TABLE>
						<TABLE  WIDTH=100% style="display:'none'" ID="divInfo">
							<TR>
								<TD CLASS="TD5">원격지 DB명</TD>
								<TD CLASS="TD6"><INPUT TYPE=TEXT NAME="txtW6" tag="25" STYLE="width: 100%" maxlength=1000></TD>
								<TD CLASS="TD5">&nbsp;</TD>
								<TD CLASS="TD6">&nbsp;</TD>
							</TR>
						</TABLE>
					</DIV>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_01%>>
		<TABLE <%=LR_SPACE_TYPE_30%>>
			<TR>
				<TD WIDTH=10>&nbsp;</TD>
				<TD><BUTTON NAME="btn1" ID="btn1"  CLASS="CLSSBTN" ONCLICK="vbscript:BtnERPGet()" Flag=1>ERP데이타 추출</BUTTON>&nbsp;
			
					<BUTTON NAME="btn3"	CLASS="CLSSBTN" ONCLICK="vbscript:BtnAllChk()"   Flag=1>모두 선택</BUTTON>&nbsp;
					<BUTTON NAME="btn4"	CLASS="CLSSBTN" ONCLICK="vbscript:BtnChkCancel()"   Flag=1>모두 취소</BUTTON>&nbsp;
				</TD>
			</TR>
		</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 tabindex="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" tabindex="-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" tabindex="-1">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24" tabindex="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm" tabindex="-1"></iframe>
</DIV>
</BODY>
</HTML>

