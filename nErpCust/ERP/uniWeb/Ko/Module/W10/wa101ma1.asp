<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% session.CodePage=949 %>


<%'======================================================================================================
'*  1. Module Name          : 법인세 
'*  2. Function Name        : 기준정보 
'*  3. Program ID           : WA101MA1
'*  4. Program Name         : WA101MA1.asp
'*  5. Program Desc         : 전자시고 Conversion 프로그램 
'*  6. Modified date(First) : 2005/02/24
'*  7. Modified date(Last)  : 2005/02/24
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
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/Cookie.vbs">   </SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit

'============================================  상수/변수 선언  ====================================

<!-- #Include file="../../inc/lgvariables.inc" -->

Const BIZ_MNU_ID = "wa101ma1"
Const BIZ_PGM_ID = "wa101mb1.asp"											 '☆: 비지니스 로직 ASP명 
Const BIZ_PGM_ID2 = "wa101mb2.asp"											 '☆: 비지니스 로직 ASP명 

Const TYPE_1 = 0
Const TYPE_2 = 1

Dim C_PGM_ID
Dim C_CHK_PGM
Dim C_TAX_DOC_CD
Dim C_PGM_NM
Dim C_ERR_TYPE
Dim C_STATUS_FLG

Dim C_SEQ_NO
Dim C_ERR_DOC
Dim C_ERR_VAL

Dim IsOpenPop    
Dim gSelframeFlg , lgCurrGrid , lgvspdData(1)   
Dim lgStrPrevKey2, IsRunEvents
Dim lgRefMode	' 레퍼런스 참조상태 : ERP 추출사용유무 

'============================================  초기화 함수  ====================================
Sub InitSpreadPosVariables()

	C_PGM_ID		= 1
	C_CHK_PGM		= 2
	C_TAX_DOC_CD	= 3	
	C_PGM_NM		= 4
	C_ERR_TYPE		= 5
	C_STATUS_FLG	= 6
	
	C_ERR_DOC		= 1
	C_ERR_VAL		= 2

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
    IsRunEvents = False
    lgCurrGrid = TYPE_1
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
End Sub

Sub InitSpreadSheet()
	Dim ret
	
    Call initSpreadPosVariables()  

	Set lgvspdData(TYPE_1) = frm1.vspdData0
	Set lgvspdData(TYPE_2) = frm1.vspdData1
	
	' -- 1번 그리드 
	With lgvspdData(TYPE_1)
	
	ggoSpread.Source = lgvspdData(TYPE_1)	
   'patch version
    ggoSpread.Spreadinit "V20041222" & TYPE_1,,parent.gAllowDragDropSpread    
    
	.ReDraw = false

    .MaxCols = C_STATUS_FLG + 1							'☜: 최대 Columns의 항상 1개 증가시킴 
	.Col = .MaxCols														'☆: 사용자 별 Hidden Column
	.ColHidden = True    
	       
    .MaxRows = 0
    ggoSpread.ClearSpreadData
    
    ggoSpread.SSSetEdit		C_PGM_ID,		"PGM", 7,,,20,1
    ggoSpread.SSSetCheck	C_CHK_PGM,		"", 3,,,True
	ggoSpread.SSSetEdit		C_TAX_DOC_CD,	"코드", 5,,,20,1
    ggoSpread.SSSetEdit		C_PGM_NM,		"서식명", 30,,,100,1
    ggoSpread.SSSetEdit		C_ERR_TYPE,		"오류여부", 10,2,,10,1
    ggoSpread.SSSetEdit		C_STATUS_FLG,		"체크불가", 10,2,,10,1
    	
	Call ggoSpread.SSSetColHidden(C_PGM_ID, C_PGM_ID, True)
	Call ggoSpread.SSSetColHidden(C_STATUS_FLG, C_STATUS_FLG, True)
	
	.ReDraw = true
	
    Call SetSpreadLock(TYPE_1)
    
    End With

	' -- 2번 그리드 
	With lgvspdData(TYPE_2)
	
	ggoSpread.Source = lgvspdData(TYPE_2)	
   'patch version
    ggoSpread.Spreadinit "V20041222" & TYPE_2,,parent.gAllowDragDropSpread    
    
	.ReDraw = false

    .MaxCols = C_ERR_VAL + 1							'☜: 최대 Columns의 항상 1개 증가시킴 
	.Col = .MaxCols														'☆: 사용자 별 Hidden Column
	.ColHidden = True    
	       
    .MaxRows = 0
    ggoSpread.ClearSpreadData
    
    ggoSpread.SSSetEdit		C_ERR_DOC,		"오류내용", 50,,,4000,1
    ggoSpread.SSSetEdit		C_ERR_VAL,		"오류값", 10,,,1000,1
	
	'Call ggoSpread.SSSetColHidden(C_SEQ_NO, C_SEQ_NO, True)
	
	.ReDraw = true
	
    Call SetSpreadLock(TYPE_2)
    
    End With
End Sub


'============================================  그리드 함수  ====================================

Sub InitSpreadComboBox()

End Sub


Sub SetSpreadLock(Byval pType)
    With lgvspdData(pType)

    .ReDraw = False
    
      ggoSpread.Source = lgvspdData(pType)
      'ggoSpread.SpreadLockWithOddEvenRowColor()
      
      If pType = TYPE_1 Then
		ggoSpread.SpreadLock C_TAX_DOC_CD, -1, C_STATUS_FLG	' 전체 적용 
	  Else
		ggoSpread.SpreadLock C_ERR_DOC, -1, C_ERR_VAL	' 전체 적용 
	  End If
				
    .ReDraw = True

    End With
End Sub


Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
 
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = lgvspdData(TYPE_1)
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
Sub BtnMakeFile()
	Call FncSave()
End Sub



Sub BtnDownLoad()

  
    Err.Clear                                                               '☜: Protect system from crashing
	Frm1.txtMode.value        =  Parent.UID_M0003

	Call ExecMyBizASP(frm1, BIZ_PGM_ID2) 

End Sub

Function subDiskOK(ByVal strVal) 
    Err.Clear
	On Error Resume Next
	Dim IntRetCD
	If strVal = "OK" Then
		IntRetCD = DisplayMsgBox("183114", "X", "X", "X")
	End If
End Function

Sub BtnERPReset()
	Call FncDelete
End Sub

'============================================  폼 함수  ====================================
Sub Form_Load()

    Call LoadInfTB19029     
                                                    <%'Load table , B_numeric_format%>
    Call ggoOper.LockField(Document, "N")                                   <%'Lock  Suitable  Field%>                         
   
    Call InitSpreadSheet                                                    <%'Setup the Spread sheet%>
    Call InitVariables                                                      <%'Initializes local global variables%>
    
    Call SetToolbar("1100000000000111")										<%'버튼 툴바 제어 %>

	' 변경한곳 
    Call InitComboBox
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.FormatDate(frm1.txtFISC_YEAR, parent.gDateFormat,3)

	Call InitData()
	Call MainQuery()
     
    
End Sub


'============================================  이벤트 함수  ====================================
Sub txtFISC_YEAR_DblClick(Button)
    If Button = 1 Then
        frm1.txtFISC_YEAR.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtFISC_YEAR.Focus
    End If
End Sub


Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)

End Sub

'==========================================================================================
' -- 0번 그리드 
Sub vspdData0_Click(ByVal Col, ByVal Row)
	lgCurrGrid = TYPE_1
	Call vspdData_Click(lgCurrGrid,  Col,  Row)
	Call vspdData1_Click(  Col,  Row)
	
End Sub

Sub vspdData0_ScriptLeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
    If Row <> NewRow And NewRow > 0 Then
		call vspdData0_Click( NewCol,  NewRow)
    End If
End Sub


Sub vspdData0_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
	lgCurrGrid = TYPE_1
	Call vspdData_ColWidthChange(lgCurrGrid, pvCol1, pvCol2)
End Sub

Sub vspdData0_GotFocus()
	lgCurrGrid = TYPE_1
	Call vspdData_GotFocus(lgCurrGrid)
End Sub

Sub vspdData0_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	lgCurrGrid = TYPE_1
	vspdData_ButtonClicked lgCurrGrid, Col, Row, ButtonDown
End Sub

' -- 1번 그리드 
Sub vspdData1_Click(ByVal Col, ByVal Row)
	lgCurrGrid = TYPE_2
	Call vspdData_Click(lgCurrGrid,  Col,  Row)
End Sub

Sub vspdData1_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
	lgCurrGrid = TYPE_2
	Call vspdData_ColWidthChange(lgCurrGrid, pvCol1, pvCol2)
End Sub

Sub vspdData1_GotFocus()
	lgCurrGrid = TYPE_2
	Call vspdData_GotFocus(lgCurrGrid)
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

Sub PrintErrDoc(ByVal Col, ByVal Row)
	
	Dim sErrVal, sErrDoc
	With lgvspdData(TYPE_2)
		.Row = Row
		.Col = C_ERR_VAL : sErrVal = .Text
		.Col = C_ERR_DOC : sErrDoc = .Text
		frm1.txtERR_DOC.value = "오류값: " & sErrVal & vbCrLf & "오류내용: " & sErrDoc
	End With
End Sub

Sub vspdData_Click(Index, ByVal Col, ByVal Row)
	lgCurrGrid = Index

	frm1.txtERR_DOC.value = ""
    Set gActiveSpdSheet = lgvspdData(Index)
	If Index = TYPE_2 Then 
		
		Call PrintErrDoc(Col, Row)
		Exit Sub
	End If

    If lgvspdData(TYPE_1).MaxRows = 0 Or Col = C_CHK_PGM Then                                                    'If there is no data.
       Exit Sub
   	End If
   	    
    If Row <= 0 Then
       ggoSpread.Source = lgvspdData(Index)
       
       If lgSortKey = 1 Then
           ggoSpread.SSSort Col               'Sort in ascending
           lgSortKey = 2
       Else
           ggoSpread.SSSort Col, lgSortKey    'Sort in descending 
           lgSortKey = 1
       End If
       
       Exit Sub
    Else
		' 디테일 그리드 호출 
		Call LayerShowHide(1)
	
		Dim strVal, sPGM_ID
    
		With lgvspdData(TYPE_1)
	
			.Row = Row	: .Col = C_PGM_ID : sPGM_ID = .Text
		End With
		
		lgvspdData(TYPE_2).MaxRows = 0
		ggoSpread.Source = lgvspdData(TYPE_2)
		ggoSpread.ClearSpreadData
		
		With frm1
			strVal = BIZ_PGM_ID & "?txtMode="            & parent.UID_M0001						         
		    strVal = strVal     & "&txtFISC_YEAR="       & .txtFISC_YEAR.Text      '☜: Query Key        
		    strVal = strVal     & "&cboREP_TYPE="        & .cboREP_TYPE.Value      '☜: Query Key   
		    strVal = strVal     & "&txtCurrGrid="        & TYPE_2      
			strVal = strVal     & "&PGM_ID="			 & sPGM_ID      
			
			Call RunMyBizASP(MyBizASP, strVal)   
		
		End With  
    End If

	
	lgvspdData(Index).Row = Row
End Sub

Sub vspdData0_DblClick(ByVal Col, ByVal Row)				
    Dim iColumnName
    
    If Row <= 0 Then
        Exit Sub
    End If
    
    If frm1.vspdData0.MaxRows = 0 Then
        Exit Sub
    End If

	With frm1.vspdData0
		.Row = Row
		.Col = C_PGM_ID
		Call PgmJump(.Value)
	End With

End Sub

Sub vspdData_ColWidthChange(Index, ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = lgvspdData(Index)
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub

Sub vspdData_GotFocus(Index)
    ggoSpread.Source = lgvspdData(Index)
    lgCurrGrid = Index
End Sub

Sub vspdData_ButtonClicked(Index, ByVal Col, ByVal Row, Byval ButtonDown)
	If IsRunEvents = True Then Exit Sub	' 밑에 타 체크박스를 꺼는 행위시 같은 이벤트가 발생함 
	
	IsRunEvents = True
	
	With lgvspdData(Index)
		Select Case Col
			Case C_CHK_PGM
				.Col = C_CHK_PGM
				.Row = Row
				If .Value = 1 Then
					.Col = C_STATUS_FLG 
					If .Value = 1 Then
						Call DisplayMsgBox("WA0004", parent.VB_INFORMATION, "X", "X")
						.Col = C_CHK_PGM
						.Value = 0
					End If
				End If
		End Select
    End With
    
    IsRunEvents = False
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
    ggoSpread.Source = lgvspdData(TYPE_1)
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "X", "X")			    <%'데이타가 변경되었습니다. 조회하시겠습니까?%>
    	If IntRetCD = vbNo Then
      	Exit Function
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
    Dim blnChange, dblSum,RetFlag
    
    FncSave = False                                                         
    blnChange = False
    Err.Clear                                                               <%'☜: Protect system from crashing%>
    'On Error Resume Next                                                   <%'☜: Protect system from crashing%>    
    
<%  '-----------------------
    'Precheck area
    '----------------------- %>
    ggoSpread.Source = lgvspdData(TYPE_1)
    If ggoSpread.SSCheckChange = True Then
		blnChange = True
    End If

	If Not  ggoSpread.SSDefaultCheck Then                                         '☜: Check contents area
	      Exit Function
	End If    
	
	
	 RetFlag = DisplayMsgBox("900018", parent.VB_YES_NO,"X","X")   '☜ 바뀐부분 
	If RetFlag = VBNO Then
		Exit Function
	End If   

<%  '-----------------------
    'Save function call area
    '----------------------- %>
    If DbSave = False Then Exit Function                                        '☜: Save db data
    
    FncSave = True                                                          
    
End Function

Function FncCopy() 
 
End Function

Function FncCancel() 
                                                 '☜: Protect system from crashing
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
	
    ggoSpread.Source = lgvspdData(TYPE_1)	
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
        strVal = strVal     & "&txtCurrGrid="        & TYPE_1      
		
		Call RunMyBizASP(MyBizASP, strVal)   
    End With                                           '☜:  Run biz logic

    DbQuery = True  
  
End Function

Function DbQueryOk()													<%'조회 성공후 실행로직 %>
	
    Dim iNameArr
    Dim iDx    
    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
    
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	
    '-----------------------
    'Reset variables area
    '-----------------------
    If lgvspdData(TYPE_1).MaxRows > 0 Then
		'-----------------------
		'Reset variables area
		'-----------------------
		lgIntFlgMode = parent.OPMD_UMODE
		'Call SetGridSpan
		
		Call SetToolbar("1100000000000111")										<%'버튼 툴바 제어 %>

	End If
	'lgvspdData(TYPE_1).focus			
	CALL vspdData0_Click(1,1)
	
End Function
Function DbQueryOk2()	

End Function

'========================================================================================================
' Name : DbSave
' Desc : This function is data query and display
'========================================================================================================
Function DbSave() 
    Dim lRow, lCol, lMaxRows, lMaxCols , i    
    Dim lGrpCnt     
    Dim retVal      
    Dim boolCheck   
    Dim lStartRow, lEndRow     
    Dim lRestGrpCnt 
    Dim strVal, strDel, sTmp
 
    DbSave = False                                                          
    
	if LayerShowHide(1) = false then
		exit Function
	end if
    strVal = ""
    strDel = ""
    lGrpCnt = 1
    

	With lgvspdData(i)
	
		ggoSpread.Source = lgvspdData(i)
		lMaxRows = .MaxRows : lMaxCols = .MaxCols
			
		' ----- 1번째 그리드 
		For lRow = 1 To .MaxRows

    
			.Row = lRow	: sTmp = "" : .Col = C_CHK_PGM

			If .Value = 1 Then
				.Col = C_PGM_ID
			  	strVal = strVal & "'" & Trim(.Text) & "',"
			End If  

		Next
							   
	End With

	frm1.txtSpread.value = Left(strVal, Len(strVal)-1)
	strDel = "" : strVal = ""


	'Frm1.txtSpread.value      = strDel & strVal
	Frm1.txtMode.value        =  Parent.UID_M0002
	'.txtUpdtUserId.value  =  Parent.gUsrID
	'.txtInsrtUserId.value =  Parent.gUsrID
				
	Call ExecMyBizASP(frm1, BIZ_PGM_ID2) 
	
    DbSave = True                                                           
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


Function PgmJump(Byval pMnuID)
	Dim objConn , PostString
	WriteCookie "gActivePgmID",pMnuID
	
	Set objConn = CreateObject("uniConnector.cGlobal") 
	PostString = objConn.GetAspPostString 
	'window.open "../../SessionTrans.asp?" & PostString 
	
	window.open "../../uniToolbar.Asp?SLX=Y&DPCP=" & pMnuID & "&arg="
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
					<TD CLASS="CLSMTABP_BAK">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 width=300>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" NOWRAP><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB_BAK"><font color=white><%=Request("strASPMnuMnuNm")%></font></td>
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
									<TD CLASS="TD6"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name=txtFISC_YEAR CLASS=FPDTYYYY title=FPDATETIME ALT="사업연도" tag="14X1" id=txtFISC_YEAR></OBJECT>');</SCRIPT>
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
						<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR HEIGHT="100%">
								<TD WIDTH=45%>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData0 WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=vaSpread1> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
								</TD>
								<TD WIDTH=55%>
									<TABLE <%=LR_SPACE_TYPE_20%>>
									<TR>
										<TD HEIGHT=70%>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData1 WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=vaSpread2> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
										</TD>
									</TR>
									<TR>
										<TD HEIGHT=30%><TEXTAREA NAME=txtERR_DOC readonly tag="24" STYLE="WIDTH: 100%; HEIGHT: 100%"></TEXTAREA>
										</TD>
									</TR>
								</TD>
							</TR>
						</TABLE>
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
				<TD><BUTTON NAME="btn1"  WIDTH=20 ONCLICK="vbscript:BtnMakeFile()" Flag=1>전자신고 파일생성</BUTTON>&nbsp;
				<BUTTON    NAME="btn1"   WIDTH=20 ONCLICK="vbscript:BtnDownLoad()" Flag=1>변환파일 내려받기</BUTTON>&nbsp;
				</TD>
			</TR>
		</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtCurrGrid" tag="24">

Frm1.txtMode.value
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

