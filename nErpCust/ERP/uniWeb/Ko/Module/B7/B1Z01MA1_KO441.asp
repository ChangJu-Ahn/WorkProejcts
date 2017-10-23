<%@ LANGUAGE="VBSCRIPT" %>
<!--'**********************************************************************************************
'*
'*  1. Module Name          : 
'*  2. Function Name        : 
'*  3. Program ID           : 
'*  4. Program Name         : SQL COmmand
'*  5. Program Desc         : 쿼리분석기
'*  6. Component List       : 
'*  7. Modified date(First) : 2000/09/17
'*  8. Modified date(Last)  : 2005/06/21
'*  9. Modifier (First)     : ahj
'* 10. Modifier (Last)      : Yim Yong Ju
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'********************************************************************************************** -->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/incSvrCcm.inc"  -->
<!-- #Include file="../../inc/incSvrHTML.inc"  -->				<!-- '⊙: 화면처리ASP에서 서버작업이 필요한 경우  -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		<!-- '☆: 해당 위치에 따라 달라짐, 상대 경로  -->

<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliMAMain.vbs">  </SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliMAEvent.vbs">  </SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliVariables.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliMAOperation.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"	 SRC="../../inc/incCliRdsQuery.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/Cookie.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="javascript" SRC="../../inc/TabScript.js">   </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incHRQuery.vbs"> </SCRIPT>

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentA.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentVariables.vbs"> </SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>

'Option Explicit    

'============================================  1.2.1 Global 상수 선언  ====================================

Const BIZ_PGM_ID			= "B1Z01MB1_KO441.asp"											 '☆: 비지니스 로직 ASP명
Const BIZ_PGM_ID_RUN 		= "B1Z010MB1_KO441.asp" 

Dim C_FirstKey
Dim C_MaxKey 
Dim arrSelCol

'============================================  1.2.2 Global 변수 선언  ===================================
'	1. 변수 표준에 따름. prefix로 g를 사용함.
'	2. Array인 경우는 ()를 반드시 사용하여 일반 변수와 구별해 됨
'========================================================================================================= 

Dim lgBlnFlgChgValue				'☜: Variable is for Dirty flag
Dim lgIntGrpCount				'☜: Group View Size를 조사할 변수
Dim lgIntFlgMode					'☜: Variable is for Operation Status

Dim lgNextNo						'☜: 화면이 Single/SingleMulti 인경우만 해당
Dim lgPrevNo						' ""

'========================================================================================================= 
Dim lgMpsFirmDate, lgLlcGivenDt											 '☜: 비지니스 로직 ASP에서 참조하므로 Dim 

Dim lgCurName()															'☆ : 개별 화면당 필요한 로칼 전역 변수 
Dim cboOldVal          
Dim IsOpenPop          
Dim lgCboKeyPress      
Dim lgOldIndex								
Dim lgOldIndex2        


'========================================================================================================= 
Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE                                               '⊙: Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                                                '⊙: Indicates that no value changed
    lgIntGrpCount = 0                                                       '⊙: Initializes Group View Size
    '-----------------------  Coding part  ------------------------------------------------------------- 
    IsOpenPop = False														'☆: 사용자 변수 초기화
    lgCboKeyPress = False
    lgLlcGivenDt=""
    
    C_MaxKey = 99
    lgStrPrevKey     = ""
    lgPageNo         = ""
    
End Sub


'========================================================================================================= 
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% Call loadInfTB19029A("I", "*", "NOCOOKIE", "MA") %>
End Sub


'==========================================  2.4.3 Set???()  ===============================================
'	Name : OpenQueryID()
'	Description : PopUp에서 전송된 값을 특정 Tag Object에 지정
'========================================================================================================= 
Sub OpenQueryID()

	Dim arrRet
	Dim Param1, Param2, Param3, Param4, Param5
	Dim iCalledAspName, IntRetCD
	dim arrField(5)
	
	If IsOpenPop = True Then Exit Sub
	
	IsOpenPop = True
	
	Param1 = ""
	Param2 = ""
	Param3 = Trim(frm1.txtQueryCd.Value)
	Param4 = parent.gDepart
	Param5 = ""
	
	iCalledAspName = AskPRAspName("B1Z01OA1_KO441")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", Parent.VB_INFORMATION, "B1Z01OA1_KO441", "X")
		IsOpenPop = False
		Exit Sub
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, Param1, Param2, Param3, Param4, Param5,arrField), _
		"dialogWidth=820px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		frm1.txtQueryCd.value = arrRet(0)
		Call DbQuery()
	End If
	
	frm1.txtQueryCd.Focus
	Set gActiveElement = document.activeElement	

End Sub



'==========================================================================================================
Sub Form_Load()

    Call InitVariables
    Call LoadInfTB19029																'⊙: Load table , B_numeric_format
    Call ggoOper.LockField(Document, "N")

    Call SetToolBar("1110100000001111")
    Call InitComboBox
    Call InitSpreadClear()
    Call InitSpreadSheet1
    
    frm1.txtDept_cd.value = parent.gDepart
	frm1.txtQueryCd.focus	
	frm1.txtMode.value = "UID_M0001"
	frm1.hCommand.value = "LOOKUP"
	
End Sub

Sub InitSpreadSheet1()
    
    With frm1
           
    ggoSpread.Source = .vspdData1
    ggoSpread.Spreadinit "V20021106", , Parent.gAllowDragDropSpread
    
    .vspdData1.ReDraw = False
    
    .vspdData1.MaxCols = 5
    .vspdData1.MaxRows = 0
	
    ggoSpread.SSSetEdit		0,		"", 0,,,100
    ggoSpread.SSSetEdit		1,		"txtSELECT", 20,,,32767
    ggoSpread.SSSetEdit		2,		"txtFROM ", 20,,,32767
    ggoSpread.SSSetEdit		3,		"txtWGERE ", 20,,,32767
    ggoSpread.SSSetEdit		4,		"txtETC ", 20,,,32767
    ggoSpread.SSSetEdit		5,		"txtREMARK ", 20,,,32767
	
	.vspdData1.ReDraw = False
	
    End With
    
End Sub


Sub InitSpreadClear()

	With frm1.vspdData
	
    ggoSpread.Source = frm1.vspdData
    ggoSpread.Spreadinit "V20021203",,parent.gAllowDragDropSpread    
        
	.ReDraw = false

    .MaxCols = 0
	.Col = .MaxCols															<%'공통콘트롤 사용 Hidden Column%>
    
    .MaxRows = 0
    ggoSpread.ClearSpreadData

	.ReDraw = true
    
    End With
    
End Sub

'========================================================================================================
' Name : InitComboBox()
' Desc : Set ComboBox
'========================================================================================================
Sub InitComboBox()
    Dim iCodeArr 
    Dim iNameArr
    Dim iDx
    
    Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("B0020", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    iCodeArr = lgF0
    iNameArr = lgF1
    
    Call SetCombo2(frm1.cboRole_type,iCodeArr, iNameArr,Chr(11))                  ''''''''DB에서 불러 condition에서
End Sub


'========================================================================================
Function FncQuery() 
    Dim IntRetCD

    FncQuery = False 
    Err.Clear
  '-----------------------
    'Check previous data area
    '----------------------- 
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "X", "X")				'⊙: "Will you destory previous data"
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

  '-----------------------
    'Erase contents area
    '----------------------- 
    Call ggoOper.ClearField(Document, "2")
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
    frm1.hCommand.value = "LOOKUP"
    Call DbQuery

    FncQuery = True
End Function

'========================================================================================
' Function Name : FncPrev
' Function Desc : This function is related to Previous Button
'========================================================================================
Function FncPrev() 
    Dim IntRetCD     
    
    FncPrev = False                                                        '⊙: Processing is NG
    
    Err.Clear                                                               '☜: Protect system from crashing

   '-----------------------
    'Query First
    '------------------------ 
    If lgIntFlgMode <> parent.OPMD_UMODE Then                                      'Check if there is retrived data
		Call DisplayMsgBox("900002","X","X","X")
		Exit Function
    End If
    
   '-----------------------
    'Check previous data area
    '------------------------ 
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO,"X","X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
  '-----------------------
    'Check condition area
    '----------------------- 
    If Not chkField(Document, "1") Then									'⊙: This function check indispensable field
       Exit Function
    End If
    
  '-----------------------
    'Erase contents area
    '----------------------- 
    Call InitVariables															'⊙: Initializes local global variables
    
  '-----------------------
    'Query function call area
    '----------------------- 
	frm1.hCommand.value = "PREV"
    Call DbQuery																'☜: Query db data
           
    FncPrev = True																'⊙: Processing is OK        
    
End Function

'========================================================================================
' Function Name : FncNext
' Function Desc : This function is related to Next Button
'========================================================================================
Function FncNext() 
    Dim IntRetCD     
    
    FncNext = False                                                        '⊙: Processing is NG
    
    Err.Clear                                                               '☜: Protect system from crashing

   '-----------------------
    'Query First
    '------------------------ 
    If lgIntFlgMode <> parent.OPMD_UMODE Then                                      'Check if there is retrived data
		Call DisplayMsgBox("900002","X","X","X")
		Exit Function
    End If
    
   '-----------------------
    'Check previous data area
    '------------------------ 
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO,"X","X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
  '-----------------------
    'Check condition area
    '----------------------- 
    If Not chkField(Document, "1") Then									'⊙: This function check indispensable field
       Exit Function
    End If
    
  '-----------------------
    'Erase contents area
    '----------------------- 
    Call InitVariables															'⊙: Initializes local global variables
    
  '-----------------------
    'Query function call area
    '----------------------- 
    
	frm1.hCommand.value = "NEXT"
    Call DbQuery																'☜: Query db data
           
    FncNext = True																'⊙: Processing is OK        
    
End Function


'========================================================================================================
' Name : OpenDept
' Desc : 부서 POPUP
'========================================================================================================
Function OpenDept(iWhere)
	Dim arrRet
	Dim arrParam(2)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	If iWhere = 0 Then
		arrParam(0) = frm1.txtDept_cd.value			            '  Code Condition
	End If
	arrParam(2) = lgUsrIntCd                            ' 자료권한 Condition  

	arrRet = window.showModalDialog(HRAskPRAspName("DeptPopupDt"), Array(window.parent,arrParam), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtDept_cd.focus
		Exit Function
	Else
		Call SetDept(arrRet, iWhere)
	End If	
	
	lgBlnFlgChgValue = True
			
End Function

'------------------------------------------  SetDept()  ------------------------------------------------
'	Name : SetDept()
'	Description : Dept Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------
Function SetDept(Byval arrRet, Byval iWhere)
		
	With frm1
		Select Case iWhere
		     Case "0"
               .txtDept_cd.value = arrRet(0)
               .txtDept_nm.value = arrRet(1)
               .txtDept_cd.focus
        End Select
	End With
End Function      

'========================================================================================
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
    Call ggoOper.LockField(Document, "N")                                       '⊙: Lock  Suitable  Field
    Call InitVariables
    Call InitSpreadClear()
    Call InitComboBox

    Call SetToolBar("1110100000001111")

	frm1.txtDept_cd.value = parent.gDepart
	frm1.txtQueryCd.focus	
	frm1.txtMode.value = "UID_M0001"
	frm1.hCommand.value = "LOOKUP"

    FncNew = True

End Function


'========================================================================================
Function FncDelete() 
    Dim IntRetCD

    FncDelete = False

  '-----------------------
    'Precheck area
    '-----------------------
    If lgIntFlgMode <> parent.OPMD_UMODE Then                                      'Check if there is retrived data
        Call DisplayMsgBox("900002", "X", "X", "X")
        Exit Function
    End If

  '-----------------------
    'Delete function call area
    '-----------------------
    IntRetCD = DisplayMsgBox("900003", parent.VB_YES_NO, "X", "X")
    IF IntRetCD = vbNO Then
		Exit Function
	End IF

    Call DbDelete
    Call InitSpreadClear()
    
    FncDelete = True
End Function


'========================================================================================
Function FncSave() 
    Dim IntRetCD 

    FncSave = False 
    Err.Clear

  '-----------------------
    'Precheck area
    '-----------------------
    If lgBlnFlgChgValue = False Then
        IntRetCD = DisplayMsgBox("900001", "X", "X", "X")                          '⊙: No data changed!!
        Exit Function
    End If

  '-----------------------
    'Check content area
    '-----------------------
    If Not chkField(Document, "2") Then                             '⊙: Check contents area
       Exit Function
    End If

  '-----------------------
    'Save function call area
    '-----------------------
    Call DbSave
    Call InitSpreadClear()

    FncSave = True

End Function

Function FncCopy() 
  '-----------------------
    'Check previous data area
    '-----------------------
    If lgBlnFlgChgValue = True Then
       IntRetCD = DisplayMsgBox("900015", parent.VB_YES_NO, "X", "X")           '⊙: "Will you destory previous data"
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

    Call InitVariables
    Call InitSpreadClear()

    Call SetToolBar("1110100000001111")

	frm1.txtQueryCd.value = ""
	frm1.txtMode.value = "UID_M0001"
	

	lgBlnFlgChgValue = True 
    FncNew = True
    
	
End Function

'========================================================================================
Function FncCancel()
     On Error Resume Next
End Function


'========================================================================================
Function FncInsertRow()
     On Error Resume Next
End Function


'========================================================================================
Function FncDeleteRow()
     On Error Resume Next
End Function


'========================================================================================
Function FncPrint()
     On Error Resume Next
    parent.FncPrint()
End Function


'========================================================================================
Function FncExcel()
    Call parent.FncExport(parent.C_SINGLE)
End Function

'========================================================================================
Function FncFind()
    Call parent.FncFind(parent.C_SINGLE, False)
End Function


'========================================================================================
Function FncExit()
	Dim IntRetCD
	FncExit = False
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO, "X", "X")			'⊙: "Will you destory previous data"
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    FncExit = True
End Function


'========================================================================================
Function DbDelete() 
    Err.Clear
    DbDelete = False

    Call LayerShowHide(1)
    Dim strVal

    strVal = BIZ_PGM_ID & "?txtMode=UID_M0004"
    strVal = strVal & "&txtQueryCd=" & Trim(frm1.txtQueryCd.value)

	Call RunMyBizASP(MyBizASP, strVal)

    DbDelete = True
End Function


'========================================================================================
Function DbDeleteOk()
	lgBlnFlgChgValue = False
	Call FncNew()
End Function


'========================================================================================
' Function Name : txtSelect_OnChange
' Function Desc : 
'========================================================================================
Sub txtSelect_OnChange()
	lgBlnFlgChgValue = True
End Sub

'========================================================================================
' Function Name : txtFrom_OnChange
' Function Desc : 
'========================================================================================
Sub txtFrom_OnChange()
	lgBlnFlgChgValue = True
End Sub

Sub cboRole_type_OnChange()
	lgBlnFlgChgValue = True
End Sub

Sub txtWhere_OnChange()
	lgBlnFlgChgValue = True
End Sub

Sub txtEtc_OnChange()
	lgBlnFlgChgValue = True
End Sub

Sub txtRemark_OnChange()
	lgBlnFlgChgValue = True
End Sub


'---------------------------------------------------------------------------------------
' 쿼리 실행 
'---------------------------------------------------------------------------------------
Function FncRun()

On Error Resume Next

	If UCASE(TRIM(frm1.txtMode.value)) <> "UID_M0002" And Trim(frm1.StrSelect_RUN.value) = "" THEN
		MSGBOX "조회를 먼저 하셔야합니다."
		Exit Function
	END IF
	
	Dim StrSelect,StrColList, TblId, Flag, FPos,LPos, XPos, StrTableNm, StrTableNm1, i
	
	ggoSpread.Source = frm1.vspdData
    Call ggoSpread.ClearSpreadData()
    
    StrSelect = "SELECT " & Trim(frm1.txtSelect.value)
    If Trim(frm1.txtFrom.value) <> "" Then
		StrSelect = StrSelect & " FROM " & Trim(frm1.txtFrom.value)
    End If
    If Trim(frm1.txtWhere.value) <> "" Then
		StrSelect = StrSelect & " Where " & Trim(frm1.txtWhere.value)
    End If
    If Trim(frm1.txtEtc.value) <> "" Then
		StrSelect = StrSelect & " " & Trim(frm1.txtEtc.value)
    End If
        
    frm1.StrSelect_RUN.value = StrSelect
    
    FPos = Instr(UCASE(StrSelect),"SELECT")
	LPos = Instr(UCASE(StrSelect),"FROM")
	XPos = Instr(UCASE(StrSelect),"WHERE")
	
    if FPos  = "0" Or LPos  = "0" Then
		MsgBox "지원하지 않는 쿼리문입니다."
		Exit Function
	End If	
	
	StrColList = Trim(MID(Trim(StrSelect), FPos + 6, LPos - (FPos + 6)))
	
    arrSelCol = Split(StrColList, ",")
    
    '컬럼 저장
	For i = 0 To UBound(arrSelCol)
		arrSelCol(i) = Trim(arrSelCol(i))
	Next
	
	'테이블
	if Instr(arrSelCol(0),"*") > 0 Then
		' where조건이 없는 경우
		if Trim(XPos) = "0" Then
			StrTableNm = Trim(UCase(MID(StrSelect, LPos + 4, Len(StrSelect))))
		Else
			StrTableNm = Trim(UCase(MID(StrSelect, LPos + 4, XPos - (LPos + 4))))
		End If
		
		if Instr(StrTableNm,chr(13)) > "0" Then
			StrTableNm = Left(StrTableNm, Instr(StrTableNm,chr(13)) - 1)
		End If
		'컬럼 개수
		Call CommonQueryRs(" info, id ", " sysobjects ", " name = '"& Trim(StrTableNm) &"' ", _
           lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) 

		C_FirstKey = Replace(lgF0,Chr(11),"")
		TblId = Replace(lgF1,Chr(11),"")
		Flag = "A"
		
	Else
		C_FirstKey = UBound(arrSelCol)+2
		TblId = "*"
		Flag = "B"
	End If	
	
	Call InitVariables 	
	Call LoadInfTB19029	
	Call InitSpreadSheet()

    If Not chkField(Document, "1") Then								              
       Exit Function
    End If

    If DbQuery_Run = False Then 
       Exit Function
    End If   

    If Err.number = 0 Then
'       FncQuery = True                                                           
    End If   

    Set gActiveElement = document.ActiveElement      	

End Function

Function DbQuery_Run() 
	Dim ColCnt
	
	Err.Clear															

	DbQuery_Run = False														

	If LayerShowHide(1) = False Then
		Exit Function
	End If

	Dim strVal
	
	IF Cint(C_FirstKey) >= CInt(C_MaxKey) Then
		ColCnt = C_MaxKey
	Else
		ColCnt = C_FirstKey
	End If
	
	frm1.txtColCnt.value	= (ColCnt - 1)
	lgPageNo				= frm1.lgPageNo.value
	lgIntStartRow			= frm1.vspdData.MaxRows + 1

	CALL ExecMyBizASP(frm1, BIZ_PGM_ID_RUN)
	
	DbQuery_Run = True							 
End Function

Sub InitSpreadSheet()
    Call SetZAdoSpreadSheet("B1Z010MA1","S","A", "V20030523", parent.C_SORT_DBAGENT,frm1.vspdData,C_MaxKey, "X", "X")
    Call SetSpreadLock()    

    'Call ggoSpread.SSSetColHidden(C_FirstKey,C_MaxKey,True)   
End Sub

'========================================
Sub SetSpreadLock()
      ggoSpread.Source = frm1.vspdData
      ggoSpread.SpreadLockWithOddEvenRowColor()
End Sub


Function SetColNm(TblId, Flag)

	Dim arrColNm, arrColId, nCol
    
	frm1.vspdData.Row = 0
	
	If isnull(TblId) = false Then
		If Flag = "A" Then
			Call CommonQueryRs(" name, colid ", " syscolumns ", " id = '"& TblId &"' order by colid ", _
			       lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) 
			 
			arrColNm = Split(lgF0, chr(11))      
			arrColId = Split(lgF1, chr(11))
    
			For nCol = 0 To UBound(arrColId)
				frm1.vspdData.Col = nCol + 1
				frm1.vspdData.Text = UCase(arrColNm(nCol))
			Next
		Else
			For nCol = 0 To UBound(arrSelCol)
				frm1.vspdData.Col = nCol + 1
				frm1.vspdData.Text = UCase(arrSelCol(nCol))
			Next
		End IF
	Else
	
		For nCol = 1 To 10
			frm1.vspdData.Col = nCol
			frm1.vspdData.Text = nCol
		Next
	
	End If
	
End Function

Function SetColNm_New(lgF0,lgF1)

	Dim arrColNm, arrColId, nCol
    
    arrColNm = Split(lgF0, chr(11))      
	arrColId = Split(lgF1, chr(11))
	
    frm1.vspdData.MaxCols = UBound(arrColId)
	frm1.vspdData.Row = 0
    
	For nCol = 0 To UBound(arrColId) - 1
		frm1.vspdData.Col = nCol + 1
		frm1.vspdData.Text = UCase(arrColNm(nCol))
	Next
	
End Function


'========================================================================================
Function DbQuery()

    Err.Clear
    DbQuery = False


    Call LayerShowHide(1)
    Dim strVal

    strVal = BIZ_PGM_ID & "?txtMode=UID_M0003"									'☜: 비지니스 처리 ASP의 상태 
    strVal = strVal & "&txtQueryCd=" & Trim(frm1.txtQueryCd.value)				'☆: 조회 조건 데이타
    strVal = strVal & "&txtCommand		=" & Trim(frm1.hCommand.value)
    strVal = strVal & "&gDepart=" & parent.gDepart
    strVal = strVal & "&PrevNextFlg=" & ""
    call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동

    DbQuery = True
End Function


'========================================================================================
Function DbQueryOk()
    Call SetToolBar("1111100011111111")
    Call ggoOper.LockField(Document, "Q")									'⊙: This function lock the suitable field
    
    lgIntFlgMode = parent.OPMD_UMODE
	frm1.txtMode.value = "UID_M0002"
	frm1.hCommand.value = "LOOKUP"
	
	
End Function

'========================================================================================
Function DbQueryOk1()
    Call SetToolBar("1111100011100111")
    Call ggoOper.LockField(Document, "Q")									'⊙: This function lock the suitable field
    Call InitSpreadClear()
    
    
	With frm1.vspdData1
	  
		.Row = 1
		.Col = 1
		frm1.txtSelect.Value = replace(.Text,chr(7),chr(13) &chr(10))
		
		.Col = 2
		frm1.txtFrom.Value = replace(.Text,chr(7),chr(13) &chr(10))
		
		.Col = 3
		frm1.txtWhere.Value = replace(.Text,chr(7),chr(13) &chr(10))
		
		.Col = 4
		frm1.txtEtc.Value = replace(.Text,chr(7),chr(13) &chr(10))
		
		.Col = 5
		frm1.txtRemark.Value = replace(.Text,chr(7),chr(13) &chr(10))								
		
    End With
    
    Call InitSpreadSheet1()
        
    lgIntFlgMode = parent.OPMD_UMODE
	frm1.txtMode.value = "UID_M0002"
	frm1.hCommand.value = "LOOKUP"
	
	
End Function


'========================================================================================
Function DbSave() 

    Err.Clear
	DbSave = False

    Dim strVal
    Call LayerShowHide(1)

	With frm1
		.txtMode.value = frm1.txtMode.value
		.txtFlgMode.value = lgIntFlgMode

	Call ExecMyBizASP(frm1, BIZ_PGM_ID)

	End With

    DbSave = True
End Function


'========================================================================================
Function DbSaveOk()
    lgBlnFlgChgValue = False
    FncQuery
End Function

Function MakeSQL()

	Dim StrSelect
    
    If Trim(frm1.txtFrom.value) <> "" Then
		StrSelect = "SELECT " & Trim(frm1.txtSelect.value)
    End If
    If Trim(frm1.txtFrom.value) <> "" Then
		StrSelect = StrSelect & " FROM " & Trim(frm1.txtFrom.value)
    End If
    If Trim(frm1.txtWhere.value) <> "" Then
		StrSelect = StrSelect & " Where " & Trim(frm1.txtWhere.value)
    End If
    If Trim(frm1.txtEtc.value) <> "" Then
		StrSelect = StrSelect & " " & Trim(frm1.txtEtc.value)
    End If
        
    frm1.StrSelect_RUN.value = StrSelect
    
End Function

Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )    
    If OldLeft <> NewLeft Then Exit Sub

	If CheckRunningBizProcess = True Then
	   Exit Sub
	End If

	If Frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(Frm1.vspdData,NewTop) Then	    
    	If lgPageNo <> "" Or frm1.lgPageNo.value <> "" Then								
           Call DisableToolBar(parent.TBC_QUERY)
           Call DbQuery_Run
    	End If
    End If    
End Sub


Function ViewSQL()

	Call MakeSQL()
	IF frm1.StrSelect_RUN.value = "" THEN
		MSGBOX "생성된 SQL구문이 없습니다."
		EXIT Function 
	END IF

	window.open "B1Z01MA1_ViewSQL.asp","","width=780;height=450;resizable=no"

		     
End Function 
</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->
</HEAD>

<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="post">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR >
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>Query 분석기</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><IMG src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23" ></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=*>
					<TABLE CELLSPACING=0 CELLPADDING=0 align=right>
							<TR>
								<td><A HREF="VBSCRIPT:ViewSQL()">SQL 구문</a>&nbsp;&nbsp;
								</td>
						    </TR>
						</TABLE>
						
					</TD>
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
					<TD HEIGHT=20 WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS="TD5" NOWRAP>Query 번호</TD>
									<TD CLASS="TD6" NOWRAP><INPUT NAME="txtQueryCd" MAXLENGTH="18" SIZE=15 ALT ="Query 번호" tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCalType" align=top TYPE="BUTTON" onclick="vbscript:Call OpenQueryID()"></TD>
									<TD CLASS="TD6">&nbsp;</TD>
									<TD CLASS="TD6">&nbsp;</TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				
				
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=100% valign=top>
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD CLASS=TD5 NOWRAP>Query 명</TD>
								<TD CLASS=TD6 NOWRAP COLSPAN=3><INPUT NAME="txtQueryNm" ALT="Query 명" MAXLENGTH="100" SIZE=100 tag = "22XXXU"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>부서</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtDept_cd" ALT="부서코드" TYPE="Text" SiZE=10 MAXLENGTH=10  tag="22XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSItemDC" ALIGN=top TYPE="BUTTON" ONCLICK="VBScript: OpenDept(0)">
			                    <INPUT NAME="txtDept_nm" ALT="부서코드명" TYPE="Text" SiZE=20 MAXLENGTH=40  tag="14">
			                    </TD>
								<TD CLASS=TD5 NOWRAP>권한</TD>
								<TD CLASS=TD6 NOWRAP><SELECT NAME="cboRole_type" ALT="권한" CLASS ="cbonormal" TAG="12X"></SELECT></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>SELECT 절</TD>
								<TD CLASS=TD6 NOWRAP COLSPAN=3>
								<TEXTAREA  NAME="txtSelect" tag="22XXXU" rows=2 cols=100  ALT="SELECT절"></TEXTAREA>
								</TD>
							</TR>
							
							<TR>
								<TD CLASS=TD5 NOWRAP>FROM 절</TD>
								<TD CLASS=TD6 NOWRAP COLSPAN=3><TEXTAREA NAME="txtFrom" rows=2 cols=100 TAG="22XXXU" ALT="FROM절"></TEXTAREA></TD>
							</TR>
							
							<TR>
								<TD CLASS=TD5 NOWRAP>WHERE 절</TD>
								<TD CLASS=TD6 NOWRAP COLSPAN=3><TEXTAREA NAME="txtWhere" rows=2 cols=100 TAG="21X" ALT="WHERE절"></TEXTAREA></TD>
							</TR>							
							
							<TR>
								<TD CLASS=TD5 NOWRAP>기타 절</TD>
								<TD CLASS=TD6 NOWRAP COLSPAN=3><TEXTAREA NAME="txtEtc" rows=2 cols=100 TAG="21X" ALT="기타 절"></TEXTAREA></TD>
							</TR>							
							
							<TR>
								<TD CLASS=TD5 NOWRAP>비고</TD>
								<TD CLASS=TD6 NOWRAP COLSPAN=3><TEXTAREA NAME="txtRemark" rows=2 cols=100 TAG="21X" ALT="비고"></TEXTAREA></TD>
							</TR>	
							<TR>
							<TD HEIGHT="60%" colspan=4>
								<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData ID = "A" WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD">
								<PARAM NAME="MaxCols" VALUE="0">
								<PARAM NAME="MaxRows" VALUE="0">
								</OBJECT>
							</TD>
							</TR>
							
							<TR>
							<TD HEIGHT=0 colspan=4>
								<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData1 ID = "A1" WIDTH=0 HEIGHT=0 tag="23" TITLE="SPREAD1">
								<PARAM NAME="MaxCols" VALUE="0">
								<PARAM NAME="MaxRows" VALUE="0">
								</OBJECT>
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
	
	<TR HEIGHT="20">
		<TD WIDTH="100%">
			<TABLE <%=LR_SPACE_TYPE_30%>>
				<TR>
					<TD Width=10> &nbsp; </TD>
					<TD><BUTTON NAME="btnExeStdCost" CLASS="CLSSBTN" onclick="FncRun()" Flag=1>실행</BUTTON></TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
		
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>>
			<IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=0 FRAMEBORDER=0 SCROLLING=YES noresize framespacing=0 tabindex="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>

<INPUT TYPE=hidden NAME="txtFlgMode" tag="24" tabindex="-1">
<INPUT TYPE=hidden NAME="txtMode" tag="24" tabindex="-1">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMajorFlg" tag="24">
<INPUT TYPE=HIDDEN NAME="txtColCnt" tag="24">
<INPUT TYPE=HIDDEN NAME="lgPageNo" tag="24">
<INPUT TYPE=HIDDEN NAME="StrSelect_RUN" VALUE="<%=StrSelect%>" tag="24">
<INPUT TYPE=hidden NAME="hCommand" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

