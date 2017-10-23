<%@ LANGUAGE="VBSCRIPT" %>
<!--
'======================================================================================================
'*  1. Module Name          : Human Resources
'*  2. Function Name        : 공제관리(국민연금소득총액신고)
'*  3. Program ID           : H5406ma1.asp
'*  4. Program Name         : H5406ma1.asp
'*  5. Program Desc         :
'*  6. Modified date(First) : 2001/05/31
'*  7. Modified date(Last)  : 2003/06/11
'*  8. Modifier (First)     : Hwang Jeong-won
'*  9. Modifier (Last)      : Lee SiNa
'* 10. Comment              :
'=======================================================================================================-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">  

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incHRQuery.vbs"></SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>
Option Explicit
'========================================================================================================
'=                       4.2 Constant variables 
'======================================================================================================== 
Const BIZ_PGM_ID = "h5406mb1.asp"												'비지니스 로직 ASP명 
Const CookieSplit = 1233
Const C_SHEETMAXROWS = 30

'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" --> 
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim IsOpenPop          

Dim C_COMP_CD
Dim C_NO
Dim C_COMP_PAGE
Dim C_RES_NO
Dim C_NAME
Dim C_WORK_MONTH
Dim C_TOT_AMT
Dim C_JISA_CODE
Dim C_EDI_CD
Dim C_EMPTY

'========================================================================================================
' Name : initSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================
Sub initSpreadPosVariables()  

    C_COMP_CD		= 1     	
    C_NO			= 2    
    C_COMP_PAGE		= 3     
    C_RES_NO		= 4     
    C_NAME			= 5
    C_WORK_MONTH	= 6 
    C_TOT_AMT		= 7  
    C_JISA_CODE		= 8  
    C_EDI_CD		= 9 
    C_EMPTY			= 10

End Sub

'======================================================================================================
'	Name : InitVariables()
'	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
'=======================================================================================================
Sub InitVariables()

    lgIntFlgMode = Parent.OPMD_CMODE            'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    
    lgStrPrevKey = ""                           'initializes Previous Key
    lgLngCurRows = 0                            'initializes Deleted Rows Count
    lgSortKey = 1
End Sub

'======================================================================================================
'	Name : SeTDefaultVal()
'	Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
'=======================================================================================================
Sub SetDefaultVal()
	Dim strYear
	Dim strMonth
	Dim strDay

	Call ExtractDateFrom("<%=GetsvrDate%>",Parent.gServerDateFormat , Parent.gServerDateType ,strYear,strMonth,strDay)
	
		
	frm1.txtEntrDt.Year = strYear 		'년월 default value setting
	frm1.txtEntrDt.Month = strMonth
	frm1.txtEntrDt.Day = strDay
	
	frm1.txtYear.Year = strYear 		'년월 default value setting
	frm1.txtYear.Month = strMonth
	frm1.txtYear.Day = strDay

	Call ggoOper.FormatDate(frm1.txtYear, Parent.gDateFormat, 3)
	frm1.txtAutoCd.value = "06"

End Sub

'======================================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'=======================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "H", "NOCOOKIE", "MA") %>
End Sub

'========================================================================================================
' Function Name : MakeKeyStream
' Function Desc : This method set focus to pos of err
'========================================================================================================
Sub MakeKeyStream(pRow)

	lgKeyStream   = Frm1.txtBizArea.Value & Parent.gColSep                 '0
    lgKeyStream   = lgKeyStream & Frm1.txtYear.text & Parent.gColSep       '1
    lgKeyStream   = lgKeyStream & Frm1.txtEntrDt.Text & Parent.gColSep     '2
	lgKeyStream   = lgKeyStream & Frm1.txtArea.Value & Parent.gColSep      '3
    lgKeyStream   = lgKeyStream & Frm1.txtCompCd.Value & Parent.gColSep    '4
    lgKeyStream   = lgKeyStream & Frm1.txtAutoCd.Value & Parent.gColSep    '5

End Sub

'======================================================================================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'=======================================================================================================
Sub InitSpreadSheet()

	Call initSpreadPosVariables()   'sbk 

	With frm1.vspdData

        ggoSpread.Source = frm1.vspdData
        ggoSpread.Spreadinit "V20021121",,parent.gAllowDragDropSpread    'sbk

	    .ReDraw = false
	
        .MaxCols = C_EMPTY + 1												<%'☜: 최대 Columns의 항상 1개 증가시킴 %>
	    .Col = .MaxCols															<%'공통콘트롤 사용 Hidden Column%>
        .ColHidden = True
    
        .MaxRows = 0
        ggoSpread.ClearSpreadData

        Call GetSpreadColumnPos("A") 'sbk
	
	    Call AppendNumberPlace("6","9","0")

        ggoSpread.SSSetEdit C_COMP_CD,		"사업장기호", 12
        ggoSpread.SSSetEdit C_NO,			"일련번호"	, 8
        ggoSpread.SSSetEdit C_COMP_PAGE,	"사업장페이지", 14, 2    
	    ggoSpread.SSSetEdit C_RES_NO,		"주민번호", 13
	    ggoSpread.SSSetEdit C_NAME,			"성명", 10
	    ggoSpread.SSSetEdit C_WORK_MONTH,	"근무월수", 10,,,2,2
'	    ggoSpread.SSSetEdit C_TOT_AMT,		"소득총액" ,9,,,9,2
		ggoSpread.SSSetFloat C_TOT_AMT,		"소득총액" ,  13,"6",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
	    ggoSpread.SSSetEdit C_JISA_CODE,	"지사코드", 10, 2
        ggoSpread.SSSetEdit C_EDI_CD,		"전산화코드", 10, 2
        ggoSpread.SSSetEdit C_EMPTY,		"공란", 8, 2
    
	    .ReDraw = true
	
        Call SetSpreadLock 
    
    End With
    
End Sub

'======================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'=======================================================================================================%>
Sub SetSpreadLock()
    With frm1
    
    .vspdData.ReDraw = False
    
    ggoSpread.SpreadLock C_COMP_CD,			-1 , -1
    ggoSpread.SpreadLock C_NO,				-1 , -1
    ggoSpread.SpreadLock C_COMP_PAGE,		-1 , -1
    ggoSpread.SpreadLock C_RES_NO,			-1 , -1
    ggoSpread.SpreadLock C_NAME,			-1 , -1

	ggoSpread.SSSetRequired C_WORK_MONTH,	-1 , -1
	ggoSpread.SSSetRequired C_TOT_AMT,		-1 , -1    
		
    ggoSpread.SpreadLock C_JISA_CODE,		-1 , -1
    ggoSpread.SpreadLock C_EDI_CD,			-1 , -1
    ggoSpread.SpreadLock C_EMPTY,			-1 , -1
    ggoSpread.SSSetProtected   .vspdData.MaxCols   , -1, -1
     
    .vspdData.ReDraw = True

    End With
End Sub

'======================================================================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'=======================================================================================================
Sub SetSpreadColor(ByVal pvStartRow,ByVal pvEndRow)
    With frm1
    
    .vspdData.ReDraw = False
   
    .vspdData.ReDraw = True
    
    End With
End Sub

'========================================================================================
' Function Name : GetSpreadColumnPos
' Description   : 
'========================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

            C_COMP_CD		= iCurColumnPos(1)
            C_NO			= iCurColumnPos(2)
            C_COMP_PAGE		= iCurColumnPos(3)
            C_RES_NO		= iCurColumnPos(4)
            C_NAME			= iCurColumnPos(5)
            C_WORK_MONTH	= iCurColumnPos(6)
            C_TOT_AMT		= iCurColumnPos(7)
            C_JISA_CODE		= iCurColumnPos(8)
            C_EDI_CD		= iCurColumnPos(9)
            C_EMPTY			= iCurColumnPos(10)            
    End Select    
End Sub

'======================================================================================================
'	Name : Form_Load()
'	Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 
'                 함수를 Call하는 부분 
'=======================================================================================================%>
Sub Form_Load()

    Call LoadInfTB19029                                                     'Load table , B_numeric_format%>
    
    Call ggoOper.LockField(Document, "N")                            'Lock  Suitable  Field%>                         
                                                                            'Format Numeric Contents Field%>                                                                            
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
            
    Call InitSpreadSheet                                                    'Setup the Spread sheet%>
    Call InitVariables                                                      'Initializes local global variables%>
    
    Call SetDefaultVal
    Call SetToolbar("1100000000011111")										'버튼 툴바 제어 %>
    Call CookiePage(0)
    
    frm1.txtBizArea.focus
    
End Sub

'======================================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'=======================================================================================================%>
Sub Form_QueryUnload(Cancel , UnloadMode )

End Sub
'======================================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'=======================================================================================================
Function FncQuery() 
    Dim IntRetCD 
    
    FncQuery = False                                                        
    
    Err.Clear                                                               'Protect system from crashing%>

    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO, "X", "X")			    '데이타가 변경되었습니다. 조회하시겠습니까?%>
    	If IntRetCD = vbNo Then
      	Exit Function
    	End If
    End If
   
   Call ggoOper.ClearField(Document, "2")									'Clear Contents  Field%>
   ggoSpread.ClearSpreadData

    Call InitVariables                                                      'Initializes local global variables%>

    If Not chkField(Document, "1") Then						         'This function check indispensable field%>
       Exit Function
    End If
    if txtArea_OnChange=false then
		exit function
	end if
	
    If Len(frm1.txtBizArea.value) <> 8 Then
		
		Call DisplayMsgBox("970029", "X", frm1.txtBizArea.alt,"X")
		Exit Function
    ElseIf Len(frm1.txtCompCd.value) <> 4 Then
		Call DisplayMsgBox("970029", "X", frm1.txtCompCd.alt,"X")
		Exit Function
	ElseIf Len(frm1.txtAutoCd.value) <> 2 Then
		Call DisplayMsgBox("970029", "X", frm1.txtAutoCd.alt,"X")
		Exit Function
	ElseIf (Len(frm1.txtYear.text) <> 4 or frm1.txtYear.text > "3000" or frm1.txtYear.text < "1900") Then
		Call DisplayMsgBox("970029", "X", frm1.txtYear.alt,"X")
		Exit Function
    End If
    
    Call MakeKeyStream("X")

    If DbQuery = False Then
        Exit Function
    End If
    FncQuery = True															
    
End Function

'======================================================================================================
' Function Name : FncDelete
' Function Desc : This function is related to Delete Button of Main ToolBar
'=======================================================================================================
Function FncDelete()

End Function

 '========================================================================================================
' Name : FncSave
' Desc : developer describe this line Called by MainSave in Common.vbs
'========================================================================================================
Function FncSave()
    Dim IntRetCD 
    
    FncSave = False                                                              '☜: Processing is NG
    
    Err.Clear                                                                    '☜: Clear err status
  
     ggoSpread.Source = frm1.vspdData
    If lgBlnFlgChgValue = False And  ggoSpread.SSCheckChange = False Then
        IntRetCD =  DisplayMsgBox("900001","X","X","X")                           '⊙: No data changed!!
        Exit Function
    End If
 	 
    If Not chkField(Document, "2") Then
       Exit Function
    End If
  
	ggoSpread.Source = frm1.vspdData
    If Not  ggoSpread.SSDefaultCheck Then                                         '⊙: Check contents area
       Exit Function
    End If
    
    Call MakeKeyStream("X")
    Call  DisableToolBar( parent.TBC_SAVE)
	If DBSave=False Then
	   Call  RestoreToolBar()
	   Exit Function
	End If
    
    FncSave = True                                                              '☜: Processing is OK
    
End Function
'========================================================================================================
' Name : DbSave
' Desc : This function is called by FncSave
'========================================================================================================
Function DbSave()
    Dim lRow        
    Dim lGrpCnt     
    Dim retVal      
	Dim strVal, strDel
    Err.Clear                                                                    '☜: Clear err status
		
	DbSave = False														         '☜: Processing is NG
		
    If LayerShowHide(1)=False Then
		Exit Function
    End If

	With frm1
		.txtMode.value        = parent.UID_M0002                                        '☜: Delete
		.txtFlgMode.value     = lgIntFlgMode
        .txtKeyStream.Value   = lgKeyStream                                      '☜: Save Key
	End With

    strVal = ""
    strDel = ""
    lGrpCnt = 1
 
	With Frm1
   
       For lRow = 1 To .vspdData.MaxRows
    
           .vspdData.Row = lRow
           .vspdData.Col = 0
        
           Select Case .vspdData.Text
               Case ggoSpread.InsertFlag                                      '☜: insert추가 
                   
																		   strVal = strVal & "C" & parent.gColSep 'array(0)
																		   strVal = strVal & lRow & parent.gColSep
                                                                           strVal = strVal & Trim(.txtYear.year) &  parent.gColSep
                    .vspdData.Col = C_COMP_CD							 : strVal = strVal & Trim(.vspdData.Text) &  parent.gColSep
                    .vspdData.Col = C_NO								 : strVal = strVal & Trim(.vspdData.Text) &  parent.gColSep
                    .vspdData.Col = C_COMP_PAGE							 : strVal = strVal & Trim(.vspdData.Text) &  parent.gColSep
                 
                    .vspdData.Col = C_RES_NO							 : strVal = strVal & Trim(.vspdData.Text) &  parent.gColSep
                    .vspdData.Col = C_NAME								 : strVal = strVal & Trim(.vspdData.Text) &  parent.gColSep
                    .vspdData.Col = C_WORK_MONTH						 : strVal = strVal & Trim(.vspdData.Text) &  parent.gColSep
                 
                    .vspdData.Col = C_TOT_AMT							 : strVal = strVal & Trim(.vspdData.Text) &  parent.gColSep
                    .vspdData.Col = C_JISA_CODE							 : strVal = strVal & Trim(.vspdData.Text) &  parent.gColSep
                    .vspdData.Col = C_EDI_CD							 : strVal = strVal & Trim(.vspdData.Text) &  parent.gRowSep
                    lGrpCnt = lGrpCnt + 1                                                               
               Case  ggoSpread.UpdateFlag                                      '☜: Update
                                                                           strVal = strVal & "U" &  parent.gColSep
                                                                           strVal = strVal & lRow &  parent.gColSep
                                                                           strVal = strVal & Trim(.txtYear.year) & parent.gColSep
                    .vspdData.Col = C_RES_NO							 : strVal = strVal & Trim(.vspdData.Text) &  parent.gColSep
                    .vspdData.Col = C_WORK_MONTH						 : strVal = strVal & Trim(.vspdData.Text) &  parent.gColSep
                    .vspdData.Col = C_TOT_AMT							 : strVal = strVal & Trim(.vspdData.Text) &  parent.gRowSep

                    lGrpCnt = lGrpCnt + 1
           End Select
       Next

	   .txtMode.value        = parent.UID_M0002
	   .txtMaxRows.value     = lGrpCnt-1	
	   .txtSpread.value      = strDel & strVal

	End With
	
	Call ExecMyBizASP(frm1, BIZ_PGM_ID)
		
    DbSave  = True                                                               '☜: Processing is NG
End Function
'========================================================================================================
' Function Name : DbSaveOk
' Function Desc : Called by MB Area when save operation is successful
'========================================================================================================
Function DbSaveOk()

    ggoSpread.Source = Frm1.vspdData    
	ggoSpread.ClearSpreadData      
    Call InitVariables															'⊙: Initializes local global variables
    
	Call DisableToolBar(parent.TBC_QUERY)
    If DbQuery = False Then
        Call RestoreToolBar()
        Exit Function
    End If
End Function

'======================================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'=======================================================================================================
Function FncCopy()

End Function

'======================================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'=======================================================================================================%>
Function FncCancel() 
    ggoSpread.Source = frm1.vspdData	
    ggoSpread.EditUndo                                                  '☜: Protect system from crashing
End Function

'======================================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'=======================================================================================================%>
Function FncInsertRow() 

End Function

'======================================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'=======================================================================================================%>
Function FncDeleteRow() 
    
End Function

'======================================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'=======================================================================================================%>
Function FncPrint()
    Call parent.FncPrint()                                                   '☜: Protect system from crashing
End Function

'======================================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'=======================================================================================================%>
Function FncExcel() 
    Call parent.FncExport(Parent.C_MULTI)											 <%'☜: 화면 유형 %>
End Function

'======================================================================================================
' Function Name : FncFind
' Function Desc : 
'=======================================================================================================%>
Function FncFind() 
    Call parent.FncFind(Parent.C_MULTI, False)                                      <%'☜:화면 유형, Tab 유무 %>
End Function

'========================================================================================
' Function Name : FncSplitColumn
' Function Desc : 
'========================================================================================
Sub FncSplitColumn()

    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   

End Sub


'========================================================================================
' Function Name : PopSaveSpreadColumnInf
' Description   : 
'========================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub

'========================================================================================
' Function Name : PopRestoreSpreadColumnInf
' Description   : 
'========================================================================================
Sub PopRestoreSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()      
	Call ggoSpread.ReOrderingSpreadData()
	Call InitData()
End Sub

'======================================================================================================
' Function Name : FncExit
' Function Desc : 
'=======================================================================================================%>
Function FncExit()
Dim IntRetCD
	
	FncExit = False
	
    ggoSpread.Source = frm1.vspdData	
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO, "X", "X")                <%'데이타가 변경되었습니다. 종료 하시겠습니까?%>
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    FncExit = True
End Function

'======================================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'=======================================================================================================%>
Function DbQuery() 
	Dim strVal
	
    DbQuery = False
    
    Err.Clear                                                               <%'☜: Protect system from crashing%>

	If LayerShowHide(1) =False Then
       Exit Function
    End If
	
    With frm1
    
		strVal = BIZ_PGM_ID & "?txtMode="            & Parent.UID_M0001						         
        strVal = strVal     & "&txtKeyStream="       & lgKeyStream                       '☜: Query Key
        strVal = strVal     & "&txtMaxRows="         & .vspdData.MaxRows
        strVal = strVal     & "&lgStrPrevKey=" & lgStrPrevKey                 '☜: Next key tag
    	    
    End With
    
    Call RunMyBizASP(MyBizASP, strVal)										<%'☜: 비지니스 ASP 를 가동 %>
    DbQuery = True
    
End Function

'======================================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'=======================================================================================================%>
Function DbQueryOk()													<%'조회 성공후 실행로직 %>
	
    lgIntFlgMode = Parent.OPMD_UMODE
    Call ggoOper.LockField(Document, "Q")									<%'This function lock the suitable field%>
    Call SetToolbar("1100100000011111")	    
	frm1.vspdData.focus	
End Function

'------------------------------------------  OpenArea()  -------------------------------------------
'	Name : OpenArea()
'	Description : 근무구역 PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenArea()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "근무구역 팝업"				<%' 팝업 명칭 %>
	arrParam(1) = "b_minor"							<%' TABLE 명칭 %>
	arrParam(2) = frm1.txtArea.value				<%' Code Condition%>
	arrParam(3) = ""								<%' Name Cindition%>
	arrParam(4) = " major_cd = " & FilterVar("H0035", "''", "S") & " "			<%' Where Condition%>
	arrParam(5) = "근무구역"					<%' 조건필드의 라벨 명칭 %>
	
    arrField(0) = "minor_cd"					<%' Field명(0)%>
    arrField(1) = "minor_nm"					<%' Field명(1)%>
    
    arrHeader(0) = "코드"					<%' Header명(0)%>
    arrHeader(1) = "코드명"					<%' Header명(1)%>
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtArea.focus
		Exit Function
	Else
		Call SetArea(arrRet)
	End If	
	
End Function

'------------------------------------------  SetArea()  --------------------------------------------
'	Name : SetArea()
'	Description : 근태코드 Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- %>
Function SetArea(Byval arrRet)
		
	With frm1		
		.txtArea.value = arrRet(0)
		.txtAreaNm.value = arrRet(1)
		.txtArea.focus
	End With
	
End Function

'==========================================================================================
'   Event Name : btnBatch_OnClick()
'   Event Desc : 파일생성 
'==========================================================================================
Function btnBatch_OnClick()
	Dim RetFlag
	Dim strVal
	Dim intRetCD

    Err.Clear                                                                   '☜: Clear err status
    
    If Not chkField(Document, "1") Then                                         ' Required로 표시된 Element들의 입력 [유/무]를 Check 한다.
       Exit Function                            
    End If
    
    If Len(frm1.txtBizArea.value) <> 8 Then
		Call DisplayMsgBox("970029", "X", frm1.txtBizArea.alt,"X")
		Exit Function
    ElseIf Len(frm1.txtCompCd.value) <> 4 Then
		Call DisplayMsgBox("970029", "X", frm1.txtCompCd.alt,"X")
		Exit Function
	ElseIf Len(frm1.txtAutoCd.value) <> 2 Then
		Call DisplayMsgBox("970029", "X", frm1.txtAutoCd.alt,"X")
		Exit Function
	ElseIf (Len(frm1.txtYear.text) <> 4 or frm1.txtYear.text > "3000" or frm1.txtYear.text < "1900") Then
		Call DisplayMsgBox("970029", "X", frm1.txtYear.alt,"X")
		Exit Function
    End If
    
    If frm1.vspdData.MaxRows <= 0 Then
		Call DisplayMsgBox("900002", "X","X","X")			 '⊙: Query First 
		Exit Function		
    End If

	RetFlag = DisplayMsgBox("900018", Parent.VB_YES_NO,"x","x")   '☜ 바뀐부분	
	If RetFlag = VBNO Then
		Exit Function
	End IF

	If LayerShowHide(1) =False Then
       Exit Function
    End If
	
    With frm1
	    
		strVal = BIZ_PGM_ID & "?txtMode="            & Parent.UID_M0003						         
        strVal = strVal     & "&txtKeyStream="       & lgKeyStream                       '☜: Query Key
        strVal = strVal     & "&txtMaxRows="         & .vspdData.MaxRows
        strVal = strVal     & "&lgStrPrevKey=" & lgStrPrevKey                 '☜: Next key tag
    End With    
    
    Call RunMyBizASP(MyBizASP, strVal)
End Function

Function subVatDiskOK(ByVal pFileName) 
Dim strVal
    Err.Clear                                                               '☜: Protect system from crashing

	strVal = BIZ_PGM_ID & "?txtMode=" & "7"							'☜: 비지니스 처리 ASP의 상태 
	strVal = strVal & "&txtFileName=" & pFileName							'☆: 조회 조건 데이타	
	Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 
End Function



'======================================================================================================
'   Event Name : vspdData_Change
'   Event Desc :
'=======================================================================================================%>
Sub vspdData_Change(ByVal Col , ByVal Row )
       
   	Frm1.vspdData.Row = Row
   	Frm1.vspdData.Col = Col

   	If Frm1.vspdData.CellType = Parent.SS_CELL_TYPE_FLOAT Then
      If UNICDbl(Frm1.vspdData.text) < UNICDbl(Frm1.vspdData.TypeFloatMin) Then
         Frm1.vspdData.text = Frm1.vspdData.TypeFloatMin
      End If
	End If
	
	ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row
End Sub

'======================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'=======================================================================================================%>
Sub vspdData_Click(ByVal Col, ByVal Row)
    
    Call SetPopupMenuItemInf("0000111111")

    gMouseClickStatus = "SPC" 

    Set gActiveSpdSheet = frm1.vspdData
   
    If frm1.vspdData.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
   	    
    If Row <= 0 Then
       ggoSpread.Source = frm1.vspdData
       
       If lgSortKey = 1 Then
           ggoSpread.SSSort Col               'Sort in ascending
           lgSortKey = 2
       Else
           ggoSpread.SSSort Col, lgSortKey    'Sort in descending 
           lgSortKey = 1
       End If
       
       Exit Sub
    End If
        
End Sub

'========================================================================================================
'   Event Name : vspdData_DblClick
'   Event Desc : 
'========================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)				
    Dim iColumnName
    
    If Row <= 0 Then
        Exit Sub
    End If
    
    If frm1.vspdData.MaxRows = 0 Then
        Exit Sub
    End If
End Sub

'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

End Sub

'========================================================================================================
'   Event Name : vspdData_MouseDown
'   Event Desc : 
'========================================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
     End If
End Sub    

'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
End Sub

'======================================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'=======================================================================================================%>
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
   	If OldLeft <> NewLeft Then
		Exit Sub
	End If
	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then
		If lgStrPrevKey <> "" Then
			If CheckRunningBizProcess = True Then
				Exit Sub
			End If	
			
			Call DisableToolBar(Parent.TBC_QUERY)
			If DBQuery = False Then
				Call RestoreToolBar()
				Exit Sub
			End If
		End If
	End If  
End Sub
'=======================================================================================================
'   Event Name : txtEntrDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtEntrDt_DblClick(Button)
    If Button = 1 Then
		Call SetFocusToDocument("M")    
        frm1.txtEntrDt.Action = 7
        frm1.txtEntrDt.focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtEntrDt_Keypress(Key)
'   Event Desc : 3rd party control에서 Enter 키를 누르면 조회 실행 
'=======================================================================================================
Sub txtEntrDt_Keypress(Key)
    If Key = 13 Then
        FncQuery()
    End If
End Sub



'=======================================================================================================
'   Event Name : txtYear_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtYear_DblClick(Button)
    If Button = 1 Then
		Call SetFocusToDocument("M")    
        frm1.txtYear.Action = 7
        frm1.txtYear.focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtYear_Keypress(Key)
'   Event Desc : 3rd party control에서 Enter 키를 누르면 조회 실행 
'=======================================================================================================
Sub txtYear_Keypress(Key)
    If Key = 13 Then
        FncQuery()
    End If
End Sub

'=======================================================================================================
'   Event Name : txtYear_Keypress(Key)
'   Event Desc : 3rd party control에서 Enter 키를 누르면 조회 실행 
'=======================================================================================================
Sub txtYear_Keypress(Key)
    If Key = 13 Then
        FncQuery()
    End If
End Sub


'========================================================================================================
'   Event Name : txtCd_OnChange
'   Event Desc :
'========================================================================================================
Function txtArea_OnChange()    

    Dim IntRetCd
	dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
    If frm1.txtArea.value = "" Then
        frm1.txtAreaNm.value = ""
		txtArea_OnChange = true
    ELSE    
        IntRetCd = CommonQueryRs(" minor_nm "," b_minor "," major_cd = " & FilterVar("H0035", "''", "S") & " AND minor_cd =  " & FilterVar(frm1.txtArea.value, "''", "S") & "" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
        If IntRetCd = false Then
            Call DisplayMsgBox("970000","X","근무구역","X")
            frm1.txtAreaNm.value=""
            frm1.txtArea.focus
			Set gActiveElement = document.ActiveElement 
            Exit Function
        Else
            frm1.txtAreaNm.value=Trim(Replace(lgF0,Chr(11),""))
			txtArea_OnChange = true
        End If
    End If
End Function
'==========================================================================================
'   Event Name : btnCb_select_OnClick
'   Event Desc : 데이터 가져오기 
'==========================================================================================
Function btnCb_select_OnClick()
	Dim RetFlag ,RetFlag2
	Dim strVal
	Dim intRetCD,strWhere, strEmp_no

    Err.Clear                                                                           '☜: Clear err status
'	If gSelframeFlg = TAB1 Then      
'		If Not chkField(Document, "1") Then                                                 'Required로 표시된 Element들의 입력 [유/무]를 Check 한다.
'		   Exit Function                            
'		End If
'	End If
		
      ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO, "X", "X")			    '데이타가 변경되었습니다. 조회하시겠습니까?%>
    	If IntRetCD = vbNo Then
      	Exit Function
    	End If
    End If
   
   Call ggoOper.ClearField(Document, "2")									'Clear Contents  Field%>
   ggoSpread.ClearSpreadData

    Call InitVariables                                                      'Initializes local global variables%>

    If Not chkField(Document, "1") Then						         'This function check indispensable field%>
       Exit Function
    End If
 		
	strWhere = " YEAR_YY = " & FilterVar(Frm1.txtYear.Year, "''", "S")
	 
 	IntRetCD = CommonQueryRs(" * "," HDB040T ", strWhere ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
 	
	If IntRetCD = True  Then
		        
		IntRetCD = DisplayMsgBox("800502", 35,"X","X")	    '이미 생성된 자료가 있습니다.삭제하시겠습니까?
		If IntRetCD = vbCancel Then
		   	Exit Function
		End If
	End If
					
'	RetFlag = DisplayMsgbox("900018", parent.VB_YES_NO,"x","x")                         '☜ 작업을 계속하시겠습니까?
'	If RetFlag = VBNO Then
'		Exit Function
'	End IF
	ggoSpread.ClearSpreadData
	
    With frm1
        Call LayerShowHide(1)					 
        lgCurrentSpd = "A"		
	    Call MakeKeyStream(lgCurrentSpd)  
        
		strVal = BIZ_PGM_ID    & "?txtMode="           & "5"						'☜: 비지니스 처리 ASP의 상태 	    	    		    
		strVal = strVal         & "&lgCurrentSpd="      & lgCurrentSpd                  '☜: Mulit의 종류 
		strVal = strVal         & "&txtKeyStream="      & lgKeyStream                   '☜: Query Key

		Call RunMyBizASP(MyBizASP, strVal)
 
    End With    
End Function

Sub DBAutoQueryOk()
    Dim lRow
    With Frm1

        .vspdData.ReDraw = false
        ggoSpread.Source = .vspdData
       For lRow = 1 To .vspdData.MaxRows
            .vspdData.Row = lRow
            .vspdData.Col = 0

            .vspdData.Text = ggoSpread.InsertFlag
        Next

'      ggoSpread.SpreadLock C_CHANG_DT, -1,C_CHANG_DT
    .vspdData.ReDraw = TRUE
    ggoSpread.ClearSpreadData "T"            
    End With    
    Call SetToolbar("1100100000011111")	    
    lgStrPrevKey = ""
    Set gActiveElement = document.ActiveElement   
End Sub


</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
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
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><IMG src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23" ></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>국민연금소득총액신고</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><IMG src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23" ></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* ALIGN=LIGHT>&nbsp;</TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
		
    <TR HEIGHT=*>
		<TD CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
			    <TR><TD <%=HEIGHT_TYPE_02%>></TD></TR>
				<TR>
					<TD HEIGHT=20>
					  <FIELDSET CLASS="CLSFLD">
					   <TABLE <%=LR_SPACE_TYPE_40%>>
						    <TR>
						        <TD CLASS=TD5 NOWRAP>사업장기호</TD>       
					            <TD CLASS=TD6><INPUT TYPE=TEXT NAME="txtBizArea" SIZE=20 MAXLENGTH=8 tag="12XXXU"  ALT="사업장기호"></TD>
							    <TD CLASS=TD5 NOWRAP>기준년도</TD>
			                    <TD CLASS=TD6><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 name=txtYear CLASS=FPDTYYYY title=FPDATETIME tag="12X1" ALT="기준년도"></OBJECT>');</SCRIPT></TD>
			                <TR>
								<TD CLASS=TD5 NOWRAP>입사기준일</TD>
			                    <TD CLASS=TD6><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 name=txtEntrDt CLASS=FPDTYYYYMMDD title=FPDATETIME tag="12X1" ALT="입사기준일"></OBJECT>');</SCRIPT></TD>
								<TD CLASS=TD5>근무구역</TD>
								<TD CLASS=TD6>
								<INPUT TYPE=TEXT NAME="txtArea" SIZE=10 MAXLENGTH=2 tag="11XXXU"  ALT="근무구역"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnArea" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenArea()">
								<INPUT TYPE=TEXT NAME="txtAreaNm" tag="14X"></TD>					           
							</TR>
						    <TR>
								<TD CLASS=TD5 NOWRAP>지사코드</TD>       
					            <TD CLASS=TD6><INPUT TYPE=TEXT NAME="txtCompCd" SIZE=10 MAXLENGTH=4 tag="12" STYLE="TEXT-ALIGN: center" ALT="지사코드"></TD>
					            <TD CLASS=TD5 NOWRAP>전산화코드</TD>       
					            <TD CLASS=TD6><INPUT TYPE=TEXT NAME="txtAutoCd" SIZE=10 MAXLENGTH=2 tag="12" STYLE="TEXT-ALIGN: center" ALT="전산화코드"></TD>
					        </TR>		
					  </TABLE>
				     </FIELDSET>
				   </TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%>></TD>
				</TR>				
				<TR>
					<TD WIDTH=100% HEIGHT=* VALIGN=TOP>
						<TABLE <%=LR_SPACE_TYPE_20%> >
							<TR>
								<TD HEIGHT="100%">
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=vaSpread> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
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
		<TD>
			<TABLE <%=LR_SPACE_TYPE_30%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD><BUTTON NAME="btnCb_select" CLASS="CLSMBTN">데이터생성</BUTTON>&nbsp;
						<BUTTON NAME="btnBatch" CLASS="CLSMBTN" Flag="1">파일생성</BUTTON></TD>
				</TR>
			</TABLE>
		</TD>
	</TR>	
	<TR>
		<TD WIDTH=100% HEIGHT=0><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=0 FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME></TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA><INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24"><%' 업무처리ASP로 넘기기 위한 변수를 담고 있는 Tag들 %>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24"><INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="hEmp_no" tag="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode"     TAG="24">
<INPUT TYPE=HIDDEN NAME="txtKeyStream"   TAG="24">
<!--
<INPUT TYPE=HIDDEN NAME="hMajorCd" tag="24">
-->
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

