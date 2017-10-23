<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% session.CodePage=949 %>

<!-- '======================================================================================================
'*  1. Module Name          : Cost
'*  2. Function Name        : 직과항목계정등록 
'*  3. Program ID           : c4009mb1.asp
'*  4. Program Name         : 직과항목계정등록 
'*  5. Program Desc         : 직과항목계정등록 
'*  6. Modified date(First) : 2005-08-30
'*  7. Modified date(Last)  : 2005-08-30
'*  8. Modifier (First)     : choe0tae 
'*  9. Modifier (Last)      : 
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'======================================================================================================= -->

<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!--
########################################################################################################
#						   3.    External File Include Part
########################################################################################################-->

<!--
========================================================================================================
=                          3.1 Server Side Script
========================================================================================================-->
<!-- #Include file="../../inc/IncSvrCcm.inc"  -->

<!--
========================================================================================================
=                          3.2 Style Sheet
======================================================================================================== -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		

<!--
========================================================================================================
=                          3.3 Client Side Script
======================================================================================================== -->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliRdsQuery.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incEB.vbs">	</SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>
Option Explicit

'########################################################################################################
'#                       4.  Data Declaration Part
'########################################################################################################

'========================================================================================================
'=                       4.1 External ASP File
'========================================================================================================

Const BIZ_PGM_ID = "c4009mb1.asp"                               'Biz Logic ASP

' -- 그리드1의 컬럼 정의 
Dim C_ACCT_CD
Dim C_ACCT_CD_POP
Dim C_ACCT_NM
Dim C_MINOR_CD			' -- 헤더 키 
Dim C_MINOR_CD_POP		' -- 배부순서 
Dim C_MINOR_NM		' -- C/C 래벨 


Const GRID_1	= 1
Const GRID_2	= 2
'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	


'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim lgQueryFlag
Dim IsOpenPop          
Dim lgCurrGrid
Dim lgCopyVersion
Dim lgErrRow, lgErrCol
'========================================================================================================
'                        5.1 Common Method-1
'========================================================================================================
'========================================================================================================

'========================================================================================================
' Name : initSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================
Sub initSpreadPosVariables()  
	
	' -- 그리드1의 컬럼 정의 
	 C_ACCT_CD					= 1
	 C_ACCT_CD_POP				= 2
	 C_ACCT_NM					= 3
	 C_MINOR_CD					= 4
	 C_MINOR_CD_POP				= 5
	 C_MINOR_NM					= 6

End Sub

'========================================================================================================
' Name : InitVariables()	
' Desc : Initialize value
'======================================================================================================== 
Sub InitVariables()

    lgIntFlgMode = Parent.OPMD_CMODE
    lgBlnFlgChgValue = False 
    lgIntGrpCount = 0  
    
    lgStrPrevKey = ""	
    lgLngCurRows = 0 
	lgSortKey = 1
	lgCurrGrid = GRID_1
	lgCopyVersion = ""
	lgErrRow = 0 : lgErrCol = 0
End Sub

'========================================================================================================
' Name : SetDefaultVal()	
' Desc : Set default value
'========================================================================================================
Sub SetDefaultVal()
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub

'========================================================================================================
' Name : LoadInfTB19029()	
' Desc : Set System Number format
' Para : 1. Currency
'        2. I(Input),Q(Query),P(Print),B(Bacth)
'        3. "*" is for Common module
'           "A" is for Accounting
'           "I" is for Inventory
'           ...
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
	<%Call LoadInfTB19029A("I","*", "NOCOOKIE", "MA") %>
End Sub

'========================================================================================================
' Name : CookiePage()
' Desc : Write or Read cookie value 
'========================================================================================================
Sub CookiePage(Kubun)
   '------ Developer Coding part (Start ) -------------------------------------------------------------- 
   '------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub

'========================================================================================================
' Name : MakeKeyStream
' Desc : This method set focus to pos of err
'========================================================================================================
Sub MakeKeyStream(pOpt)
   
   '------ Developer Coding part (Start ) -------------------------------------------------------------- 
       
   '------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub        


'========================================================================================================
' Name : InitSpreadSheet
' Desc : This method initializes spread sheet column property
'========================================================================================================
Sub InitSpreadSheet()
	' -- 그리드 컬럼 위치 초기화 
	Call initSpreadPosVariables()    
	
	'Call AppendNumberPlace("6","3","0")
	'Call AppendNumberPlace("7","2","0")
	' -- 그리드 1 정의 
	With frm1.vspdData
	
	.MaxCols = C_MINOR_NM + 1

	.Col = .MaxCols	
    .ColHidden = True
    
    ggoSpread.Source = frm1.vspdData
    ggoSpread.Spreadinit "V20030825",,"" ',,parent.gAllowDragDropSpread 

	Call ggoSpread.ClearSpreadData()    '☜: Clear spreadsheet data 

	.ReDraw = false

	Call GetSpreadColumnPos("A")
	
    ggoSpread.SSSetEdit		C_ACCT_CD		,"계정코드",18,,, 18,2
    ggoSpread.SSSetButton	C_ACCT_CD_POP
    ggoSpread.SSSetEdit		C_ACCT_NM		,"계정명"	,30
    ggoSpread.SSSetEdit		C_MINOR_CD		,"직과항목코드",10,,, 3,2
    ggoSpread.SSSetButton	C_MINOR_CD_POP
    ggoSpread.SSSetEdit		C_MINOR_NM		,"직과항목명"	,30

	'Call ggoSpread.SSSetColHidden(C_MINOR_CD,C_MINOR_CD,True)
	
'	.rowheight(-1000) = 20	' 높이 재지정 

	.ReDraw = true
	
    Call SetSpreadLock 
    'Call InitComboBox
    
    End With
    
    
End Sub

'======================================================================================================
' Name : SetSpreadLock
' Desc : This method set color and protect cell in spread sheet
'======================================================================================================
Sub SetSpreadLock()
    With frm1.vspdData
    
    ggoSpread.Source = frm1.vspdData
    
    .ReDraw = False
    ggoSpread.SpreadLock		C_ACCT_CD			,-1	,C_ACCT_CD
    ggoSpread.SpreadLock		C_ACCT_NM			,-1	,C_ACCT_NM		
    ggoSpread.SpreadLock		C_ACCT_CD_POP			,-1	,C_ACCT_CD_POP		
    ggoSpread.SSSetRequired		C_MINOR_CD			,-1	, .MaxRows
	ggoSpread.SpreadLock		C_MINOR_NM			,-1	,C_MINOR_NM
    .ReDraw = True

    End With
End Sub


'======================================================================================================
' Name : SetSpreadColor
' Desc : This method set color and protect cell by cell in spread sheet
'======================================================================================================
Sub SetSpreadColor(ByVal pvStartRow,ByVal pvEndRow)
    With frm1
    
    ggoSpread.Source = frm1.vspdData
    .vspdData.ReDraw = False
								      'Col          Row				Row2    
	ggoSpread.SSSetRequired		C_ACCT_CD			,pvStartRow		,pvEndRow
	ggoSpread.SSSetProtected	C_ACCT_NM			,pvStartRow		,pvEndRow    
	ggoSpread.SSSetRequired		C_MINOR_CD			,pvStartRow		,pvEndRow    
	ggoSpread.SSSetProtected	C_MINOR_NM			,pvStartRow		,pvEndRow    
	
    .vspdData.ReDraw = True
    
    End With
End Sub

'======================================================================================================
' Name : SubSetErrPos
' Desc : This method set focus to position of error
'      : This method is called in MB area
'======================================================================================================
Sub SubSetErrPos(iPosArr)
    Dim iDx, oGrid, j, sAcctCd, sMinorCd
    Dim iRow
    If iPosArr = "" Then Exit Sub
    iPosArr = Split(iPosArr,Parent.gColSep)		' 리턴문자열: 그리드n/gColSep/상태플래그/gColSep/계정코드/gColSep/직과항목코드 
    
	Set oGrid = frm1.vspdData

	With oGrid
		
	For iRow = 1 To  .MaxRows 
	    .Col = 0
	    .Row = iRow
		    
		' -- 그리드2 일 경우 
		.Col = C_ACCT_CD		: sAcctCd	= Trim(.value)
		.Col = C_MINOR_CD		: sMinorCd = Trim(.value)
		If sAcctCd = iPosArr(2) And sMinorCd = Trim(iPosArr(3)) Then	' -- 에러행번호와 SEQ_NO가 같다면 
			.Col = C_MINOR_CD	: .Action  = 0	
			lgErrRow = iRow		' -- 에러난 행지정 
			Exit Sub
		End If
					
	Next
        
    End With 

End Sub

'======================================================================================================
Sub SubSetErrPos2(Byval iPosArr)
    Dim iDx, oGrid, j, iSeqNo, iSubSeqNo
    Dim iRow
    If iPosArr = "" Then Exit Sub
    iPosArr = Split(iPosArr,Parent.gColSep)		' 리턴문자열: ACCT_CD/MINOR_CD/ERR_CODE
    
    If IsNumeric(iPosArr(2)) Then
    
		Select Case CInt(iPosArr(4))
			Case 1	' -- MINOR_CD is null
				Call DisplayMsgBox("970000", "x",iPosArr(1),"x")
			Case 2	' -- CTRL_CD is null
				Call DisplayMsgBox("970000", "x",iPosArr(2),"x")
			Case 3	' -- SP_NM is null
				Call DisplayMsgBox("970000", "x",iPosArr(1),"x")
		End Select
	End If
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
            
			' -- 그리드1의 컬럼 정의 
			 C_ACCT_CD					= iCurColumnPos(1)	
			 C_ACCT_CD_POP				= iCurColumnPos(2)
			 C_ACCT_NM					= iCurColumnPos(3)
			 C_MINOR_CD					= iCurColumnPos(4)	
			 C_MINOR_CD_POP				= iCurColumnPos(5)
			 C_MINOR_NM					= iCurColumnPos(6)

		
    End Select    
End Sub

'******************************************  2.4 POP-UP 처리함수  ****************************************
'	기능: 기준 POP-UP
'   설명: 기준 POP-UP에 관한 Open은 Include한다. 
'	      하나의 ASP에서 Popup이 중복되면 하나는 링크시켜 사용하고 나머지는 재정의하여 사용한다.
'********************************************************************************************************* 
 
'------------------------------------------  OpenCalType()  ----------------------------------------------
'	Name :InitComboBox
'	Description : 
'--------------------------------------------------------------------------------------------------------- 
Sub InitComboBox
    
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
     
End Sub

' -- 직과항목 팝업시.
Function OpenAcct(Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	Select Case iWhere
		Case 1

			If Not chkField(Document, "1") Then
			   Exit Function
			End If

			Dim IntRetCD , blnChange1, blnChange2
    
			Err.Clear
    
			ggoSpread.Source = frm1.vspdData
			blnChange1 = ggoSpread.SSCheckChange

			ggoSpread.Source = frm1.vspdData2
			blnChange2 = ggoSpread.SSCheckChange
    
			If blnChange1 = True Or blnChange2 = True Then
				IntRetCD = DisplayMsgBox("900017", Parent.VB_YES_NO,"x","x")
				If IntRetCD = vbNo Then
			      	Exit Function
				End If
			End If
	End Select

	IsOpenPop = True

	arrParam(0) = "계정 팝업"
	arrParam(1) = "A_ACCT"	
	
	If iWhere = 0 Then	' -- 그냥 팝업 
		arrParam(2) = Trim(frm1.txtACCT_CD.Value)
	End If
	
	arrParam(3) = ""
	arrParam(4) = "TEMP_FG_3 IN (" & FilterVar("M2", "''", "S") & "," &  FilterVar("M3", "''", "S") & "," & FilterVar("M4", "''", "S") & ")"
	
	If iWhere = 1 And frm1.txtACCT_CD.value <> "" Then		' -- 카피 팝업시 추가조건 
		arrParam(4) = arrParam(4) & " AND ACCT_CD <> " & FilterVar(frm1.txtACCT_CD.value, "''", "S")
	End If

	arrParam(5) = "계정"
	
    arrField(0) = "ACCT_CD"
    arrField(1) = "ACCT_NM"
    
    arrHeader(0) = "계정"
    arrHeader(1) = "계정명"
    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtACCT_CD.focus
		Exit Function
	Else
		Call SetMinor(arrRet, iWhere)
	End If
		
End Function

' -- 직과항목 팝업후 
Function SetMinor(byval arrRet, Byval iWhere)
	Select Case iWhere
		Case 0
			frm1.txtACCT_CD.focus
			frm1.txtACCT_CD.Value    = arrRet(0)
			frm1.txtACCT_NM.Value    = arrRet(1)				

		Case 1
			IF LayerShowHide(1) = False Then
				Exit Function
			END IF

			Dim strVal
	
			With frm1
				strVal = BIZ_COPY_PGM_ID & "?txtMode=" & Parent.UID_M0001
				strVal = strVal & "&txtACCT_CD=" & Trim(.txtACCT_CD.value)	
				strVal = strVal & "&hCopyVerCd=" & arrRet(0)
				
				Call RunMyBizASP(MyBizASP, strVal)
   
			End With
	End Select
    
End Function


' -- 그리드1에서 팝업 클릭시 
Function OpenPopUp(Byval iWhere, Byval strCode, Byval strCode1)
	Dim arrRet, sTmp, iWidth
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	With frm1
	
	iWidth = 500	' -- 팝업Width
	
	Select Case iWhere

		Case C_ACCT_CD_POP
			arrParam(0) = "계정 팝업"
			arrParam(1) = "A_ACCT"	
			arrParam(2) = strCode
			arrParam(3) = ""		
			arrParam(4) = "TEMP_FG_3 IN (" & FilterVar("M2", "''", "S") & "," & FilterVar("M3", "''", "S") & "," & FilterVar("M4", "''", "S") & ")"
			arrParam(5) = "계정" 

			arrField(0) = "ACCT_CD"	
			arrField(1) = "ACCT_NM"
    
			arrHeader(0) = "계정"	
			arrHeader(1) = "계정명"	

		Case C_MINOR_CD_POP
			arrParam(0) = "직과항목 팝업"
			arrParam(1) = "c_dir_dstb_fctr_s a(nolock) inner join b_minor b(nolock) on a.minor_cd=b.minor_cd and b.major_cd = " & FilterVar("C4010", "''", "S")
			arrParam(2) = strCode
			arrParam(3) = ""		
			arrParam(4) = ""
			arrParam(5) = "직과항목코드" 

			arrField(0) = "a.MINOR_CD"	
			arrField(1) = "b.MINOR_NM"
    
			arrHeader(0) = "직과항목코드"	
			arrHeader(1) = "직과항목명"	
			
	End Select
    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=" & CStr(iWidth) & "px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetPopUp(arrRet, iWhere)
	End If	

	End With
End Function


Function SetPopUp(Byval arrRet, Byval iWhere)
	Dim sTmp
	
	With frm1.vspdData
	
	Select Case iWhere
		
		Case C_MINOR_CD_POP
			
			.Col = C_MINOR_CD	: .Text = arrRet(0)
			.Col = C_MINOR_NM	: .Text = arrRet(1)

			Call vspdData_Change(C_MINOR_CD, .ActiveRow)

		Case C_ACCT_CD_POP
			
			.Col = C_ACCT_CD	: .Text = arrRet(0)
			.Col = C_ACCT_NM	: .Text = arrRet(1)

			Call vspdData_Change(C_ACCT_CD, .ActiveRow)

	End Select
		
	lgBlnFlgChgValue = True
	
	End With
End Function

' -- 문자형리턴 
Function GetGridTxt(Byref pObj, Byval pCol, Byval pRow)
	With pObj
		.Col = pCol	: .Row = pRow
		GetGridTxt = Trim(.Text)
	End With
End Function

' -- 값 리턴 
Function GetGridVal(Byref pObj, Byval pCol, Byval pRow)
	With pObj
		.Col = pCol	: .Row = pRow
		GetGridVal = .Value
	End With
End Function

Sub SetGridTxt(Byref pObj, Byval pCol, Byval pRow, Byval pVal)
	With pObj
		.Col = pCol	: .Row = pRow
		.Text = pVal
	End With
End Sub

Sub SetGridVal(Byref pObj, Byval pCol, Byval pRow, Byval pVal)
	With pObj
		.Col = pCol	: .Row = pRow
		.Value = pVal
	End With
End Sub
'========================================================================================================
'========================================================================================================
'                        5.2 Common Method-2
'========================================================================================================
'========================================================================================================

'========================================================================================================
' Name : Form_Load
' Desc : This sub is called from window_OnLoad in Common.vbs automatically
'========================================================================================================
Sub Form_Load()
	
    Call LoadInfTB19029
    Call ggoOper.LockField(Document, "N") 
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)

    Call InitSpreadSheet
    Call InitVariables
    
'	Call InitComboBox
    Call SetDefaultVal
    Call SetToolbar("111011010011111")	
    frm1.txtACCT_CD.focus
   	Set gActiveElement = document.activeElement			    
     
End Sub

'========================================================================================================
' Name : Form_QueryUnload
' Desc : This sub is called from window_Unonload in Common.vbs automatically
'========================================================================================================
Sub Form_QueryUnload(Cancel, UnloadMode)
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    '--------- Developer Coding Part (End) ------------------------------------------------------------
End Sub

'========================================================================================================
'========================================================================================================
'                        5.4 User-defined Method
'========================================================================================================

'==========================================================================================
'   Event Desc : Grid의 Max Count 넣어준다 
'==========================================================================================
Function InsertParentMinorCd(Byval pSeqNo, Byval pCol, Byval pRow1, Byval pRow2)
	Dim iRow
	With frm1.vspdData2
		For iRow = pRow1 To pRow2
			.Row = iRow	: .Col = pCol	: .Text = pSeqNo
		Next
	End With
end Function

'==========================================================================================
'   Event Desc : 그리드 보이기/숨기기 
'==========================================================================================
Function ShowRowHidden(Byval pMinorCd)
	Dim iRow, iMaxRows, iFirstRow, sMinorCd
	
	With frm1.vspdData2 
	
	iMaxRows = .MaxRows : iFirstRow = 0
	
	.ReDraw = False
	
	For iRow = 1 To iMaxRows
	
		.Row = iRow	: .Col = C_MINOR_CD	: sMinorCd = Trim(.Value)
		 
		If sMinorCd = pMinorCd Then
			.RowHidden = False
			If iFirstRow = 0 Then iFirstRow = iRow
		Else
			.RowHidden = True
		End If 

	Next
	
	.ReDraw = True
	
	ShowRowHidden = iFirstRow
	
	End With
	
End Function

'==========================================================================================
'   Event Desc : 2번 그리드 전체 삭제 루틴 : FncCancel 시 
'==========================================================================================
Function CancelChildGrid2()
	Dim iCol, iRow, iMaxRows, sMinorCd, lDelRows, sFlag, iChildSeqNo
	
	With frm1.vspdData2
		iMaxRows = .MaxRows
		
		' -- 부모키 조회 
		sMinorCd = GetGridTxt(frm1.vspdData, C_MINOR_CD, frm1.vspdData.ActiveRow)

		ggoSpread.Source = frm1.vspdData2
		
		For iRow = iMaxRows To 1 Step -1	' -- 취소시 역순으로..
			
			.Col = C_MINOR_CD : .Row = iRow
			
			If Trim(.Text) = sMinorCd Then	' 부모순번이 같은 행 
				.Col = 0 : sFlag = .Text

				If sFlag = ggoSpread.DeleteFlag Or sFlag = ggoSpread.InsertFlag Then	' 삭제가 아닐경우에만..
					.SetActiveCell C_CTRL_CD, iRow
					
					lDelRows = ggoSpread.EditUndo
					
				End If
			End If
		Next
		
	End With
End Function

'==========================================================================================
'   Event Desc : 2번 그리드 전체 삭제 루틴 : FncDelete 시 
'==========================================================================================
Function DeleteChildGrid2()
	Dim iCol, iRow, iMaxRows, sMinorCd, sMinorCd2, lDelRows, sFlag, iChildSeqNo, i, iSelBlockRow, iSelBlockRow2
	
	With frm1.vspdData2
		iMaxRows = .MaxRows		' -- 자식 갯수 
		iSelBlockRow	= frm1.vspdData.SelBlockRow
		iSelBlockRow2	= frm1.vspdData.SelBlockRow2 
		
		' -- 멀티 부모 삭제 seq_no 기억 
		For i = iSelBlockRow To iSelBlockRow2
			frm1.vspdData.Col = C_MINOR_CD : frm1.vspdData.Row = i : sMinorCd = sMinorCd & Trim(frm1.vspdData.Text) & "|"
		Next

		ggoSpread.Source = frm1.vspdData2
		For iRow = 1 To iMaxRows
			
			.Col = C_MINOR_CD : .Row = iRow : sMinorCd2 = Trim(.Text) & "|"
			
			If Instr(1, sMinorCd, sMinorCd2) > 0 Then	' 부모순번이 같은 행 
				.Col = 0 : sFlag = .Text	' -- 현재상태 체크 
				
				If sFlag <> ggoSpread.DeleteFlag  Then	' 삭제가 아닐경우에만..
					.SetActiveCell C_CTRL_CD, iRow
						
					lDelRows = ggoSpread.DeleteRow
						
					' -- 계산필드가 있으면, 이벤트 발생시켜야함 
					'Call vspdData2_Change(C_W9, .ActiveRow)
				End If
			End If
		Next

	End With
End Function

' -- 직과항목 코드 존재 체크 
Function ExistsKey(Byval pAcctCd, Byval pMinorcd, Byval pRow)
	Dim i, iMaxRows, sAcctCd, sMinorCd
	
	With frm1.vspdData
		iMaxRows = .MaxRows

		.Redraw = False
		
		For i = 1 To iMaxRows
			.Row = i
			
			.Col = C_MINOR_CD	: sMinorCd = Trim(.Text)
			.Col = C_ACCT_CD	: sAcctCd = Trim(.Text)
			
			If ( sMinorCd = pMinorcd And pAcctCd = sAcctCd ) And i <> pRow Then
				ExistsKey = True
				Exit Function
			End If
		Next
		.Redraw = True
		
		ExistsKey = False
	End With
End Function


'========================================================================================================
'                        5.5 Tag Event
'========================================================================================================

'==========================================================================================
'   Event Name : vspdData_Click
'   Event Desc : This event is spread sheet data changed
'==========================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
	lgCurrGrid = GRID_1
	
	If lgIntFlgMode <> Parent.OPMD_UMODE Then
		Call SetPopupMenuItemInf("1001111111")
	Else 
		Call SetPopupMenuItemInf("1101111111")
	End If	

    gMouseClickStatus = "SPC"	'Split 상태코드 
    Set gActiveSpdSheet = frm1.vspdData

    
    If frm1.vspdData.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
   	    
    If Row <= 0 Then
        ggoSpread.Source = frm1.vspdData
        If lgSortKey = 1 Then
            ggoSpread.SSSort
            lgSortKey = 2
        Else
            ggoSpread.SSSort ,lgSortKey
            lgSortKey = 1
        End If    
    End If

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub


Sub vspdData_ScriptLeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
	'MsgBox "hello"
    If Row <> NewRow And NewRow > 0 Then

		Dim iLastRow	' -- 보이는 마지막 행 
		Dim sMinorCd
	
		With frm1.vspdData 
		
	
		End With
    
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
'   Event Name : vspdData_DblClick
'   Event Desc : 
'========================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)				
    
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
    	
End Sub


'==========================================================================================
'   Event Desc : Spread Split 상태코드 
'==========================================================================================
Sub vspdData_MouseDown(Button, Shift, X, Y)
	If Button = 2 And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If
	
	lgCurrGrid = GRID_1
End Sub

'========================================================================================================
'   Event Name : vspdData_Change
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_Change(ByVal Col, ByVal Row)
    Frm1.vspdData.Row = Row
    Frm1.vspdData.Col = Col
	
'	Call CheckMinNumSpread(frm1.vspdData, Col, Row)

	ggoSpread.Source = frm1.vspdData
	ggoSpread.UpdateRow Row
	lgBlnFlgChgValue = True

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	Dim sSelectSQL, sFromSQL, sWhereSQL, sVal, sCd, sCdNm, sTmp
	
	With frm1.vspdData
		.Row = Row	: .Col = Col : sVal = UCase(Trim(.Text))
		
		Select Case Col
		
			Case C_MINOR_CD	' -- 직과항목 
			
				.Col = C_ACCT_CD : sTmp = UCase(Trim(.Text))
				
				If ExistsKey(sTmp, sVal, Row) Then 
					Call DisplayMsgBox("236309", "x",sVal,sTmp)
					.Col = Col : .Row = Row : .Text = ""
					Call SetFocusToDocument("M")
					.Focus
					Exit sub
				End If
			
				sSelectSQL	= "a.MINOR_CD, b.MINOR_NM"
				sFromSQL	= "c_dir_dstb_fctr_s a(nolock) inner join b_minor b(nolock) on a.minor_cd=b.minor_cd and b.major_cd = " & FilterVar("C4010", "''", "S")
				sWhereSQL	= "a.MINOR_CD = " & FilterVar(sVal, "''", "S")

			Case C_ACCT_CD	' -- 계정 
			
				.Col = C_MINOR_CD : sTmp = UCase(Trim(.Text))
				
				If ExistsKey(sVal, sTmp, Row) Then 
					Call DisplayMsgBox("236309", "x",sVal,sTmp)
					.Col = Col : .Row = Row : .Text = ""
					Call SetFocusToDocument("M")
					.Focus
					Exit sub
				End If
			
				sSelectSQL	= "ACCT_CD, ACCT_NM"
				sFromSQL	= "A_ACCT" 
				sWhereSQL	= "TEMP_FG_3 IN (" & FilterVar("M2", "''", "S") & "," & FilterVar("M3", "''", "S") & "," & FilterVar("M4", "''", "S") & ")"
				sWhereSQL	= sWhereSQL & " AND ACCT_CD = " & FilterVar(sVal, "''", "S")
				
		End Select
	
		If sWhereSQL <> "" Then
			' -- DB 콜 
			If CommonQueryRs(sSelectSQL, sFromSQL , sWhereSQL, lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) Then
				sCd		= Replace(lgF0, Chr(11), "")
				sCdNm	= Replace(lgF1, Chr(11), "")
				
				.Row = Row
				' -- 존재시 코드명을 출력한다.
				Select Case Col
					Case C_MINOR_CD, C_ACCT_CD
						.Col = Col + 2	
						.Text = sCdNm

						.Col = Col
						.Text = sCd

				End Select
			Else
				' -- 비존재시 메시지 처리 
				If sVal <> "" Then
					Call DisplayMsgBox("970000", "x",sVal,"x")
					Call SetFocusToDocument("M")
					.Focus
				End If
				
				' -- 명 들을 지운다 
				Select Case Col
					Case C_MINOR_CD
						'.Col = Col		: .Text = ""
						.Col = Col + 2	: .Text = ""
				End Select
				
			End If
		
		End If
		
	End With
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
	
End Sub

' -- 그리드1 팝업 버튼 클릭 
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	Dim sCode, sCode2
	
	With frm1 
	
	    ggoSpread.Source = frm1.vspdData
   
		If Row = 0 Then Exit Sub
		
		Select Case Col
			Case C_MINOR_CD_POP, C_ACCT_CD_POP
				.vspdData.Col = Col - 1
				.vspdData.Row = Row
				
				sCode = UCase(Trim(.vspdData.Text))
				
				Call OpenPopup(Col, sCode, sCode2)
		End Select
        Call SetActiveCell(.vspdData,Col-1,.vspdData.ActiveRow ,"M","X","X")   	
    End With
    
End Sub

'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    'If Col <= C_SNm Or NewCol <= C_SNm Then
     '   Cancel = True
      '  Exit Sub
   ' End If
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
End Sub


Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
	IF CheckRunningBizProcess = True Then
		Exit Sub
	END IF
	
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
End Sub

'========================================================================================================
' Name : FncQuery
' Desc : This function is called from MainQuery in Common.vbs
'========================================================================================================
Function FncQuery() 
    Dim IntRetCD , blnChange1, blnChange2
    
    FncQuery = False
    
    Err.Clear

    If Not chkField(Document, "1") Then
       Exit Function
    End If
    
	ggoSpread.Source = frm1.vspdData
    blnChange1 = ggoSpread.SSCheckChange

    If blnChange1 = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO,"x","x")
    	If IntRetCD = vbNo Then
	      	Exit Function
    	End If
    End If
    
    If ChkKeyField=False then Exit Function 
    
    Call ggoOper.ClearField(Document, "2")
    ggoSpread.Source = frm1.vspdData
	ggoSpread.ClearSpreadData

'	Call InitSpreadSheet
    Call InitVariables 	
    'Call InitComboBox

    Call SetToolbar("1110110100101111")

    IF DbQuery = False Then
		Exit Function
	END IF
       
    FncQuery = True		
    
End Function

'========================================================================================================
' Name : FncNew
' Desc : This function is called from MainNew in Common.vbs
'========================================================================================================
Function FncNew() 
    Dim IntRetCD 
    
    FncNew = False 
    
    Err.Clear     

    If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
    	IntRetCD = DisplayMsgBox("900015", Parent.VB_YES_NO,"x","x")   
			If IntRetCD = vbNo Then
				Exit Function
			End If
    End If
    
	Call SetToolbar("111011010011111")
	
    Call ggoOper.ClearField(Document, "1") 
    Call ggoOper.ClearField(Document, "2") 
    ggoSpread.Source = frm1.vspdData
	ggoSpread.ClearSpreadData

    Call ggoOper.LockField(Document, "N") 
    Call InitVariables 
    Call SetDefaultVal
    
    FncNew = True 

End Function

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

'========================================================================================================
' Name : FncSave
' Desc : This function is called from MainSave in Common.vbs
'========================================================================================================
Function FncSave() 
    Dim IntRetCD , blnChange1, blnChange2, iRow, sMinorCd
    
    FncSave = False
    
    Err.Clear     

    ggoSpread.Source = frm1.vspddata
    blnChange1 = ggoSpread.SSCheckChange
    
    If blnChange1 = False Then	' -- 둘다 미수정 
        IntRetCD = DisplayMsgBox("900001","x","x","x")  
        Exit Function
    End If
    
    ggoSpread.Source = frm1.vspddata
    If Not ggoSpread.SSDefaultCheck Then      
       Exit Function
    End If

    IF DbSave = False Then
		Exit function
	END IF

    FncSave = True      
    
End Function


'========================================================================================================
' Name : FncCopy
' Desc : This function is called from MainCopy in Common.vbs
' Keep : Make sure to clear primary key area
'========================================================================================================
Function FncCopy() 
	Dim iSeqNo, iSubSeqNo, iOldCol
	
    if frm1.vspdData.maxrows = 0 then exit function 

	With frm1.vspdData
	
		.ReDraw = False
			
		iOldCol = .ActiveCol
		ggoSpread.Source = frm1.vspdData	
		ggoSpread.CopyRow

		.Col = C_ACCT_CD : .Text = ""
		.Col = C_ACCT_NM : .Text = ""

		SetSpreadColor  .ActiveRow , .ActiveRow

		.ReDraw = True

		.SetActiveCell iOldCol, .vspdData.ActiveRow			
		.vspdData.focus				
    End With
End Function


Function FncCancel() 
    Dim lDelRows

	With frm1.vspdData 
		.focus
		ggoSpread.Source = frm1.vspdData 

		ggoSpread.Source = frm1.vspdData 
		lDelRows = ggoSpread.EditUndo
					
	End With
	
	lgBlnFlgChgValue = True
End Function


'========================================================================================================
' Name : FncInsertRow
' Desc : This function is called by MainInsertRow in Common.vbs
'========================================================================================================
Function FncInsertRow(ByVal pvRowCnt)
    Dim IntRetCD, sMinorCd, iOldCol
    Dim imRow
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

	With frm1

	'------ Developer Coding part (Start ) --------------------------------------------------------------
	
	iOldCol = .vspdData.ActiveCol
	.vspdData.focus
			
	ggoSpread.Source = .vspdData
	ggoSpread.InsertRow  .vspdData.ActiveRow, imRow
			
	SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1

	frm1.vspdData.SetActiveCell iOldCol, .vspdData.ActiveRow
	.vspdData.focus
	'------ Developer Coding part (End )   -------------------------------------------------------------- 

	End With
    
    Set gActiveElement = document.ActiveElement  
    If Err.number = 0 Then
       FncInsertRow = True                                                          '☜: Processing is OK
    End If   
End Function


Function FncDeleteRow() 
    Dim lDelRows

	With frm1.vspdData 
		.focus
		ggoSpread.Source = frm1.vspdData 

		lDelRows = ggoSpread.DeleteRow
			
		lgCurrGrid = 2 : Call DeleteChildGrid2()
					
	End With
	
	lgBlnFlgChgValue = True
End Function

Function FncPrint()
    Call parent.FncPrint() 
End Function

Function FncPrev() 
End Function

Function FncNext() 
End Function

Function FncExcel() 
    Call parent.FncExport(Parent.C_MULTI)
End Function

Function FncFind() 
    Call parent.FncFind(Parent.C_MULTI, False) 
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
    Call InitComboBox
	Call ggoSpread.ReOrderingSpreadData()
	Call InitData()
End Sub

'========================================================================================================
' Name : FncExit
' Desc : This function is called by MainExit in Common.vbs
'========================================================================================================
Function FncExit()
Dim IntRetCD
	
	FncExit = False
	
    ggoSpread.Source = frm1.vspdData	
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO,"x","x")
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    FncExit = True
End Function

'========================================================================================================
'========================================================================================================
'                        5.3 Common Method-3
'========================================================================================================
'========================================================================================================

'========================================================================================================
' Name : DbQuery
' Desc : This sub is called by FncQuery
'========================================================================================================
Function DbQuery() 
	Dim strVal

    DbQuery = False
    IF LayerShowHide(1) = False Then
		Exit Function
	END IF

    Err.Clear	
    
    With frm1
		'If lgIntFlgMode = Parent.OPMD_UMODE Then
			strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
			'strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey				
			strVal = strVal & "&txtACCT_CD=" & Trim(.txtACCT_CD.value)	
			'strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
		'Else
		'	strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001	
		'	strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
			'strVal = strVal & "&txtDstbFctr=" & Trim(.txtDstbFctr.value)
			'strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
		'End If
		
		Call RunMyBizASP(MyBizASP, strVal)
   
    End With
    
    DbQuery = True

End Function

'========================================================================================================
' Name : DbQueryOk
' Desc : Called by MB Area when query operation is successful
'========================================================================================================
Function DbQueryOk()	
	
	If lgCopyVersion = "" Then
		lgIntFlgMode = Parent.OPMD_UMODE
    
		'Call ggoOper.LockField(Document, "Q")

		Call SetToolbar("111011110011111")
	Else
		' -- 타버전카피 
		Call ggoOper.ClearField(Document, "1") 
		'Call ggoOper.LockField(Document, "N")
		
		lgIntFlgMode = Parent.OPMD_CMODE
		
		Call SetToolbar("111011010011111")
		
		' -- 그리드를 모두 입력으로 바꾼다.
		Call ChangeNewFlag(frm1.vspdData)
		Call ChangeNewFlag(frm1.vspdData2)
	End If

   	Call SetSpreadLock

	Frm1.vspdData.Focus
	
    Set gActiveElement = document.ActiveElement   
   	
End Function

'========================================================================================================
' Name : InitData()
' Desc : Reset ComboBox
'========================================================================================================
Sub InitData()
End Sub

Function CheckNullIs0(Byval pVal)
	If pVal = "" Then
		CheckNullIs0 = "0"
	Else
		CheckNullIs0 = pVal
	End If
End Function

Function CheckNullIsX(Byval pVal)
	If pVal = "" Then
		CheckNullIs0 = "*"
	Else
		CheckNullIs0 = pVal
	End If
End Function

'========================================================================================================
' Name : DbSave
' Desc : This sub is called by FncSave
'========================================================================================================
Function DbSave() 
    Dim lRow        
    Dim lGrpCnt     
	Dim strVal, strDel
    Dim iColSep 
    Dim iRowSep  
    Dim sSQLI1, sSQLI2, sSQLU1, sSQLU2, sSQLD1, sSQLD2, sMinorCd, tmpA, tmpB
	
    DbSave = False 
    IF LayerShowHide(1) = False Then
		Exit Function
	END IF
	
	iColSep = Parent.gColSep
	iRowSep = Parent.gRowSep	

	lGrpCnt = 1
	strVal = ""
	strDel = ""
	
	If lgIntFlgMode = Parent.OPMD_CMODE Then
		sMinorCd = UCase(Trim(frm1.txtACCT_CD.value))
	Else
		sMinorCd = UCase(Trim(frm1.hMINOR_CD.value))
		frm1.txtACCT_CD.value = sMinorCd
	End If

	With frm1.vspdData
	
		ggoSpread.Source = frm1.vspdData	
		
		For lRow = 1 To .MaxRows
    
			.Row = lRow	: .Col = 0
        
			Select Case .Text

	            Case ggoSpread.InsertFlag	
					strVal = "C" & iColSep 
					
					.Col = .MaxCols				: strVal = strVal & Trim(lRow) & iColSep
					.Col = C_ACCT_CD			: strVal = strVal & Trim(.Value) & iColSep
					.Col = C_MINOR_CD			: strVal = strVal & Trim(.Value) & iColSep &  Parent.gRowSep

					sSQLI1 = sSQLI1 + strVal
					lGrpCnt = lGrpCnt + 1

	            Case ggoSpread.UpdateFlag		
	            
					strVal = "U" & iColSep 
					
					.Col = .MaxCols				: strVal = strVal & Trim(lRow) & iColSep
					.Col = C_ACCT_CD			: strVal = strVal & Trim(.Value) & iColSep
					.Col = C_MINOR_CD			: strVal = strVal & Trim(.Value) & iColSep &  Parent.gRowSep

					sSQLU1 = sSQLU1 + strVal
					lGrpCnt = lGrpCnt + 1

	            Case ggoSpread.DeleteFlag		

					strVal = "D" & iColSep 
					
					.Col = .MaxCols				: strVal = strVal & Trim(lRow) & iColSep
					.Col = C_ACCT_CD			: strVal = strVal & Trim(.Value) & iColSep
					.Col = C_MINOR_CD			: strVal = strVal & Trim(.Value) & iColSep &  Parent.gRowSep
					
					sSQLD1 = sSQLD1 + strVal
					lGrpCnt = lGrpCnt + 1
                
	        End Select

		Next

	End With

	strDel = "" : strVal = ""
		
	frm1.txtMode.value = Parent.UID_M0002
	'frm1.txtMaxRows.value = lGrpCnt-1

	frm1.txtSpreadI1.value = sSQLI1
	frm1.txtSpreadU1.value = sSQLU1
	frm1.txtSpreadD1.value = sSQLD1
	
	Call ExecMyBizASP(frm1, BIZ_PGM_ID)
	
    DbSave = True    
    
End Function

'========================================================================================================
' Name : DbSaveOk
' Desc : Called by MB Area when save operation is successful
'========================================================================================================
Function DbSaveOk()	
   
	Call InitVariables
	frm1.vspdData.MaxRows = 0

	Call MainQuery()
		
End Function

'========================================================================================
Function DbDelete() 
    Err.Clear
    DbDelete = False

    Dim strVal

    strVal = BIZ_PGM_ID & "?txtMode="	& parent.UID_M0003							'☜: 비지니스 처리 ASP의 상태 
    strVal = strVal     & "&txtACCT_CD=" & frm1.txtACCT_CD.value					    '☜: Query Key        
	
	Call RunMyBizASP(MyBizASP, strVal)

    DbDelete = True
End Function


'========================================================================================
Function DbDeleteOk()
	Call InitVariables
	Call FncNew()
End Function
'=================================================================================
'	Name : ChkKeyField()
'	Description : check the valid data
'=========================================================================================================
Function ChkKeyField()
	Dim strDataCd, strDataNm
    Dim strWhere , strFrom 
    Dim lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6
    
    Err.Clear                                       

	ChkKeyField = true		
'check txtACCT_NM
	If Trim(frm1.txtACCT_CD.value) <> "" Then
		strWhere = " acct_cd= " & FilterVar(frm1.txtACCT_CD.value, "''", "S") & "  "
		strWhere = strWhere & "	and TEMP_FG_3 IN (" & FilterVar("M2", "''", "S") & "," &  FilterVar("M3", "''", "S") & "," & FilterVar("M4", "''", "S") & ")"

		Call CommonQueryRs(" acct_nm ","	 a_acct ", strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
		IF Len(lgF0) < 1 Then 
			Call DisplayMsgBox("970000","X",frm1.txtACCT_CD.alt,"X")
			frm1.txtACCT_CD.focus 
			frm1.txtACCT_NM.value = ""
			ChkKeyField = False
			Exit Function
		End If
	
		strDataNm = split(lgF0,chr(11))
		frm1.txtACCT_NM.value = strDataNm(0)
	Else
		frm1.txtACCT_NM.value=""
	End If

End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->	
</HEAD>

<BODY TABINDEX="-1" SCROLL="no" oncontextmenu="javascript:return false">
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
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white><%=Request("strASPMnuMnuNm")%></font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* align=right>&nbsp;</TD>
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
									<TD CLASS="TD5">계정</TD>
									<TD CLASS="TD656"><INPUT ClASS="clstxt" TYPE=TEXT NAME="txtACCT_CD" SIZE=18 MAXLENGTH=18 tag="15XXXU" ALT="계정"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDstbFctr" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenAcct(0)">
									<input NAME="txtACCT_NM" TYPE="TEXT"  tag="14XXX" size="30">
									</TD>
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
						<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD HEIGHT="50%" NOWRAP>
								<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> ID=vspdData NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD"> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
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
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>>
			<IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=No noresize framespacing=0 TABINDEX= "-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpreadI1" tag="24" TABINDEX= "-1"></TEXTAREA>
<TEXTAREA CLASS="hidden" NAME="txtSpreadU1" tag="24" TABINDEX= "-1"></TEXTAREA>
<TEXTAREA CLASS="hidden" NAME="txtSpreadD1" tag="24" TABINDEX= "-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="hMINOR_CD" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="hCopyVerCd" tag="24" TABINDEX= "-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

