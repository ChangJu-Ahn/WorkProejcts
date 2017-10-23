<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% session.CodePage=949 %>

<!-- '======================================================================================================
'*  1. Module Name          : Cost
'*  2. Function Name        : 직과항목등록 
'*  3. Program ID           : c4002ma1.asp
'*  4. Program Name         : 직과항목등록 
'*  5. Program Desc         : 직과항목등록 
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

Const BIZ_PGM_ID = "c4008mb1.asp"                               'Biz Logic ASP

' -- 그리드1의 컬럼 정의 
Dim C_MINOR_CD			' -- 헤더 키 
Dim C_MINOR_CD_POP		' -- 배부순서 
Dim C_MINOR_NM		' -- C/C 래벨 
Dim C_SP_NM		' -- C/C 래벨 

' -- 그리드2의 보이는 컬럼 정의 
Dim C_CTRL_CD	' -- C/C 래벨 
Dim C_CTRL_CD_POP
Dim C_CTRL_NM			' -- RECV C/C


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
	 C_MINOR_CD					= 1
	 C_MINOR_CD_POP				= 2
	 C_MINOR_NM					= 3
	 C_SP_NM					= 4

	' -- 그리드2의 보이는 컬럼 정의 
	 C_CTRL_CD					= 2		' -- 부모 C_MINOR_CD 포함되어 2부터임		
	 C_CTRL_CD_POP				= 3		
	 C_CTRL_NM					= 4

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
	
	.MaxCols = C_SP_NM + 1

	.Col = .MaxCols	
    .ColHidden = True
    
    ggoSpread.Source = frm1.vspdData
    ggoSpread.Spreadinit "V20030825",,"" ',,parent.gAllowDragDropSpread 

	Call ggoSpread.ClearSpreadData()    '☜: Clear spreadsheet data 

	.ReDraw = false

	Call GetSpreadColumnPos("A")
	
    ggoSpread.SSSetEdit		C_MINOR_CD		,"직과항목코드",10,,, 3,2
    ggoSpread.SSSetButton	C_MINOR_CD_POP
    ggoSpread.SSSetEdit		C_MINOR_NM		,"직과항목명"	,20
    ggoSpread.SSSetEdit		C_SP_NM	,		"Stored Procedure 명",40


	.ReDraw = true
	
    Call SetSpreadLock 
   
    End With
    
    
    ' -- 그리드 2 정의 
    With frm1.vspdData2
	
	.MaxCols = C_CTRL_NM + 1

	.Col = .MaxCols	
    .ColHidden = True
    
    ggoSpread.Source = frm1.vspdData2
    ggoSpread.Spreadinit "V20030825",, "" ',,parent.gAllowDragDropSpread 

	Call ggoSpread.ClearSpreadData()    '☜: Clear spreadsheet data 

	.ReDraw = false

	Call GetSpreadColumnPos("B")
	
	ggoSpread.SSSetEdit		C_MINOR_CD					,"직과항목코드",10	' -- 히든 

	ggoSpread.SSSetEdit		C_CTRL_CD					,"관리항목"	, 10,,,10,2
	ggoSpread.SSSetButton	C_CTRL_CD_POP    
    ggoSpread.SSSetEdit		C_CTRL_NM					,"관리항목명" ,30

	Call ggoSpread.SSSetColHidden(C_MINOR_CD,C_MINOR_CD,True)
	
	.ReDraw = true
	
    Call SetSpreadLock2 
    
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
    ggoSpread.SpreadLock		C_MINOR_CD			,-1	,C_MINOR_CD
    ggoSpread.SpreadLock			C_MINOR_CD_POP		,-1	,C_MINOR_CD_POP
	ggoSpread.SSSetRequired		C_SP_NM				,-1	,C_SP_NM
	ggoSpread.SpreadLock		C_MINOR_NM			,-1	,C_MINOR_NM
    .ReDraw = True

    End With
End Sub

Sub SetSpreadLock2()
    With frm1.vspdData2
    
    ggoSpread.Source = frm1.vspdData2
    
    .ReDraw = False
    ggoSpread.SpreadLock		C_MINOR_CD			,-1	,C_MINOR_CD
    ggoSpread.SpreadLock		C_CTRL_CD			,-1	,C_CTRL_CD
    ggoSpread.SpreadLock			C_CTRL_CD_POP		,-1	,C_CTRL_CD_POP
	ggoSpread.SpreadLock		C_CTRL_NM			,-1	,C_CTRL_NM
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
	ggoSpread.SSSetRequired		C_MINOR_CD			,pvStartRow		,pvEndRow    
	ggoSpread.SSSetRequired		C_SP_NM				,pvStartRow		,pvEndRow
	ggoSpread.SSSetProtected	C_MINOR_NM			,pvStartRow		,pvEndRow    
	
    .vspdData.ReDraw = True
    
    End With
End Sub

Sub SetSpreadColor2(ByVal pvStartRow,ByVal pvEndRow)
    With frm1
    
    ggoSpread.Source = frm1.vspdData2
    .vspdData2.ReDraw = False
									      'Col          Row				Row2
	ggoSpread.SSSetProtected	C_MINOR_CD			,pvStartRow		,pvEndRow    
	ggoSpread.SSSetRequired		C_CTRL_CD			,pvStartRow		,pvEndRow    
	ggoSpread.SSSetProtected	C_CTRL_NM			,pvStartRow		,pvEndRow
	
    .vspdData2.ReDraw = True
    
    End With
End Sub

'======================================================================================================
' Name : SubSetErrPos
' Desc : This method set focus to position of error
'      : This method is called in MB area
'======================================================================================================
Sub SubSetErrPos(iPosArr)
    Dim iDx, oGrid, j, iSeqNo, iSubSeqNo
    Dim iRow
    If iPosArr = "" Then Exit Sub
    iPosArr = Split(iPosArr,Parent.gColSep)		' 리턴문자열: 그리드n/gColSep/상태플래그/gColSep/에러행번호(C:SEQ_NO번호)/gColSep/SUB_SEQ_NO
    If IsNumeric(iPosArr(0)) Then
       iDx = CDbl(iPosArr(2))	' 행번호/SEQ_NO번호 
       
		If iPosArr(0) = "1" Then	' 그리드n 지정 
			Set oGrid = frm1.vspdData
		Else
			Set oGrid = frm1.vspdData2
		End If
       
		With oGrid
		
		For iRow = 1 To  .MaxRows 
		    .Col = 0
		    .Row = iRow
		    
			If iPosArr(0) = "1" Then	' -- 그리드1일 경우 
				.Col = C_MINOR_CD	: iSeqNo = Trim(.value)
				If iSeqNo = iDx Then	' -- 에러행번호와 SEQ_NO가 같다면 
					Call ClickGrid1(iSeqNo)
					Exit Sub
				End If
				' -- 에러행번호와 SEQ_NO가 다르므로 다음 For문 실행 
			Else
				' -- 그리드2 일 경우 
				.Col = C_MINOR_CD		: iSeqNo	= Trim(.value)
				.Col = C_CTRL_CD		: iSubSeqNo = Trim(.value)
				If iSeqNo = iDx And iSubSeqNo = Trim(iPosArr(3)) Then	' -- 에러행번호와 SEQ_NO가 같다면 
					.Col = C_CTRL_NM	: .Action  = 0	
					lgErrRow = iRow		' -- 에러난 행지정 
					Call ClickGrid1(iSeqNo)
					Exit Sub
				End If
			End If
					
		Next
        
        End With 
    End If   
End Sub

'======================================================================================================
Sub SubSetErrPos2(Byval iPosArr)
    Dim iDx, oGrid, j, iSeqNo, iSubSeqNo
    Dim iRow
    If iPosArr = "" Then Exit Sub
    iPosArr = Split(iPosArr,Parent.gColSep)		' 리턴문자열: MINOR_CD/CTRL_CD/SP_NM/ERR_CODE
    
    If IsNumeric(iPosArr(4)) Then
    
		Select Case CInt(iPosArr(4))
			Case 1	' -- MINOR_CD is null
				Call DisplayMsgBox("970000", "x",iPosArr(1),"x")
			Case 2	' -- CTRL_CD is null
				Call DisplayMsgBox("970000", "x",iPosArr(2),"x")
			Case 4	' -- SP_NM is null
				Call DisplayMsgBox("970000", "x",iPosArr(3),"x")
			Case 5	' -- MINOR_CD+SP_NM is null
				Call DisplayMsgBox("970000", "x",iPosArr(1),"x")
			Case 6	' -- CTRL_CD+SP_NM is null
				Call DisplayMsgBox("970000", "x",iPosArr(3),"x")
			Case 7	' -- ALL is null
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
			 C_MINOR_CD					= iCurColumnPos(1)	
			 C_MINOR_CD_POP				= iCurColumnPos(2)
			 C_MINOR_NM					= iCurColumnPos(3)
			 C_SP_NM					= iCurColumnPos(4)

		Case "B"

            ggoSpread.Source = frm1.vspdData2
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
		
			' -- 그리드2의 보이는 컬럼 정의 
			C_CTRL_CD					= iCurColumnPos(2)		' -- 부모키(MINOR_CD) 이후부터 
			C_CTRL_CD_POP				= iCurColumnPos(3)		
			C_CTRL_NM					= iCurColumnPos(4)	
		
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
   
'   ggoSpread.SetCombo "10" & vbtab & "20" & vbtab & "30" & vbtab & "50" , C_ItemAcct
'    ggoSpread.SetCombo "제품" & vbtab & "반제품" & vbtab & "원자재"& vbtab & "상품", C_ItemAcctNm
'    ggoSpread.SetCombo "M" & vbtab & "O" & vbtab & "P", C_ProcurType
'    ggoSpread.SetCombo "사내가공품" & vbtab & "외주가공품" & vbtab & "구매품", C_ProcurTypeNm
   
    'Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR ", "MAJOR_CD=" & FilterVar("C2201", "''", "S") & " "  ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

    'ggoSpread.SetCombo Replace(lgF0,Chr(11),vbTab), C_SENDER_COST_NM			'COLM_DATA_TYPE
    'ggoSpread.SetCombo Replace(lgF1,Chr(11),vbTab), C_GP_LEVEL
     
End Sub

' -- 직과항목 팝업시.
Function OpenMinor(Byval iWhere)
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

	arrParam(0) = "직과항목 팝업"
	arrParam(1) = "B_MINOR"	
	
	If iWhere = 0 Then	' -- 그냥 팝업 
		arrParam(2) = Trim(frm1.txtMINOR_CD.Value)
	End If
	
	arrParam(3) = ""
	arrParam(4) = "MAJOR_CD = " & FilterVar("C4010", "''", "S")
	
	If iWhere = 1 And frm1.txtMINOR_CD.value <> "" Then		' -- 카피 팝업시 추가조건 
		arrParam(4) = arrParam(4) & " AND MINOR_CD <> " & FilterVar(frm1.txtMINOR_CD.value, "''", "S")
	End If

	arrParam(5) = "직과항목"
	
    arrField(0) = "MINOR_CD"
    arrField(1) = "MINOR_NM"
    
    arrHeader(0) = "직과항목"
    arrHeader(1) = "직과항목명"
    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtMINOR_CD.focus
		Exit Function
	Else
		Call SetMinor(arrRet, iWhere)
	End If
		
End Function

' -- 직과항목 팝업후 
Function SetMinor(byval arrRet, Byval iWhere)
	Select Case iWhere
		Case 0
			frm1.txtMINOR_CD.focus
			frm1.txtMINOR_CD.Value    = arrRet(0)
			frm1.txtMINOR_NM.Value    = arrRet(1)				

		Case 1
			IF LayerShowHide(1) = False Then
				Exit Function
			END IF

			Dim strVal
	
			With frm1
				strVal = BIZ_COPY_PGM_ID & "?txtMode=" & Parent.UID_M0001
				strVal = strVal & "&txtMINOR_CD=" & Trim(.txtMINOR_CD.value)	
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
		Case C_MINOR_CD_POP
			arrParam(0) = "직과항목 팝업"
			arrParam(1) = "B_MINOR"	
			arrParam(2) = strCode
			arrParam(3) = ""		
			arrParam(4) = "MAJOR_CD=" & FilterVar("C4010", "''", "S")
			arrParam(5) = "직과항목코드" 

			arrField(0) = "MINOR_CD"	
			arrField(1) = "MINOR_NM"
    
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
	
		
		
	Select Case iWhere
		
		Case C_MINOR_CD_POP
			With frm1.vspdData
				.Col = C_MINOR_CD	: .Text = arrRet(0)
				.Col = C_MINOR_NM	: .Text = arrRet(1)

				Call vspdData_Change(C_MINOR_CD, .ActiveRow)
			End With
			
	End Select
		
	lgBlnFlgChgValue = True
	
End Function

' -- 그리드1에서 팝업 클릭시 
Function OpenPopUp2(Byval iWhere, Byval strCode, Byval strCode1)
	Dim arrRet, sTmp, iWidth
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	With frm1
	
	iWidth = 500	' -- 팝업Width
	
	Select Case iWhere
			
		Case C_CTRL_CD_POP
			arrParam(0) = "관리항목 팝업"
			arrParam(1) = "A_CTRL_ITEM A(NOLOCK) INNER JOIN A_ACCT_CTRL_ASSN B(NOLOCK) ON A.CTRL_CD=B.CTRL_CD "	
			arrParam(2) = strCode
			arrParam(3) = ""	
			arrParam(4) = ""
			arrParam(5) = "관리항목" 

			'arrField(0) = "ED15" & Parent.gColSep & "CODE"
			arrField(0) = "A.CTRL_CD"	
			arrField(1) = "A.CTRL_NM"
    
			arrHeader(0) = "관리항목"	
			arrHeader(1) = "관리항목명"	

	End Select
    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=" & CStr(iWidth) & "px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetPopUp2(arrRet, iWhere)
	End If	

	End With
End Function

Function SetPopUp2(Byval arrRet, Byval iWhere)
	Dim sTmp
	
		
		
	Select Case iWhere
		
		Case C_CTRL_CD_POP
			
			With frm1.vspdData2
				.Col = C_CTRL_CD	: .Text = arrRet(0)
				.Col = C_CTRL_NM	: .Text = arrRet(1)

				ggoSpread.Source = frm1.vspdData2
				ggoSpread.UpdateRow frm1.vspdData2.ActiveRow
			
			End With
	End Select
		
	lgBlnFlgChgValue = True
	

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
    frm1.txtMINOR_CD.focus
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
Function ExistsMinorCd(Byval pMinorcd, Byval pRow)
	Dim i, iMaxRows
	
	With frm1.vspdData
		iMaxRows = .MaxRows

		.Redraw = False
		.Col = C_MINOR_CD

		For i = 1 To iMaxRows
			.Row = i
			If Trim(.Text) = pMinorcd And i <> pRow Then
				ExistsMinorCd = True
				Exit Function
			End If
		Next
		.Redraw = True
		
		ExistsMinorCd = False
	End With
End Function

' -- 직과항목 코드 존재 체크 
Function ExistsCtrlCd(Byval pMinorcd, Byval pCtrlCd, Byval pRow)
	Dim i, iMaxRows, sMinorCd, sCtrlCd
	
	With frm1.vspdData2
		iMaxRows = .MaxRows

		.Redraw = False
		

		For i = 1 To iMaxRows
			.Row = i
			
			.Col = C_MINOR_CD : sMinorCd = Trim(.Text)
			.Col = C_CTRL_CD : sCtrlCd = Trim(.Text)
			
			If pMinorcd = sMinorCd And pCtrlCd = sCtrlCd And i <> pRow Then
				ExistsCtrlCd = True
				Exit Function
			End If
		Next
		.Redraw = True
		
		ExistsCtrlCd = False
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

Sub vspdData2_Click(ByVal Col, ByVal Row)
	lgCurrGrid = GRID_2
	
	If lgIntFlgMode <> Parent.OPMD_UMODE Then
		Call SetPopupMenuItemInf("1001111111")
	Else 
		Call SetPopupMenuItemInf("1101111111")
	End If	

    gMouseClickStatus = "SPC"	'Split 상태코드 
    Set gActiveSpdSheet = frm1.vspdData2

    
    If frm1.vspdData2.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
   	    
    If Row <= 0 Then
        ggoSpread.Source = frm1.vspdData2
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
		
			sMinorCd = GetGridTxt(frm1.vspdData, C_MINOR_CD, NewRow)
			
			' -- 그리드2를 그리드1의 키값에 맞는 행만 보이게 한다.
			iLastRow = ShowRowHidden(sMinorCd)
			
			If lgErrRow <> 0 Then iLastRow = lgErrRow
			frm1.vspdData2.SetActiveCell C_CTRL_CD, iLastRow
			'frm1.vspdData2.Focus
	
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

Sub vspdData2_MouseDown(Button, Shift, X, Y)
	If Button = 2 And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If
	
	lgCurrGrid = GRID_2
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
		
			Case C_MINOR_CD	' -- c/c 래벨 
			
				If ExistsMinorCd(sVal, Row) Then 
					Call DisplayMsgBox("970001", "x",sVal,"x")
					.Col = Col : .Row = Row : .Text = ""
					Call SetFocusToDocument("M")
					.Focus
					Exit sub
				End If
			
				sSelectSQL	= "MINOR_CD, MINOR_NM"
				sFromSQL	= "B_MINOR" 
				sWhereSQL	= "MAJOR_CD=" & FilterVar("C4010", "''", "S") & " AND MINOR_CD = " & FilterVar(sVal, "''", "S")
				
		End Select
	
		If sWhereSQL <> "" Then
			' -- DB 콜 
			If CommonQueryRs(sSelectSQL, sFromSQL , sWhereSQL, lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) Then
				sCd		= Replace(lgF0, Chr(11), "")
				sCdNm	= Replace(lgF1, Chr(11), "")
				
				.Row = Row
				' -- 존재시 코드명을 출력한다.
				Select Case Col
					Case C_MINOR_CD
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
		
		' -- 디테일 히든 그리드에 복사해준다 
		Select Case Col		' -- 수정된 그리드1 컬럼 
			Case C_MINOR_CD_POP, C_MINOR_CD
				Call ChangeGrid2HiddenByGrid1	
		End Select
	End With
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
	
End Sub

Sub vspdData2_Change(ByVal Col, ByVal Row)
    Frm1.vspdData2.Row = Row
    Frm1.vspdData2.Col = Col
	
	'Call CheckMinNumSpread(frm1.vspdData2, Col, Row)

	ggoSpread.Source = frm1.vspdData2
	ggoSpread.UpdateRow Row
	lgBlnFlgChgValue = True

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	Dim sSelectCd, sFromSQL, sWhereSQL, sVal, sCd, sCdNm, sTmp, sMinorCd
	
	sFromSQL = " dbo.ufn_c_getListOfPopup_C4002MA1"
	
	With frm1.vspdData2
		.Row = Row	: .Col = Col : sVal = UCase(Trim(.Value))
		
		Select Case Col
		
			Case C_CTRL_CD	' -- c/c 래벨 
			
				sMinorCd = GetGridTxt(frm1.vspdData2, C_MINOR_CD, Row)
				' -- 키 존재 체크 
				If ExistsCtrlCd(sMinorCd, sVal, Row) Then
					Call DisplayMsgBox("970001", "x",sVal,"x")
					.Col = Col : .Row = Row : .Text = ""
					Call SetFocusToDocument("M")
					.Focus
					Exit sub
				End If
			
			
				sSelectCd	= " TOP 1 A.CTRL_CD, A.CTRL_NM"
				sFromSQL	= " A_CTRL_ITEM A (NOLOCK) INNER JOIN  A_ACCT_CTRL_ASSN B (NOLOCK) ON A.CTRL_CD = B.CTRL_CD "
				sWhereSQL	= " A.CTRL_CD = " & FilterVar(sVal, "''", "S")
				
		End Select

		If sWhereSQL <> "" Then	
			' -- DB 콜 
			If CommonQueryRs(sSelectCd, sFromSQL , sWhereSQL, lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) Then
				sCd		= Replace(lgF0, Chr(11), "")
				sCdNm	= Replace(lgF1, Chr(11), "")
				
				.Row = Row
				' -- 존재시 코드명을 출력한다.
				Select Case Col
					Case C_CTRL_CD
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
					Case C_CTRL_CD
						'.Col = Col		: .Text = ""
						.Col = Col + 2	: .Text = ""
				End Select
				
			End If
		End If		
	End With
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
	
End Sub

' -- 그리드 1에 의해 그리드2가 변경되어야 할곳 
Function ChangeGrid2HiddenByGrid1()
	Dim sCCLvl, sCCCd, sGPCd, sAcctCd, iRow, iMaxRows, sMinorCd
	With frm1.vspdData
		.Row = .ActiveRow	
		.Col = C_MINOR_CD			: sMinorCd	= Trim(.text)
	End With
	
	With frm1.vspdData2
		iMaxRows = .MaxRows
		ggoSpread.Source = frm1.vspdData2
		
		.ReDraw = False
		For iRow = 1 To iMaxRows
			.Row = iRow
			.Col = C_MINOR_CD
			If .RowHidden = False Then	' -- 보이는 행만 
				.Col = C_MINOR_CD	: .Text = sMinorCd

				ggoSpread.UpdateRow iRow
			End If
		Next
		.ReDraw = True
	End With
End Function

' -- 그리드1 팝업 버튼 클릭 
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	Dim sCode, sCode2
	
	With frm1 
	
	    ggoSpread.Source = frm1.vspdData
   
		If Row = 0 Then Exit Sub
		
		Select Case Col
			Case C_MINOR_CD_POP
				.vspdData.Col = Col - 1
				.vspdData.Row = Row
				
				sCode = UCase(Trim(.vspdData.Text))
				
				Call OpenPopup(Col, sCode, sCode2)
		End Select
        Call SetActiveCell(.vspdData,Col-1,.vspdData.ActiveRow ,"M","X","X")   	
    End With
    
End Sub

' -- 그리드2 팝업 버튼 클릭 
Sub vspdData2_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	Dim sCode, sCode2
	
	With frm1 
	
	    ggoSpread.Source = frm1.vspdData2
   
		If Row = 0 Then Exit Sub
		
		Select Case Col
			Case C_CTRL_CD_POP
				.vspdData2.Col = Col - 1
				.vspdData2.Row = Row
				
				sCode = UCase(Trim(.vspdData2.Text))
				
				Call OpenPopup2(Col, sCode, sCode2)
		End Select
        Call SetActiveCell(.vspdData2,Col-1,.vspdData2.ActiveRow ,"M","X","X")   	
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
    
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then
    	If lgStrPrevKey <> "" Then
	      	DbQuery
    	End If

    End if
    
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

    ggoSpread.Source = frm1.vspdData2
    blnChange2 = ggoSpread.SSCheckChange
    
    If blnChange1 = True Or blnChange2 = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO,"x","x")
    	If IntRetCD = vbNo Then
	      	Exit Function
    	End If
    End If
    
    If ChkKeyField=False then Exit function 
    
    
    Call ggoOper.ClearField(Document, "2")
    ggoSpread.Source = frm1.vspdData
	ggoSpread.ClearSpreadData

    ggoSpread.Source = frm1.vspdData2
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

    ggoSpread.Source = frm1.vspdData2
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
    
    ggoSpread.Source = frm1.vspddata2
    blnChange2 = ggoSpread.SSCheckChange
    
    If blnChange1 = False And blnChange2 = False Then	' -- 둘다 미수정 
        IntRetCD = DisplayMsgBox("900001","x","x","x")  
        Exit Function
    End If
    
    ggoSpread.Source = frm1.vspddata
    If Not ggoSpread.SSDefaultCheck Then      
       Exit Function
    End If

    ggoSpread.Source = frm1.vspddata2
    iRow = 0
    If Not ggoSpread.SSDefaultCheck(,iRow) Then      
		If iRow <> 0 Then
			lgErrRow = iRow
			frm1.vspdData2.Row = iRow
			frm1.vspdData2.Col = C_MINOR_CD
			sMinorCd = Trim(frm1.vspdData2.Value)
			Call ClickGrid1(sMinorCd)
		End If
		Exit Function
    End If
    
   
    IF DbSave = False Then
		Exit function
	END IF

    FncSave = True      
    
End Function

' --- 그리드 1 의 C_MINOR_CD의 값이 pSeqNo 이면 클릭해준다 
Function ClickGrid1(Byval pSeqNo)
	Dim iRow, iMaxRows
	
	With frm1.vspdData
		iMaxRows = .MaxRows
		For iRow = 1 To iMaxRows
			.Row = iRow	: .Col = C_MINOR_CD
			If Trim(.Value) = pSeqNo Then
				.Col = C_SP_NM	: .Action = 0
				Call vspdData_Click(frm1.vspdData.ActiveCol, iRow)
				Exit Function
			End If
		Next
	End With
End Function

' --- 그리드 2 의 C_MINOR_CD의 값과 C_CTRL_CD값이 pSeqNo, pSubSeqNo 이면 그리드1의 pSeqNo를 클릭해준다 
Function ClickGrid2(Byval pSeqNo, Byval pSubSeqNo)
	Dim iRow, iMaxRows, iSeqNo, iSubSeqNo
	
	With frm1.vspdData2
		iMaxRows = .MaxRows
		For iRow = 1 To iMaxRows
			.Row = iRow	: .Col = C_MINOR_CD		: iSeqNo	= Trim(.value)
			.Row = iRow	: .Col = C_CTRL_CD		: iSubSeqNo = Trim(.value)
			
			If iSeqNo = pSeqNo And iSubSeqNo = pSubSeqNo Then
				Call ClickGrid1(pSeqNo)
				Exit Function
			End If
		Next
	End With
End Function

'========================================================================================================
' Name : FncCopy
' Desc : This function is called from MainCopy in Common.vbs
' Keep : Make sure to clear primary key area
'========================================================================================================
Function FncCopy() 
	Dim iSeqNo, iSubSeqNo, iOldCol
	
    if frm1.vspdData.maxrows = 0 then exit function 

	With frm1
	
	Select Case lgCurrGrid

		Case GRID_1
			.vspdData.ReDraw = False
			
			iOldCol = .vspdData.ActiveCol
			ggoSpread.Source = .vspdData	
			ggoSpread.CopyRow
			SetSpreadColor  .vspdData.ActiveRow , .vspdData.ActiveRow

			.vspdData.ReDraw = True

			.vspdData.SetActiveCell iOldCol, .vspdData.ActiveRow
			
			.vspdData.Col = C_MINOR_CD : .vspdData.Text = ""
			.vspdData.Col = C_MINOR_NM : .vspdData.Text = ""
			
			Call vspdData_ScriptLeaveCell(iOldCol, .vspdData.ActiveRow-1, iOldCol,.vspdData.ActiveRow, False)
			
			.vspdData.focus
			
		Case GRID_2
			.vspdData.Col = C_MINOR_CD : .vspdData.Row = .vspdData.ActiveRow : iSeqNo = Trim(.vspdData.Value)
			
			.vspdData2.ReDraw = False
			
			ggoSpread.Source = frm1.vspdData2	
			ggoSpread.CopyRow
			SetSpreadColor2 .vspdData2.ActiveRow ,.vspdData2.ActiveRow

			.vspdData2.Col = C_CTRL_CD : .vspdData2.Text = ""
			.vspdData2.Col = C_CTRL_NM : .vspdData2.Text = ""
			
			Call InsertParentMinorCd(iSeqNo, C_MINOR_CD, .vspdData2.ActiveRow, .vspdData2.ActiveRow)
						
			.vspdData2.ReDraw = True
	End Select
	
    End With

	
End Function


Function FncCancel() 
    Dim lDelRows

	Select Case lgCurrGrid 
		CAse  1	
			With frm1.vspdData 
				.focus
				ggoSpread.Source = frm1.vspdData 

				' -- 하위 그리드 부터 취소함 
				lgCurrGrid = 2 : Call CancelChildGrid2()
				
				ggoSpread.Source = frm1.vspdData 
				lDelRows = ggoSpread.EditUndo

				' -- 계산행이 있다면 넣어줘야됨					
				'Call vspdData_Change(C_W3, .ActiveRow)
'				lgOldRow = 0	' -- 같은행 발생을 초기화                                                    
				Call vspdData_ScriptLeaveCell(.ActiveCol, .ActiveRow+1, .ActiveCol, .ActiveRow, False)
					
			End With
		CAse 2
			With frm1.vspdData2 
				.focus
				ggoSpread.Source = frm1.vspdData2

				lDelRows = ggoSpread.EditUndo
					
				' -- 계산 컬럼이 있는 경우 이벤트 호출되어야 함 
				'Call vspdData2_Change(C_W9, .ActiveRow)
			End With    
	End Select
	
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
	If lgCurrGrid = GRID_2 And frm1.vspdData.MaxRows = 0 Then lgCurrGrid = GRID_1
	
	Select Case lgCurrGrid
		Case GRID_1
			iOldCol = .vspdData.ActiveCol
			.vspdData.focus
			
			ggoSpread.Source = .vspdData
			ggoSpread.InsertRow  .vspdData.ActiveRow, imRow
			
			SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1
			
			'iSeqNo = MaxSpreadVal(.vspdData, C_MINOR_CD, .vspdData.ActiveRow)
			
			'Call InsertSeqNo(.vspdData, iSeqNo, C_MINOR_CD, .vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1)

			'If imRow = 1 Then
			'	lgCurrGrid = GRID_2
			'	Call FncInsertRow(1)
			'	lgCurrGrid = GRID_1
			'End If
			
			'Call vspdData_Click(iOldCol, .vspdData.ActiveRow)
			Call vspdData_ScriptLeaveCell(.vspdData.ActiveCol, .vspdData.ActiveRow-1, .vspdData.ActiveCol, .vspdData.ActiveRow, False)
			
			frm1.vspdData.SetActiveCell iOldCol, .vspdData.ActiveRow
			.vspdData.focus

		Case GRID_2
			' -- 부모그리드의  현재행의 seq_no를 읽어온다.
			.vspdData.Col = C_MINOR_CD : .vspdData.Row = .vspdData.ActiveRow : sMinorCd = Trim(.vspdData.Text)
			
			If sMinorCd = "" Then
				Call DisplayMsgBox("970021","x",frm1.txtMINOR_CD.alt,"x")  
				Exit Function
			End If
			
			.vspdData2.ReDraw = False
			.vspdData2.focus
			ggoSpread.Source = .vspdData2
			ggoSpread.InsertRow  .vspdData2.ActiveRow, imRow
			SetSpreadColor2 .vspdData2.ActiveRow, .vspdData2.ActiveRow + imRow - 1
			
			'iSubSeqNo = MaxSpreadVal2(.vspdData2, iSeqNo, C_CTRL_CD, .vspdData2.ActiveRow)
			Call InsertParentMinorCd(sMinorCd, C_MINOR_CD, .vspdData2.ActiveRow, .vspdData2.ActiveRow + imRow - 1)

			.vspdData2.ReDraw = True
			
			frm1.vspdData2.SetActiveCell C_CTRL_CD, .vspdData2.ActiveRow
			.vspdData2.focus

	End Select		
	
	'------ Developer Coding part (End )   -------------------------------------------------------------- 

	End With
    
    Set gActiveElement = document.ActiveElement  
    If Err.number = 0 Then
       FncInsertRow = True                                                          '☜: Processing is OK
    End If   
End Function


Function FncDeleteRow() 
    Dim lDelRows

	Select Case lgCurrGrid 
		Case  1	
			With frm1.vspdData 
				.focus
				ggoSpread.Source = frm1.vspdData 

				lDelRows = ggoSpread.DeleteRow
				
				' -- 계산 컬럼이 존재시 이벤트 호출되어야 함		
				'Call vspdData_Change(C_W3, .ActiveRow)
					
				lgCurrGrid = 2 : Call DeleteChildGrid2()
					
			End With
		Case 2
			With frm1.vspdData2 
				.focus
				ggoSpread.Source = frm1.vspdData2

				lDelRows = ggoSpread.DeleteRow
				
				' -- 계산 컬럼이 존재시 이벤트 호출되어야 함	
				'Call vspdData2_Change(C_W9, .ActiveRow)

			End With    
	End Select
	
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
			strVal = strVal & "&txtMINOR_CD=" & Trim(.txtMINOR_CD.value)	
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
   	Call SetSpreadLock2
	Frm1.vspdData.Focus
	
    Set gActiveElement = document.ActiveElement   
   	
   	'Call vspdData_Click(C_MINOR_CD, 1)
   	Call vspdData_ScriptLeaveCell(C_SP_NM, 0, C_SP_NM, 1, False)
   	
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

	
	If lgIntFlgMode = Parent.OPMD_CMODE Then
		sMinorCd = UCase(Trim(frm1.txtMINOR_CD.value))
	Else
		sMinorCd = UCase(Trim(frm1.hMINOR_CD.value))
		frm1.txtMINOR_CD.value = sMinorCd
	End If

	With frm1.vspdData
	
		ggoSpread.Source = frm1.vspdData	
		
		For lRow = 1 To .MaxRows
			strVal = ""
			strDel = ""
			.Row = lRow	: .Col = 0
        
			Select Case .Text

	            Case ggoSpread.InsertFlag	
					strVal = "C" & iColSep 
					
					.Col = .MaxCols				: strVal = strVal & Trim(lRow) & iColSep
					.Col = C_MINOR_CD			: strVal = strVal & Trim(.Value) & iColSep
					.Col = C_SP_NM				: strVal = strVal & Trim(.Value) & iColSep &  Parent.gRowSep

					sSQLI1 = sSQLI1 + strVal
					lGrpCnt = lGrpCnt + 1

	            Case ggoSpread.UpdateFlag		
	            
					strVal = "U" & iColSep 
					
					.Col = .MaxCols				: strVal = strVal & Trim(lRow) & iColSep
					.Col = C_MINOR_CD			: strVal = strVal & Trim(.Value) & iColSep
					.Col = C_SP_NM				: strVal = strVal & Trim(.Value) & iColSep &  Parent.gRowSep

					sSQLU1 = sSQLU1 + strVal
					lGrpCnt = lGrpCnt + 1

	            Case ggoSpread.DeleteFlag		

					strVal = "D" & iColSep 
					
					.Col = .MaxCols				: strVal = strVal & Trim(lRow) & iColSep
					.Col = C_MINOR_CD			: strVal = strVal & Trim(.Value) & iColSep &  Parent.gRowSep
					
					sSQLD1 = sSQLD1 + strVal
					lGrpCnt = lGrpCnt + 1
                
	        End Select

		Next

	End With


	
	With frm1.vspdData2
	
		ggoSpread.Source = frm1.vspdData2	
		
		For lRow = 1 To .MaxRows
			strVal = ""
			strDel = ""
			.Row = lRow	: .Col = 0
        
			Select Case .Text

	            Case ggoSpread.InsertFlag	
					strVal = "C" & iColSep 
					
					.Col = .MaxCols					: strVal = strVal & Trim(.Value) & iColSep
					.Col = C_MINOR_CD				: strVal = strVal & Trim(.Value) & iColSep
					.Col = C_CTRL_CD				: strVal = strVal & Trim(.Value) & iColSep &  Parent.gRowSep

					sSQLI2 = sSQLI2 + strVal
					lGrpCnt = lGrpCnt + 1

	            Case ggoSpread.UpdateFlag		
	            
					strVal = "U" & iColSep 
					
					.Col = .MaxCols					: strVal = strVal & Trim(.Value) & iColSep
					.Col = C_MINOR_CD				: strVal = strVal & Trim(.Value) & iColSep
					.Col = C_CTRL_CD				: strVal = strVal & Trim(.Value) & iColSep &  Parent.gRowSep

					sSQLU2 = sSQLU2 + strVal
					lGrpCnt = lGrpCnt + 1

	            Case ggoSpread.DeleteFlag		

					strVal = "D" & iColSep 
					
					.Col = .MaxCols					: strVal = strVal & Trim(.Value) & iColSep
					.Col = C_MINOR_CD				: strVal = strVal & Trim(.Value) & iColSep
					.Col = C_CTRL_CD				: strVal = strVal & Trim(.Value) & iColSep &  Parent.gRowSep

					sSQLD2 = sSQLD2 + strVal
					lGrpCnt = lGrpCnt + 1
                
	        End Select
                
		Next

	End With
		
	frm1.txtMode.value = Parent.UID_M0002
	'frm1.txtMaxRows.value = lGrpCnt-1

	frm1.txtSpreadI1.value = sSQLI1
	frm1.txtSpreadU1.value = sSQLU1
	frm1.txtSpreadD1.value = sSQLD1
	frm1.txtSpreadI2.value = sSQLI2
	frm1.txtSpreadU2.value = sSQLU2
	frm1.txtSpreadD2.value = sSQLD2
	
	
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
	frm1.vspdData2.MaxRows = 0
	Call MainQuery()
		
End Function

'========================================================================================
Function DbDelete() 
    Err.Clear
    DbDelete = False

    Dim strVal

    strVal = BIZ_PGM_ID & "?txtMode="	& parent.UID_M0003							'☜: 비지니스 처리 ASP의 상태 
    strVal = strVal     & "&txtMINOR_CD=" & frm1.txtMINOR_CD.value					    '☜: Query Key        
	
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
'check plant
	If Trim(frm1.txtMINOR_CD.value) <> "" Then
		strWhere = " minor_cd= " & FilterVar(frm1.txtMINOR_CD.value, "''", "S") & "  "
		strWhere = strWhere &	"	and major_cd=" & filterVar("C4010","","S")

		Call CommonQueryRs(" minor_nm ","	 b_minor ", strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
		IF Len(lgF0) < 1 Then 
			Call DisplayMsgBox("970000","X",frm1.txtMINOR_CD.alt,"X")
			frm1.txtMINOR_CD.focus 
			frm1.txtMINOR_NM.value = ""
			ChkKeyField = False
			Exit Function
		End If
	
		strDataNm = split(lgF0,chr(11))
		frm1.txtMINOR_NM.value = strDataNm(0)
	Else
		frm1.txtMINOR_NM.value=""
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
					<TD WIDTH=* align=right>&nbsp;&nbsp;</TD>
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
									<TD CLASS="TD5">직과항목</TD>
									<TD CLASS="TD656"><INPUT ClASS="clstxt" TYPE=TEXT NAME="txtMINOR_CD" SIZE=10 MAXLENGTH=3 tag="15XXXU" ALT="직과항목"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDstbFctr" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenMinor(0)">
									<input NAME="txtMINOR_NM" TYPE="TEXT"  tag="14XXX" size="30">
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
								<TD WIDTH="60%" NOWRAP>
								<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> ID=vspdData NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD"> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
								</TD>
								<TD WIDTH="40%" NOWRAP>
								<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> ID=vspdData2 NAME=vspdData2 WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD"> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
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
<TEXTAREA CLASS="hidden" NAME="txtSpreadI2" tag="24" TABINDEX= "-1"></TEXTAREA>
<TEXTAREA CLASS="hidden" NAME="txtSpreadU1" tag="24" TABINDEX= "-1"></TEXTAREA>
<TEXTAREA CLASS="hidden" NAME="txtSpreadU2" tag="24" TABINDEX= "-1"></TEXTAREA>
<TEXTAREA CLASS="hidden" NAME="txtSpreadD1" tag="24" TABINDEX= "-1"></TEXTAREA>
<TEXTAREA CLASS="hidden" NAME="txtSpreadD2" tag="24" TABINDEX= "-1"></TEXTAREA>
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

