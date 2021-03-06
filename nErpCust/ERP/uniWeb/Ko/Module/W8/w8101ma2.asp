<%@ LANGUAGE="VBSCRIPT" %>
<!--
'======================================================================================================
'*  1. Module Name          : 법인세 
'*  2. Function Name        : 법인세 조정 
'*  3. Program ID           : W8101MA2
'*  4. Program Name         : W8101MA2.asp
'*  5. Program Desc         : 이월결손금 팝업 
'*  6. Modified date(First) : 2005/03/03
'*  7. Modified date(Last)  : 2005/03/03
'*  8. Modifier (First)     : 최영태 
'*  9. Modifier (Last)      : 최영태 
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================
 -->
<HTML>
<HEAD>
<TITLE>이월결손금 팝업</TITLE>

<!--
########################################################################################################
#						   3.    External File Include Part
########################################################################################################-->

<!--
========================================================================================================
=                          3.1 Server Side Script
========================================================================================================-->

<!-- #Include file="../../inc/IncSvrMain.asp" -->
<%
'=========================  로긴중인 유저의 법인코드를 출력하기 위해  ======================
    Call LoadBasisGlobalInf()
%>

<!--
========================================================================================================
=                          3.2 Style Sheet
======================================================================================================== -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		

<!--
========================================================================================================
=                          3.3 Client Side Script
======================================================================================================== -->
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliPAMain.vbs">			</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliPAEvent.vbs">			</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliPAOperation.vbs">		</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliVariables.vbs">		</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliDBAgentA.vbs">			</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliDBAgentVariables.vbs">	</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliRdsQuery.vbs">			</SCRIPT>
<SCRIPT LANGUAGE ="JavaScript"SRC = "../../inc/incImage.js">			</SCRIPT>

<Script Language="VBScript">
Option Explicit                                                        '☜: indicates that All variables must be declared in advance
	

'########################################################################################################
'#                       4.  Data Declaration Part
'########################################################################################################

'========================================================================================================
'=                       4.1 External ASP File
'========================================================================================================
Const BIZ_PGM_ID 		= "w8101mb2.asp"                              '☆: Biz Logic ASP Name

'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
Const C_SHEETMAXROWS_D  = 100                                          '☆: Server에서 한번에 fetch할 최대 데이타 건수 
Const C_MaxKey          = 18												'☆: SpreadSheet의 키의 갯수 

'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
Dim lgIsOpenPop                                          
Dim lgPopUpR                                              

Dim  IsOpenPop                                                  '☜: 마크                                  

Dim arrReturn
Dim arrParent
Dim arrParam

arrParent = window.dialogArguments
Set PopupParent = arrParent(0)
arrParam = arrParent(1)

	 '------ Set Parameters from Parent ASP ------ 
Dim C_SEQ_NO
Dim C_W1
Dim C_W2
Dim C_W3


	top.document.title = "이월결손금 팝업"

'########################################################################################################
'#                       5.Method Declaration Part
'########################################################################################################

'========================================================================================================
'========================================================================================================
'                        5.1 Common Method-1
'========================================================================================================
'========================================================================================================

'========================================================================================================
' Name : InitVariables()	
' Desc : Initialize value
'========================================================================================================	
Sub InitVariables()
    Redim arrReturn(0)

    lgStrPrevKey     = ""
    lgPageNo         = ""
    lgIntFlgMode     = PopupParent.OPMD_CMODE                          'Indicates that current mode is Create mode
	lgBlnFlgChgValue = False
    lgSortKey        = 1

	Self.Returnvalue = arrReturn
 
End Sub

'============================================  초기화 함수  ====================================
Sub InitSpreadPosVariables()
	C_SEQ_NO			= 1
	C_W1				= 2
    C_W2				= 3
	C_W3				= 4
End Sub

'========================================================================================================
' Name : SetDefaultVal()	
' Desc : Set default value
'========================================================================================================
Sub SetDefaultVal()
 
End Sub


'======================================================================================================
'   Function Name : EscPopup(Byval arrRet)
'   Function Desc : 
'=======================================================================================================
Function EscPopup(Byval iWhere)
	With frm1
		Select Case iWhere
			Case 0		
				.txtDeptCd.focus
			Case 1
				.txtDealBpCd.focus
		End Select
	End With
End Function
'======================================================================================================
'   Function Name : SetPopup(Byval arrRet)
'   Function Desc : 
'=======================================================================================================
Function SetPopUp(Byval arrRet, Byval iWhere)
	With frm1
		Select Case iWhere
			Case 0		
				.txtDeptCd.value = arrRet(0)
				.txtDeptNm.value = arrRet(1)
				.txtDeptCd.focus
			Case 1
				.txtDealBpCd.value = arrRet(0)
				.txtDealBpNm.value = arrRet(1)
				.txtDealBpCd.focus
		End Select
	End With
End Function
'========================================  2.3 LoadInfTB19029()  =========================================
' Name : LoadInfTB19029()	
' Desc : Set System Number format
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "A","NOCOOKIE","RA") %>                                '☆: 
	<% Call LoadBNumericFormatA("Q", "A", "NOCOOKIE", "RA") %>
End Sub

'========================================================================================================
'	Name : CookiePage()
'	Description : JUMP시 Load화면으로 조건부로 Value
'========================================================================================================
Function CookiePage(ByVal Kubun)

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 

	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
		
End Function

'========================================================================================================
' Name : InitComboBox()
' Desc : Set ComboBox
'========================================================================================================
Sub InitComboBox()

		
End Sub

'========================================================================================================
' Name : InitSpreadSheet
' Desc : This method initializes spread sheet column property
'========================================================================================================
Sub InitSpreadSheet()
    Call initSpreadPosVariables()  

	With frm1.vspdData
	
		ggoSpread.Source = frm1.vspdData	
		'patch version
		'ggoSpread.Spreadinit "V20041222",,PopupParent.gAllowDragDropSpread    
			 
		.ReDraw = false

		.MaxCols = C_W3 + 1							'☜: 최대 Columns의 항상 1개 증가시킴 
		.Col = .MaxCols														'☆: 사용자 별 Hidden Column
		.ColHidden = True    
				       
		.MaxRows = 0
		ggoSpread.ClearSpreadData
			 
		ggoSpread.SSSetEdit		C_SEQ_NO,	"SEQ_NO",	10,,,15,1
		ggoSpread.SSSetDate		C_W1,	"발생일자",	10,		2,		Parent.gDateFormat,	-1
		ggoSpread.SSSetFloat	C_W2,	"이월결손공제가능액",	17,		"8",		ggStrIntegeralPart,		ggStrDeciPointPart,		PopupParent.gComNum1000,		PopupParent.gComNumDec ,,,,"0","" 
		ggoSpread.SSSetFloat	C_W3,	"당기이월결손금공제액",	17,		"8",		ggStrIntegeralPart,		ggStrDeciPointPart,		PopupParent.gComNum1000,		PopupParent.gComNumDec ,,,,"0","" 
		
		.Col = C_SEQ_NO														'☆: 사용자 별 Hidden Column
		.ColHidden = True  
		'Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
		'Call ggoSpread.SSSetColHidden(C_SEQ_NO,C_SEQ_NO,True)
							
		'Call InitSpreadComboBox()
			
		.ReDraw = true
			
		Call SetSpreadLock 
    
    End With
    
End Sub

'========================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================================
Sub SetSpreadLock()
    With frm1

    .vspdData.ReDraw = False
    
    ggoSpread.Source = frm1.vspdData
	ggoSpread.SpreadLockWithOddEvenRowColor()

    .vspdData.ReDraw = True

    End With
End Sub

'===========================================  2.3.1 OkClick()  ==========================================
'=	Name : OkClick()																					=
'=	Description : Return Array to Opener Window when OK button click									=
'=				  이 부분에서 컬럼 추가하고 데이타 전송이 일어나야 합니다.   							=
'========================================================================================================
Function OKClick()
	If frm1.vspdData.MaxRows > 0 Then
		Call DisplayMsgBox("W80003", PopupParent.VB_INFORMATION, "", "") 

		Call FncSave
	Else
		Call SetReturnVal
	End If
End Function

Sub SetReturnVal()	' 리턴데이타 
	
	With frm1.vspdData
		If .MaxRows > 0 Then 				
			Redim arrReturn(1)
			.Row			= .MaxRows
			.Col			= C_W3	
			arrReturn(0)	= .Value
		End if			
	End With

	Self.Returnvalue = arrReturn
	Self.Close()
End Sub

'=========================================  2.3.2 CancelClick()  ========================================
'=	Name : CancelClick()																				=
'=	Description : Return Array to Opener Window for Cancel button click 								=
'========================================================================================================
Function CancelClick()
	Self.Close()			
End Function

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
    Call LoadInfTB19029()
    Call ggoOper.FormatField(Document, "1", PopupParent.ggStrIntegeralPart, PopupParent.ggStrDeciPointPart, PopupParent.gDateFormat, PopupParent.gComNum1000, PopupParent.gComNumDec,,,PopupParent.ggStrMinPart,PopupParent.ggStrMaxPart)
    Call ggoOper.LockField(Document, "N")
    
	Call InitVariables()														
	'Call SetDefaultVal()	
	Call InitSpreadSheet()
    'Call InitComboBox()
    
    Call FncQuery
End Sub
'========================================================================================================
' Name : Form_QueryUnload
' Desc : This sub is called from window_Unonload in Common.vbs automatically
'========================================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )
 
End Sub

'========================================================================================================
' Name : FncQuery
' Desc : This function is called from MainQuery in Common.vbs
'========================================================================================================
Function FncQuery() 

	Dim IntRetCD
    FncQuery = False                                            
    
    Err.Clear                                                   
    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")						
    Call InitVariables() 											
    ggoSpread.Source = frm1.vspdData
	ggospread.ClearSpreadData		'Buffer Clear
    '-----------------------
    'Check condition area
    '-----------------------

    '-----------------------
    'Query function call area
    '-----------------------
    If DbQuery = False Then Exit Function

    FncQuery = True													
End Function

'========================================================================================================
' Name : FncNew
' Desc : This function is called from MainNew in Common.vbs
'========================================================================================================
Function FncNew()
    Dim IntRetCD 
    
    FncNew = False																 '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Set gActiveElement = document.ActiveElement   
    FncNew = True																 '☜: Processing is OK
End Function
	
'========================================================================================================
' Name : FncDelete
' Desc : This function is called from MainDelete in Common.vbs
'========================================================================================================
Function FncDelete()
    Dim intRetCD
    
    FncDelete = False                                                            '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Set gActiveElement = document.ActiveElement       
    FncDelete = True                                                             '☜: Processing is OK
End Function


'========================================================================================================
' Name : FncSave
' Desc : This function is called from MainSave in Common.vbs
'========================================================================================================
Function FncSave()
    Dim IntRetCD 
    
    FncSave = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status    
   
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	If DbSave = False Then Exit Function  
	'------ Developer Coding part (End )   -------------------------------------------------------------- 

    Set gActiveElement = document.ActiveElement   
    FncSave = True                                                               '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncCopy
' Desc : This function is called from MainCopy in Common.vbs
' Keep : Make sure to clear primary key area
'========================================================================================================
Function FncCopy()
	Dim IntRetCD

    FncCopy = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
	
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 

	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Set gActiveElement = document.ActiveElement   
    FncCopy = True                                                               '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncCancel
' Desc : This function is called by MainCancel in Common.vbs
'========================================================================================================
Function FncCancel() 
    FncCancel = False                                                            '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Set gActiveElement = document.ActiveElement   
    FncCancel = False                                                            '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncInsertRow
' Desc : This function is called by MainInsertRow in Common.vbs
'========================================================================================================
Function FncInsertRow()
    FncInsertRow = False                                                         '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Set gActiveElement = document.ActiveElement   
    FncInsertRow = True                                                          '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncDeleteRow
' Desc : This function is called by MainDeleteRow in Common.vbs
'========================================================================================================
Function FncDeleteRow()
    Dim lDelRows
    FncDeleteRow = False                                                         '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Set gActiveElement = document.ActiveElement   
    FncDeleteRow = True                                                          '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncPrint
' Desc : This function is called by MainPrint in Common.vbs
'========================================================================================================
Function FncPrint()
    FncPrint = False                                                             '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    
    '--------- Developer Coding Part (End) ------------------------------------------------------------
	Call Parent.FncPrint()                                                       '☜: Protect system from crashing
    FncPrint = True                                                              '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncPrev
' Desc : This function is called by MainPrev in Common.vbs
'========================================================================================================
Function FncPrev() 
    FncPrev = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    
    '--------- Developer Coding Part (End) ------------------------------------------------------------
    Set gActiveElement = document.ActiveElement   	
    FncPrev = True                                                               '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncNext
' Desc : This function is called by MainPrev in Common.vbs
'========================================================================================================
Function FncNext() 
    FncNext = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    
    '--------- Developer Coding Part (End) ------------------------------------------------------------
    Set gActiveElement = document.ActiveElement   	
    FncNext = True                                                               '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncExcel
' Desc : This function is called by MainExcel in Common.vbs
'========================================================================================================
Function FncExcel() 
    FncExcel = False                                                             '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

	'Call Parent.FncExport(PopupParent.C_MULTI)

    FncExcel = True                                                              '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncFind
' Desc : This function is called by MainFind in Common.vbs
'========================================================================================================
Function FncFind() 
    FncFind = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

	'Call Parent.FncFind(PopupParent.C_MULTI, True)

    FncFind = True                                                               '☜: Processing is OK
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

'========================================================================================================
' Name : FncExit
' Desc : This function is called by MainExit in Common.vbs
'========================================================================================================
Function FncExit()
	Dim IntRetCD

    FncExit = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900016", PopupParent.VB_YES_NO,"X","X")
		If IntRetCD = vbNo Then
		    Exit Function
		End If
    End If
    
    '--------- Developer Coding Part (End) ------------------------------------------------------------

    FncExit = True                                                               '☜: Processing is OK
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
	Dim strVal, iRet, iMaxRows, arrSeqNO, arrW1, arrW2, dblW2, dblBasicAmt, dblW4, iRow

    Err.Clear                                                       
    DbQuery = False
    
    With frm1.vspdData
    
	iRet = CommonQueryRs("SEQ_NO, W1, W2"," dbo.ufn_TB_3_Get_W07('" & arrParam(0) & "','" & arrParam(1) & "','" & arrParam(2) & "')", "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

	If iRet = True Then
		arrSeqNO	= Split(lgF0, chr(11))
		arrW1		= Split(lgF1, chr(11))
		arrW2		= Split(lgF2, chr(11))
		iMaxRows	= UBound(arrSeqNO)

		ggoSpread.InsertRow ,iMaxRows
		.Row = .MaxRows		: .Col = C_W1	: .CellType = 1	: .TypeHAlign = 2
		
		For iRow = 0 To iMaxRows -1
			.Row = iRow + 1
			.Col = 0		: .Value = iRow + 1
			.Col = C_SEQ_NO	: .Value = arrSeqNo(iRow)
			.Col = C_W1		: .Text = arrW1(iRow)
			.Col = C_W2		: .Value = arrW2(iRow)
		Next
		
		' 기준금액 결정 
		.Row = .MaxRows		: .Col = C_W2	: dblW2 = UNICDbl(.value)
		If  dblW2 > arrParam(3) Then
			dblBasicAmt = arrParam(3)
		Else
			dblBasicAmt = dblW2
		End If
		
		For iRow = 1 To iMaxRows -1
			.Row = iRow
			.Col = C_W2		: dblW2 = UNICDbl(.value)
			If dblBasicAmt - dblW2 > 0 Then
				.Col = C_W3		: .value = dblW2
				dblBasicAmt = dblBasicAmt - dblW2	
			Else
				.Col = C_W3		: .value = dblBasicAmt
				dblBasicAmt = 0
			End If
		Next
		
		Call FncSumSheet(frm1.vspdData, C_W3, 1, .MaxRows - 1, true, .MaxRows, C_W3, "V")
	End If					
    
    End With
    
    DbQuery = True

End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================
Function DbQueryOk()												

	lgBlnFlgChgValue = False
    lgIntFlgMode     = PopupParent.OPMD_UMODE												'⊙: Indicates that current mode is Update mode
	
	If frm1.vspdData.MaxRows > 0 Then
		frm1.vspdData.Focus
	End If
		
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
    Dim strVal, strDel
 
    DbSave = False                                                          
 
     strVal = ""
    strDel = ""
    lGrpCnt = 1
    
	With frm1.vspdData
	
		ggoSpread.Source = frm1.vspdData
		lMaxRows = .MaxRows : lMaxCols = .MaxCols
			
		' ----- 1번째 그리드 
		For lRow = 1 To .MaxRows
    
	       .Row = lRow
	       .Col = 0

	        strVal = strVal & "U"  &  PopupParent.gColSep
		       
		  ' 모든 그리드 데이타 보냄     
			For lCol = 1 To lMaxCols
				.Col = lCol : strVal = strVal & Trim(.Text) &  PopupParent.gColSep
			Next
			strVal = strVal & Trim(.Text) &  PopupParent.gRowSep
		Next
		
	End With

	frm1.txtSpread.value =  strDel & strVal

    strVal = BIZ_PGM_ID
	strVal = strVal & "?txtMode=" & PopupParent.UID_M0002					      
	strVal = strVal & "&txtFISC_YEAR=" & arrParam(1)
	strVal = strVal & "&cboREP_TYPE=" & arrParam(2)
						
	Call ExecMyBizASP(frm1, strVal) 
	
    DbSave = True                                                           
End Function


Function DbSaveOk()		

    Call SetReturnVal()
End Function

'========================================================================================================
'========================================================================================================
'                        5.4 User-defined Method
'========================================================================================================
'========================================================================================================


'========================================================================================================
'	Name : OpenGroupPopup()
'	Description : Group Condition PopUp
'========================================================================================================
Function OpenGroupPopup()

	Dim arrRet
	Dim arrParam
	Dim TInf(5)
	Dim ii
	
	On Error Resume Next
	
	ReDim arrParam(PopupParent.C_MaxSelList * 2 - 1 )

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True
	
    TInf(0) = gMethodText
  
	For ii = 0 to PopupParent.C_MaxSelList * 2 - 1 Step 2
      arrParam(ii + 0 ) = lgPopUpR(ii / 2  , 0)
      arrParam(ii + 1 ) = lgPopUpR(ii / 2  , 1)
    Next  
      
  
	arrRet = window.showModalDialog("../../ComAsp/ADOGrpSortPopup.asp",Array(GetSQLSortFieldCD("A"),GetSQLSortFieldNM("A"),arrParam,TInf),"dialogWidth=420px; dialogHeight=250px;; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If arrRet(0) = "0" Then
		If Err.Number <> 0 Then
			Err.Clear 
		End If
		Exit Function
	Else
	
	   For ii = 0 to PopupParent.C_MaxSelList * 2 - 1 Step 2
           lgPopUpR(ii / 2 ,0) = arrRet(ii + 1)  
           lgPopUpR(ii / 2 ,1) = arrRet(ii + 2)
       Next    
	   
       Call InitVariables
       Call InitSpreadSheet
   End If
End Function

'========================================================================================================
'========================================================================================================
'                        5.5 Tag Event
'========================================================================================================
'========================================================================================================

'========================================================================================================
' Function Name : vspdData_MouseDown
' Function Desc : 
'========================================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)

   If Button = 2 And gMouseClickStatus = "SPC" Then
      gMouseClickStatus = "SPCR"
   End If
End Sub    

'========================================================================================================
'   Event Name : vspdData_GotFocus
'   Event Desc : This event is spread sheet data changed
'========================================================================================================
Sub vspdData_GotFocus()
    ggoSpread.Source = frm1.vspdData
End Sub

'========================================================================================================
'   Event Name : vspdData_DblClick
'   Event Desc :
'========================================================================================================
Function vspdData_DblClick(ByVal Col, ByVal Row)
	If Row = 0 Then              		' Title cell을 dblclick했거나....
		Exit Function
	End If
	If Frm1.vspdData.MaxRows = 0 Then  	'NO Data
		Exit Function
	End If
	Call OKClick
End Function
	
'========================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'========================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
    
End Sub
	
'========================================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
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
           If DbQuery = False Then
              Exit Sub
           End if
    	End If
    End If
    
End Sub

Function vspdData_KeyPress(KeyAscii)

End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>

<!--
'########################################################################################################
'#						6. TAG 부																		#
'######################################################################################################## 
-->
<BODY SCROLL=NO TABINDEX="-1">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE CELLSPACING=0 CLASS="basicTB">
	<TR><TD HEIGHT=2>
	</TD></TR>
	<TR><TD HEIGHT=100%>
			<script language =javascript src='./js/w8101ma2_vaSpread1_vspdData.js'></script>
	</TD></TR>
	<TR><TD HEIGHT=20>
		<TABLE CLASS="basicTB" CELLSPACING=0>
			<TR>
				<TD WIDTH=70% NOWRAP>&nbsp;&nbsp;
				<IMG SRC="../../image/query_d.gif"  Style="CURSOR: hand" ALT="Search" NAME="Search" ONCLICK="FncQuery()"    onMouseOut="javascript:MM_swapImgRestore()"  onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../image/Query.gif',1)"></IMG></TD>
				<TD WIDTH=30% ALIGN=RIGHT>
				<IMG SRC="../../image/ok_d.gif"     Style="CURSOR: hand" ALT="OK"     NAME="pop1"   ONCLICK="OkClick()"     onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../image/OK.gif',1)"></IMG>&nbsp;&nbsp;
				<IMG SRC="../../image/cancel_d.gif" Style="CURSOR: hand" ALT="CANCEL" NAME="pop2"   ONCLICK="CancelClick()" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../image/Cancel.gif',1)"></IMG>&nbsp;&nbsp;</TD>
			</TR>
		</TABLE>
	</TD></TR>
	<TR>
		<TD WIDTH=100% HEIGHT=10><IFRAME NAME="MyBizASP" SRC="../blank.htm" WIDTH=100% HEIGHT=10 FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<DIV ID="MousePT" NAME="MousePT">
	<IFRAME name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../inc/cursor.htm"></IFRAME>
</DIV>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" style="display:'none'"></TEXTAREA>
</FORM>
</BODY>
</HTML>


