<%@ LANGUAGE="VBSCRIPT" %>
<!--
<%
'********************************************************************************************************
'*  1. Module Name          : Procuremant																*
'*  2. Function Name        :																			*
'*  3. Program ID           : M9211PA1.asp																*
'*  4. Program Name         :																			*
'*  5. Program Desc         : 															*
'*  6. Comproxy List        : + B19029LookupNumericFormat												*
'*  7. Modified date(First) : 2000/03/21																*
'*  8. Modified date(Last)  : 2000/03/21																*
'*  9. Modifier (First)     :																			*
'* 10. Modifier (Last)      : KO MYOUNG JIN																			*
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"									*
'*                            this mark(⊙) Means that "may  change"									*
'*                            this mark(☆) Means that "must change"									*
'* 13. History              : 1. 2000/03/21 : 화면 design												*
'******************************************************************************************************
%>
-->
<HTML>
<HEAD>
<TITLE>입고번호</TITLE>
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<!--
'#########################################################################################################
'												1. 선 언 부 
'##########################################################################################################
'******************************************  1.1 Inc 선언   **********************************************
'	기능: Inc. Include
'*********************************************************************************************************
-->

<!--
'==========================================  1.1.1 Style Sheet  ======================================
'==========================================================================================================
-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--
'==========================================  1.1.2 공통 Include   ======================================
'==========================================================================================================
-->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentA.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript" SRC="../../inc/incImage.js"></SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit                                                             '☜: indicates that All variables must be declared in advance
                                                                            ' 명시적으로 변수를 선언 
<%'****************************************  1.2 Global 변수/상수 선언  ***********************************
'	1. Constant는 반드시 대문자 표기.
'**********************************************************************************************************%>
Const BIZ_PGM_QRY_ID 		= "m9211pb1.asp"                              '☆: Biz Logic ASP Name

'========================================================================================================
'=									1.2 Constant variables 
'========================================================================================================
Const C_SHEETMAXROWS_D  = 30                                          '☆: Fetch max count at once
Const C_MaxKey          = 1                                           '☆: key count of SpreadSheet

'========================================================================================================
'=									1.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	

'========================================================================================================
'=									1.4 User-defind Variables
'========================================================================================================

Dim C_MVMTNO
Dim C_MVMTTypeCd
Dim C_MVMTTypeNm
Dim C_MVMTDT
Dim C_PlantCd
Dim C_PlantNm
Dim C_PURGRPCd
Dim C_PURGRPNm

Dim arrReturn
Dim arrParam					
Dim arrField
Dim PlantCd
Dim arrParent

Dim gblnWinEvent
Dim lgKeyPos                                                '☜: Key위치                               
Dim lgKeyPosVal                                             '☜: Key위치 Value                         
Dim IscookieSplit 
'Dim PopupParent

arrParent = window.dialogArguments
Set PopupParent = arrParent(0)

Dim StartDate,EndDate

EndDate = UNIConvDateAToB("<%=GetSvrDate%>", PopupParent.gServerDateFormat, PopupParent.gDateFormat)
StartDate = UnIDateAdd("m", -1, EndDate, PopupParent.gDateFormat)
top.document.title = "입고번호"

'--------------- 개발자 coding part(변수선언,End)-------------------------------------------------------------
<% '==========================================  1.2.3 Global Variable값 정의  ===============================
'========================================================================================================= %>
<% '----------------  공통 Global 변수값 정의  ----------------------------------------------------------- %>

<% '++++++++++++++++  Insert Your Code for Global Variables Assign  ++++++++++++++++++++++++++++++++++++++ %>
 Dim IsOpenPop						' Popup
 Dim arrValue(3)                    ' Popup되는 창으로 넘길때 인수를 배열로 넘김 

<% '#########################################################################################################
'												2. Function부 
'
'	내용 : 개발자가 정의한 함수, 즉 Event관련 함수를 제외한 모든 사용자 정의 함수 기슬 
'	공통으로 적용 사항 : 1. Sub 또는 Function을 호출할 때 반드시 Call을 쓴다.
'		     	     	 2. Sub, Function 이름에 _를 쓰지 않도록 한다. (Event와 구별하기 위함) 
'######################################################################################################### %>


'========================================================================================================
' Name : InitSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================
Sub InitSpreadPosVariables()

		C_MVMTNO		= 1
		C_MVMTTypeCd	= 2
		C_MVMTTypeNm	= 3
		C_MVMTDT		= 4
		C_PlantCd		= 5
		C_PlantNm		= 6
		C_PURGRPCd		= 7
		C_PURGRPNm		= 8
	
End Sub



<% '==========================================  2.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
'========================================================================================================= %>
Sub InitVariables()
	Dim arrParent

	lgStrPrevKeyIndex = ""
	
	lgIntFlgMode = PopupParent.OPMD_CMODE
	gblnWinEvent = False
	
    lgSortKey = 1                                       '⊙: initializes sort direction
	Redim arrReturn(0)
	Self.Returnvalue = arrReturn

End Sub
<% '==========================================  2.2 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : 서버로 부터 필드 정의 정보를 가져옴 
'                 lgSort...로 시작하는 변수 영역에 sort대상 목록을 저장 
'                 IsPopUpR 변수영역에 sort 정보의 기본이 되는 값 저장 
'                 프로그램 ID를 넣고 go버튼을 누르거나 menu tree에서 클릭하는 순간 넘어옴                  
'========================================================================================================= %>
Sub SetDefaultVal()

<%'--------------- 개발자 coding part(실행로직,Start)--------------------------------------------------%>
	frm1.txtFrRcptDt.Text = StartDate
	frm1.txtToRcptDt.Text = EndDate
<%'--------------- 개발자 coding part(실행로직,End)----------------------------------------------------%>

End Sub



<%'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== %>
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029A("I", "*", "NOCOOKIE", "PA") %>     
End Sub

<%
'++++++++++++++++++++++++++++++++++++++++++  2.3 개발자 정의 함수  ++++++++++++++++++++++++++++++++++++++
'+	개발자 정의 Function, Procedure																		+
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
%>
<%
'===========================================  2.3.1 OkClick()  ==========================================
'=	Name : OkClick()																					=
'=	Description : Return Array to Opener Window when OK button click									=
'========================================================================================================
%>	
	Function OKClick()
		Dim intColCnt
		With frm1.vspdData	
			Redim arrReturn(.MaxCols - 1)
			If .MaxRows > 0 Then 
			.Row = .ActiveRow
			'For intColCnt = 0 To .MaxCols - 1
			'	.Col = intColCnt + 1
				.Col = C_MVMTNO
				arrReturn(0) = .Text
			'Next
			end if
		End With
		
		Self.Returnvalue = arrReturn
		Self.Close()
		
	End Function
<%
'=========================================  2.3.2 CancelClick()  ========================================
'=	Name : CancelClick()																				=
'=	Description : Return Array to Opener Window for Cancel button click 								=
'========================================================================================================
%>
Function CancelClick()
		Redim arrReturn(0)
		arrReturn(0) = ""
		self.Returnvalue = arrReturn
		Self.Close()
End Function

<%
'=========================================  2.3.3 Mouse Pointer 처리 함수 ===============================
'========================================================================================================
%>
	Function MousePointer(pstr1)
	      Select case UCase(pstr1)
	            case "PON"
					window.document.search.style.cursor = "wait"
	            case "POFF"
					window.document.search.style.cursor = ""
	      End Select
	End Function
<%
'==========================================================================================
'   Event Name : txtFrRcptDt
'   Event Desc :
'==========================================================================================
%>
Sub txtFrRcptDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtFrRcptDt.Action = 7
		Call SetFocusToDocument("M")	
        frm1.txtFrRcptDt.Focus
	End If
End Sub

<%
'==========================================================================================
'   Event Name : txtToRcptDt
'   Event Desc :
'==========================================================================================
%>
Sub txtToRcptDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtToRcptDt.Action = 7
		Call SetFocusToDocument("M")	
        frm1.txtToRcptDt.Focus
	End If
End Sub	
	
<% 
'*******************************************  2.4 POP-UP 처리함수  **************************************
'*	기능: POP-UP																						*
'*	Description : POP-UP Call하는 함수 및 Return Value setting 처리										*
'********************************************************************************************************
%>
<%
'===========================================  2.4.1 POP-UP Open 함수()  =================================
'=	Name : Open???()																					=
'=	Description : POP-UP Open																			=
'========================================================================================================
%>
<% '------------------------------------------  OpenPoType()  -------------------------------------------------
'	Name : OpenPoType()
'	Description : OpenPoType PopUp
'--------------------------------------------------------------------------------------------------------- %>
Function OpenConSItemDC(Byval iWhere)

	Dim arrRet
		Dim arrParam(5), arrField(6), arrHeader(6)

		If gblnWinEvent = True Then Exit Function

		gblnWinEvent = True
		
		Select Case iWhere	
				
	Case 1
						
		arrParam(0) = "출고공장"				
		arrParam(1) = "B_Biz_Partner"
	
		arrParam(2) = Trim(frm1.txtSupplierCd.Value)
		arrParam(3) = ""							
	
		'arrParam(4) = "Bp_Type in ('S','CS') AND usage_flag='Y'"	
		arrParam(4) = "Bp_Type <> " & FilterVar("C", "''", "S") & "  AND USAGE_FLAG = " & FilterVar("Y", "''", "S") & "  AND IN_OUT_FLAG = " & FilterVar("I", "''", "S") & " "	
		arrParam(5) = "출고공장"				
	
		arrField(0) = "BP_CD"					
		arrField(1) = "BP_NM"					

		arrHeader(0) = "출고공장"				
		arrHeader(1) = "출고공장명"	
	    
	Case 2					
	
		arrParam(0) = "구매그룹"	
		arrParam(1) = "B_Pur_Grp"					
		arrParam(2) = Trim(frm1.txtGroupCd.Value)
		'	arrParam(3) = Trim(frm1.txtGroupNm.Value)				
		arrParam(4) = " B_Pur_Grp.USAGE_FLG=" & FilterVar("Y", "''", "S") & "  "			
		arrParam(5) = "구매그룹"					
		arrField(0) = "PUR_GRP"	
		arrField(1) = "PUR_GRP_NM"	
		    
		arrHeader(0) = "구매그룹"		
		arrHeader(1) = "구매그룹명"	
	
	case 3

		arrParam(0) = "입고형태"	
		arrParam(1) = "(select distinct  IO_Type_Cd, io_type_nm from  M_CONFIG_PROCESS a,  m_mvmt_type b"
		arrParam(1) = arrParam(1) & " where a.rcpt_type = b.io_type_cd    and a.sto_flg = " & FilterVar("Y", "''", "S") & "  AND a.USAGE_FLG=" & FilterVar("Y", "''", "S") & " ) c "
	
		arrParam(2) = Trim(frm1.txtMvmtType.Value)

		'arrParam(4) = "a.rcpt_type = b.io_type_cd    and a.sto_flg = 'Y' AND a.USAGE_FLG='Y' "
		arrParam(5) = "입고형태"			
	
		arrField(0) = " c.IO_Type_Cd"
		arrField(1) = " c.IO_Type_NM"
    
		arrHeader(0) = "입고형태"		
		arrHeader(1) = "입고형태명"
		'arrParam(0) = "입고형태"	
		'arrParam(1) = "M_Mvmt_type"
			
		'arrParam(2) = Trim(frm1.txtMvmtType.Value)
		'arrParam(3) = trim(frm1.txtMvmtTypeNm.Value)
	
		'arrParam(4) = "((RCPT_FLG='Y' AND RET_FLG='N') or (RET_FLG='N' And SUBCONTRA_FLG='N')) AND USAGE_FLG='Y' "
		'arrParam(5) = "입고형태"			
			
		'arrField(0) = "IO_Type_Cd"	
		'arrField(1) = "IO_Type_NM"	
		    
		'arrHeader(0) = "입고형태"		
		'arrHeader(1) = "입고형태명"
	
	End Select

	arrParam(0) = arrParam(5)								<%' 팝업 명칭 %>

	Select Case iWhere
	Case 1,2,3
		arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	End Select

        
        arrParam(0) = arrParam(5)	
        
		gblnWinEvent = False

		If arrRet(0) = "" Then
			Exit Function
		Else
			Call SetConSItemDC(arrRet, iWhere)
		End If	
		
End Function

'==========================================  2.4.2  Set???()  ==========================================
'	Name : Set???()
'	Description : PopUp에서 전송된 값을 특정 Tag Object에 지정 
'=======================================================================================================

'-------------------------------------------------------------------------------------------------------
'	Name : SetConSItemDC()
'	Description : OpenConSItemDC Popup에서 Return되는 값 setting
'-------------------------------------------------------------------------------------------------------
Function SetConSItemDC(Byval arrRet, Byval iWhere)
	With frm1
		Select Case iWhere				  
	    Case 1
			.txtSupplierCd.Value = arrRet(0)
			.txtSupplierNm.Value = arrRet(1)		  
		Case 2
	        .txtGroupCd.Value = arrRet(0)
		    .txtGroupNm.Value = arrRet(1)		
		case 3
		    .txtMvmtType.Value = arrRet(0) 
		    .txtMvmtTypeNm.Value = arrRet(1)
		End Select	
	End With
End Function


'==========================================  2.2.3 InitSpreadSheet()  ===================================
' Name : InitSpreadSheet
' Desc : This method initializes spread sheet column property
'========================================================================================================
Sub InitSpreadSheet()
	
	Call InitSpreadPosVariables
	
	ggoSpread.Source = frm1.vspdData
	ggoSpread.Spreadinit "V20021125",, PopupParent.gAllowDragDropSpread
		
	With frm1.vspdData
	
	.ReDraw = false
    .MaxCols = C_PURGRPNm+1
    .Col = .MaxCols:	.ColHidden = True
    .MaxRows = 0

    Call GetSpreadColumnPos("A")
    
    ggoSpread.SSSetEdit 		C_MVMTNO, "입고번호", 20
    ggoSpread.SSSetEdit 		C_MVMTTypeCd, "입고형태", 10,,,4,2
    ggoSpread.SSSetEdit 		C_MVMTTypeNm, "입고형태명", 20
    ggoSpread.SSSetDate 		C_MVMTDT, "입고일자", 10, 2, PopupParent.gDateFormat
    ggoSpread.SSSetEdit 		C_PlantCd, "출고공장", 15,,,4,2
    ggoSpread.SSSetEdit		    C_PlantNm, "출고공장명", 20        '품목규격 추가 
    ggoSpread.SSSetEdit 		C_PURGRPCd, "구매그룹", 20
    ggoSpread.SSSetEdit 		C_PURGRPNm, "구매그룹명", 20
    
	Call ggoSpread.MakePairsColumn(C_MVMTTypeCd,C_MVMTTypeNm)
	Call ggoSpread.MakePairsColumn(C_PlantCd,C_PlantNm)
	Call ggoSpread.MakePairsColumn(C_PURGRPCd,C_PURGRPNm)

    ggoSpread.SSSetSplit(1)										'frozen 기능추가 
    
    Call SetSpreadLock
    
    
	.ReDraw = true
	
    End With
	            
End Sub

'============================================ 2.2.4 SetSpreadLock()  ====================================
' Name : SetSpreadLock
' Desc : This method set color and protect in spread sheet celles
'========================================================================================================
Sub SetSpreadLock()
	ggoSpread.Source = frm1.vspdData
			
	frm1.vspdData.ReDraw = False
	ggoSpread.SpreadLock -1, -1
	frm1.vspdData.ReDraw = True
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
            
            C_MVMTNO  		 = iCurColumnPos(1)
			C_MVMTTypeCd 	 = iCurColumnPos(2)
			C_MVMTTypeNm	 = iCurColumnPos(3)
			C_MVMTDT 		 = iCurColumnPos(4)
			C_PlantCd 		 = iCurColumnPos(5)
			C_PlantNm 		 = iCurColumnPos(6)
			C_PURGRPCd 		 = iCurColumnPos(7)
			C_PURGRPNm       = iCurColumnPos(8)
	
	End Select

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
    frm1.vspdData.Redraw = False
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()
	Call ggoSpread.ReOrderingSpreadData()
End Sub

'#########################################################################################################
'												3. Event부 
'	기능: Event 함수에 관한 처리 
'	설명: Window처리, Single처리, Grid처리 작업.
'         여기서 Validation Check, Calcuration 작업이 가능한 Event가 발생.
'         각 Object단위로 Grouping한다.
'########################################################################################################
'******************************************  3.1 Window 처리  *******************************************
'	Window에 발생 하는 모든 Even 처리	
'********************************************************************************************************
'==========================================  3.1.1 Form_Load()  =========================================
'	Name : Form_Load()
'	Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분 
'========================================================================================================
Sub Form_Load()
    Call LoadInfTB19029													'⊙: Load table , B_numeric_format
	
    'Html에서 tag 숫자가 1과 2로 시작하는 부분 각각Format
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)

	Call ggoOper.LockField(Document, "N")                         '⊙: Lock  Suitable  Field
    
	Call InitVariables											  '⊙: Initializes local global variables
	Call SetDefaultVal	
	Call InitSpreadSheet()
	Call MM_preloadImages("../../../CShared/image/Query.gif","../../../CShared/image/OK.gif","../../../CShared/image/Cancel.gif")
	
	Call FncQuery() 

End Sub

'=========================================  3.1.2 Form_QueryUnload()  ===================================
'   Event Name : Form_QueryUnload																		=
'   Event Desc :																						=
'========================================================================================================
Sub Form_QueryUnload(Cancel, UnloadMode)
End Sub

'*********************************************  3.2 Tag 처리  *******************************************
'*	Document의 TAG에서 발생 하는 Event 처리																*
'*	Event의 경우 아래에 기술한 Event이외의 사용을 자제하며 필요시 추가 가능하나							*
'*	Event간 충돌을 고려하여 작성한다.																	*
'********************************************************************************************************
'*********************************************  3.3 Object Tag 처리  ************************************
'*	Object에서 발생 하는 Event 처리																		*
'********************************************************************************************************
'=========================================  3.3.1 vspdData_DblClick()  ==================================
'=	Event Name : vspdData_DblClick																		=
'=	Event Desc :																						=
'========================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)
	If Row = 0 Or Frm1.vspdData.MaxRows = 0 Then 
         Exit Sub
    End If
    	
	If frm1.vspdData.MaxRows > 0 Then
		If frm1.vspdData.ActiveRow = Row Or frm1.vspdData.ActiveRow > 0 Then
			Call OKClick
		End If
	End If
End Sub

'========================================  3.3.2 vspdData_KeyPress()  ===================================
'=	Event Name : vspdData_KeyPress																		=
'=	Event Desc :																						=
'========================================================================================================
    Function vspdData_KeyPress(KeyAscii)
         On Error Resume Next
         If KeyAscii = 13 And frm1.vspdData.ActiveRow > 0 Then    'Frm1없으면 frm1삭제 
            Call OKClick()
         ElseIf KeyAscii = 27 Then
            Call CancelClick()
         End If
    End Function

'======================================  3.3.3 vspdData_TopLeftChange()  ================================
'=	Event Name : vspdData_TopLeftChange																	=
'=	Event Desc :																						=
'========================================================================================================
	Sub vspdData_TopLeftChange(ByVal OldLeft, ByVal OldTop, ByVal NewLeft, ByVal NewTop)
	
		If OldLeft <> NewLeft Then
		    Exit Sub
		End If		

		If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	    '☜: 재쿼리 체크	
			If lgPageNo <> "" Then		                                                    '다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
				If DbQuery = False Then
					Exit Sub
				End if
			End If
		End If		 
	End Sub


'=======================================================================================================
'   Event Name : OCX_KeyDown()
'   Event Desc : 조회조건부의 OCX_KeyDown시 EnterKey일 경우는 Query
'=======================================================================================================
Sub txtFrRcptDt_Keypress(KeyAscii)
	On Error Resume Next
	If KeyAscii = 27 Then
		Call CancelClick()
	Elseif KeyAscii = 13 Then
		Call FncQuery()
	End if
End Sub

Sub txtToRcptDt_Keypress(KeyAscii)
	On Error Resume Next
	If KeyAscii = 27 Then
		Call CancelClick()
	Elseif KeyAscii = 13 Then
		Call FncQuery()
	End if
End Sub
	
Sub vspdSort(ByVal SortCol, ByVal intKey)
	With frm1.vspdData
		.BlockMode = True
		.Col = 0
		.Col2 = .MaxCols
		.Row = 1
		.Row2 = .MaxRows
    
		'Row기준 Sort
		.SortBy = 0
    
		'Sort기준 Column
		.SortKey(1) = SortCol
    
		'정렬방법 
		.SortKeyOrder(1) = intKey					'0: 정렬None 1 :오름차순  2: 내림차순 
		.Action = 25								'SS_ACTION_SORT : VB number
    
		.BlockMode = False
    End With
End Sub

'======================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'======================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
  gMouseClickStatus = "SPC"   
  Set gActiveSpdSheet = frm1.vspdData
    If Row = 0 Then
        ggoSpread.Source = frm1.vspdData
        If lgSortKey = 1 Then
            ggoSpread.SSSort Col
            lgSortKey = 2
        Else
            ggoSpread.SSSort Col, lgSortKey
            lgSortKey = 1
        End If
        Exit Sub
    End If
    gMouseClickStatus = "SPC"

	If Row < 1 Then Exit Sub'

	IscookieSplit = ""
	
	Dim ii

     frm1.vspdData.Col = C_MVMTNO
     frm1.vspdData.Row = Row'
	 IscookieSplit = IscookieSplit & Trim(frm1.vspdData.text) & PopupParent.gRowSep
	 Call SetPopupMenuItemInf("0000111111")  
End Sub

'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub

'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    
    Call GetSpreadColumnPos("A")
End Sub

'=======================================================================================================
'   Event Name : vspdData_MouseDown
'   Event Desc : 
'=======================================================================================================
Sub vspdData_MouseDown(Button, Shift, X, Y)
	If Button = 2 And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If
End Sub

'========================================================================================================
'   Event Name : vspdData_LeaveCell
'   Event Desc :
'========================================================================================================
Sub vspdData_LeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
	With frm1.vspdData
		If Row >= NewRow Then
			Exit Sub
		End If

		If NewRow = .MaxRows Then
			If lgStrPrevKeyIndex <> "" Then							'⊙: 다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
				'DbQuery
				If DbQuery = False Then
					Call RestoreToolBar()
					Exit Sub
				End If
			End If
		End If
	End With
End Sub


'#########################################################################################################
'												4. Common Function부 
'	기능: Common Function
'	설명: 환율처리함수, VAT 처리 함수 
'#########################################################################################################
'#########################################################################################################
'												5. Interface부 
'	기능: Interface
'	설명: 각각의 Toolbar에 대한 처리를 행한다. 
'	      Toolbar의 위치순서대로 기술하는 것으로 한다. 
'	<< 공통변수 정의 부분 >>
' 	공통변수 : Global Variables는 아니지만 각각의 Sub나 Function에서 자주 사용하는 변수로 변수명은 
'				통일하도록 한다.
' 	1. 공통컨트롤을 Call하는 변수 
'    	   ADF (ADS, ADC, ADF는 그대로 사용)
'    	   - ADF는 Set하고 사용한 뒤 바로 Nothing 하도록 한다.
' 	2. 공통컨트롤에서 Return된 값을 받는 변수 
'    		strRetMsg
'#########################################################################################################
'*******************************  5.1 Toolbar(Main)에서 호출되는 Function *******************************
'	설명 : Fnc함수명 으로 시작하는 모든 Function
'********************************************************************************************************* %>
Function FncQuery() 
    
    FncQuery = False                                                        '⊙: Processing is NG
    
    Err.Clear                                                               '☜: Protect system from crashing
	
	'** ValidDateCheck(pObjFromDt, pObjToDt) : 'pObjToDt'이 'pObjFromDt'보다 크거나 같아야 할때 **
	If ValidDateCheck(frm1.txtFrRcptDt, frm1.txtToRcptDt) = False Then Exit Function
	
    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")	         						'⊙: Clear Contents  Field
    Call InitVariables 														'⊙: Initializes local global variables
    
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then								'⊙: This function check indispensable field
       Exit Function
    End If

    '-----------------------
    'Query function call area
    '-----------------------	
	If DbQuery = False Then Exit Function									

    FncQuery = True		
    
End Function

'========================================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================================

Function DbQuery() 

	Err.Clear														'☜: Protect system from crashing
	DbQuery = False													'⊙: Processing is NG
	
	If LayerShowHide(1) = False Then
		Exit Function
	End If
	
	Dim strVal
    
    With frm1
		
		If lgIntFlgMode = PopupParent.OPMD_UMODE Then
		   
		    strVal = BIZ_PGM_QRY_ID & "?txtMode=" & PopupParent.UID_M0001		    
		    strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		    strVal = strVal & "&txtMvmtType=" & Trim(frm1.hdnMvmtType.value)
		    strVal = strVal & "&txtSupplier=" & Trim(frm1.hdnSupplier.value)
			strVal = strVal & "&txtFrRcptDt=" & Trim(frm1.hdnFrRcptDt.value)
			strVal = strVal & "&txtToRcptDt=" & Trim(frm1.hdnToRcptDt.value)
		    strVal = strVal & "&txtGroup=" & Trim(frm1.hdnGroup.value)
		    strVal = strVal & "&txtInspFlag=" & Trim(frm1.hdnInspFlag.value)		
		
		else
			strVal = BIZ_PGM_QRY_ID & "?txtMode=" & PopupParent.UID_M0001
		    strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		    strVal = strVal & "&txtMvmtType=" & Trim(frm1.txtMvmtType.value)
		    strVal = strVal & "&txtSupplier=" & Trim(frm1.txtSupplierCd.value)
			strVal = strVal & "&txtFrRcptDt=" & Trim(frm1.txtFrRcptDt.text)
			strVal = strVal & "&txtToRcptDt=" & Trim(frm1.txtToRcptDt.text)
		    strVal = strVal & "&txtGroup=" & Trim(frm1.txtGroupCd.Value)
		    strVal = strVal & "&txtInspFlag=" & frm1.hdnInspFlag.value	
		
		End if
		strVal = strVal & "&lgPageNo="		 & lgPageNo						'☜: Next key tag 
        strVal = strVal & "&lgMaxCount="     & C_SHEETMAXROWS_D             '☜: 한번에 가져올수 있는 데이타 건수  
        
		        
        'strVal = strVal & "&lgPageNo="		 & lgPageNo						'☜: Next key tag 
        'strVal = strVal & "&lgMaxCount="     & C_SHEETMAXROWS_D             '☜: 한번에 가져올수 있는 데이타 건수  
		'strVal = strVal & "&lgSelectListDT=" & lgSelectListDT
		
        'strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList(UBound(gFieldNM),lgPopUpR,gFieldCD,gNextSeq,gTypeCD(0),C_MaxSelList)
		'strVal = strVal & "&lgSelectList="   & EnCoding(lgSelectList)

        Call RunMyBizASP(MyBizASP, strVal)		    						'☜: 비지니스 ASP 를 가동 
        
    End With
    
    DbQuery = True    

End Function

'=========================================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'=========================================================================================================
Function DbQueryOk()	    												'☆: 조회 성공후 실행로직 

	lgIntFlgMode = PopupParent.OPMD_UMODE
	
	If frm1.vspdData.MaxRows > 0 Then
		frm1.vspdData.Focus
		frm1.vspdData.Row = 1	
		frm1.vspdData.SelModeSelected = True		
	Else
		frm1.vspdData.focus
	End If

End Function

</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
</HEAD>
<%
'#########################################################################################################
'       					6. Tag부 
'#########################################################################################################
 %>
<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_20%>>
	<TR>
		<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
	</TR>
	<TR>
		<TD HEIGHT=20 WIDTH=100%>
			<FIELDSET CLASS="CLSFLD">
				<TABLE <%=LR_SPACE_TYPE_40%>>
					<TR>
						<TD CLASS="TD5" NOWRAP>입고형태</TD>
						<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT Alt="입고형태" NAME="txtMvmtType" SIZE=10 MAXLENGTH=5 tag="11NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnMoveType" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenConSItemDC 3">
											   <INPUT TYPE=TEXT Alt="입고형태" NAME="txtMvmtTypeNm" SIZE=20 tag="14X"></TD>
						<TD CLASS="TD5" NOWRAP>입고일</TD>
						<TD CLASS="TD6" NOWRAP>
							<table cellspacing=0 cellpadding=0>
								<tr NOWRAP>
									<td NOWRAP>
										<script language =javascript src='./js/m9211pa1_fpDateTime1_txtFrRcptDt.js'></script>
									</td>
									<td NOWRAP>~</td>
									<td NOWRAP>
										<script language =javascript src='./js/m9211pa1_fpDateTime1_txtToRcptDt.js'></script>
									</td>
								<tr>
							</table>
						</TD>
					</TR>	
					<TR>	
						<TD CLASS="TD5" NOWRAP>출고공장</TD>
						<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT AlT="출고공장" NAME="txtSupplierCd" SIZE=10 MAXLENGTH=10 tag="11NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnGroup1Cd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenConSItemDC 1">
					   			 	     	   <INPUT TYPE=TEXT AlT="출고공장명" ID="txtSupplierNm" NAME="arrCond" tag="14X"></TD>
						<TD CLASS="TD5" NOWRAP>구매그룹</TD>
						<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT AlT="구매그룹" NAME="txtGroupCd" SIZE=10 MAXLENGTH=4 tag="11NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnGroup1Cd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenConSItemDC 2">
										 	   <INPUT TYPE=TEXT AlT="구매그룹" ID="txtGroupNm" NAME="arrCond" tag="14X"></TD>
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
				<TR>
					<TD HEIGHT=100% WIDTH=100% COLSPAN=4>					
						<script language =javascript src='./js/m9211pa1_vaSpread1_vspdData.js'></script>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
    <tr>
		<TD <%=HEIGHT_TYPE_01%>></TD>
    </tr>
	<TR HEIGHT="20">
		<TD WIDTH="100%">
			<TABLE <%=LR_SPACE_TYPE_30%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD WIDTH=70% NOWRAP>     <IMG SRC="../../../CShared/image/query_d.gif"  Style="CURSOR: hand" ALT="Search" NAME="Search" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Query.gif',1)"  ONCLICK="FncQuery()"   ></IMG></TD>
					<TD WIDTH=30% ALIGN=RIGHT><IMG SRC="../../../CShared/image/ok_d.gif"     Style="CURSOR: hand" ALT="OK"     NAME="pop1"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/OK.gif',1)"     ONCLICK="OkClick()"    ></IMG>
							                  <IMG SRC="../../../CShared/image/cancel_d.gif" Style="CURSOR: hand" ALT="CANCEL" NAME="pop2"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)" ONCLICK="CancelClick()"></IMG>					</TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=No noresize framespacing=0 TABINDEX=-1></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="hdnInspFlag" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnMvmtType" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnSupplier" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnFrRcptDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnToRcptDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnGroup" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT" STYLE="visible:true">
	<IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm"></IFRAME>
</DIV>
</BODY>
</HTML>
