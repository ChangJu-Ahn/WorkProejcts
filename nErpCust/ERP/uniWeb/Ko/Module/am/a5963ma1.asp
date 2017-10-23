<%@ LANGUAGE="VBSCRIPT" %>
<!--'======================================================================================================
'*  1. Module Name          : Accounting
'*  2. Function Name        : 
'*  3. Program ID           : A5963MA1
'*  4. Program Name         : 퇴직급여 추계액 등록 
'*  5. Program Desc         : 회계관리 / 월차계산 / 
'*  6. Component List       :
'*  7. Modified date(First) : 2002/01/15
'*  8. Modified date(Last)  : 2003/05/30
'*  9. Modifier (First)     : song sang min
'* 10. Modifier (Last)      : Jung Sung Ki
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'=======================================================================================================-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>


<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliRdsQuery.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs">   </SCRIPT>

<Script Language="VBScript">
Option Explicit                                                        '☜: indicates that All variables must be declared in advance

'========================================================================================================

Const BIZ_PGM_ID      = "a5963mb1.asp"						           '☆: Biz Logic ASP Name

'========================================================================================================

Dim C_DEPT_CD 	
Dim C_DEPT_CD_PB
Dim C_DEPT_CD_NM
'hidden--------------------------------
Dim c_org_change_id
Dim C_INTERNAL_CD	

Dim C_ACCT_CODE 	
'hidden--------------------------------
Dim C_ACCT_NM       
Dim C_AMOUNT1		
Dim C_AMOUNT2   	
Dim c_AMOUNT3      	
Dim c_AMOUNT4      	
Dim C_biz_area_cd   


Const COOKIE_SPLIT       = 4877	                                      'Cookie Split String

'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->
'========================================================================================================
Dim lgIsOpenPop
Dim IsOpenPop




'========================================================================================================
Sub InitSpreadPosVariables()
	 C_DEPT_CD 		= 1		
	 C_DEPT_CD_PB   	= 2		
	 C_DEPT_CD_NM   	= 3		
	'hidden--------------------------------
	 c_org_change_id   = 4
	 C_INTERNAL_CD		= 5

	 C_ACCT_CODE 		= 6	
	'hidden--------------------------------
	 C_ACCT_NM         = 7	
	 C_AMOUNT1			= 8	
	 C_AMOUNT2   		= 9	
	 c_AMOUNT3      	= 10	
	 c_AMOUNT4      	= 11	
	 C_biz_area_cd     = 12


End Sub

'========================================================================================================
Sub InitVariables()
	lgIntFlgMode      = Parent.OPMD_CMODE						        '⊙: Indicates that current mode is Create mode
	lgBlnFlgChgValue  = False								    '⊙: Indicates that no value changed
	lgIntGrpCount     = 0										'⊙: Initializes Group View Size
   	lgStrPrevKey      = ""                                      '⊙: initializes Previous Key
   	lgStrPrevKeyIndex = ""                                      '⊙: initializes Previous Key Index
   	lgSortKey         = 1                                       '⊙: initializes sort direction

	'------ Developer Coding part (Start ) --------------------------------------------------------------
	lgStrPrevKey = ""                                           'initializes Previous Key
   	lgLngCurRows = 0                                            'initializes Deleted Rows Count
	'------ Developer Coding part (End )   --------------------------------------------------------------
End Sub

'========================================================================================================
Sub SetDefaultVal()


	Dim StartDate
	Dim strYear, strMonth, strDay

	StartDate	= "<%=GetSvrDate%>"                           'Get Server DB Date

	'------ Developer Coding part (Start ) --------------------------------------------------------------
	Call ExtractDateFrom(StartDate,Parent.gServerDateFormat,Parent.gServerDateType,strYear,strMonth,strDay)
	frm1.fpdtWk_yymm.Text	= UniConvDateAToB(StartDate ,parent.gServerDateFormat,parent.gDateFormat) 
	Call ggoOper.FormatDate(frm1.fpdtWk_yymm, Parent.gDateFormat,2)
	frm1.fpdtWk_yymm.focus
	'------ Developer Coding part (End )   --------------------------------------------------------------
End Sub

'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "*","NOCOOKIE", "MA") %>
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
    
    Dim strYYYYMM
    Dim strYear,strMonth,strDay
    lgKeyStream = ""
    '------ Developer Coding part (Start ) --------------------------------------------------------------
    Call ExtractDateFrom(frm1.fpdtWk_yymm.Text,frm1.fpdtWk_yymm.UserDefinedFormat,Parent.gComDateType,strYear,strMonth,strDay)
    
    strYYYYMM =   strYear & strMonth
    lgKeyStream = lgKeyStream & Trim(strYYYYMM)    & Parent.gColSep '월차코드 
    lgKeyStream = lgKeyStream & Trim(Frm1.txtCurrencyCode.Value) & Parent.gColSep '사업장코드 
    
End Sub

'========================================================================================================
' Name : InitComboBox()
' Desc : Set ComboBox
'========================================================================================================
Sub InitComboBox()
    Dim iCodeArr
	Dim iNameArr
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0071", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    iCodeArr = lgF0
    iNameArr = lgF1
    ggoSpread.source = frm1.vspdData
    ggoSpread.SetCombo Replace(iNameArr,Chr(11),vbTab), C_ACCT_NM
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
	
End Sub

'======================================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : 컬럼을 클릭할 경우 발생하는 콤보 박스 
'=======================================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)

	Dim BizAreaCd
    BizAreaCd = Trim(frm1.txtCurrencyCode.value)
	Select Case Col				'추가부분을 위해..select로..
	    Case C_DEPT_CD_PB        'Cost center
	        frm1.vspdData.Col = C_DEPT_CD
	        If BizAreaCd = "" then
	            Call OpenDept(frm1.vspdData.Text,1, Row)
	        Else
	            Call OpenDept(frm1.vspdData.Text,2, Row )
	        End If
	End Select
	Call SetActiveCell(frm1.vspdData,Col - 1,frm1.vspdData.ActiveRow ,"M","X","X")
End Sub


'======================================================================================================
'	Name : SetMajor()
'	Description : Item Popup에서 Return되는 값 setting
'=======================================================================================================

Function OpenCost(strCode, iWhere, Row)

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)


	If IsOpenPop = True Then Exit Function
    

	IsOpenPop = True

	Select Case iWhere
	    Case 1
	    	arrParam(0) = "부서 팝업"						' 팝업 명칭 
	    	arrParam(1) = "B_ACCT_DEPT A, B_COST_CENTER B"		' TABLE 명칭 
	    	arrParam(2) = Trim(strCode) 		' Code Condition
	    	arrParam(3) = "" 			' Name Condition
	    	arrParam(4) = "A.COST_CD = B.COST_CD AND A.org_change_id =(select max(org_change_id) from b_acct_dept) and B.biz_area_cd = " & FilterVar(Frm1.txtCurrencyCode.value, "''", "S")  		   ' Where Condition
	    	arrParam(5) = "부서코드"		' TextBox 명칭 

	    	arrField(0) = "A.DEPT_cd"		 	    ' Field명(0)
	    	arrField(1) = "A.DEPT_nm"    		    ' Field명(1)%>
			arrField(2) = "A.org_change_id"    		    ' Field명(1)%>
			arrField(3) = "B.biz_AREA_cd"
			
			arrHeader(0) = "부서 코드"	' Header명(0)%>
	    	arrHeader(1) = "부서 명"	' Header명(0)%>
	    	arrHeader(2) = "조직변경 아이디"	' Header명(1)%>
			arrHeader(3) = "사업장코드"	' Header명(1)%>
			
	    Case 2
		   arrParam(0) = "계정타입 팝업"			' 팝업 명칭 
	       arrParam(1) = "B_MAJOR A, B_MINOR B"	            <%' TABLE 명칭 %>
	       arrParam(2) = Trim(strCode)	     	    <%' Code Condition%>
	       arrParam(3) = "" 		                <%' Name Cindition%>
	       arrParam(4) = "A.MAJOR_CD =B.MAJOR_cD AND  A.MAJOR_cD = " & FilterVar("H0071", "''", "S") & " "              <%' Where Condition%>
	       arrParam(5) = "계정타입"

           arrField(0) = "B.MINOR_CD"	     	  	<%' Field명(1)%>
           arrField(1) = "B.MINOR_NM"			    <%' Field명(0)%>

           arrHeader(0) = "계정타입"	  		    <%' Header명(0)%>
           arrHeader(1) = "계정타입명"		  	    <%' Header명(1)%>

	End Select

        	
  	        

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
    "dialogWidth=450px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

  	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetCost(arrRet, iWhere, Row)
	End If

End Function


'------------------------------------------  SetCode()  --------------------------------------------------
'	Name : SetCode()
'	Description : OpenCode Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------
Function SetCost(arrRet, iWhere, Row)

	With frm1
        .vspdData.Row = Row
		Select Case iWhere
		    Case 1
		        .vspdData.Col = C_DEPT_CD
		    	.vspdData.text = arrRet(0)
		    	.vspdData.Col = C_DEPT_CD_NM
		    	.vspdData.text = arrRet(1)
		    	.vspdData.Col = c_org_change_id
		    	.vspdData.text = arrRet(2)
		    	.vspdData.Col = C_biz_area_cd
		    	.vspdData.text = arrRet(3)
		    Case 2
			    .vspdData.Col = C_ACCT_CODE
		    	.vspdData.text = arrRet(0)
		    	.vspdData.Col = C_ACCT_NM
		    	.vspdData.text = arrRet(1)
		End Select

		lgBlnFlgChgValue = True

	End With
	ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row

End Function


'------------------------------------------  OpenDept()  --------------------------------------------------
'	Name : OpenDept()
'	Description : OpenCode Popup
'---------------------------------------------------------------------------------------------------------
Function OpenDept(Byval strCode, iWhere, Row)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim strDate
	'------ Developer Coding part (Start ) --------------------------------------------------------------
	'------ Developer Coding part (Start ) --------------------------------------------------------------
	
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = strCode									'  Code Condition
	
	strDate = FilterVar(frm1.fpdtWk_yymm.year,"2999","SNM") & "-" & Right("0" & FilterVar(frm1.fpdtWk_yymm.Month,"12","SNM"),2) & "-" & "01"
	strDate = DateAdd("D",-1, DateAdd("M",1,cdate(strDate)))
   	
   	arrParam(1) = UNIDateClientFormat(strDate)
	arrParam(2) = lgUsrIntCd								' 자료권한 Condition  

	'' T : protected F: 필수 
	'If lgIntFlgMode = Parent.OPMD_UMODE then
	arrParam(3) = "T"									' 결의일자 상태 Condition  
	'Else
	'	arrParam(3) = "F"									' 결의일자 상태 Condition  
	'End If
	
	arrParam(4) = iWhere
	arrParam(5) = Trim(frm1.txtCurrencyCode.value)
	
	
	arrRet = window.showModalDialog("../../comasp/DeptPopupDt3.asp", Array(window.parent ,arrParam), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False
	
	If arrRet(0) = "" Then
'		Call GridSetFocus(iWhere)
		Exit Function
	Else
		Call SetDept(arrRet, iWhere, Row)
	End If	
			
End Function

'------------------------------------------  SetDept()  ------------------------------------------------
'	Name : SetDept()
'	Description : Dept Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------
Function SetDept(Byval arrRet, Byval iWhere, Byval iRow)
		
	With frm1
		.vspdData.Row = iRow
		Select Case iWhere
		     Case 1,2
                .vspdData.Col = C_DEPT_CD
				.vspdData.text = arrRet(0)
				.vspdData.Col = C_DEPT_CD_NM
				.vspdData.text = arrRet(1)
				.vspdData.Col = C_BIZ_AREA_CD
				.vspdData.text = arrRet(2)
				.vspdData.Col = C_ORG_CHANGE_ID
				.vspdData.text = arrRet(3)
				.vspdData.Col = C_INTERNAL_CD	
				.vspdData.text = arrRet(4)
        End Select
	    Call vspdData_Change(C_DEPT_CD,iRow)
'		Call GridSetFocus(iWhere)
	End With
End Function  



'========================================================================================================
Sub InitData()
End Sub

'========================================================================================================
Sub InitSpreadSheet()

	Call initSpreadPosVariables()      '1.3 [initSpreadPosVariables] 호출 Logic 추가 
	
	ggoSpread.Source = frm1.vspdData
	ggoSpread.Spreadinit "V20021103",,Parent.gAllowDragDropSpread            '2.1 [Spreadinit]에 Source Version No. 와 DragDrop지원여부를 설정 
	
	With frm1.vspdData
	
		.ReDraw = false
	
    	.MaxCols   = C_biz_area_cd + 1                                                  ' ☜:☜: Add 1 to Maxcols
	    .Col       = .MaxCols                                                        ' ☜:☜: Hide maxcols
        .ColHidden = True           
       
		ggoSpread.Source= frm1.vspdData
		ggoSpread.ClearSpreadData

        call GetSpreadColumnPos("A")
 		Call AppendNumberPlace("6","15","2")
        
        ggoSpread.SSSetEdit     C_DEPT_CD        ,     "부서 코드"  ,10,,,5,2
        ggoSpread.SSSetButton   C_DEPT_CD_PB      
        ggoSpread.SSSetEdit     C_DEPT_CD_NM     ,     "부서명"   	 ,18,,,20,2
        ggoSpread.SSSetEdit     c_org_change_id  ,	"조직변경ID"  ,10,,,10,2
        ggoSpread.SSSetEdit     C_INTERNAL_CD	,	"내부부서코드"  ,10,,,10,2
        ggoSpread.SSSetEdit     C_ACCT_CODE    ,     "계정타입코드"   	 ,5,,,10,2
        ggoSpread.SSSetCombo    C_ACCT_NM      ,     "계정타입"    ,12,2
        ggoSpread.SSSetFloat    C_AMOUNT1      ,     "예상액"     ,15,6,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
        ggoSpread.SSSetFloat    C_AMOUNT2      ,     "전월누계액"     ,18,6,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
        ggoSpread.SSSetFloat    C_AMOUNT3      ,     "당월지급액"     ,18,6,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
        ggoSpread.SSSetFloat    C_AMOUNT4      ,     "당월반영액"     ,18,6,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
        ggoSpread.SSSetEdit     C_biz_area_cd    ,    "사업장"   	 ,10,,,10,2
        
        Call ggoSpread.SSSetColHidden(C_ACCT_CODE,C_ACCT_CODE,True)
        Call ggoSpread.SSSetColHidden(c_org_change_id,c_org_change_id,True)
        Call ggoSpread.SSSetColHidden(C_INTERNAL_CD,C_INTERNAL_CD,True)
        Call ggoSpread.SSSetColHidden(C_biz_area_cd,C_biz_area_cd,True)
        call ggoSpread.MakePairsColumn(C_DEPT_CD,C_DEPT_CD_NM)	
        call ggoSpread.MakePairsColumn(C_ACCT_CODE,C_ACCT_NM)	
        
	   .ReDraw = true

       Call SetSpreadLock
    End With
End Sub

'======================================================================================================
' Name : SetSpreadLock
' Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadLock()
    With frm1
                     .vspdData.ReDraw = False
                      ggoSpread.SpreadLock      C_DEPT_CD, -1,  C_DEPT_CD
                      ggoSpread.SpreadLock      C_DEPT_CD_NM, -1,  C_DEPT_CD_NM
                      ggoSpread.SpreadLock      C_DEPT_CD_PB, -1,  C_DEPT_CD_PB
                      ggoSpread.SpreadLock      C_ACCT_NM, -1,  C_ACCT_NM
                      ggoSpread.SSSetRequired	C_AMOUNT1, -1, C_AMOUNT1
     				  ggoSpread.SpreadLock      C_AMOUNT4, -1, C_AMOUNT4
     				  ggoSpread.SpreadLock      C_AMOUNT3, -1, C_AMOUNT3
     				  ggoSpread.SpreadLock      C_AMOUNT2, -1, C_AMOUNT2
     				  ggoSpread.SpreadLock	.vspdData.MaxCols, -1,.vspdData.MaxCols
     			      .vspdData.ReDraw = True
                  End With
End Sub

'======================================================================================================
'	Name : OpenCurrency()
'	Description : Major PopUp
'======================================================================================================
Function OpenCurrency()
	Dim arrRet
	Dim arrParam(6), arrField(5), arrHeader(5)


	If IsOpenPop = True Then Exit Function

	IsOpenPop = True


	arrParam(0) = "사업장 팝업"		    	    <%' 팝업 명칭 %>
	arrParam(1) = "B_BIZ_AREA"					 	<%' TABLE 명칭 %>
	arrParam(2) = frm1.txtCurrencyCode.Value	    	<%' Code Cindition%>
	arrParam(3) = ""								<%' Name Condition%>
	arrParam(4) = ""
    arrParam(5) = "사업장"

    arrField(0) = "BIZ_AREA_CD"					<%' Field명(0)%>
    arrField(1) = "BIZ_AREA_NM"	     			<%' Field명(1)%>

    arrHeader(0) = "사업장 코드"				<%' Header명(0)%>
    arrHeader(1) = "사업장명"				<%' Header명(1)%>


	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtCurrencyCode.focus
		Exit Function
	Else
		Call SetMajor(arrRet)

	End If

End Function
'======================================================================================================
'	Name : SetMajor()
'	Description : Item Popup에서 Return되는 값 setting
'=======================================================================================================
Function SetMajor(Byval arrRet)
	With frm1
		.txtCurrencyCode.focus
		.txtCurrencyCode.value = arrRet(0)
		.txtCurrency.value = arrRet(1)
	End With

End Function

'======================================================================================================
Sub  SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)

	With Frm1
		.vspdData.ReDraw = False
		ggoSpread.SSSetRequired		C_DEPT_CD, pvStartRow, pvEndRow
		ggoSpread.SSSetRequired		C_ACCT_NM, pvStartRow, pvEndRow
		ggoSpread.SSSetRequired		C_AMOUNT1  ,pvStartRow, pvEndRow
		ggoSpread.SpreadLock		C_DEPT_CD_NM,pvStartRow,C_DEPT_CD_NM
		ggoSpread.SpreadLock      	C_AMOUNT4, pvStartRow,C_AMOUNT4      
		ggoSpread.SpreadLock		C_AMOUNT3, pvStartRow,C_AMOUNT3
		ggoSpread.SpreadLock		C_AMOUNT2, pvStartRow,C_AMOUNT2   
		.vspdData.ReDraw = True
	End With
End Sub

'======================================================================================================
' Name : SubSetErrPos
' Desc : This method set focus to pos of err
'======================================================================================================
Sub SubSetErrPos(iPosArr)
    Dim iDx
    Dim iRow
    iPosArr = Split(iPosArr,Parent.gColSep)
    If IsNumeric(iPosArr(0)) Then
       iRow = CInt(iPosArr(0))
       For iDx = 1 To  frm1.vspdData.MaxCols - 1
           Frm1.vspdData.Col = iDx
           Frm1.vspdData.Row = iRow
           If Frm1.vspdData.ColHidden <> True And Frm1.vspdData.BackColor <> Parent.UC_PROTECTED Then
              Frm1.vspdData.Col = iDx
              Frm1.vspdData.Row = iRow
              Frm1.vspdData.Action = 0
              Exit For
           End If

       Next

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
			C_DEPT_CD          = iCurColumnPos(1)
			C_DEPT_CD_PB          = iCurColumnPos(2)
			C_DEPT_CD_NM       = iCurColumnPos(3)    
			c_org_change_id        = iCurColumnPos(4)
			C_INTERNAL_CD      = iCurColumnPos(5)
			C_ACCT_CODE = iCurColumnPos(6)
			C_ACCT_NM    = iCurColumnPos(7)
			C_AMOUNT1 = iCurColumnPos(8)
			C_AMOUNT2 = iCurColumnPos(9)
			c_AMOUNT3     = iCurColumnPos(10)
			c_AMOUNT4  = iCurColumnPos(11)
			C_biz_area_cd       = iCurColumnPos(12)
			
    End Select    
    
End Sub

'========================================================================================================
Sub Form_Load()
    Err.Clear
	Call LoadInfTB19029
    Call ggoOper.LockField(Document, "N")
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
	Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
	
    Call InitSpreadSheet                                                    <%'Setup the Spread sheet%>
    Call InitVariables                                                      <%'Initializes local global variables%>

    Call SetDefaultVal
    Call InitComboBox
    if frm1.vspdData.MaxRows = 0 then
       Call SetToolbar("1100110100101111")									<%'버튼 툴바 제어 %>
    else
       Call SetToolbar("1100111100111111")									<%'버튼 툴바 제어 %>
    end if
    Call CookiePage(0)

    '------ Developer Coding part (End )   --------------------------------------------------------------
End Sub

'========================================================================================================
Sub Form_QueryUnload(Cancel, UnloadMode)

End Sub

'========================================================================================================
Function FncQuery()
    Dim IntRetCD

    FncQuery = False															  '☜: Processing is NG
    Err.Clear                                                                     '☜: Clear err status

    ggoSpread.Source = Frm1.vspdData
    If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO,"X","X")			          '☜: "Will you destory previous data"
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

	
	If txtBizAreaCdChange = False  Then Exit Function
	
    Call ggoOper.ClearField(Document, "2")										  '⊙: Clear Contents  Field
'    Call SetDefaultVal
    Call InitVariables															  '⊙: Initializes local global variables

    If Not chkField(Document, "1") Then									          '⊙: This function check indispensable field
       Exit Function
    End If

	'------ Developer Coding part (Start ) --------------------------------------------------------------
    Call MakeKeyStream("X")

	Call ggoOper.SetReqAttr(Frm1.txtCurrencyCode, "N")
	Call ggoOper.SetReqAttr(Frm1.fpdtWk_yymm, "N")	

   '------ Developer Coding part (End )   --------------------------------------------------------------
    If DbQuery = False Then                                                       '☜: Query db data
       Exit Function
    End If
    
    Set gActiveElement = document.ActiveElement
    FncQuery = True                                                               '☜: Processing is OK
End Function
'========================================================================================================
' Name : FncNew
' Desc : This function is called from MainNew in Common.vbs
'========================================================================================================

Function FncNew()
    Dim IntRetCD

    FncNew = False																  '☜: Processing is NG
    Err.Clear                                                                     '☜: Clear err status

    If lgBlnFlgChgValue = True Then
       IntRetCD = DisplayMsgBox("900015", Parent.VB_YES_NO,"x","x")					  '☜: Data is changed.  Do you want to make it new?
       If IntRetCD = vbNo Then
          Exit Function
       End If
    End If

    Call ggoOper.ClearField(Document, "1")                                        '☜: Clear Condition Field
    Call ggoOper.ClearField(Document, "2")                                        '☜: Clear Contents  Field
    Call ggoOper.LockField(Document , "N")                                        '☜: Lock  Field
	'------ Developer Coding part (Start ) --------------------------------------------------------------

	if frm1.vspdData.MaxRows = 0 then
       Call SetToolbar("1100110100101111")									<%'버튼 툴바 제어 %>
    else
       Call SetToolbar("1100111100111111")									<%'버튼 툴바 제어 %>
    end if
    Call SetDefaultVal
    Call InitVariables

	'------ Developer Coding part (End )   --------------------------------------------------------------
    Set gActiveElement = document.ActiveElement
    FncNew = True															      '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncDelete
' Desc : This function is called from MainDelete in Common.vbs
'========================================================================================================
Function FncDelete()
    Dim intRetCD

    FncDelete = False                                                            '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

    If lgIntFlgMode <> Parent.OPMD_UMODE Then                                           'Check if there is retrived data
        Call DisplayMsgBox("900002","X","X","X")                                 '☜: Please do Display first.
        Exit Function
    End If

    IntRetCD = DisplayMsgBox("900003", Parent.VB_YES_NO,"X","X")		                 '☜: Do you want to delete?
	If IntRetCD = vbNo Then
		Exit Function
	End If

	'------ Developer Coding part (Start ) --------------------------------------------------------------

	'------ Developer Coding part (End )   --------------------------------------------------------------
    If DbDelete = False Then                                                     '☜: Query db data
       Exit Function
    End If

    Set gActiveElement = document.ActiveElement
    FncDelete = True                                                             '☜: Processing is OK
End Function

'========================================================================================================
Function FncSave()
    Dim IntRetCD
    
    FncSave = False                                                              '☜: Processing is NG

    Err.Clear                                                                    '☜: Clear err status

    ggoSpread.Source = frm1.vspdData
    If lgBlnFlgChgValue = False And ggoSpread.SSCheckChange = False Then
        IntRetCD = DisplayMsgBox("900001","X","X","X")                           '⊙: No data changed!!
        Exit Function
    End If

    If Not chkField(Document, "2") Then
       Exit Function
    End If

	ggoSpread.Source = frm1.vspdData
    If Not ggoSpread.SSDefaultCheck Then                                         '⊙: Check contents area
       Exit Function
    End If

	'------ Developer Coding part (Start ) --------------------------------------------------------------
    
     Call MakeKeyStream("X")
    
	'------ Developer Coding part (End )   --------------------------------------------------------------
    If DbSave = False Then                                                       '☜: Query db data
       Call LayerShowHide(0)
       Exit Function
    End If
    
    Set gActiveElement = document.ActiveElement
    FncSave = True                                                               '☜: Processing is OK
End Function

'========================================================================================================
Function FncCopy()
	Dim IntRetCD

    FncCopy = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

    If Frm1.vspdData.MaxRows < 1 Then
       Exit Function
    End If

	With Frm1

		If .vspdData.ActiveRow > 0 Then
			.vspdData.ReDraw = False

			ggoSpread.Source = frm1.vspdData
			ggoSpread.CopyRow
			SetSpreadColor frm1.vspdData.ActiveRow, frm1.vspdData.ActiveRow
	'------ Developer Coding part (Start ) --------------------------------------------------------------
	' Clear key field
	'----------------------------------------------------------------------------------------------------
			frm1.vspdData.Row = frm1.vspdData.ActiveRow
			frm1.vspdData.col = C_DEPT_CD 
			frm1.vspdData.text = ""
			frm1.vspdData.Row = frm1.vspdData.ActiveRow
			frm1.vspdData.col = C_DEPT_CD_NM 
			frm1.vspdData.text = ""
			
	'------ Developer Coding part (End )   --------------------------------------------------------------

			.vspdData.ReDraw = True
			.vspdData.focus
		End If
	End With
    Set gActiveElement = document.ActiveElement
    FncCopy = True                                                                '☜: Processing is OK
End Function

'========================================================================================================
Function FncCancel()
    FncCancel = False															 '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

    ggoSpread.Source = frm1.vspdData
    ggoSpread.EditUndo
	'------ Developer Coding part (Start ) --------------------------------------------------------------
	'------ Developer Coding part (End )   --------------------------------------------------------------
    Set gActiveElement = document.ActiveElement
    FncCancel = True                                                             '☜: Processing is OK
End Function

'========================================================================================================
Function  FncInsertRow(ByVal pvRowCnt) 
	Dim imRow
	Dim ii
	Dim iCurRowPos
	  Dim lRow, IntRetCD
    
    On Error Resume Next															'☜: If process fails
    Err.Clear																		'☜: Clear error stat	

    
	FncInsertRow = False															'☜: Processing is NG
	
	If IsNumeric(Trim(pvRowCnt)) Then
		imRow = Cint(pvRowCnt)
	Else
	    imRow = AskSpdSheetAddRowCount()
    
		If imRow = "" Then
		    Exit Function
		End If
	End If		


                                                             '☜: Clear err status
    IF Trim(Frm1.txtCurrencyCode.value) = "" THEN				'☜: 사업장 선택 
	IntRetCD = DisplayMsgBox("169803","X","X","X")                           		'⊙: No data changed!!
	Frm1.txtCurrencyCode.focus
	Set gActiveElement = document.ActiveElement
        Exit Function
    END if
    With frm1
        .vspdData.ReDraw = False
        .vspdData.focus
        ggoSpread.Source = .vspdData
        ggoSpread.InsertRow .vspdData.ActiveRow, imRow
        SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow + imrow -1
        
       .vspdData.ReDraw = True
    End With
	'------ Developer Coding part (Start ) --------------------------------------------------------------
	
	'------ Developer Coding part (End )   --------------------------------------------------------------
    Set gActiveElement = document.ActiveElement
    FncInsertRow = True                                                          '☜: Processing is OK
	'Call ggoOper.SetReqAttr(Frm1.txtCurrencyCode, "Q")
	'Call ggoOper.SetReqAttr(Frm1.fpdtWk_yymm, "Q")
End Function

'========================================================================================================
Function FncDeleteRow()
    Dim lDelRows

    FncDeleteRow = False														 '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

    If Frm1.vspdData.MaxRows < 1 then
       Exit function
	End if
    With Frm1.vspdData
    	.focus
    	ggoSpread.Source = frm1.vspdData
    	lDelRows = ggoSpread.DeleteRow
    End With
	'------ Developer Coding part (Start ) --------------------------------------------------------------
	'------ Developer Coding part (End )   --------------------------------------------------------------
    Set gActiveElement = document.ActiveElement
    FncDeleteRow = True                                                          '☜: Processing is OK
End Function

'========================================================================================================
Function FncPrint()
    FncPrint = False	                                                         '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
	'------ Developer Coding part (Start ) --------------------------------------------------------------
	'------ Developer Coding part (End )   --------------------------------------------------------------
	Call Parent.FncPrint()                                                       '☜: Protect system from crashing
    FncPrint = True	                                                             '☜: Processing is OK
End Function


'========================================================================================================
Function FncExcel()
    FncExcel = False                                                             '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

	Call Parent.FncExport(Parent.C_MULTI)

    FncExcel = True                                                              '☜: Processing is OK
End Function

'========================================================================================================
Function FncFind()
    FncFind = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

	Call Parent.FncFind(Parent.C_MULTI, True)

    FncFind = True                                                               '☜: Processing is OK
End Function

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

    FncExit = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

	ggoSpread.Source = frm1.vspdData
    If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO,"X","X")			         '⊙: Data is changed.  Do you want to exit?
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    FncExit = True                                                               '☜: Processing is OK
End Function

'========================================================================================================
Function DbQuery()
    Dim strVal
    Err.Clear                                                                    '☜: Clear err status

    DbQuery = False                                                              '☜: Processing is NG

    if LayerShowHide(1) = false then
	    Exit Function
	end if                                                       '☜: Show Processing Message

    strVal = BIZ_PGM_ID & "?txtMode="            & Parent.UID_M0001                     '☜: Query
    strVal = strVal     & "&txtKeyStream="       & lgKeyStream                   '☜: Query Key
    strVal = strVal     & "&lgStrPrevKeyIndex=" & lgStrPrevKeyIndex             '☜: Next key tag
    strVal = strVal     & "&txtMaxRows="         & Frm1.vspdData.MaxRows         '☜: Max fetched data
	'------ Developer Coding part (Start)  --------------------------------------------------------------

	'------ Developer Coding part (End )   --------------------------------------------------------------
    Call RunMyBizASP(MyBizASP, strVal)                                           '☜:  Run biz logic

    DbQuery = True                                                               '☜: Processing is OK

    Set gActiveElement = document.ActiveElement
End Function
'========================================================================================================
Function DbSave()
    Dim lRow        
    Dim lGrpCnt     
	Dim strVal, strDel
	Dim pP21011
    Dim retVal
    Dim boolCheck
    Dim lStartRow
    Dim lEndRow
    Dim lRestGrpCnt
    DIm IntRetCD
    Dim strYYYYMM
    Dim strYear,strMonth,strDay
 
	
    Err.Clear                                                                      '☜: Clear err status
    DbSave = False                                                                 '☜: Processing is NG
	
	if LayerShowHide(1) = false then
	    Exit Function
	end if 
	
	'------ Developer Coding part (Start)  -------------------------------------------------------------- 
  	Call ExtractDateFrom(frm1.fpdtWk_yymm.Text,frm1.fpdtWk_yymm.UserDefinedFormat,Parent.gComDateType,strYear,strMonth,strDay)
    strYYYYMM = strYear & strMonth
  	
  	With Frm1
		.txtMode.value      = Parent.UID_M0002                                            '☜: Delete
		.txtFlgMode.value     = lgIntFlgMode
		.txtKeyStream.value = lgKeyStream
	End With

    strVal  = ""
    strDel  = ""
    lGrpCnt = 1

	With Frm1
       For lRow = 1 To .vspdData.MaxRows

           .vspdData.Row = lRow
           .vspdData.Col = 0

            Select Case .vspdData.Text

               Case ggoSpread.InsertFlag                                      '☜: Update
                                                		   		strVal = strVal & "C" & Parent.gColSep '0
                                            		       		strVal = strVal & lRow & Parent.gColSep  '1
                                        		           		strval = strval & strYYYYMM& Parent.gColSep '2
                     .vspdData.Col = C_DEPT_CD    	          : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep '3
                     .vspdData.Col = C_ACCT_CODE       		  : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep '4
                     .vspdData.Col = C_AMOUNT1       		  : strVal = strVal & Trim(.vspdData.text) & Parent.gColSep '5
                     .vspdData.Col = C_AMOUNT2 	      	  	  : strVal = strVal & Trim(.vspdData.text) & Parent.gColSep '6
                     .vspdData.Col = C_AMOUNT3  			  : strVal = strVal & Trim(.vspdData.text) & Parent.gColSep '7
                     .vspdData.Col = c_org_change_id          : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep '8
					 .vspdData.Col = C_INTERNAL_CD            : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep '8
                     .vspdData.Col = C_biz_area_cd            : strVal = strVal & Trim(.vspdData.Text) & Parent.gRowSep '9
                     lGrpCnt = lGrpCnt + 1
               Case ggoSpread.UpdateFlag                                      '☜: Update
                  	                                            strVal = strVal & "U" & Parent.gColSep  '0
                    	                                        strVal = strVal & lRow & Parent.gColSep  '1
                      	                                        strval = strval & strYYYYMM& Parent.gColSep '2
                     .vspdData.Col = C_DEPT_CD    	          : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep '3
                     .vspdData.Col = C_ACCT_CODE       		  : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep '4
                     .vspdData.Col = C_AMOUNT1       		  : strVal = strVal & Trim(.vspdData.text) & Parent.gColSep '5
                     .vspdData.Col = C_AMOUNT2 	      	  	  : strVal = strVal & Trim(.vspdData.text) & Parent.gColSep '6
                     .vspdData.Col = C_AMOUNT3  			  : strVal = strVal & Trim(.vspdData.text) & Parent.gColSep '7
                     .vspdData.Col = C_org_change_id		  : strVal = strVal & Trim(.vspdData.text) & Parent.gColSep '8
                     .vspdData.Col = C_INTERNAL_CD			  : strVal = strVal & Trim(.vspdData.text) & Parent.gRowSep '8
                    lGrpCnt = lGrpCnt + 1
               Case ggoSpread.DeleteFlag                                      '☜: Delete

                                                  				strDel = strDel & "D" & Parent.gColSep  '0
                                                  				strDel = strDel & lRow & Parent.gColSep   '1
                                                  				strDel = strDel & strYYYYMM& Parent.gColSep '2
                     .vspdData.Col = C_DEPT_CD    	          : strDel = strDel & Trim(.vspdData.Text) & Parent.gColSep '3
                     .vspdData.Col = C_ACCT_CODE       		  : strDel = strDel & Trim(.vspdData.Text) & Parent.gColSep '4
                     .vspdData.Col = C_AMOUNT1       		  : strDel = strDel & Trim(.vspdData.text) & Parent.gColSep '5
                     .vspdData.Col = C_AMOUNT2 	      	  	  : strDel = strDel & Trim(.vspdData.text) & Parent.gColSep '6
                     .vspdData.Col = C_AMOUNT3  			  : strDel = strDel & Trim(.vspdData.text) & Parent.gRowSep '7
                    lGrpCnt = lGrpCnt + 1
           End Select
       Next

	   .txtMaxRows.value     = lGrpCnt-1
	   .txtSpread.value      = strDel & strVal
	End With
	
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
	Call ExecMyBizASP(frm1, BIZ_PGM_ID)

    DbSave = True                                                                '☜: Processing is OK
    Set gActiveElement = document.ActiveElement   
End Function
'========================================================================================================
Function DbDelete()

    Err.Clear                                                                    '☜: Clear err status
    DbDelete = False                                                             '☜: Processing is NG
    'Call LayerShowHide(1)                                                        '☜: Show Processing Message

	'------ Developer Coding part (Start)  --------------------------------------------------------------
	'------ Developer Coding part (End )   --------------------------------------------------------------

    DbDelete = True                                                             '☜: Processing is OK
	'Call RunMyBizASP(MyBizASP, strVal)                                           '☜: Run Biz logic
End Function

'========================================================================================================
Sub DbQueryOk()

	lgIntFlgMode      = Parent.OPMD_UMODE                                                   '⊙: Indicates that current mode is Create mode
	'------ Developer Coding part (Start)  --------------------------------------------------------------


    if frm1.vspdData.MaxRows = 0 then
       Call SetToolbar("1100110100101111")
    else
       Call SetToolbar("1100111100111111")
    end if

	'------ Developer Coding part (End )   --------------------------------------------------------------
    Call InitData()
	Call ggoOper.LockField(Document, "Q")
	frm1.vspdData.focus
	Set gActiveElement = document.ActiveElement
End Sub

'========================================================================================================
Sub DbSaveOk()
	Call InitVariables
	'------ Developer Coding part (Start)  --------------------------------------------------------------
  
    ggoSpread.Source = Frm1.vspdData
    Frm1.vspdData.MaxRows = 0
    if frm1.vspdData.MaxRows = 0 then
       Call SetToolbar("1100110100101111")
    else
       Call SetToolbar("1100111100111111")
    end if
	'------ Developer Coding part (End )   --------------------------------------------------------------
   
    Set gActiveElement = document.ActiveElement
	MainQuery()
End Sub

'========================================================================================================
Sub DbDeleteOk()
  	'------ Developer Coding part (Start)  --------------------------------------------------------------
    FncQuery()
	if frm1.vspdData.MaxRows = 0 then
       Call SetToolbar("1100110100101111")
    else
       Call SetToolbar("1100111100111111")
    end if
	'------ Developer Coding part (End )   --------------------------------------------------------------
	Call FncNew()
End Sub


'========================================================================================================
'   Event Name : vspdData_Change
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_Change(Col , Row)

    Dim iDx
    Dim IntRetCD,EFlag
    Dim strSelect
	Dim strFrom
	Dim strWhere 	
    Dim arrVal1
	Dim arrVal2
	Dim ii
	Dim jj
	Dim strDate
    DIM DEPT_CD
    DIM ACCT_CD
    DIM ACCT_NM
    EFlag = False

   	Row = Frm1.vspdData.ActiveRow
   	Frm1.vspdData.Row = Frm1.vspdData.ActiveRow
   	Frm1.vspdData.Col = Col
	
	'------ Developer Coding part (Start ) --------------------------------------------------------------
	If Col = C_Dept_Cd or Col = c_acct_nm Then
		If DupCheck(Row) = false Then
			Call DisplayMsgBox("970001","X","자료","X") 				           
				Frm1.vspdData.Row = Row
				Frm1.vspdData.Col = C_DEPT_CD
				frm1.vspdData.Text=""
				frm1.vspdData.Col = C_DEPT_CD_NM
				frm1.vspdData.Text=""
				frm1.vspdData.Col = C_BIZ_AREA_CD
				frm1.vspdData.Text=""
				frm1.vspdData.Col = C_ORG_CHANGE_ID
				frm1.vspdData.Text=""
				frm1.vspdData.Col = C_INTERNAL_CD
				frm1.vspdData.Text=""
				frm1.vspdData.Col = Col
				frm1.vspdData.Action=0
				'call fncCanCel()
			Exit Sub
		end If
	End If
	Select Case Col
		Case C_Dept_Cd
			Frm1.vspdData.Col = C_DEPT_CD
			DEPT_CD = Frm1.vspdData.Text
				If DEPT_CD = "" Then
					frm1.vspdData.Col = C_DEPT_CD_NM
					frm1.vspdData.Text=""
					frm1.vspdData.Col = C_BIZ_AREA_CD
					frm1.vspdData.Text=""
					frm1.vspdData.Col = C_ORG_CHANGE_ID
					frm1.vspdData.Text=""
					frm1.vspdData.Col = C_INTERNAL_CD
					frm1.vspdData.Text=""
					Frm1.vspdData.Col = Col
					Frm1.vspdData.Action = 0
					Set gActiveElement = document.activeElement  
				Else
					strDate = FilterVar(frm1.fpdtWk_yymm.year,"2999","SNM") & "-" & Right("0" & FilterVar(frm1.fpdtWk_yymm.Month,"12","SNM"),2) & "-" & "01"
					strDate = DateAdd("D",-1, DateAdd("M",1,cdate(strDate)))
					strDate = UNIConvDate(strDate)
				
			
					frm1.vspdData.Col = C_DEPT_CD
					Frm1.vspdData.Row = Row
				
					strSelect	=			 " a.dept_cd,a.dept_nm, a.org_change_id, a.internal_cd, b.biz_area_cd "    		
					strFrom		=			 " b_acct_dept a, b_cost_center b "		
					strWhere	= " a.cost_cd = b.cost_cd " 	 
					strWhere	= strWhere & " and a.dept_Cd = " & FilterVar(LTrim(RTrim(frm1.vspdData.Text)), "''", "S")
					strWhere	= strWhere & " and a.org_change_id = (select distinct org_change_id "			
					strWhere	= strWhere & " from b_acct_dept where org_change_dt = ( select max(org_change_dt)"
					strWhere	= strWhere & " from b_acct_dept where org_change_dt <= " & FilterVar(strDate, "''", "S") & "))"

					If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) = False Then
	
						IntRetCD = DisplayMsgBox("124600","X","X","X")  
						frm1.vspdData.Col = C_DEPT_CD_NM
						frm1.vspdData.Text=""
						frm1.vspdData.Col = C_BIZ_AREA_CD
						frm1.vspdData.Text=""
						frm1.vspdData.Col = C_ORG_CHANGE_ID
						frm1.vspdData.Text=""
						frm1.vspdData.Col = C_INTERNAL_CD
						frm1.vspdData.Text=""
						Frm1.vspdData.Col = Col
						Frm1.vspdData.Action = 0
						Set gActiveElement = document.activeElement  
											    
					Else 
						
						arrVal1 = Split(lgF2By2, Chr(11) & Chr(12))			
						jj = Ubound(arrVal1,1)
									
						For ii = 0 to jj - 1
							arrVal2 = Split(arrVal1(ii), chr(11))			
							frm1.vspdData.Col = C_DEPT_CD_NM
							frm1.vspdData.Text=Trim(arrVal2(2))
							frm1.vspdData.Col = C_BIZ_AREA_CD
							frm1.vspdData.Text=Trim(arrVal2(5))
							frm1.vspdData.Col = C_ORG_CHANGE_ID
							frm1.vspdData.Text=Trim(arrVal2(3))
							frm1.vspdData.Col = C_INTERNAL_CD
							frm1.vspdData.Text=Trim(arrVal2(4))
						Next	
					End If
				End If
	  Case c_acct_nm		
			Frm1.vspdData.Col = c_acct_nm
			Frm1.vspdData.Row = Row
			ACCT_NM = Frm1.vspdData.Text
				If ACCT_NM <>"" Then
				    IntRetCD = CommonQueryRs("B.MINOR_CD","B_MAJOR A, B_MINOR B","A.MAJOR_CD =B.MAJOR_cD AND  A.MAJOR_cD = " & FilterVar("H0071", "''", "S") & "  AND B.MINOR_NM = " & FilterVar(ACCT_NM, "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
				    If IntRetCD = False Then
					    Call DisplayMsgBox("110100","x","X","X")
					    Frm1.vspdData.Col = c_acct_code
					    Frm1.vspdData.Text = ""
					    Frm1.vspdData.Col = c_acct_nm
					    Frm1.vspdData.Text = ""
					    Frm1.vspdData.Col = Col
					    Frm1.vspdData.Action = 0
					    Set gActiveElement = document.activeElement
					    EFlag = True
				    Else
					    Frm1.vspdData.Col = c_acct_code
					    Frm1.vspdData.Text = Trim(Replace(lgF0,Chr(11),""))
			    End If
			End If
	end select 	
	'------ Developer Coding part (End   ) --------------------------------------------------------------

	Call CheckMinNumSpread(frm1.vspdData, Col, Row)


	ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row

    '------ Developer Coding part (Start ) --------------------------------------------------------------
    '데이터 확인시 틀린데이터에 대해 undo 해준다.
    Frm1.vspdData.Row = Row
    Frm1.vspdData.Col = 0

    If EFlag And Frm1.vspdData.Text <> ggoSpread.InsertFlag Then
		Call FncCancel()
	End If
	'------ Developer Coding part (End   ) --------------------------------------------------------------
End Sub

'==================================================
	'동일한 데이타가 있는지 체크 
'==================================================
Function DupCheck(row)
	Dim strSelect, strFrom, strWhere
	Dim Rs0,Rs1
	Dim i,j
	Dim strDeptCd
	Dim strBizareacd
	Dim strAcctType
	Dim strPayType
	Dim tmpDeptCd
	Dim tmpBizareacd
	Dim tmpAcctType
	
	
	
	Err.Clear
	DupCheck = False
   
    
	With frm1
			.vspdData.Row = Row
			
			.vspddata.col = C_DEPT_CD
			strDeptCd = Trim(.vspddata.text)
				
			.vspddata.col = C_BIZ_AREA_CD
			strBizareacd = Trim(.vspddata.text)
				
			.vspddata.col = c_acct_nm
			strAcctType = Trim(.vspddata.text)
			
			
			For i=1 to .vspdData.MaxRows
				If i<> Row Then
					.vspdData.Row = i
					.vspddata.col = C_DEPT_CD
					tmpDeptCd = Trim(.vspddata.text)
				
					.vspddata.col = C_BIZ_AREA_CD
					tmpBizareacd = Trim(.vspddata.text)
				
					.vspddata.col = c_acct_nm
					tmpAcctType = Trim(.vspddata.text)
				
					If strDeptCd = tmpDeptCd and strBizareacd = tmpBizareacd and strAcctType = tmpAcctType Then
						Exit Function
					End If
				End If
			Next
		
	End With

	DupCheck = True
End Function	
	



'========================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'========================================================================================================
Sub vspdData_Click(Col, Row)

	If lgIntFlgMode = Parent.OPMD_CMODE Then
		Call SetPopupMenuItemInf("1001111111")
	Else
		Call SetPopupMenuItemInf("1101111111")
	End If
    gMouseClickStatus = "SPC"

    Set gActiveSpdSheet = frm1.vspdData

    If frm1.vspdData.MaxRows <= 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
   	    
    If Row <= 0 Then
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
    
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    	frm1.vspdData.Row = Row
	'	frm1.vspdData.Col = C_MajorCd
	
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
    
	
End Sub

'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub

'======================================================================================================
'   Event Name : vspdData_MouseDown
'   Event Desc : 특정 column를 click할때 
'======================================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If
End Sub

'========================================================================================================
'   Event Name : vspdData_DblClick
'   Event Desc : 
'========================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)				
    Dim iColumnName
	
	If Row<=0 then
		Exit Sub
	End If
	If Frm1.vspdData.MaxRows =0 then
		Exit Sub
	End if
End Sub

'========================================================================================================
Sub vspdData_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    'If Col <= C_SNm Or NewCol <= C_SNm Then
    '    Cancel = True
    '    Exit Sub
    'End If
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    
    Call GetSpreadColumnPos("A")
End Sub
'========================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )

    If OldLeft <> NewLeft Then
        Exit Sub
    End If

    if Frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then
       If lgStrPrevKeyIndex <> "" Then
          lgCurrentSpd = "M"
          Call MakeKeyStream("X")
          Call DisableToolBar(Parent.TBC_QUERY)
			If DbQuery = False Then
				Call RestoreToolBar()
				Exit Sub
			End if
       End If
    End if

End Sub


'======================================================================================================
'   Event Name : vspdData_ComboSelChange
'   Event Desc : Combo 변경 이벤트 
'=======================================================================================================
Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
	Dim intIndex

	With frm1.vspdData

		.Row = Row

        Select Case Col
            Case c_acct_nm
               Call vspdData_change(c_acct_nm, Row)
		End Select
	End With

   	ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row

End Sub

'========================================================================================================
'   Event Name : cboYesNo_OnChange
'   Event Desc :
'========================================================================================================
Sub cboYesNo_OnChange()
    lgBlnFlgChgValue = True
End Sub


'========================================================================================================
'   Event Name : txtCurrencyCode_OnChange
'   Event Desc :
'========================================================================================================
Sub txtCurrencyCode_OnChange()
   If txtBizAreaCdChange = False  Then Exit Sub
End Sub


'========================================================================================================
'   Event Name : txtCurrencyCode_OnChange
'   Event Desc :
'========================================================================================================
Function txtBizAreaCdChange()
    Dim IntRetCd
	txtBizAreaCdChange = False
    If  frm1.txtCurrencyCode.value = "" Then
		frm1.txtCurrency.value=""
		frm1.txtCurrencyCode.focus
	
    Else
        IntRetCD= CommonQueryRs(" BIZ_AREA_NM "," B_BIZ_AREA "," BIZ_AREA_CD =  " & FilterVar(Frm1.txtCurrencyCode.value, "''", "S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) 
			If IntRetCD=False And Trim(frm1.txtCurrencyCode.value) <>"" Then
			    Call DisplayMsgBox("124200","X","X","X") 
				frm1.txtCurrency.value=""
				frm1.txtCurrencyCode.focus
			    Set gActiveElement = document.activeElement  
			    Exit Function
			Else
			    frm1.txtCurrency.value=Trim(Replace(lgF0,Chr(11),""))
			    frm1.txtCurrencyCode.focus
			    Set gActiveElement = document.activeElement  
			End If
    End if
	txtBizAreaCdChange = True
End Function

'========================================================================================================
' Name : fpdtWk_yymm_DblClick
' Desc :
'=======================================================================================================
Sub fpdtWk_yymm_DblClick(Button)
	If Button = 1 Then
		frm1.fpdtWk_yymm.Action = 7
 		Call SetFocusToDocument("M")
		Frm1.fpdtWk_yymm.Focus
	End If
End Sub
'======================================================================================================
' Name : fpdtWk_yymm_KeyPress
' Desc : Call Mainquery
'=======================================================================================================
Sub fpdtWk_yymm_KeyPress(Key)
    If key = 13 Then
        Call FncQuery
		End If
End Sub
</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->
</HEAD>

<!--
'########################################################################################################
'#						6. TAG 부																		#
'########################################################################################################
-->
<BODY SCROLL="NO" TABINDEX="-1">
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>퇴직급여추계액등록</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
                    <TD WIDTH=* ALIGN=RIGHT>&nbsp;</A></TD>
					<TD WIDTH=10>&nbsp;</TD>
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
									<TD CLASS=TD5 NOWRAP WIDTH=14%>년월</TD>
									<TD CLASS=TD6 NOWRAP WIDTH=86%><script language =javascript src='./js/a5963ma1_fpDateTime3_fpdtWk_yymm.js'></script></TD>
									<TD CLASS=TD5 NOWRAP>사업장</TD>
									<TD CLASS=TD6 NOWRAP>
									<INPUT TYPE=TEXT NAME="txtCurrencyCode" SIZE=10 MAXLENGTH=20 tag="12XXXU"  ALT="사업장" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCode1" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenCurrency()">
									<INPUT TYPE=TEXT NAME="txtCurrency" SIZE=22 MAXLENGTH=50 tag="14XXXU"  ALT="사업장명">
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
					<TD HEIGHT="80%" WIDTH="100%" COLSPAN=4>
						<script language =javascript src='./js/a5963ma1_vaSpread1_vspdData.js'></script>
					</TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_30%>>
							<TR>
						   	<TD CLASS=TD5 NOWRAP>당월예상액 합계</TD>
						   	<TD CLASS=TD6 NOWRAP> <script language =javascript src='./js/a5963ma1_fpDoubleSingle1_txtAmt1.js'></script></TD>
						   	<TD CLASS=TD5 NOWRAP>전월누계액 합계</TD>
						   	<TD CLASS=TD6 NOWRAP> <script language =javascript src='./js/a5963ma1_fpDoubleSingle2_txtAmt2.js'></script></TD>
							</TR>
							<TR>
						   	<TD CLASS=TD5 NOWRAP>당월지급액 합계</TD>
						   	<TD CLASS=TD6 NOWRAP> <script language =javascript src='./js/a5963ma1_fpDoubleSingle3_txtAmt3.js'></script></TD>
						   	<TD CLASS=TD5 NOWRAP>당월반영액 합계</TD>
						   	<TD CLASS=TD6 NOWRAP> <script language =javascript src='./js/a5963ma1_fpDoubleSingle4_txtAmt4.js'></script></TD>
							</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR >
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=no	 noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread"       TAG="24" TABINDEX="-1"></TEXTAREA>
<INPUT TYPE=HIDDEN       NAME="txtMode"         TAG="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN       NAME="txtKeyStream"    TAG="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN       NAME="txtUpdtUserId"   TAG="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN       NAME="txtInsrtUserId"  TAG="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN       NAME="txtFlgMode"      TAG="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN       NAME="txtPrevNext"     TAG="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN       NAME="txtMaxRows"      TAG="24" TABINDEX="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm" TABINDEX="-1"></IFRAME>
</DIV>
</BODY>
</HTML>

