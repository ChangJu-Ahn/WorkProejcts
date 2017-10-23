<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          	: Human Resources
*  2. Function Name        	: 급여자동기표처리이력조회 
*  3. Program ID           	: H6022ma1
*  4. Program Name         	: H6022ma1
*  5. Program Desc         	: 급여관리 
*  6. Comproxy List        	:
*  7. Modified date(First) 	: 2003/05/21
*  8. Modified date(Last)  	: 2003/06/13
*  9. Modifier (First)     	: SBK
* 10. Modifier (Last)     	: Lee SiNa
* 11. Comment              	:
=======================================================================================================-->
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
<SCRIPT LANGUAGE="VBScript"	  SRC="../../inc/incHRQuery.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs"></SCRIPT>

<Script Language="VBScript">
Option Explicit
'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
Const BIZ_PGM_ID = "H6022mb1.asp"						           '☆: Biz Logic ASP Name
Const BIZ_PGM_JUMP_ID = "H6021ba1" 

'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim gSelframeFlg                                                 
Dim gblnWinEvent                                                 
Dim lgBlnFlawChgFlg	
Dim gtxtChargeType
Dim IsOpenPop
Dim lsInternal_cd
Dim lgStrComDateType		                                            'Company Date Type을 저장(년월 Mask에 사용함.)
Dim lgIsOpenPop                                          

Dim C_PAY_YYMM
Dim C_PROV_TYPE
Dim C_PROV_TYPE_NM
Dim C_PROV_DT
Dim C_BIZ_AREA_NM
Dim C_TRANS_FLAG
Dim C_TRANS_DT
Dim C_GL_NO

'========================================================================================================
' Name : initSpreadPosVariables()	
' Desc : Initialize the position
'========================================================================================================
Sub InitSpreadPosVariables()	 
	 C_PAY_YYMM	    = 1
	 C_PROV_TYPE	= 2  
	 C_PROV_TYPE_NM	= 3
	 C_PROV_DT	    = 4
	 C_BIZ_AREA_NM  = 5
	 C_TRANS_FLAG	= 6
	 C_TRANS_DT	    = 7
	 C_GL_NO		= 8
End Sub

'========================================================================================================
' Name : InitVariables()	
' Desc : Initialize value
'========================================================================================================
Sub InitVariables()
	lgIntFlgMode      =  parent.OPMD_CMODE						        '⊙: Indicates that current mode is Create mode
	lgBlnFlgChgValue  = False								    '⊙: Indicates that no value changed
	lgIntGrpCount     = 0										'⊙: Initializes Group View Size
	lgStrPrevKey      = ""                                      '⊙: initializes Previous Key
	lgSortKey         = 1                                       '⊙: initializes sort direction
	gblnWinEvent      = False
	lgBlnFlawChgFlg   = False
End Sub

'========================================================================================================
' Name : SetDefaultVal()	
' Desc : Set default value
'========================================================================================================
Sub SetDefaultVal()
	Dim strYear
	Dim strMonth
	Dim strDay

	Call ExtractDateFrom("<%=GetsvrDate%>",Parent.gServerDateFormat , Parent.gServerDateType ,strYear,strMonth,strDay)
	
	frm1.txtFrProv_dt.Text = UNIConvDateAToB("<%=GetSvrDate%>",Parent.gServerDateFormat , Parent.gDateFormat)

	frm1.txtFrProv_dt.focus
	frm1.txtFrProv_dt.Year = strYear 		 '년월일 default value setting
	frm1.txtFrProv_dt.Month = "01" 
	frm1.txtFrProv_dt.Day = "01" 

	frm1.txtToProv_dt.Text = UNIConvDateAToB("<%=GetSvrDate%>", Parent.gServerDateFormat, Parent.gDateFormat)
	frm1.txtToProv_dt.Text = UNIGetLastDay(frm1.txtToProv_dt.Text, Parent.gDateFormat)

End Sub
	
'========================================================================================================
' Name : LoadInfTB19029()	
' Desc : Set System Number format
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "H","NOCOOKIE","MA") %>
End Sub
'========================================================================================================
'	Name : CookiePage()
'	Description : Item Popup에서 Return되는 값 setting
'========================================================================================================
Function CookiePage(ByVal flgs)
	Const CookieSplit = 4877						
	Dim strTemp

	If flgs = 1 Then
        frm1.vspdData.Row = frm1.vspdData.ActiveRow
	    
        frm1.vspdData.Col = C_PROV_DT               
        WriteCookie "PROV_DT" , frm1.vspdData.Text

        frm1.vspdData.Col = C_PROV_TYPE
        WriteCookie "PROV_TYPE", frm1.vspdData.Text

        frm1.vspdData.Col = C_PROV_TYPE_NM
        WriteCookie "PROV_TYPE_NM", frm1.vspdData.Text
	    
        frm1.vspdData.Col = C_TRANS_DT 
        WriteCookie "TRANS_DT"   , frm1.vspdData.Text

	ElseIf flgs = 0 Then

		strTemp = ReadCookie(CookieSplit)
		If strTemp = "" then Exit Function
	
		WriteCookie "C_PROV_DT" , ""
	    WriteCookie "C_PROV_TYPE"      , ""
        WriteCookie "C_PROV_TYPE_NM"   , ""
        WriteCookie "C_TRANS_DT"   , ""
		
	    Call MainQuery()
			
	End If
End Function
'--------------------------	Description : 상세조회 클릭시 에러 체크사항 -------------------------------
FUNCTION PgmJumpCheck()         
    If frm1.vspdData.ActiveRow =  0 Then
		Call DisplayMsgBox("800167","X","X","X")
		frm1.txtFrProv_dt.focus		
	    Exit Function
	Else
        PgmJump(BIZ_PGM_JUMP_ID)   
	End If	   
End Function

'========================================================================================================
' Function Name : MakeKeyStream
' Function Desc : This method set focus to pos of err
'========================================================================================================
Sub MakeKeyStream(pOpt)
    Dim rbo_sort
    
    lgKeyStream  = UNIConvDate(frm1.txtFrProv_dt.Text) & parent.gColSep
    lgKeyStream  = lgKeyStream & UNIConvDate(frm1.txtToProv_dt.Text) & parent.gColSep
    lgKeyStream  = lgKeyStream & frm1.txtProv_type.value
End Sub 

'========================================================================================================
' Name : InitData()
' Desc : Reset Combox
'========================================================================================================
Sub InitData()
End Sub

'========================================================================================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================================
Sub InitSpreadSheet()

    Dim strMaskYM	

	If Date_DefMask(strMaskYM) = False Then
		strMaskYM = "9999" & lgStrComDateType & "99"
	End If	
	
	Call initSpreadPosVariables()	

    With frm1.vspdData

	    ggoSpread.Source = frm1.vspdData	
        ggoSpread.Spreadinit "V20030523",,parent.gAllowDragDropSpread    

	    .ReDraw = false
        .MaxCols = C_GL_NO + 1											'☜: 최대 Columns의 항상 1개 증가시킴 

	    .Col = .MaxCols													'공통콘트롤 사용 Hidden Column
        .ColHidden = True
        
        .MaxRows = 0	
        ggoSpread.ClearSpreadData

         Call  GetSpreadColumnPos("A")

         ggoSpread.SSSetMask   C_PAY_YYMM,        "급/상여년월",  12, 2, strMaskYM
         ggoSpread.SSSetEdit   C_PROV_TYPE,    "",  2
         ggoSpread.SSSetEdit   C_PROV_TYPE_NM,    "지급구분",  18
         ggoSpread.SSSetEdit   C_PROV_DT,         "급/상여지급일",14, 2
         ggoSpread.SSSetEdit   C_BIZ_AREA_NM,     "사업장",   25
         ggoSpread.SSSetEdit   C_TRANS_FLAG,      "전표생성여부", 12
         ggoSpread.SSSetEdit   C_TRANS_DT,        "전표생성일",   14, 2
         ggoSpread.SSSetEdit   C_GL_NO,           "전표번호",     20,,,50,2

         Call ggoSpread.SSSetColHidden(C_PROV_TYPE,C_PROV_TYPE,True)
		
		.ReDraw = true
		
		Call SetSpreadLock 
	
	End With

End Sub

'======================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadLock()
      ggoSpread.Source = frm1.vspdData
      ggoSpread.SpreadLockWithOddEvenRowColor()
End Sub

'======================================================================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
    With frm1         
        .vspdData.ReDraw = False
         ggoSpread.SSSetProtected	C_PAY_YYMM	, pvStartRow, pvEndRow
		 ggoSpread.SSSetProtected	C_PROV_TYPE, pvStartRow, pvEndRow
		 ggoSpread.SSSetProtected	C_PROV_TYPE_NM, pvStartRow, pvEndRow
         ggoSpread.SSSetProtected	C_PROV_DT	, pvStartRow, pvEndRow  
         ggoSpread.SSSetProtected	C_BIZ_AREA_NM	, pvStartRow, pvEndRow  
         ggoSpread.SSSetProtected	C_TRANS_FLAG	, pvStartRow, pvEndRow
         ggoSpread.SSSetProtected	C_TRANS_DT	, pvStartRow, pvEndRow  
		 ggoSpread.SSSetProtected	C_GL_NO		, pvStartRow, pvEndRow
        .vspdData.ReDraw = True
    End With
End Sub

'======================================================================================================
' Function Name : SubSetErrPos
' Function Desc : This method set focus to pos of err
'======================================================================================================
Sub SubSetErrPos(iPosArr)
    Dim iDx
    Dim iRow
    iPosArr = Split(iPosArr, parent.gColSep)
    If IsNumeric(iPosArr(0)) Then
       iRow = CInt(iPosArr(0))
       For iDx = 1 To  frm1.vspdData.MaxCols - 1
           Frm1.vspdData.Col = iDx
           Frm1.vspdData.Row = iRow
           If Frm1.vspdData.ColHidden <> True And Frm1.vspdData.BackColor <>  parent.UC_PROTECTED Then
              Frm1.vspdData.Col = iDx
              Frm1.vspdData.Row = iRow
              Frm1.vspdData.Action = 0 ' go to 
              Exit For
           End If
       Next
    End If   
End Sub

Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)                

            C_PAY_YYMM	    = iCurColumnPos(1)
			C_PROV_TYPE     = iCurColumnPos(2)  
			C_PROV_TYPE_NM  = iCurColumnPos(3)
			C_PROV_DT	    = iCurColumnPos(4)
			C_BIZ_AREA_NM	= iCurColumnPos(5)
			C_TRANS_FLAG	= iCurColumnPos(6)
			C_TRANS_DT	= iCurColumnPos(7)
			C_GL_NO		= iCurColumnPos(8)       
   
    End Select    
End Sub

'========================================================================================================
' Name : Form_Load
' Desc : developer describe this line Called by Window_OnLoad() evnt
'========================================================================================================
Sub Form_Load()

    Err.Clear                                                                       '☜: Clear err status
	Call LoadInfTB19029                                                             '☜: Load table , B_numeric_format
		
    Call ggoOper.FormatField(Document, "A", ggStrIntegeralPart,  ggStrDeciPointPart, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")											'⊙: Lock Field
    
    Call InitSpreadSheet                                                            'Setup the Spread sheet
    Call InitVariables                                                              'Initializes local global variables
    
    Call FuncGetAuth(gStrRequestMenuID,  parent.gUsrID, lgUsrIntCd)                                ' 자료권한:lgUsrIntCd ("%", "1%")
    
    Call SetDefaultVal
    Call SetToolbar("1100000000011111")												'⊙: Set ToolBar
    
    Call CookiePage (0)                                                             '☜: Check Cookie
    
End Sub
	
'========================================================================================================
' Name : Form_QueryUnload
' Desc : developer describe this line Called by Window_OnUnLoad() evnt
'========================================================================================================
Sub Form_QueryUnload(Cancel, UnloadMode)

End Sub

'========================================================================================================
' Name : FncQuery
' Desc : developer describe this line Called by MainQuery in Common.vbs
'========================================================================================================
Function FncQuery()
    Dim IntRetCD 
    
    FncQuery = False                                                            '☜: Processing is NG
    
    Err.Clear                                                                   '☜: Protect system from crashing

    ggoSpread.Source = Frm1.vspdData
    If   ggoSpread.SSCheckChange = True Then
		IntRetCD =  DisplayMsgBox("900013",  parent.VB_YES_NO,"X","X")			        '☜: "Will you destory previous data"		
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If    

    ggoSpread.ClearSpreadData
    														'⊙: Initializes local global variables
    If Not chkField(Document, "1") Then									        '⊙: This function check indispensable field
       Exit Function
    End If

    If txtprov_type_Onchange() Then       
        Exit Function
    End if

    If CompareDateByFormat(frm1.txtFrProv_dt.Text,frm1.txtToProv_dt.Text,frm1.txtFrProv_dt.Alt,frm1.txtToProv_dt.Alt,"970025",parent.gDateFormat,parent.gComDateType,True) = False Then
        frm1.txtFrProv_dt.focus
        Set gActiveElement = document.activeElement

        Exit Function
    End if 

    Call InitVariables	
    Call MakeKeyStream("X")

    Call  DisableToolBar( parent.TBC_QUERY)
    If DbQuery = False Then
        Call  RestoreToolBar()
        Exit Function
    End If
          
    FncQuery = True																'☜: Processing is OK

End Function
	
'========================================================================================================
' Name : FncPrint
' Desc : developer describe this line Called by MainDeleteRow in Common.vbs
'========================================================================================================
Function FncPrint()
	Call Parent.FncPrint()                                                    '☜: Protect system from crashing
End Function

'========================================================================================================
' Name : FncExcel
' Desc : developer describe this line Called by MainExcel in Common.vbs
'========================================================================================================
Function FncExcel() 
	Call Parent.FncExport( parent.C_SINGLE)
End Function

'========================================================================================================
' Name : FncFind
' Desc : developer describe this line Called by MainFind in Common.vbs
'========================================================================================================
Function FncFind() 
	Call Parent.FncFind( parent.C_SINGLE, True)
End Function

'========================================================================================================
' Name : FncExit
' Desc : developer describe this line Called by MainExit in Common.vbs
'========================================================================================================
Function FncExit()

	Dim IntRetCD
	FncExit = False
	 ggoSpread.Source = frm1.vspdData
    If lgBlnFlgChgValue = True Or  ggoSpread.SSCheckChange = True Then
		IntRetCD =  DisplayMsgBox("900016",  parent.VB_YES_NO,"X","X")			 '⊙: Data is changed.  Do you want to exit? 
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    FncExit = True

End Function

'========================================================================================================
' Name : DbQuery
' Desc : This function is called by FncQuery
'========================================================================================================
Function DbQuery()
    Dim strVal
    Err.Clear                                                                    '☜: Clear err status

    DbQuery = False                                                              '☜: Processing is NG

    If   LayerShowHide(1) = False Then
     	Exit Function
    End If

    strVal = BIZ_PGM_ID & "?txtMode="            & parent.UID_M0001                     '☜: Query
    strVal = strVal     & "&txtKeyStream="       & lgKeyStream                   '☜: Query Key
    strVal = strVal     & "&lgStrPrevKey="       & lgStrPrevKey             '☜: Next key tag
    strVal = strVal     & "&txtMaxRows="         & Frm1.vspdData.MaxRows         '☜: Max fetched data

    Call RunMyBizASP(MyBizASP, strVal)                                           '☜:  Run biz logic
	
    DbQuery = True                                                               '☜: Processing is NG
    
    frm1.vspdData.focus
End Function

'========================================================================================================
' Function Name : DbQueryOk
' Function Desc : Called by MB Area when query operation is successful
'========================================================================================================
Function DbQueryOk()
    Dim IRow

	lgIntFlgMode      =  parent.OPMD_UMODE                                               '⊙: Indicates that current mode is Create mode

    Call  ggoOper.LockField(Document, "Q")

    Set gActiveElement = document.ActiveElement   
    lgBlnFlgChgValue = False
    frm1.vspdData.Focus

End Function
	
Sub cboProv_type_OnChange()
    lgBlnFlgChgValue = True
End Sub

'========================================== OpenPopupTempGl() ============================================
'	Name : OpenPopuptempGL()
'	Description : CtrlItem Popup에서 Return되는 값 setting (결의전표 팝업)
'=========================================================================================================
Function OpenPopuptempGL()
	Dim arrRet
	Dim arrParam(1)	
	Dim iCalledAspName
	Dim IntRetCD
	
	iCalledAspName = AskPRAspName("a5130ra1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a5130ra1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	If lgIsOpenPop = True Then Exit Function
	
	With frm1.vspdData
		.Row = .ActiveRow
		.Col = C_GL_NO
		arrParam(0) = Trim(.Text)							        '결의전표번호 
	    arrParam(1) = ""											'Reference번호	
	End With

	lgIsOpenPop = True
   
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	lgIsOpenPop = False
End Function

'========================================== OpenPopupGL()  =============================================
'	Name : OpenPopupGL()
'	Description : CtrlItem Popup에서 Return되는 값 setting (회계전표 팝업)
'=======================================================================================================
Function OpenPopupGL()
	Dim arrRet
	Dim arrParam(1)	
	Dim iCalledAspName
	Dim IntRetCD
	
	iCalledAspName = AskPRAspName("a5120ra1")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a5120ra1", "X")
		IsOpenPop = False
		Exit Function
	End If

	If lgIsOpenPop = True Then Exit Function
	
	With frm1.vspdData
		.Row = .ActiveRow
		.Col = C_GL_NO
		arrParam(0) = Trim(.Text)							        '회계전표번호 
	    arrParam(1) = ""											'Reference번호	
	End With

	lgIsOpenPop = True
   
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	lgIsOpenPop = False
End Function

'========================================================================================================
' Name : OpenCondAreaPopup()
' Desc : developer describe this line 
'========================================================================================================
Function OpenCondAreaPopup(Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True  Then  
	   Exit Function
	End If   

	IsOpenPop = True
	Select Case iWhere
	    Case "1"
	        arrParam(0) = "지급구분팝업"			' 팝업 명칭 
	        arrParam(1) = "B_MINOR"				 		' TABLE 명칭 
	        arrParam(2) = frm1.txtProv_type.value		    ' Code Condition
	        arrParam(3) = ""'frm1.txtAllow_nm.value		' Name Cindition
	        arrParam(4) = "MAJOR_CD = " & FilterVar("H0040", "''", "S") & ""          ' Where Condition
	        arrParam(5) = "지급구분"			    ' TextBox 명칭 
	
            arrField(0) = "MINOR_CD"					' Field명(0)
            arrField(1) = "MINOR_NM"				    ' Field명(1)
    
            arrHeader(0) = "지급구분"				' Header명(0)
            arrHeader(1) = "지급구분명"			    ' Header명(1)
	End Select
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtProv_type.focus
		Exit Function
	Else
		Call SubSetCondArea(arrRet,iWhere)
	End If	
	
End Function

'======================================================================================================
'	Name : SetCondArea()
'	Description : Item Popup에서 Return되는 값 setting
'=======================================================================================================
Sub SubSetCondArea(Byval arrRet, Byval iWhere)
	With Frm1
		Select Case iWhere
		    Case "1"
		        .txtProv_type.value = arrRet(0)
		        .txtProv_type_Nm.value = arrRet(1)		
		        .txtProv_type.focus
        End Select
	End With
End Sub

'========================================================================================================
'   Event Name : txtprov_type_change
'   Event Desc :
'========================================================================================================
Function txtprov_type_Onchange()
    Dim IntRetCd
    
    If frm1.txtprov_type.value = "" Then
		frm1.txtprov_type_nm.value = ""
    Else
        IntRetCd = CommonQueryRs(" MINOR_NM "," B_MINOR "," MAJOR_CD=" & FilterVar("H0040", "''", "S") & " AND MINOR_CD= " & FilterVar(frm1.txtprov_type.value, "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    
        If IntRetCd = false then
			Call DisplayMsgBox("800248","X","X","X")	
			 frm1.txtprov_type_nm.value = ""
             frm1.txtprov_type.focus
            Set gActiveElement = document.ActiveElement
            txtprov_type_Onchange = true 
            
            Exit Function          
        Else
			frm1.txtprov_type_nm.value = Trim(Replace(lgF0,Chr(11),""))
        End if 
    End if  
End Function

'========================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'========================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)

    Call SetPopupMenuItemInf("0000111111")
    gMouseClickStatus = "SPC" 
    Set gActiveSpdSheet = frm1.vspdData

	if frm1.vspddata.MaxRows <= 0 then
		exit sub
	end if
	
	if Row <=0 then
		ggoSpread.Source = frm1.vspdData
		if lgSortkey = 1 then
			ggoSpread.SSSort Col
			lgSortKey = 2
		else
			ggoSpread.SSSort Col, lgSortkey
			lgSortKey = 1
		end if
		Exit sub
	end if
	frm1.vspdData.Row = Row     
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
'   Event Name : vspdData_DblClick
'   Event Desc : 
'========================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)
    Dim iColumnName
    if Row <= 0 then
		exit sub
	end if
	if Frm1.vspdData.MaxRows = 0 then
		exit sub
	end if
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

Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And  gMouseClickStatus = "SPC" Then
        gMouseClickStatus = "SPCR"
     End If
End Sub

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

'========================================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
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

Sub txtFrProv_dt_DblClick(Button) 
    If Button = 1 Then
		Call SetFocusToDocument("M")    
        frm1.txtFrProv_dt.Action = 7
        frm1.txtFrProv_dt.focus
    End If
End Sub

Sub txtFrProv_dt_Keypress(Key) 
    If Key = 13 Then
        Call MainQuery()
    End If
End Sub

Sub txtToProv_dt_DblClick(Button) 
    If Button = 1 Then
		Call SetFocusToDocument("M")    
        frm1.txtToProv_dt.Action = 7
        frm1.txtToProv_dt.focus
    End If
End Sub

Sub txtToProv_dt_Keypress(Key) 
    If Key = 13 Then
        Call MainQuery()
    End If
End Sub

'========================================================================================================
' Function Name : Date_DefMask()
' Function Desc : This method initializes spread sheet column property
'========================================================================================================
Function Date_DefMask(strMaskYM)
    Dim i,j
    Dim ArrMask,StrComDateType
	
	Date_DefMask = False
	
	strMaskYM = ""
	
	ArrMask = Split(parent.gDateFormat,parent.gComDateType)
	
	If parent.gComDateType = "/" Then 
		lgStrComDateType = "/" & parent.gComDateType
	Else
		lgStrComDateType = parent.gComDateType
	End If
		
	If IsArray(ArrMask) Then
		For i=0 To Ubound(ArrMask)		
			If Instr(UCase(ArrMask(i)),"D") = False Then
				If strMaskYM <> "" Then
					strMaskYM = strMaskYM & lgStrComDateType
				End If
				If Instr(UCase(ArrMask(i)),"M") And Len(ArrMask(i)) >= 3 Then
					strMaskYM = strMaskYM & "U"
					For j=0 To Len(ArrMask(i)) - 2
						strMaskYM = strMaskYM & "L"
					Next
				Else
					strMaskYM = strMaskYM & ArrMask(i)
				End If
			End If
		Next		
	Else
		Date_DefMask = False
		Exit Function
	End If	

	strMaskYM = Replace(UCase(strMaskYM),"Y","9")
	strMaskYM = Replace(UCase(strMaskYM),"M","9")

	Date_DefMask = True 
	
End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>

<BODY SCROLL="NO" TABINDEX="-1">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR >
		<TD <%=HEIGHT_TYPE_00%>></TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_40%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>급/상여자동기표처리이력조회</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* align=right><A href="vbscript:OpenPopupGL()">회계전표</A>&nbsp;|&nbsp;<A href="vbscript:OpenPopupTempGL()">결의전표</A> </TD>					
					<TD WIDTH=10></TD>
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
					<TD WIDTH=100% HEIGHT=20 VALIGN=TOP>
       					<FIELDSET CLASS="CLSFLD">
       						<TABLE <%=LR_SPACE_TYPE_40%>>
                                <TR>
									<TD	CLASS="TD5"	NOWRAP>지급일자</TD>
									<TD	CLASS="TD6"	NOWRAP><script language =javascript src='./js/h6022ma1_txtFrProv_dt_txtFrProv_dt.js'></script>&nbsp;~&nbsp;
									                       <script language =javascript src='./js/h6022ma1_txtToProv_dt_txtToProv_dt.js'></script>
									<TD	CLASS="TD5"	NOWRAP>지급구분</TD>
									<TD	CLASS="TD6"	NOWRAP><INPUT NAME="txtProv_type" MAXLENGTH="1"	 SIZE="10" ALT ="지급구분" TAG="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" ALIGN=top TYPE="BUTTON" ONCLICK="VBScript:	OpenCondAreaPopup(1)">
														   <INPUT NAME="txtProv_type_nm" MAXLENGTH="20"	SIZE="20" ALT ="지급구분명" tag="14"></TD>
              					</TR>
                            </TABLE>
						</FIELDSET>						        
				</TR>		   
				<TR>
				    <TD <%=HEIGHT_TYPE_03%>></TD>
				</TR>
				<TR>
		         	<TD WIDTH=100% HEIGHT=* valign=top>
		                <TABLE <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD HEIGHT=100% WIDTH=100%> 
					                <script language =javascript src='./js/h6022ma1_vaSpread_vspdData.js'></script>
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
	         		<TD WIDTH="*" ALIGN=RIGHT><a href = "VBSCRIPT:PgmJumpCheck()" ONCLICK="VBSCRIPT:CookiePage 1">급여자동기표처리</a></TD>
	                <TD WIDTH=10>&nbsp;</TD>
	            </TR>
	        </TABLE>
	    </TD>
	</TR>
	<TR >
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode"        TAG="24">
<INPUT TYPE=HIDDEN NAME="txtKeyStream"   TAG="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId"  TAG="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode"     TAG="24">
<INPUT TYPE=HIDDEN NAME="txtPrevNext"    TAG="24">

<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA><INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">

</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm"></IFRAME>
</DIV>
</BODY>
</HTML>

