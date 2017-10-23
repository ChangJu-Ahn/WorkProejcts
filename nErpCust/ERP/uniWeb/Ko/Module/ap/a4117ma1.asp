<%@ LANGUAGE="VBSCRIPT" %>

<!--
'=======================================================================================================
'*  1. Module Name          : Accounting
'*  2. Function Name        : 
'*  3. Program ID           : A4117ma1
'*  4. Program Name         : 채무잔액정리 
'*  5. Program Desc         : 
'*  6. Modified date(First) : 2000/09/25
'*  7. Modified date(Last)  : 2000/12/20
'*  8. Modifier (First)     : 장성희 
'*  9. Modifier (Last)      : hersheys
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'======================================================================================================= -->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!--
'=======================================================================================================
'												1. 선 언 부 
'=======================================================================================================

'=======================================================================================================
'                                               1.1 Inc 선언   
'	기능: Inc. Include
'======================================================================================================= -->
<!-- #Include file="../../inc/incSvrCcm.inc"  -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAMain.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAEvent.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliVariables.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAOperation.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliRdsQuery.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"     SRC="../ag/AcctCtrl.vbs">				</SCRIPT>
<SCRIPT LANGUAGE=vbscript>

Option Explicit																	'☜: indicates that All variables must be declared in advance

'=======================================================================================================
'                                               1.2 Global 변수/상수 선언  
'	.Constant는 반드시 대문자 표기.
'	.변수 표준에 따름. prefix로 g를 사용함.
'	.Array인 경우는 ()를 반드시 사용하여 일반 변수와 구별해 됨 
'=======================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	

Const BIZ_PGM_ID         = "a4117mb1.asp"									' F_PrPaym_Sttl 의 CRUD

Const TAB1 = 1																		'☜: Tab의 위치 
Const TAB2 = 2

Dim C_ItemSeq   
Dim C_AdjustDt  
Dim C_AcctCd    
Dim C_AcctPopUp 
Dim C_AcctNm	
Dim C_AdjustAmt   
Dim C_AdjustLocAmt
Dim C_DocCur     
Dim C_DocCurPopUp
Dim C_AdjustDESC
Dim C_Temp_GlNo
Dim C_GlNo 
Dim C_RefNo
Dim C_DrCRFG

Dim  lgStrPrevKeyDtl
Dim  lgStrPrevKey2
Dim  lgStrPrevKey3
Dim  lgPrevNo
Dim  lgNextNo
Dim  lgCurrRow

Dim  IsOpenPop	                'Popup
Dim  gSelframeFlg

' 권한관리 추가 
Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' 사업장 
Dim lgInternalCd, lgDeptCd, lgDeptNm			' 내부부서 
Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' 내부부서(하위포함)
Dim lgAuthUsrID, lgAuthUsrNm					' 개인 

'*******************************************************************************************************
'=======================================================================================================
' Area Name   : 4.5.1 Common Group -1
' Description : This part declares 1st common function group
'=======================================================================================================
'*******************************************************************************************************




'========================================================================================================= 
' Name : initSpreadPosVariables()
' Description : 그리드(스프래드) 컬럼 관련 변수 초기화 
'========================================================================================================= 
Sub initSpreadPosVariables()
     C_ItemSeq      = 1																'☆: Spread Sheet 의 Columns 인덱스 
	 C_AdjustDt     = 2
	 C_AcctCd       = 3
	 C_AcctPopUp    = 4
	 C_AcctNm		= 5
	 C_AdjustAmt    = 6
	 C_AdjustLocAmt = 7
	 C_DocCur       = 8
	 C_DocCurPopUp  = 9
	 C_AdjustDESC	= 10
	 C_Temp_GlNo    = 11
	 C_GlNo         = 12
	 C_RefNo        = 13
	 C_DrCRFG       = 14
End Sub

'=======================================================================================================
'	Name : InitVariables()
'	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
'=======================================================================================================
Sub  InitVariables()
    lgIntFlgMode = parent.OPMD_CMODE												'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False														'Indicates that no value changed
    lgIntGrpCount = 0																'initializes Group View Size
        
    lgStrPrevKey = 0																'initializes Previous Key
    lgStrPrevKeyDtl = 0																'initializes Previous Key
    lgLngCurRows = 0																'initializes Deleted Rows Count
    frm1.hOrgChangeId.value = parent.gChangeOrgId
End Sub

'=======================================================================================================
'	Name : SetDefaultVal()
'	Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
'=======================================================================================================
Sub  SetDefaultVal()
	
End Sub

'======================================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'=======================================================================================================
Sub  LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% Call LoadInfTB19029A("I","*","NOCOOKIE","MA") %>
<% Call LoadBNumericFormatA("I", "*","NOCOOKIE","MA") %>
End Sub

'======================================================================================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'=======================================================================================================
Sub  InitSpreadSheet()
    Call initSpreadPosVariables()  
    
    With frm1
	    ggoSpread.Source = .vspdData
        ggoSpread.Spreadinit "V20021127",,parent.gAllowDragDropSpread   
	    .vspdData.ReDraw = False    

	    .vspdData.MaxCols = C_DrCRFG + 1 												'☜: 최대 Columns의 항상 1개 증가시킴 
	    .vspdData.Col = .vspdData.MaxCols													'공통콘트롤 사용 Hidden Column
	    .vspdData.ColHidden = True
	    .vspdData.MaxRows = 0	    
		    
	    Call AppendNumberPlace("6","3","0")
	    Call GetSpreadColumnPos("A")

		ggoSpread.SSSetFloat  C_ItemSeq     , "NO"            , 6,"6",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,2,,,"0","999"         
	    ggoSpread.SSSetDate	 C_AdjustDt    , "청산일자"      ,11, 2, gDateFormat
	    ggoSpread.SSSetEdit	 C_AcctCd      , "계정코드"      ,11,,,20,2
	    ggoSpread.SSSetButton C_AcctPopUp
	    ggoSpread.SSSetEdit	 C_AcctNm      , "계정코드명"    ,20,,,20, 2
	    ggoSpread.SSSetFloat  C_AdjustAmt   , "청산금액"      ,       15, parent.ggAmtOfMoneyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
	    ggoSpread.SSSetFloat  C_AdjustLocAmt, "청산금액(자국)",15, parent.ggAmtOfMoneyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
	    ggoSpread.SSSetEdit   C_DocCur      , "화폐단위"      ,11,,,3,2
	    ggoSpread.SSSetButton C_DocCurPopUp
	    ggoSpread.SSSetEdit	 C_AdjustDESC  , "비고"          ,20,,,128	    
		ggoSpread.SSSetEdit	 C_Temp_GlNo   , "결의전표번호"  ,15,,,18,2	    
	    ggoSpread.SSSetEdit	 C_GlNo        , "전표번호"      ,11,,,18,2
	    ggoSpread.SSSetEdit	 C_RefNo       , "참조번호"      ,14,,,30, 2
	    ggoSpread.SSSetEdit   C_DrCRFG      ," "    ,5	    

		Call ggoSpread.MakePairsColumn(C_AcctCd,C_AcctPopUp)
        Call ggoSpread.SSSetColHidden(C_DocCur,C_DocCurPopUp,True)
	    Call ggoSpread.SSSetColHidden(C_DrCRFG,C_DrCRFG,True)	            
        
	    .vspdData.ReDraw = True
	End With
	
   	Call SetSpreadLock()
End Sub

'=======================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'=======================================================================================================
Sub  SetSpreadLock()
     With frm1
		ggoSpread.Source = .vspdData
		.vspdData.Redraw = False

     	ggoSpread.SpreadLock    C_ItemSeq     ,-1, C_ItemSeq
   		ggoSpread.SSSetRequired C_AdjustDt    ,-1, C_AdjustDt
   		ggoSpread.SpreadLock    C_AcctCd      ,-1, C_AcctCd
   		ggoSpread.SpreadLock    C_AcctPopUp   ,-1, C_AcctPopUp   						
		ggoSpread.SpreadLock    C_AcctNm      ,-1, C_AcctNm
	    ggoSpread.SSSetRequired C_AdjustAmt   ,-1, C_AdjustAmt 
		ggoSpread.SpreadUnLock  C_AdjustLocAmt,-1, C_AdjustLocAmt						
		ggoSpread.SpreadLock    C_DocCur      ,-1, C_DocCur
   		ggoSpread.SpreadLock    C_DocCurPopUp ,-1, C_DocCurPopUp
   		ggoSpread.SpreadUnLock  C_AdjustDESC  ,-1, C_AdjustDESC
   		ggoSpread.SpreadLock    C_Temp_GlNo   ,-1, C_Temp_GlNo   						
		ggoSpread.SpreadLock    C_GlNo        ,-1, C_GlNo
	    ggoSpread.SpreadLock    C_RefNo       ,-1, C_RefNo 
	
		.vspdData.Redraw = True
	End With
End Sub

'======================================================================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'=======================================================================================================
Sub  SetSpreadColor(ByVal pvStartRow,ByVal pvEndRow)
    With frm1
		ggoSpread.Source = .vspdData
		.vspdData.ReDraw = False
		
  		ggoSpread.SSSetProtected C_ItemSeq,     pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_AcctNm,      pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_Temp_GlNo,   pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_GlNo,        pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_RefNo,       pvStartRow, pvEndRow		
		ggoSpread.SSSetRequired  C_AcctCd,      pvStartRow, pvEndRow   
        ggoSpread.SSSetRequired  C_AdjustDt,    pvStartRow, pvEndRow 
        ggoSpread.SSSetRequired  C_AdjustAmt,   pvStartRow, pvEndRow
   
		.vspdData.ReDraw = True
    End With
End Sub

'======================================================================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'=======================================================================================================
Sub  SetSpread2ColorAP()
	Dim i

    With frm1
		ggoSpread.Source = .vspdData2
		.vspdData2.ReDraw = False
		                
		For i = 1 To .vspddata2.maxrows
					 
			ggoSpread.SSSetProtected C_DtlSeq, i, i
			ggoSpread.SSSetProtected C_CtrlCd, i, i
			ggoSpread.SSSetProtected C_CtrlNm, i, i
			ggoSpread.SSSetProtected C_CtrlValNm, i, i

			.vspddata2.Col = C_DrFg			
			If (.vspddata2.text = "C" ) _
                            Or .vspddata2.text = "Y" Or .vspddata2.text = "DC" Then
				ggoSpread.SSSetRequired C_CtrlVal, i, i	' 
			End If
		Next
		.vspdData2.ReDraw = True
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
             C_ItemSeq       = iCurColumnPos(1)
	         C_AdjustDt      = iCurColumnPos(2)
	         C_AcctCd        = iCurColumnPos(3)
	         C_AcctPopUp     = iCurColumnPos(4)
	         C_AcctNm		 = iCurColumnPos(5)
	         C_AdjustAmt     = iCurColumnPos(6)
	         C_AdjustLocAmt  = iCurColumnPos(7)
	         C_DocCur        = iCurColumnPos(8)
	         C_DocCurPopUp   = iCurColumnPos(9)
	         C_AdjustDESC	 = iCurColumnPos(10)
	         C_Temp_GlNo     = iCurColumnPos(11)
	         C_GlNo          = iCurColumnPos(12)
	         C_RefNo         = iCurColumnPos(13)
	         C_DrCRFG		 = iCurColumnPos(14)
    End Select    
End Sub

'=======================================================================================================
'	Name : OpenPopupGL()
'	Description : 
'=======================================================================================================
Function OpenPopupGL()
	Dim arrRet
	Dim arrParam(8)	
	Dim iCalledAspName
	Dim IntRetCD
	
	If IsOpenPop = True Then Exit Function
	
	iCalledAspName = AskPRAspName("A5120RA1")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "A5120RA1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	With frm1.vspdData
		.Row = .ActiveRow
		.Col = C_GlNo
		
		arrParam(0) = Trim(.Text)	'gl_no
		If CDbl(.Row) < 1 Then
 			arrParam(0) = ""
 		End If				
		arrParam(1) = ""			'Reference NO
	End With						

	IsOpenPop = True

	' 권한관리 추가 
	arrParam(5) = lgAuthBizAreaCd
	arrParam(6) = lgInternalCd
	arrParam(7) = lgSubInternalCd
	arrParam(8) = lgAuthUsrID
	   
   arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		     
	
	IsOpenPop = False
End Function

'=======================================================================================================
'	Name : OpenPopupTempGL()
'	Description : 
'=======================================================================================================
Function OpenPopuptempGL()
	Dim arrRet
	Dim arrParam(8)	
	Dim iCalledAspName
	Dim IntRetCD

	If IsOpenPop = True Then Exit Function
	
	iCalledAspName = AskPRAspName("A5130RA1")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "A5130RA1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	With frm1.vspdData
		.Row = .ActiveRow
		.Col = C_Temp_GlNo
		
		arrParam(0) = Trim(.Text)	'Temp_gl_no
		If CDbl(.Row) < 1 Then
 			arrParam(0) = ""
 		End If				
		arrParam(1) = ""			'Reference NO
	End With

	IsOpenPop = True

	' 권한관리 추가 
	arrParam(5) = lgAuthBizAreaCd
	arrParam(6) = lgInternalCd
	arrParam(7) = lgSubInternalCd
	arrParam(8) = lgAuthUsrID
		
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
   
	IsOpenPop = False
End Function

'=======================================================================================================
'	Name : OpenApNo()
'	Description : Prepayment No PopUp
'=======================================================================================================
Function OpenApNo()
	Dim arrRet
	Dim arrParam(8)	
	Dim iCalledAspName
	Dim IntRetCD
	
	If IsOpenPop = True Then Exit Function
	
	iCalledAspName = AskPRAspName("A4117RA1")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "A4117RA1", "X")
		IsOpenPop = False
		Exit Function
	End If
	IsOpenPop = True

	' 권한관리 추가 
	arrParam(5) = lgAuthBizAreaCd
	arrParam(6) = lgInternalCd
	arrParam(7) = lgSubInternalCd
	arrParam(8) = lgAuthUsrID
	  
  arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False	
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetApInfo(arrRet)
	End If	
End Function

'======================================================================================================
'   Function Name : SetApInfo(Byval arrRet)
'   Function Desc : 
'=======================================================================================================
Function SetApInfo(Byval arrRet)
	frm1.txtApNo.value = arrRet(0)			
	frm1.txtApNo.focus
End Function

'======================================================================================================
'   Function Name : OpenAcctCd(Byval strCode, Byval iWhere)
'   Function Desc : 
'======================================================================================================
Function  OpenAcctInfo(Byval strCode)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "계정코드팝업"										' 팝업 명칭 
	arrParam(1) = "A_ACCT,A_ACCT_GP"										' TABLE 명칭 
	arrParam(2) = Trim(strCode)							 					' Code Condition
	arrParam(3) = ""														' Name Cindition
	arrParam(4) = "A_ACCT.GP_CD = A_ACCT_GP.GP_CD AND A_ACCT.DEL_FG <> " & FilterVar("Y", "''", "S") & " "	' Where Condition
	arrParam(5) = "계정코드"			
		
	arrField(0) = "A_ACCT.ACCT_CD"											' Field명(0)
	arrField(1) = "A_ACCT.ACCT_NM"											' Field명(1)
	arrField(2) = "A_ACCT_GP.GP_CD"											' Field명(2)
	arrField(3) = "A_ACCT_GP.GP_NM"											' Field명(3)
	    
	arrHeader(0) = "계정코드"											' Header명(0)
	arrHeader(1) = "계정코드명"											' Header명(1)				
	arrHeader(2) = "그룹코드"										' Header명(2)
	arrHeader(3) = "그룹명"										' Header명(3)				
    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Call SetActiveCell(frm1.vspdData,C_AcctCd,frm1.vspdData.ActiveRow ,"M","X","X")
		Exit Function
	Else
		Call SetAcctCd(arrRet)
	End If	
End Function

'======================================================================================================
'   Function Name : SetAcctCd(Byval arrRet,byval iWhere)
'   Function Desc : 
'======================================================================================================
Function SetAcctCd(Byval arrRet)
	With frm1
		.vspdData.Col  = C_AcctCd
		.vspdData.Text = arrRet(0)
		.vspdData.Col  = C_AcctNm
		.vspdData.Text = arrRet(1)
			
		Call vspdData_Change(C_AcctCd, .vspddata.activerow)		 ' 변경이 읽어났다고 알려줌 
		Call SetActiveCell(.vspdData,C_AcctCd,.vspdData.ActiveRow ,"M","X","X")
	    lgBlnFlgChgValue = True
	End With
End Function

'=======================================================================================================
'	Name : OpenCurrency()
'	Description : Currency PopUp
'=======================================================================================================
Function OpenCurrency(Byval strCode)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
    dim field_fg
    
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "거래통화 팝업"	
	arrParam(1) = "B_CURRENCY"				
	arrParam(2) = Trim(frm1.txtDocCur.Value)
	arrParam(3) = ""
	arrParam(4) = ""
	arrParam(5) = "거래통화"
	
    arrField(0) = "CURRENCY"	
    arrField(1) = "CURRENCY_DESC"	
    
    arrHeader(0) = "거래통화"		
    arrHeader(1) = "거래통화명"
    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Call SetActiveCell(frm1.vspdData,C_DocCur,frm1.vspdData.ActiveRow ,"M","X","X")
		Exit Function
	Else
		Call SetCurrency(arrRet)
	End If	
End Function

'=======================================================================================================
'	Name : SetCurrency()
'	Description : Currency Popup에서 Return되는 값 setting
'=======================================================================================================
Function SetCurrency(ByVal arrRet)
	With frm1
		.vspdData.Col  = C_DocCur
		.vspdData.Text = arrRet(0)
			
		Call vspdData_Change(C_DocCur , .vspdData.Row)		 ' 변경이 읽어났다고 알려줌 
		Call SetActiveCell(.vspdData,C_DocCur,.vspdData.ActiveRow ,"M","X","X")
	    lgBlnFlgChgValue = True
	End With
End Function





'*******************************************************************************************************
'=======================================================================================================
' Area Name   : 4.5.2 Common Group-2
' Description : This part declares 2nd common function group
'=======================================================================================================
'*******************************************************************************************************





'======================================================================================================
'	Name : Form_Load()
'	Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 
'                 함수를 Call하는 부분 
'======================================================================================================
Sub  Form_Load()
    Call LoadInfTB19029()																'Load table , B_numeric_format

    Call ggoOper.ClearField(Document, "1")												'⊙: Condition field clear
    Call ggoOper.FormatField(Document, "1", ggStrIntegeralPart, ggStrDeciPointPart, _
							gDateFormat, parent.gComNum1000, parent.gComNumDec)
     Call ggoOper.FormatField(Document, "2", ggStrIntegeralPart, ggStrDeciPointPart, _
						gDateFormat, parent.gComNum1000, parent.gComNumDec)
	Call ggoOper.LockField(Document, "N")												'Lock  Suitable  Field

    Call InitSpreadSheet()																'Setup the Spread sheet
	Call InitCtrlSpread()
	Call InitCtrlHSpread()	    
    Call InitVariables()																'Initializes local global variables
    Call SetDefaultVal()
    
    Call SetToolbar("1110100100001111")													'버튼 툴바 제어 
    
    lgBlnFlgChgValue = False            
	frm1.txtApNo.focus
	

	' 권한관리 추가 
	Dim xmlDoc
	
	Call GetDataAuthXML(parent.gUsrID, gStrRequestMenuID, xmlDoc) 
	
	' 사업장 
	lgAuthBizAreaCd	= xmlDoc.selectSingleNode("/root/data/data_biz_area_cd").Text
	lgAuthBizAreaNm	= xmlDoc.selectSingleNode("/root/data/data_biz_area_nm").Text

	' 내부부서 
	lgInternalCd	= xmlDoc.selectSingleNode("/root/data/data_internal_cd").Text
	lgDeptCd		= xmlDoc.selectSingleNode("/root/data/data_dept_cd").Text
	lgDeptNm		= xmlDoc.selectSingleNode("/root/data/data_dept_nm").Text
	
	' 내부부서(하위포함)
	lgSubInternalCd	= xmlDoc.selectSingleNode("/root/data/data_sub_internal_cd").Text
	lgSubDeptCd		= xmlDoc.selectSingleNode("/root/data/data_sub_dept_cd").Text
	lgSubDeptNm		= xmlDoc.selectSingleNode("/root/data/data_sub_dept_nm").Text
	
	' 개인 
	lgAuthUsrID		= xmlDoc.selectSingleNode("/root/data/data_auth_usr_id").Text
	lgAuthUsrNm		= xmlDoc.selectSingleNode("/root/data/data_auth_usr_nm").Text
	
	Set xmlDoc = Nothing	
End Sub

'======================================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'======================================================================================================
Sub  Form_QueryUnload(Cancel , UnloadMode )
    
End Sub

'======================================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'======================================================================================================
Function  FncQuery() 
    Dim IntRetCD 
    Dim var1, var2
    
    FncQuery = False                                                        
    
    Err.Clear                                                               
	'-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then												'This function check indispensable field
		Exit Function
    End If
    '-----------------------
    'Check previous data area
    '-----------------------
    ggoSpread.Source = frm1.vspdData
    var1 = ggoSpread.SSCheckChange
    ggoSpread.Source = frm1.vspdData2
    var2 = ggoSpread.SSCheckChange
    
    If lgBlnFlgChgValue = True Or var1 = True Or var2 = True Then		
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO,"X","X")	    
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")									'Clear Contents  Field
    Call InitVariables()															'Initializes local global variables
	ggoSpread.Source = frm1.vspdData
	ggoSpread.ClearSpreadData
	ggoSpread.Source = frm1.vspdData2
	ggoSpread.ClearSpreadData		
	ggoSpread.Source = frm1.vspdData3
	ggoSpread.ClearSpreadData
    '-----------------------
    'Query function call area
    '-----------------------
    Call DbQuery()																	'☜: Query db data
    
    FncQuery = True																
    		
	Set gActiveElement = document.activeElement    

End Function

'=======================================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'========================================================================================================
Function  FncNew() 
    Dim IntRetCD 
	Dim var1, var2
	    
    FncNew = False                                                          
    
    ggoSpread.Source = frm1.vspdData
    var1 = ggoSpread.SSCheckChange
    ggoSpread.Source = frm1.vspdData2
    var2 = ggoSpread.SSCheckChange
	'-----------------------
    'Check previous data area
    '----------------------- 
    If lgBlnFlgChgValue = True Or var1 = True Or var2 = True Then
        IntRetCD = DisplayMsgBox("900015", parent.VB_YES_NO,"X","X")               
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
	'-----------------------
    'Erase condition area
    'Erase contents area
    '-----------------------	
    Call ggoOper.ClearField(Document, "1")									'⊙: Clear Condition Field    
    Call ggoOper.ClearField(Document, "2")									'Clear Condition Field
    Call ggoOper.LockField(Document, "N")									'Lock  Suitable  Field
    Call InitVariables()															'Initializes local global variables
    Call SetDefaultVal()    

    Call txtDocCur_OnChange()    
    
	ggoSpread.Source = frm1.vspdData
	ggoSpread.ClearSpreadData
	ggoSpread.Source = frm1.vspdData2
	ggoSpread.ClearSpreadData		
	ggoSpread.Source = frm1.vspdData3
	ggoSpread.ClearSpreadData

    frm1.txtApNo.Value = ""
    frm1.txtApNo.focus
    
    lgBlnFlgChgValue = False        
    
    FncNew = True                                                          
    		
	Set gActiveElement = document.activeElement    

End Function

'=======================================================================================================
' Function Name : FncDelete
' Function Desc : This function is related to Delete Button of Main ToolBar
'=======================================================================================================
Function  FncDelete() 
    Dim IntRetCD
    
    FncDelete = False                                                      
    '-----------------------
    'Precheck area
    '-----------------------
    If lgIntFlgMode <> parent.OPMD_UMODE Then										'Check if there is retrived data
        IntRetCD = DisplayMsgBox("900002","X","X","X")                               
        Exit Function
    End If
    
    IntRetCD = DisplayMsgBox("900003", parent.VB_YES_NO,"X","X")		            'Will you destory previous data"
	If IntRetCD = vbNo Then
		Exit Function
	End If
    '-----------------------
    'Delete function call area
    '-----------------------
    Call DbDelete()																	'☜: Delete db data
    
    FncDelete = True                                                        

    		
	Set gActiveElement = document.activeElement    

End Function

'=======================================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================================
Function  FncSave() 
    Dim IntRetCD 
	Dim var1,var2
	
    FncSave = False                                                         
    
    Err.Clear                                                               
	
    ggoSpread.Source = frm1.vspdData
    var1 = ggoSpread.SSCheckChange

    ggoSpread.Source = frm1.vspdData2
    var2 = ggoSpread.SSCheckChange

    If lgBlnFlgChgValue = False And var1 = False And var2 = False  Then				'⊙: Check If data is chaged
        IntRetCD = DisplayMsgBox("900001","X","X","X")								'⊙: Display Message(There is no changed data.)
		Exit Function		
    End If

	If Not chkField(Document, "2") Then												'⊙: Check required field(Single area)
		Exit Function
    End If
    
    ggoSpread.Source = frm1.vspdData
    If Not ggoSpread.SSDefaultCheck Then										'⊙: Check contents area
		Exit Function
    End If

	If CheckSpread3 = False then
	IntRetCD = DisplayMsgBox("110420","X","X","X")									'필수입력 check!!
        Exit Function
    End If
    '-----------------------
    'Save function call area
    '-----------------------
    Call DbSave()																	'☜: Save db data
    FncSave = True   
        		
	Set gActiveElement = document.activeElement    
                                                    
End Function

'=======================================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================================
Function  FncCopy() 
	Dim  IntRetCD
	 
	If frm1.vspdData.MaxRows < 1 Then exit Function
	 
	frm1.vspdData.ReDraw = False	
	With frm1
		.vspdData.ReDraw = False
	
		ggoSpread.Source = frm1.vspdData	
		ggoSpread.CopyRow
		
		Call MaxSpreadVal(frm1.vspdData, C_ItemSeq , frm1.vspdData.ActiveRow)
		Call SetSpreadColor(frm1.vspdData.ActiveRow,  frm1.vspdData.ActiveRow)
    
		.vspdData.ReDraw = True
	End With
	    		
	Set gActiveElement = document.activeElement    

End Function

'=======================================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================================
Function  FncCancel() 
    If frm1.vspdData.MaxRows < 1 Then exit Function
    
    With frm1.vspdData
        .Row = .ActiveRow
        .Col = 0
        If .Text = ggoSpread.InsertFlag Then
			.Col = C_AcctCd
			If Len(Trim(.text)) > 0 Then  
				.Col = C_ItemSeq
				Call DeleteHSheet(.Text)
			End If
        End If
   
        ggoSpread.Source = frm1.vspdData	
        ggoSpread.EditUndo
		Call DoSum()
		If frm1.vspdData.MaxRows < 1 Then exit Function
		
        .Row = .ActiveRow
        .Col = 0
        
		If .Row = 0 Then 
			Exit Function
		End If

        If .Text = ggoSpread.InsertFlag Then
			.Col = C_AcctCd
			If Len(Trim(.text)) > 0 Then 
				.Col = C_ItemSeq
				frm1.hItemSeq.value = .Text
				ggoSpread.Source = frm1.vspdData2
				ggoSpread.ClearSpreadData
				Call DbQuery3(.ActiveRow)
			End If	
        Else
            .Col = C_ItemSeq
            frm1.hItemSeq.value = .Text
			ggoSpread.Source = frm1.vspdData2
			ggoSpread.ClearSpreadData
            Call DbQuery2(.ActiveRow)
        End If
    End With
    		
	Set gActiveElement = document.activeElement    

End Function

'=======================================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================================
Function  FncInsertRow(Byval pvRowcnt) 
	Dim iCurRowPos
	Dim imRow
    Dim ii
    
	On Error Resume Next                                                        '☜: If process fails
    Err.Clear   
	
    FncInsertRow = False														'☜: Processing is NG

    If IsNumeric(Trim(pvRowcnt)) Then 
       imRow  = Cint(pvRowcnt)
    Else
		imRow = AskSpdSheetAddRowCount()
		If imRow = "" Then
		    Exit Function
		End If
    End If
        
	With frm1.vspdData
		iCurRowPos = .ActiveRow
        .ReDraw = False
        ggoSpread.Source = frm1.vspdData
		ggoSpread.InsertRow ,imRow
		
		For ii = .ActiveRow To  .ActiveRow + imRow - 1
			Call MaxSpreadVal(frm1.vspdData, C_ItemSeq, ii)
		Next  
		.Col = 2																' 컬럼의 절대 위치로 이동      
		.Row = 	ii - 1
		.Action = 0
		
		.Col = C_DrCRFG
		.Row = ii
		.Text = "CR"		

        Call SetSpreadColor(iCurRowPos + 1, iCurRowPos + imRow)        
        .ReDraw = True
    End With

    Call ggoOper.LockField(Document, "I")								'This function lock the suitable field
    
    If Err.number = 0 Then
       FncInsertRow = True                                                      '☜: Processing is OK
    End If   
	
    Set gActiveElement = document.ActiveElement   
End Function

'=======================================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================================
Function  FncDeleteRow() 
   	Dim lDelRows
	Dim iDelRowCnt, i
    Dim DelItemSeq

	If frm1.vspdData.MaxRows < 1 Then Exit Function
	
    With frm1.vspdData 
        .Row = .ActiveRow
		.Col = C_ItemSeq 
	    DelItemSeq = .Text
    
    	ggoSpread.Source = frm1.vspdData 

    	lDelRows = ggoSpread.DeleteRow
   End With

    Call DeleteHsheet(DelItemSeq)
	Call DoSum()
	    		
	Set gActiveElement = document.activeElement    

End Function

'=======================================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================================
Function  FncPrint() 
    On Error Resume Next        
    parent.FncPrint()      
        		
	Set gActiveElement = document.activeElement    
                                 
End Function

'=======================================================================================================
' Function Name : FncPrev
' Function Desc : This function is related to Previous Button
'========================================================================================================
Function  FncPrev() 
    On Error Resume Next                                               
End Function

'=======================================================================================================
' Function Name : FncNext
' Function Desc : This function is related to Next Button
'=======================================================================================================
Function  FncNext() 
    On Error Resume Next                                               
End Function

'=======================================================================================================
' Function Name : FncFind
' Function Desc : 화면 속성, Tab유무 
'========================================================================================================
Function  FncFind() 
    Call parent.FncFind(parent.C_SINGLEMULTI , True)    
        		
	Set gActiveElement = document.activeElement    
                      
End Function

'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================
Function  FncExcel() 
	Call parent.FncExport(parent.C_SINGLEMULTI)
	    		
	Set gActiveElement = document.activeElement    

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

'=======================================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================================
Function  FncExit()
	Dim IntRetCD
	Dim var1,var2
	
	FncExit = False

	ggoSpread.Source = frm1.vspdData
    var1 = ggoSpread.SSCheckChange

    ggoSpread.Source = frm1.vspdData2
    var2 = ggoSpread.SSCheckChange
    
    If lgBlnFlgChgValue = True or var1 = True Or var2 = True Then					'⊙: Check If data is chaged
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO,"X","X")                '데이타가 변경되었습니다. 종료 하시겠습니까?
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

    FncExit = True
        		
	Set gActiveElement = document.activeElement    

End Function





'*******************************************************************************************************
'=======================================================================================================
' Area Name   : 4.5.3 Common Group - 3
' Description : This part declares 3rd common function group
'=======================================================================================================
'*******************************************************************************************************





'=======================================================================================================
' Function Name : DbDelete
' Function Desc : This function delete data
'========================================================================================================
Function  DbDelete() 
    DbDelete = True                                                         
End Function

'=======================================================================================================
' Function Name : DbDeleteOk
' Function Desc : DbDelete가 성공적일때 수행 
'========================================================================================================
Function DbDeleteOk()			
	Call ggoOper.ClearField(Document, "2")									'Clear Condition Field
    Call ggoOper.LockField(Document, "N")									'Lock  Suitable  Field
    Call InitVariables()															'Initializes local global variables
    
	ggoSpread.Source = frm1.vspdData
	ggoSpread.ClearSpreadData
	ggoSpread.Source = frm1.vspdData2
	ggoSpread.ClearSpreadData		
	ggoSpread.Source = frm1.vspdData3
	ggoSpread.ClearSpreadData
End Function

'=======================================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'=======================================================================================================
Function DbQueryOk()																'☆: 조회 성공후 실행로직	
	With frm1
        '-----------------------
        'Reset variables area
        '-----------------------
        lgIntFlgMode = parent.OPMD_UMODE											'Indicates that current mode is Update mode
        
        Call ggoOper.LockField(Document, "Q")								'This function lock the suitable field
        Call SetToolbar("1110111100001111")											'버튼 툴바 제어 
		
        If .vspdData.MaxRows > 0 Then
            .vspdData.Row = 1
            .vspdData.Col = C_ItemSeq
            .hItemSeq.Value = .vspdData.Text 
            Call DbQuery2(1)
        End If
    End With
    lgBlnFlgChgValue = FALSE

	Call DoSum()
	Call txtDocCur_OnChange()
	
	lgBlnFlgChgValue = False	
End Function

'=======================================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'=======================================================================================================
Function  DbQuery() 
    Dim strVal

    DbQuery = False                                                             
    Call LayerShowHide(1)
   
    With frm1
		If lgIntFlgMode = parent.OPMD_UMODE Then
			strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001					'☜: 
			strVal = strVal & "&txtApNo=" & Trim(.txtApNo.value)	'조회 조건 데이타 
			strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
			strVal = strVal & "&lgStrPrevKeyDtl=" & lgStrPrevKeyDtl
			strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
		Else
			strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001					'☜: 
			strVal = strVal & "&txtApNo=" & Trim(.txtApNo.value)		'조회 조건 데이타 
			strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
			strVal = strVal & "&lgStrPrevKeyDtl=" & lgStrPrevKeyDtl
			strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
		End If
    End With

	' 권한관리 추가 
	strVal = strVal & "&lgAuthBizAreaCd="	& lgAuthBizAreaCd			' 사업장 
	strVal = strVal & "&lgInternalCd="		& lgInternalCd				' 내부부서 
	strVal = strVal & "&lgSubInternalCd="	& lgSubInternalCd			' 내부부서(하위포함)
	strVal = strVal & "&lgAuthUsrID="		& lgAuthUsrID				' 개인 

	Call RunMyBizASP(MyBizASP, strVal)										    '☜: 비지니스 ASP 를 가동 

    DbQuery = True                                                              
End Function

'=======================================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function
'=======================================================================================================
Function  DbQueryOk1()
     With frm1
        '-----------------------
        'Reset variables area
        '-----------------------
        lgIntFlgMode = parent.OPMD_UMODE												'Indicates that current mode is Update mode
        
        Call ggoOper.LockField(Document, "Q")									'This function lock the suitable field
        Call SetToolbar("1110111100001111")											'버튼 툴바 제어 
        
        If .vspdData.MaxRows > 0 Then
            .vspdData.Row = 1
            .vspdData.Col = C_ItemSeq
            .hItemSeq.Value = .vspdData.Text 
            Call DbQuery2(1)
        End If
    End With
End Function

'=======================================================================================================
' Function Name : DbSave
' Function Desc : This function is data query and display
'=======================================================================================================
Function  DbSave() 
    Dim lngRows 
    Dim lGrpcnt
    DIM strVal     
    Dim strDel
    Dim RowD
    DIM GrpCntD
    DIM strValD
    DIM strItemSEQ	'관리항목 파라미터 
    
	DbSave = False                                                          
    Call LayerShowHide(1)
    
    On Error Resume Next                                                   
	Err.Clear 
	
	With frm1
		.txtFlgMode.value = lgIntFlgMode
		.txtMode.value = parent.UID_M0002
	End With

    '-----------------------
    'Data manipulate area
    '-----------------------
    ' Data 연결 규칙 
    ' 0: Sheet명, 1: Flag , 2: Row위치, 3~N: 각 데이타 

    lGrpCnt = 1
    strVal = ""
    strDel = ""
    
    GrpCntD = 1: strValD = ""	'관리항목 파라미터 
    
    ggoSpread.Source = frm1.vspdData
	With frm1.vspdData
		For lngRows = 1 To .MaxRows
			.Row = lngRows
			.Col = 0

			If .Text = ggoSpread.InsertFlag Then
				strVal = strVal & "C" & parent.gColSep & lngRows & parent.gColSep				'C=Create, Sheet가 2개 이므로 구별 
			ElseIf .Text = ggoSpread.UpdateFlag Then
				strVal = strVal & "U" & parent.gColSep & lngRows & parent.gColSep				'U=Update
			ElseIf .Text = ggoSpread.DeleteFlag Then
				strVal = strVal & "D" & parent.gColSep & lngRows & parent.gColSep				'D=Delete
			End If
	
			Select Case .Text
			    Case ggoSpread.InsertFlag, ggoSpread.UpdateFlag
			    
			        .Col = C_ItemSeq		'2
			        strVal = strVal & Trim(.Text) & parent.gColSep
			        strItemSEQ = Trim(.Text)	'ITEM_SEQ 
			        .Col = C_AdjustDt		'3
			        strVal = strVal & UNIConvDate(Trim(.Text)) & parent.gColSep
					.Col = C_AcctCd		'4
			        strVal = strVal & Trim(.Text) & parent.gColSep
			        .Col = C_AdjustAmt	'5
			        strVal = strVal & UNIConvNum(Trim(.Text),0) & parent.gColSep
			        .Col = C_AdjustLocAmt	'6
			        strVal = strVal & UNIConvNum(Trim(.Text),0) & parent.gColSep		        
			        .Col = C_DocCur     '7
			        strVal = strVal & frm1.txtDocCur.value & parent.gColSep
   			        .Col = C_RefNo		'8
			        strVal = strVal & Trim(.Text) & parent.gColSep
			        .Col = C_AdjustDESC		'9
			        strVal = strVal & Trim(.Text) & parent.gRowSep		        
			        
			        lGrpCnt = lGrpCnt + 1
			        
			        '=======================================================================
			        '2001.06.18 Song,MunGil 관리항목도 입력/수정된 걸로 간주하고 생성함.
			        '=======================================================================
			        With frm1.vspdData3
						For RowD = 1 To .MaxRows
							.Row = RowD
							.Col = C_DtlSeq
							
							If strItemSEQ = Trim(.Text) Then
								strValD = strValD & "C" & parent.gColSep & RowD & parent.gColSep
								.Col = 1 
								strValD = strValD & Trim(.Text) & parent.gColSep
								.Col = 2
								strValD = strValD & Trim(.Text) & parent.gColSep
								.Col = 3
								strValD = strValD & Trim(.Text) & parent.gColSep
								.Col = 5
								strValD = strValD & Trim(.Text) & parent.gRowSep
								
								GrpCntD = GrpCntD + 1
							End If
						Next
					End With				
			    Case ggoSpread.DeleteFlag
					
			        .Col = C_ItemSeq		'2
			        strVal = strVal & Trim(.Text) & parent.gColSep				        '마지막 데이타는 Row 분리기호를 넣는다 
			        .Col = C_RefNo		'3
			        strVal = strVal & Trim(.Text) & parent.gRowSep

					lGrpcnt = lGrpcnt + 1             
			End Select
		Next
	End With
	
	frm1.txtMaxRows.value = lGrpCnt-1										'Spread Sheet의 변경된 최대갯수 
	frm1.txtSpread.value =  strDel & strVal									'Spread Sheet 내용을 저장 
		
	frm1.txtMaxRows3.value = GrpCntD - 1									'Spread Sheet의 변경된 최대갯수 
	frm1.txtSpread3.value  = strValD				

	'권한관리추가 start
	frm1.txthAuthBizAreaCd.value =  lgAuthBizAreaCd
	frm1.txthInternalCd.value =  lgInternalCd
	frm1.txthSubInternalCd.value = lgSubInternalCd
	frm1.txthAuthUsrID.value = lgAuthUsrID		
	'권한관리추가 end
		
	Call ExecMyBizASP(frm1, BIZ_PGM_ID)										'저장 비지니스 ASP 를 가동 

    DbSave = True
End Function

'=======================================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncSave에 있는것을 옮김 
'=======================================================================================================
Function  DbSaveOk()													'☆: 저장 성공후 실행 로직 
    ggoSpread.SSDeleteFlag 1
    
	Call ggoOper.ClearField(Document, "2")										'Clear Contents  Field
    Call InitVariables()															'Initializes local global variables
	ggoSpread.Source = frm1.vspdData
	ggoSpread.ClearSpreadData
	ggoSpread.Source = frm1.vspdData2
	ggoSpread.ClearSpreadData		
	ggoSpread.Source = frm1.vspdData3
	ggoSpread.ClearSpreadData
	CAll DbQuery()
End Function




'*******************************************************************************************************
'=======================================================================================================
' Area Name   : 4.5.4 User-defined Method Part
' Description : This part declares user-defined method
'=======================================================================================================
'*******************************************************************************************************




'=======================================================================================================
' Function Name : DbQuery2																				
' Function Desc : This function is data query and display												
'=======================================================================================================
Function DbQuery2(ByVal Row)
	Dim strVal	
	Dim lngRows
	Dim strSelect
	Dim strFrom
	Dim strWhere 	
	Dim strTableid
	Dim strColid
	Dim strColNm	
	Dim strMajorCd	
	Dim strNmwhere
	Dim i,Indx1
	Dim arrVal,arrTemp
	
	Err.Clear
	
	With frm1
	    .vspdData.Row = Row
	    .vspdData.Col = C_ItemSeq
	    .hItemSeq.Value = .vspdData.Text

	    If Trim(.hItemSeq.Value) = "" Then
	        Exit Function
	    End If
	    
        If CopyFromData(.hItemSeq.Value) = True Then
            Call SetSpread2ColorAp() 	
            Exit Function
        End If
    	
		Call LayerShowHide(1)
	
		DbQuery2 = False
		
		.vspdData.Row = Row
		.vspdData.Col = C_ItemSeq
		
		strSelect =				" C.DTL_SEQ,  A.CTRL_CD, A.CTRL_NM , C.CTRL_VAL, '',"
		strSelect = strSelect & " CASE  WHEN A.ColM_DATA_TYPE = " & FilterVar("D", "''", "S") & "  THEN " & FilterVar("(Format : YYYY-MM-DD)", "''", "S") & "  END , "
		strSelect = strSelect & " " & .VSPDDATA.value & " , LTrim(ISNULL(A.TBL_ID,'')), LTrim(ISNULL(A.DATA_COLM_ID,'')), "
		strSelect = strSelect & " LTrim(ISNULL(A.DATA_COLM_NM,'')),  LTrim(ISNULL(A.ColM_DATA_TYPE,'')), LTrim(ISNULL(A.DATA_LEN,'')), "
		strSelect = strSelect & " CASE WHEN B.DR_FG = " & FilterVar("Y", "''", "S") & "  AND  B.CR_FG = " & FilterVar("Y", "''", "S") & "  THEN " & FilterVar("DC", "''", "S") & "  "
		strSelect = strSelect & " WHEN B.DR_FG = " & FilterVar("Y", "''", "S") & "  AND  B.CR_FG = " & FilterVar("N", "''", "S") & "  THEN " & FilterVar("D", "''", "S") & "  "
		strSelect = strSelect & " WHEN B.DR_FG = " & FilterVar("N", "''", "S") & "  AND  B.CR_FG = " & FilterVar("Y", "''", "S") & "  THEN " & FilterVar("C", "''", "S") & "  "
		strSelect = strSelect & " END	, " & .hItemSeq.Value & ",  "
		strSelect = strSelect & " LTrim(ISNULL(A.MAJOR_CD,'')), CHAR(8) "
  		
		strFrom = " A_CTRL_ITEM	A (NOLOCK), A_ACCT_CTRL_ASSN B (NOLOCK), A_AP_ADJUST_DTL C (NOLOCK), A_AP_ADJUST D (NOLOCK) "
		
		.vspdData.Col = C_REFNo
					
		strWhere =			  " D.ADJUST_NO =  " & FilterVar(UCase(.VSPDDATA.value), "''", "S") & " "		
		strWhere = strWhere & " AND D.ADJUST_NO  =  C.ADJUST_NO  "		
		strWhere = strWhere & "	AND D.ACCT_CD = B.ACCT_CD "
		strWhere = strWhere & " AND C.CTRL_CD = B.CTRL_CD "
		strWhere = strWhere & " AND D.ACCT_CD = B.ACCT_CD "
		strWhere = strWhere & " AND B.CTRL_CD = A.CTRL_CD "
		strWhere = strWhere & " ORDER BY C.DTL_SEQ "
   					
		frm1.vspdData2.ReDraw = False
			
		If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) Then   
			ggoSpread.Source = frm1.vspdData2
			arrTemp =  Split(lgF2By2,Chr(12))
			For Indx1 = 0 To Ubound(arrTemp) - 1
				arrTemp(indx1) = Replace(arrTemp(indx1), Chr(8), indx1 + 1)
			Next
			lgF2By2 = Join(arrTemp,Chr(12))				
			ggoSpread.SSShowData lgF2By2							
			
			For lngRows = 1 To frm1.vspdData2.Maxrows
				frm1.vspddata2.Row = lngRows	
				frm1.vspddata2.Col = C_Tableid 
				If Trim(frm1.vspddata2.text) <> "" Then
					frm1.vspddata2.Col = C_Tableid
					strTableid = frm1.vspddata2.text
					frm1.vspddata2.Col = C_Colid
					strColid = frm1.vspddata2.text
					frm1.vspddata2.Col = C_ColNm
					strColNm = frm1.vspddata2.text	
					frm1.vspddata2.Col = C_MajorCd					
					strMajorCd = frm1.vspddata2.text	
					
					frm1.vspddata2.Col = C_CtrlVal
					
					strNmwhere = strColid & " =   " & FilterVar(frm1.vspddata2.text , "''", "S") & " " 
					
					If Trim(strMajorCd) <> "" Then
						strNmwhere = strNmwhere & " AND MAJOR_CD =  " & FilterVar(strMajorCd , "''", "S") & " "
					End If				 
					
					If CommonQueryRs( strColNm , strTableid ,  strNmwhere , lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then   
						frm1.vspddata2.Col = C_CtrlValNm
						arrVal = Split(lgF0, Chr(11))  
						frm1.vspddata2.text = arrVal(0)
					End If
				End If								
				
				strVal = strVal & Chr(11) & .hItemSeq.Value 

			    frm1.vspddata2.Col = C_DtlSeq
				strVal = strVal & Chr(11) & frm1.vspddata2.text
				frm1.vspddata2.Col = C_CtrlCd
				strVal = strVal & Chr(11) & frm1.vspddata2.text
				frm1.vspddata2.Col = C_CtrlNm
				strVal = strVal & Chr(11) & frm1.vspddata2.text
				frm1.vspddata2.Col = C_CtrlVal
				strVal = strVal & Chr(11) & frm1.vspddata2.text
				frm1.vspddata2.Col = C_CtrlPB
				strVal = strVal & Chr(11) & frm1.vspddata2.text
				frm1.vspddata2.Col = C_CtrlValNm
				strVal = strVal & Chr(11) & frm1.vspddata2.text
				frm1.vspddata2.Col = C_Seq
				strVal = strVal & Chr(11) & frm1.vspddata2.text
				frm1.vspddata2.Col = C_Tableid
				strVal = strVal & Chr(11) & frm1.vspddata2.text
				frm1.vspddata2.Col = C_Colid
				strVal = strVal & Chr(11) & frm1.vspddata2.text
				frm1.vspddata2.Col = C_ColNm
				strVal = strVal & Chr(11) & frm1.vspddata2.text
				frm1.vspddata2.Col = C_Datatype
				strVal = strVal & Chr(11) & frm1.vspddata2.text
				frm1.vspddata2.Col = C_DataLen
				strVal = strVal & Chr(11) & frm1.vspddata2.text
				frm1.vspddata2.Col = C_DRFg
				strVal = strVal & Chr(11) & frm1.vspddata2.text
				frm1.vspddata2.Col = C_MajorCd
				strVal = strVal & Chr(11) & frm1.vspddata2.text
				frm1.vspdData2.Col = C_MajorCd+1 				
				.vspdData2.Text = lngRows
				strVal = strVal & Chr(11) & frm1.vspddata2.text						
											
				strVal = strVal & Chr(11) & Chr(12)									
			Next					
			ggoSpread.Source = frm1.vspdData3
			ggoSpread.SSShowData strVal	
		End If 		
	
		Call SetSpread2ColorAp() 	
	End With
	
	Call LayerShowHide(0)
	
	frm1.vspdData2.ReDraw = True
	
	DbQuery2 = True
End Function

'=======================================================================================================
' Function Name : DbQueryOk2
' Function Desc : DbQuery2가 성공적일 경우 MyBizASP 에서 호출되는 Function
'=======================================================================================================
Function  DbQueryOk2()
	Call SetSpread2ColorAP()
    Call txtDocCur_OnChange()
    
    lgBlnFlgChgValue = False    
End Function

'======================================================================================================
'   Name : DoSum()
'   Desc : Sum sheet Data
'=======================================================================================================
Sub DoSum()
	Dim dblToAdjustAmt
	Dim dblToAdjustLocAmt

	With frm1.vspdData
		dblToAdjustAmt = FncSumSheet1(frm1.vspdData,C_AdjustAmt, 1, .MaxRows, False, -1, -1, "V")
		dblToAdjustLocAmt = FncSumSheet1(frm1.vspdData,C_AdjustLocAmt, 1, .MaxRows, False, -1, -1, "V")
		
		If CommonQueryRs( "CURRENCY_DESC" , "B_CURRENCY" ,  " CURRENCY =  " & FilterVar(frm1.txtDocCur.value , "''", "S") & "" , lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then   			        							
			frm1.txtTotAdjustAmt.text	= UNIConvNumPCToCompanyByCurrency(dblToAdjustAmt,frm1.txtDocCur.value,parent.ggAmtOfMoneyNo, "X", "X")
		End If	
        frm1.txtTotAdjustLocAmt.text = UNIConvNumPCToCompanyByCurrency(dblToAdjustLocAmt,parent.gCurrency,parent.ggAmtOfMoneyNo, parent.gLocRndPolicyNo, "X")		
	End With	
End Sub

'==========================================================================================
'   Event Name : txtDocCur_OnChange
'   Event Desc : 
'==========================================================================================
Sub txtDocCur_OnChange()
    lgBlnFlgChgValue = True
    If CommonQueryRs( "CURRENCY_DESC" , "B_CURRENCY" ,  " CURRENCY =  " & FilterVar(frm1.txtDocCur.value , "''", "S") & "" , lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then   			        							
		Call CurFormatNumericOCX()
		Call CurFormatNumSprSheet()
	End If	   
End Sub

'===================================== CurFormatNumericOCX()  =======================================
'	Name : CurFormatNumericOCX()
'	Description : 화면에서 일괄적으로 Rounding되는 Numeric OCX
'====================================================================================================
Sub CurFormatNumericOCX()
	With frm1
		' 채무액 
		ggoOper.FormatFieldByObjectOfCur .txtApAmt, .txtDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, gDateFormat, parent.gComNum1000, parent.gComNumDec
		' 잔액 
		ggoOper.FormatFieldByObjectOfCur .txtBalAmt, .txtDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, gDateFormat, parent.gComNum1000, parent.gComNumDec
		' 청산금액 
		ggoOper.FormatFieldByObjectOfCur .txtTotAdjustAmt, .txtDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, gDateFormat, parent.gComNum1000, parent.gComNumDec
	End With
End Sub

'===================================== CurFormatNumSprSheet()  ======================================
'	Name : CurFormatNumSprSheet()
'	Description : 화면에서 일괄적으로 Rounding되는 Numeric Spread Sheet
'====================================================================================================
Sub CurFormatNumSprSheet()
	With frm1
		ggoSpread.Source = frm1.vspdData
		' 청산금액 
		ggoSpread.SSSetFloatByCellOfCur C_AdjustAmt,-1, .txtDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gComNum1000, parent.gComNumDec
	End With
End Sub

'=======================================================================================================
'   Event Name : InputCtrlVal
'   Event Desc :
'=======================================================================================================  
Sub InputCtrlVal(ByVal Row)
	Dim strAcctCd		
	Dim ii
			
	lgBlnFlgChgValue = True
		
	ggoSpread.Source = frm1.vspdData
	frm1.vspdData.Col = C_AcctCd
	frm1.vspdData.Row = Row		
	strAcctCd	= Trim(frm1.vspdData.text)		
		
	frm1.vspdData.Col = C_AdjustDt
	frm1.vspdData.Row = Row			
		
	Call AutoInputDetail(strAcctCd,Trim(frm1.txtDeptCd.value), frm1.vspdData.text, Row)
	For ii = 1 To frm1.vspdData2.MaxRows
		frm1.vspddata2.Col = C_CtrlVal
		frm1.vspddata2.Row = ii
					
		If Trim(frm1.vspddata2.text) <> "" Then
			Call CopyToHSheet2(frm1.vspdData.ActiveRow,ii)			 			
		End if
	Next
End Sub




'*******************************************************************************************************
'=======================================================================================================
' Area Name   : 4.5.5 Spread Popup method 
' Description : This part declares spread popup method
'=======================================================================================================
'*******************************************************************************************************




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
Sub  PopRestoreSpreadColumnInf()
	Dim indx

	On Error Resume Next
	Err.Clear 		
	
	ggoSpread.Source = gActiveSpdSheet
	Select Case Trim(UCase(gActiveSpdSheet.Name))
		Case "VSPDDATA" 
			Call PrevspdDataRestore(gActiveSpdSheet)
			Call ggoSpread.RestoreSpreadInf()
			Call InitSpreadSheet()
			Call ggoSpread.ReOrderingSpreadData()
			Call SetSpreadLock()
			Call SetSpread2ColorAp()						
		Case "VSPDDATA2"
			Call PrevspdData2Restore(gActiveSpdSheet)   
			Call ggoSpread.RestoreSpreadInf()
			Call InitCtrlSpread()			'관리항목 그리드 초기화 
			Call ggoSpread.ReOrderingSpreadData()
			Call SetSpread2ColorAp()  
	End Select
	
	If frm1.vspdData2.MaxRows <= 0 Then
		Call DbQuery2(frm1.vspdData.ActiveRow)
	End If	
End Sub

'===================================== PrevspdDataRestore()  ========================================
' Name : PrevspdDataRestore()
' Description : 그리드 복원시 관리항목 복원 
'====================================================================================================
Sub PrevspdDataRestore(pActiveSheetName)
	Dim indx, indx1

	For indx = 0 To frm1.vspdData.MaxRows
        frm1.vspdData.Row    = indx
        frm1.vspdData.Col    = 0
		
		If frm1.vspdData.Text <> "" Then
			Select Case frm1.vspdData.Text			
				Case ggoSpread.InsertFlag					
					frm1.vspdData.Col = C_ItemSeq					
					Call DeleteHsheet(frm1.vspdData.Text)					
				Case ggoSpread.UpdateFlag		
					For indx1 = 0 To frm1.vspdData3.MaxRows					
						frm1.vspdData3.Row = indx1
						frm1.vspdData3.Col = 0
						Select Case frm1.vspdData3.Text 
							Case ggoSpread.UpdateFlag
								frm1.vspdData.Col = C_ItemSeq
								frm1.vspdData3.Col = 1					
								If UCase(Trim(frm1.vspdData.Text)) = UCase(Trim(frm1.vspdData3.Text)) Then
									Call DeleteHsheet(frm1.vspdData.Text)	
									frm1.vspdData.Col = C_REFNo																		
									Call FncRestoreDbQuery2(indx, frm1.vspdData.ActiveRow, UCase(Trim(frm1.vspddata.value)))
								End If
						End Select
					Next
				Case ggoSpread.DeleteFlag
					frm1.vspdData.Col = C_REFNo
					Call fncRestoreDbQuery2(indx, frm1.vspdData.ActiveRow, UCase(Trim(frm1.vspddata.value)))
			End Select
		End If
	Next
	
	ggoSpread.Source = pActiveSheetName
End Sub

'===================================== PrevspdDataRestore()  ========================================
' Name : PrevspdData2Restore()
' Description : 그리드 복원시 관리항목 복원 
'====================================================================================================
Sub PrevspdData2Restore(pActiveSheetName)
	Dim indx, indx1

	For indx = 0 to frm1.vspdData2.MaxRows
        frm1.vspdData2.Row    = indx
        frm1.vspdData2.Col    = 0

		If frm1.vspdData2.Text <> "" Then
			Select Case frm1.vspdData2.Text
				Case ggoSpread.InsertFlag
					frm1.vspdData2.Col = C_HItemSeq
					For indx1 = 0 to frm1.vspdData.MaxRows
						frm1.vspdData.Row = indx1
						frm1.vspdData.Col = C_ItemSeq
						If frm1.vspdData.Text = frm1.vspdData2.Text Then
							Call DeleteHsheet(frm1.vspdData.Text)
							ggoSpread.Source = frm1.vspdData	
					        ggoSpread.EditUndo							
						End If
					Next
				Case ggoSpread.UpdateFlag
					frm1.vspdData2.Col = C_HItemSeq
					For indx1 = 0 to frm1.vspdData.MaxRows
						frm1.vspdData.Row = indx1
						frm1.vspdData.Col = C_ItemSeq
						If frm1.vspdData.Text = frm1.vspdData2.Text Then
							Call DeleteHsheet(frm1.vspdData.Text)
							ggoSpread.Source = frm1.vspdData
							ggoSpread.EditUndo
							frm1.vspdData.Col = C_REFNo
							Call fncRestoreDbQuery2(indx1, frm1.vspdData.ActiveRow, UCase(Trim(frm1.vspddata.value))) 
						End If
					Next
				Case ggoSpread.DeleteFlag

			End Select
		End If
	Next
	ggoSpread.Source = pActiveSheetName
End Sub

'========================================================================================================
' Name : fncRestoreDbQuery2																				
' Desc : This function is data query and display												
'========================================================================================================
Function fncRestoreDbQuery2(Row, CurrRow, Byval pInvalue1)
	Dim strItemSeq
	Dim strSelect, strFrom, strWhere
	Dim arrTempRow, arrTempCol
	Dim Indx1
	Dim strTableid, strColid, strColNm, strMajorCd
	Dim strNmwhere
	Dim arrVal
	Dim strVal

	On Error Resume Next
	Err.Clear

	fncRestoreDbQuery2 = False

	Call DisableToolBar(parent.TBC_QUERY)
	Call LayerShowHide(1)
	With frm1
		.vspdData.row = Row
	    .vspdData.col = C_ItemSeq
		strItemSeq    = .vspdData.Text

	    If Trim(strItemSeq) = "" Then
	        Exit Function
	    End If
    	
		Call LayerShowHide(1)
	
		DbQuery2 = False
		
		.vspdData.Row = Row
		.vspdData.Col = C_ItemSeq
		
		strSelect =				" C.DTL_SEQ,  A.CTRL_CD, A.CTRL_NM , C.CTRL_VAL, '',"
		strSelect = strSelect & " CASE  WHEN A.ColM_DATA_TYPE = " & FilterVar("D", "''", "S") & "  THEN " & FilterVar("(Format : YYYY-MM-DD)", "''", "S") & "  END , "
		strSelect = strSelect & " " & .VSPDDATA.value & " , LTrim(ISNULL(A.TBL_ID,'')), LTrim(ISNULL(A.DATA_COLM_ID,'')), "
		strSelect = strSelect & " LTrim(ISNULL(A.DATA_COLM_NM,'')),  LTrim(ISNULL(A.ColM_DATA_TYPE,'')), LTrim(ISNULL(A.DATA_LEN,'')), "
		strSelect = strSelect & " CASE WHEN B.DR_FG = " & FilterVar("Y", "''", "S") & "  AND  B.CR_FG = " & FilterVar("Y", "''", "S") & "  THEN " & FilterVar("DC", "''", "S") & "  "
		strSelect = strSelect & " WHEN B.DR_FG = " & FilterVar("Y", "''", "S") & "  AND  B.CR_FG = " & FilterVar("N", "''", "S") & "  THEN " & FilterVar("D", "''", "S") & "  "
		strSelect = strSelect & " WHEN B.DR_FG = " & FilterVar("N", "''", "S") & "  AND  B.CR_FG = " & FilterVar("Y", "''", "S") & "  THEN " & FilterVar("C", "''", "S") & "  "
		strSelect = strSelect & " END	, " & strItemSeq & ",  "
		strSelect = strSelect & " LTrim(ISNULL(A.MAJOR_CD,'')), CHAR(8) "
  		
		strFrom = " A_CTRL_ITEM	A (NOLOCK), A_ACCT_CTRL_ASSN B (NOLOCK), A_AP_ADJUST_DTL C (NOLOCK), A_AP_ADJUST D (NOLOCK) "
		
		.vspdData.Col = C_REFNo
					
		strWhere =			  " D.ADJUST_NO =  " & FilterVar(UCase(.vspddata.value), "''", "S") & " "		
		strWhere = strWhere & " AND D.ADJUST_NO  =  C.ADJUST_NO  "		
		strWhere = strWhere & "	AND D.ACCT_CD = B.ACCT_CD "
		strWhere = strWhere & " AND C.CTRL_CD = B.CTRL_CD "
		strWhere = strWhere & " AND D.ACCT_CD = B.ACCT_CD "
		strWhere = strWhere & " AND B.CTRL_CD = A.CTRL_CD "
		strWhere = strWhere & " ORDER BY C.DTL_SEQ "		
		
		If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) Then
			arrTempRow =  Split(lgF2By2, Chr(12))
			For Indx1 = 0 To Ubound(arrTempRow) - 1
				arrTempCol = split(arrTempRow(indx1), Chr(11))
				If Trim(arrTempCol(8)) <> "" Then
					strTableid = arrTempCol(8)
					strColid   = arrTempCol(9)
					strColNm   = arrTempCol(10)
					strMajorCd = arrTempCol(15)
					
					strNmwhere = strColid & " =   " & FilterVar(arrTempCol(C_CtrlVal), "''", "S") & "  " 

					If Trim(strMajorCd) <> "" Then
						strNmwhere = strNmwhere & " AND MAJOR_CD =  " & FilterVar(strMajorCd, "''", "S") & "  "
					End If

					If CommonQueryRs( strColNm , strTableid ,  strNmwhere , lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then
						arrVal = Split(lgF0, Chr(11))
						arrTempCol(6) = arrVal(0)
					End If
				End If

				strVal = strVal & Chr(11) & strItemSeq
				strVal = strVal & Chr(11) & arrTempCol(1)
				strVal = strVal & Chr(11) & arrTempCol(2)
				strVal = strVal & Chr(11) & arrTempCol(3)
				strVal = strVal & Chr(11) & arrTempCol(4)
				strVal = strVal & Chr(11) & arrTempCol(5)
				strVal = strVal & Chr(11) & arrTempCol(6)
				strVal = strVal & Chr(11) & arrTempCol(7)
				strVal = strVal & Chr(11) & arrTempCol(8)
				strVal = strVal & Chr(11) & arrTempCol(9)
				strVal = strVal & Chr(11) & arrTempCol(10)
				strVal = strVal & Chr(11) & arrTempCol(11)
				strVal = strVal & Chr(11) & arrTempCol(12)
				strVal = strVal & Chr(11) & arrTempCol(13)
				strVal = strVal & Chr(11) & arrTempCol(15)
				strVal = strVal & Chr(11) & Indx1 + 1
				strVal = strVal & Chr(11) & Chr(12)
			Next
			ggoSpread.Source = .vspdData3
			ggoSpread.SSShowData strVal	
		End If 		

		If Row = CurrRow Then
			Call CopyFromData (strItemSeq)
		End If

		Call LayerShowHide(0)
		Call RestoreToolBar()
	End With

	If Err.number = 0 Then
		fncRestoreDbQuery2 = True
	End If
End Function



'*******************************************************************************************************
'=======================================================================================================
' Area Name   : 4.5.6 Spread OCX Tag Event
' Description : This part declares Spread OCX Tag Event
'=======================================================================================================
'*******************************************************************************************************





'=======================================================================================================
'   Event Name : vspdData_onfocus
'   Event Desc :
'=======================================================================================================
Sub  vspdData_onfocus()

End Sub

'=======================================================================================================
'   Event Name : vspdData2_onfocus
'   Event Desc :
'=======================================================================================================
Sub  vspdData2_onfocus()

End Sub

'======================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'======================================================================================================
Sub  vspdData_Click(ByVal Col, ByVal Row)
	Call SetPopupMenuItemInf("1101111111")
	
    gMouseClickStatus = "SPC"	'Split 상태코드 
    Set gActiveSpdSheet = frm1.vspdData
    
    If frm1.vspdData.MaxRows < 1 Then Exit Sub
    
    ggoSpread.Source = frm1.vspdData
	frm1.vspddata.Row = frm1.vspddata.ActiveRow

	If Row <= 0 Then
        ggoSpread.Source = frm1.vspdData
        If lgSortKey = 1 Then
            ggoSpread.SSSort
            lgSortKey = 2
            frm1.vspddata.Row = 1
            Call DbQuery2(1)
        Else
            ggoSpread.SSSort ,lgSortKey
            lgSortKey = 1
            frm1.vspddata.Row = 1            
            Call DbQuery2(1)            
        End If 
		Exit Sub   
    End If

	If frm1.vspdData.ActiveCol = C_AcctPopUp Then
	    Exit Sub
    End If
 
  	frm1.vspdData.Col = C_AcctCd
	
    If Len(frm1.vspdData.Text) > 0 Then

	Else
		ggoSpread.Source = frm1.vspdData2
		ggoSpread.ClearSpreadData		
	End If	
End Sub

'==========================================================================================
'   Event Desc : Spread Split 상태코드 
'==========================================================================================
Sub vspdData_MouseDown(Button, Shift, X, Y)
	If Button = 2 And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If
End Sub

'=======================================================================================================
'   Event Name : vspdData_LeaveCell
'   Event Desc :
'=======================================================================================================
Sub  vspdData_ScriptLeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
    If Row <> NewRow And NewRow > 0 Then
        With frm1        
            .vspdData.Row = NewRow
            .vspdData.Col = C_ItemSeq
            .hItemSeq.value = .vspdData.Text
			ggoSpread.Source = frm1.vspdData2
			ggoSpread.ClearSpreadData		
        End With
        
        frm1.vspddata.Col = 0
        If frm1.vspddata.Text = ggoSpread.DeleteFlag Then
			Exit Sub       
		End If

        lgCurrRow = NewRow
        
        Call DbQuery2(lgCurrRow)
    End If
End Sub

'======================================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : 버튼 컬럼을 클릭할 경우 발생 
'======================================================================================================
Sub  vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	Dim strTemp
	Dim intPos1

	With frm1.vspdData 
        ggoSpread.Source = frm1.vspdData
       
        If Row > 0 And Col = C_AcctPopUp Then
            .Col = Col - 1
            .Row = Row
            
            Call OpenAcctInfo(.Text)
        End If
        
        If Row > 0 And Col = C_DocCurPopUp Then
            .Col = Col - 1
            .Row = Row
            
            Call OpenCurrency(.Text)
        End If    
    End With
End Sub

'======================================================================================================
'   Event Name :vspddata_EditChange
'   Event Desc :
'======================================================================================================
Sub  vspdData_EditChange(ByVal Col , ByVal Row )
                
End Sub

'======================================================================================================
'   Event Name : vspdData_Change
'   Event Desc :
'======================================================================================================
Sub  vspdData_Change(ByVal Col, ByVal Row )
	Dim AdjustAmt
	
    ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row
    
    frm1.vspdData.Row = Row
    frm1.vspdData.Col = 0
    
	Select Case Col 
		Case C_AcctCd 
			If frm1.vspdData.Text = ggoSpread.InsertFlag Then
				frm1.vspdData.Col   = C_ItemSeq
				frm1.hItemSeq.value = frm1.vspdData.Text  
				frm1.vspdData.Col   = C_AcctCd        
				
				If Len(frm1.vspdData.Text) > 0 Then
					frm1.vspdData.Row = Row
					frm1.vspdData.Col = C_ItemSeq	
					DeleteHsheet frm1.vspdData.Text
					
					Call DbQuery3(Row)
					Call InputCtrlVal(Row)	
					Call SetSpread2ColorAP()		 
				End If 
			End If
		Case C_AdjustAmt
				frm1.vspdData.Col = C_AdjustAmt
				AdjustAmt = frm1.vspdData.Text
				
				If (UNICDbl(frm1.txtApAmt.Text) > 0 And parent.UNICDbl(AdjustAmt) < 0) Or (UNICDbl(frm1.txtApAmt.Text) < 0 And parent.UNICDbl(AdjustAmt) > 0) then
					frm1.vspdData.Text = UNIConvNumPCToCompanyByCurrency(UNICDbl(frm1.vspdData.Text) * (-1),frm1.txtDocCur.value,parent.ggAmtOfMoneyNo, "X", "X")
				End If
				frm1.vspdData.Col  = C_AdjustLocAmt        
				frm1.vspdData.text = "" 					
				Call DoSum()
		Case C_AdjustLocAmt
			Call DoSum()
		Case C_AdjustDt
			With frm1
				.vspdData.Col  = C_AdjustLocAmt	
				.vspdData.Text = ""    		
			End With	
			Call DoSum()					
	End Select
End Sub

'======================================================================================================
'   Event Name :vspddata_DblClick
'   Event Desc :
'======================================================================================================
Sub  vspddata_DblClick( ByVal Col , ByVal Row )
    Dim iColumnName
    
    If Row <= 0 Then
       Exit Sub
    End If
    
    If frm1.vspdData.MaxRows = 0 Then
       Exit Sub
    End If       
End Sub

'======================================================================================================
'   Event Name :vspddata_KeyPress
'   Event Desc :
'======================================================================================================
Sub  vspddata_KeyPress(KeyAscii )
     
End Sub

'======================================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'======================================================================================================
Sub  vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    Dim intListGrvCnt 
    Dim LngLastRow    
    Dim LngMaxRow     
    
    If OldLeft <> NewLeft Then
        Exit Sub
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

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  --> 
</HEAD>

<!--
'======================================================================================================
'       					6. Tag부 
'	기능: Tag부분 설정 
'====================================================================================================== -->
<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="post">
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
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 >
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><IMG height=23 src="../../../CShared/image/table/seltab_up_left.gif" width=9></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white><%=Request("strASPMnuMnuNm")%></font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><IMG height=23 src="../../../CShared/image/table/seltab_up_right.gif" width=10></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* ALIGN=RIGHT><A HREF="VBSCRIPT:OpenPopupTempGL()">결의전표</A>&nbsp;|&nbsp;<A HREF="VBSCRIPT:OpenPopupGL()">회계전표</A></TD>
					<TD WIDTH=10>&nbsp;</TD>
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
									<TD CLASS="TD5" NOWRAP>채무번호</TD>
									<TD CLASS="TD656" NOWRAP><INPUT NAME="txtApNo" MAXLENGTH=18 STYLE="TEXT-ALIGN: left" tag ="12XXXU" ALT="채무번호"><IMG align=top name=btnPrpaymNo src="../../../CShared/image/btnPopup.gif"  TYPE="BUTTON" onclick="vbscript:OpenApNo"></TD>								
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
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD CLASS="TD5" NOWRAP>부서</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtDeptCd" SIZE=10 MAXLENGTH=10 STYLE="TEXT-ALIGN: left" tag="24" ALT="부서">&nbsp;<INPUT TYPE=TEXT NAME="txtDeptNm" SIZE=25 tag="24" ALT="부서명"></TD>
								<TD CLASS="TD5" NOWRAP>채무일자</TD>
								<TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name=txtApDt CLASS=FPDTYYYYMMDD title="FPDATETIME" ALT="채무일자" tag="24" ></OBJECT>');</SCRIPT></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>지급처</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtBpCd" SIZE=10 MAXLENGTH=10 STYLE="TEXT-ALIGN: left" tag="24" ALT="지급처">&nbsp;<INPUT TYPE=TEXT NAME="txtBpNm" SIZE=25 tag="24" ALT="지급처명"></TD>
								<TD CLASS="TD5" NOWRAP>참조번호</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtRefNo" SIZE=30 MAXLENGTH=30 STYLE="TEXT-ALIGN: left" tag="24" ALT="참조번호"></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>거래통화</TD>
								<TD CLASS="TD6" NOWRAP><INPUT NAME="txtDocCur" ALT="거래통화" TYPE="Text" SIZE=10 STYLE="TEXT-ALIGN: left" tag="24" ></TD>
								<TD CLASS="TD5" NOWRAP>전표번호</TD>
						        <TD CLASS="TD6" NOWRAP><INPUT NAME="txtGlNo" ALT="전표번호" TYPE="Text" SIZE=19 STYLE="TEXT-ALIGN: left" tag="24" ></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>채무액</TD>
						        <TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name=txtApAmt CLASS=FPDS140 title=FPDOUBLESINGLE ALT="채무액" tag="24X2" ></OBJECT>');</SCRIPT></TD>
								<TD CLASS="TD5" NOWRAP>채무액(자국)</TD>
						        <TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name=txtApLocAmt CLASS=FPDS140 title=FPDOUBLESINGLE ALT="채무액(자국)" tag="24X2" ></OBJECT>');</SCRIPT></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>잔액</TD>
						        <TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name=txtBalAmt CLASS=FPDS140 title=FPDOUBLESINGLE ALT="잔액" tag="24X2" ></OBJECT>');</SCRIPT></TD>
								<TD CLASS="TD5" NOWRAP>잔액(자국)</TD>
						        <TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name=txtBalLocAmt CLASS=FPDS140 title=FPDOUBLESINGLE ALT="잔액(자국)" tag="24X2" ></OBJECT>');</SCRIPT></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>적요</TD>
								<TD CLASS="TD6" NOWRAP COLSPAN=3><INPUT TYPE=TEXT NAME="txtApDesc" SIZE=90 MAXLENGTH=128 tag="24" ALT="적요"></TD>
							</TR>						
							<TR HEIGHT="50%">
								<TD WIDTH="100%" COLSPAN=4>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> HEIGHT="100%" tag="2" TITLE="SPREAD" name=vspdData width="100%" id=OBJECT1> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"></OBJECT>');</SCRIPT>
								</TD>
							</TR>
							<TR>
								<TD  COLSPAN="4">
									<TABLE <%=LR_SPACE_TYPE_60%>>
										<TR>														
								<TD CLASS="TD5" NOWRAP>청산금액</TD>
						        <TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name=txtTotAdjustAmt CLASS=FPDS140 title=FPDOUBLESINGLE ALT="청산금액" tag="24X2" ></OBJECT>');</SCRIPT></TD>
						        <TD class=TD5 STYLE="WIDTH : 0px;"></TD>
								<TD CLASS="TD5" NOWRAP>청산금액(자국)</TD>
						        <TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name=txtTotAdjustLocAmt CLASS=FPDS140 title=FPDOUBLESINGLE ALT="청산금액(자국)" tag="24X2" ></OBJECT>');</SCRIPT></TD>
										</TR>
									</TABLE>
								</TD>														        
							</TR>
							<TR HEIGHT="50%">
								<TD WIDTH="100%" COLSPAN=4>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> HEIGHT="100%" tag="2" TITLE="SPREAD" name=vspdData2 width="100%" id=OBJECT2> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"></OBJECT>');</SCRIPT>
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
		<TD WIDTH="100%" HEIGHT=<%=BizSize%>>
			<IFRAME NAME="MyBizASP" WIDTH="100%" HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA TYPE=hidden Class=hidden name=txtSpread  tag="24" TABINDEX="-1"></TEXTAREA>
<TEXTAREA TYPE=hidden Class=hidden name=txtSpread3 tag="24" TABINDEX="-1"></TEXTAREA>
<INPUT TYPE=hidden NAME="txtMode"      tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="hUpdtUserId"  tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="hInsrtUserId" tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="txtMaxRows"   tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="txtFlgMode"   tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="hRcptNo"    tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="hOrgChangeId" tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="hItemSeq"     tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="hAcctCd"      tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="txtMaxRows3"  tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="hbankcd"      tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="hbanknm"      tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="hbankacct"    tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="hClsAmt"      tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="hClsLocAmt"   tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="hConfFg"      tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="hGlNo"        tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="hNoteNo"      tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="hAcctCd"      tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="hAcctNm"      tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="hSttlAmt"     tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="hSttlLocAmt"  tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="txthAuthBizAreaCd"	tag="24" Tabindex="-1">
<INPUT TYPE=hidden NAME="txthInternalCd"	tag="24" Tabindex="-1">
<INPUT TYPE=hidden NAME="txthSubInternalCd"	tag="24" Tabindex="-1">
<INPUT TYPE=hidden NAME="txthAuthUsrID"		tag="24" Tabindex="-1">
<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT TYPE=hidden CLASSID=<%=gCLSIDFPSPD%> HEIGHT=0 name=vspdData3 tag="2" width="100%" TABINDEX="-1"> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"></OBJECT>');</SCRIPT>
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>            

