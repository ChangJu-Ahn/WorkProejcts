<%@ LANGUAGE="VBSCRIPT" %>

<!--
'=======================================================================================================
'*  1. Module Name          : Accounting
'*  2. Function Name        : A_RECEIPT
'*  3. Program ID           : A3119ma1
'*  4. Program Name         : 가수금잔액정리 
'*  5. Program Desc         : 입금청산 
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
<!-- #Include file="../../inc/incSvrHTML.inc"  -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAMain.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAEvent.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliVariables.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAOperation.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliRdsQuery.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"     SRC="../ar/AcctCtrl3.vbs">				</SCRIPT>
<SCRIPT LANGUAGE=vbscript>

Option Explicit																	'☜: indicates that All variables must be declared in advance

'=======================================================================================================
'                                               1.2 Global 변수/상수 선언  
'	.Constant는 반드시 대문자 표기.
'	.변수 표준에 따름. prefix로 g를 사용함.
'	.Array인 경우는 ()를 반드시 사용하여 일반 변수와 구별해 됨 
'=======================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'@PGM_ID
Const BIZ_PGM_ID         = "a3119mb1.asp"									' F_PrPaym_Sttl 의 CRUD

Const COM_BIZ_EXCHRATE_ID = "../../inc/GetExchRate.asp"								'☆: 환율정보 비지니스 로직 ASP명 

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


Dim  lgStrPrevKeyDtl
Dim  lgStrPrevKey2
Dim  lgStrPrevKey3
Dim  lgCurrRow
Dim  lgPrevNo
Dim  lgNextNo

Dim  IsOpenPop	                'Popup
Dim  gSelframeFlg
Dim  dtToday
dtToday = "<%=GetSvrDate%>"

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
   
End Sub

'=======================================================================================================
'	Name : InitVariables()
'	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
'=======================================================================================================
Sub  InitVariables()
    lgIntFlgMode = parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
        
    lgStrPrevKey = 0                            'initializes Previous Key
    lgStrPrevKeyDtl = 0                         'initializes Previous Key
    lgLngCurRows = 0                            'initializes Deleted Rows Count
    frm1.hOrgChangeId.value = parent.gChangeOrgId
End Sub

'=======================================================================================================
'	Name : SetDefaultVal()
'	Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
'=======================================================================================================
Sub  SetDefaultVal()
	frm1.txtadjustDt.text = UniConvDateAToB(dtToday, parent.gServerDateFormat,gDateFormat)
	frm1.txtDocCur.value = parent.gcurrency	
	 Call txtDocCur_OnChange()   
End Sub

'======================================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'=======================================================================================================
Sub  LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% Call LoadInfTB19029A("I","*","NOCOOKIE","MA") %>
<% Call LoadBNumericFormatA("I", "*", "NOCOOKIE" , "MA") %>
End Sub

'======================================================================================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'=======================================================================================================
Sub  InitSpreadSheet()
    frm1.txtadjustDt.text = UniConvDateAToB(dtToday, parent.gServerDateFormat,gDateFormat)
End Sub

'=======================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'=======================================================================================================
Sub  SetSpreadLock()
    With frm1
	
	End With
End Sub

'======================================================================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'=======================================================================================================
Sub SetSpreadColor(ByVal pvStartRow,ByVal pvEndRow)
    With frm1
	
    End With
End Sub

'======================================================================================================
' Function Name : SetSpread2ColorAR
' Function Desc : This method set color and protect in spread sheet celles
'=======================================================================================================
Sub  SetSpread2ColorAR()
	Dim i

    With frm1
		ggoSpread.Source = .vspdData2

		.vspdData2.ReDraw = False
		                
		for i = 1 to .vspddata2.maxrows
			ggoSpread.SSSetProtected C_DtlSeq   , i, i
			ggoSpread.SSSetProtected C_CtrlCd   , i, i
			ggoSpread.SSSetProtected C_CtrlNm   , i, i
			ggoSpread.SSSetProtected C_CtrlValNm, i, i

			.vspddata2.Col = C_DrFg		
		
			If (.vspddata2.text = "C" And .vspddata2.text <> "") _
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
   
End Sub
'=========================================================================================================
'	Name : OpenAdjustNo()
'	Description : Ref 화면을 call한다. : 채권발생정보 
'========================================================================================================= 
Function OpenAdjustNo()
	
	Dim arrRet
	Dim arrParam(8)	
	Dim iCalledAspName
	Dim IntRetCD

	If IsOpenPop = True Then Exit Function
	
	iCalledAspName = AskPRAspName("a3506ra2")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a3506ra2", "X")
		IsOpenPop = False
		Exit Function
	End If
   
	IsOpenPop = True

	arrParam(0) = ""				' 검색조건이 있을경우 파라미터 
	arrParam(1) = ""				
	arrParam(2) = ""			
	arrParam(3) = "M"


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
		Call SetAdjustNo(arrRet)
	End If
End Function
'======================================================================================================
'   Function Name : SetAdjustNo(Byval arrRet)
'   Function Desc : 
'=======================================================================================================
Function SetAdjustNo(Byval arrRet)
	With frm1
		frm1.txtAdjustNo.value	= arrRet(0)
		frm1.txtAdJustNo.focus
	End With
End Function

'=======================================================================================================
'	Name : Openpopupgl()
'	Description : 
'=======================================================================================================
Function OpenPopupGL()
	Dim arrRet
	Dim arrParam(8)	
	Dim RetFlag
	Dim iCalledAspName
	Dim IntRetCD
	
	If IsOpenPop = True Then Exit Function
	
	iCalledAspName = AskPRAspName("a5120ra1")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a5120ra1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	If Trim(frm1.txtGlNo.value) = "" Then
		RetFlag = DisplayMsgBox("970000","X" , frm1.txtGlNo.Alt, "X") 	
		IsOpenPop = False
		Exit Function
	End If
	arrParam(0) = Trim(frm1.txtGlNo.value)
	arrParam(1) = ""			'Reference번호 


	' 권한관리 추가 
	arrParam(5) = lgAuthBizAreaCd
	arrParam(6) = lgInternalCd
	arrParam(7) = lgSubInternalCd
	arrParam(8) = lgAuthUsrID
	
	IsOpenPop = True
   
	arrRet = window.showModalDialog(iCalledAspName,Array(window.parent, arrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
End Function

'=======================================================================================================
'	Name : OpenPopuptempGL()
'	Description : 
'=======================================================================================================
Function OpenPopuptempGL()
	Dim arrRet
	Dim arrParam(8)	
	Dim RetFlag
	Dim iCalledAspName
	Dim IntRetCD
	
	If IsOpenPop = True Then Exit Function
	
	iCalledAspName = AskPRAspName("a5130ra1")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a5130ra1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	If Trim(frm1.txtTempGlNo.value) = "" Then
		RetFlag = DisplayMsgBox("970000","X" , frm1.txtTempGlNo.Alt, "X") 	
		IsOpenPop = False
		Exit Function
	End If
	
	arrParam(0) = Trim(frm1.txtTempGlNo.value)	'결의전표번호 
	arrParam(1) = ""			'Reference번호 

	' 권한관리 추가 
	arrParam(5) = lgAuthBizAreaCd
	arrParam(6) = lgInternalCd
	arrParam(7) = lgSubInternalCd
	arrParam(8) = lgAuthUsrID
		
	IsOpenPop = True
   
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
End Function

'=======================================================================================================
'	Name : OpenRcptNo()
'	Description : Prepayment No PopUp
'=======================================================================================================
Function OpenRcptNo()
	Dim arrRet
	Dim arrParam(8)
	Dim iCalledAspName
	Dim IntRetCD
	
	If IsOpenPop = True  Then Exit Function
	
	iCalledAspName = AskPRAspName("a3119ra1")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a3119ra1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	If Trim(frm1.txtAdJustNo.value) <> "" And lgIntFlgMode = parent.OPMD_UMODE Then Exit Function

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
		Call SetRcptNo(arrRet)
	End If	
End Function

'=======================================================================================================
'	Name : SetRcptNo()
'	Description : PrpaymNo Popup에서 Return되는 값 setting
'***	아래의 sql 을 해당 db 에서 실행해보아 첨자 그대로 사용한다.	이 설명은 개발이후에 지운다.
'***	select field_nm + '---'  + field_cd+ '---arrRet('  +  rtrim(convert(char(3), key_tag-1 )) + ')'  
'***	from z_ado_field_inf
'***	where pgm_id = 'a3119ra1'
'***	and lang_Cd = 'ko'
'***	order by seq_no

'=======================================================================================================
Function SetRcptNo(byval arrRet)
	With frm1
		.txtRcptDt.text  = arrRet(0)
		.txtBpCd.value = arrRet(1)
		.txtBpNM.value  = arrRet(2)
		.txtDeptCd.value  = arrRet(3)
		.txtDeptNm.value = arrRet(4)
		.txtRefNo.value = arrRet(5)
		.txtRcptNo.value = arrRet(6)
		.txtDocCur.value  = arrRet(7)
		.txtXchRate.text = arrRet(8)
		.txtRcptAmt.text  = arrRet(9)
		.txtRcptLocAmt.text = arrRet(10)
		.txtBalAmt.text  = arrRet(11)
		.txtBalLocAmt.text = arrRet(12)
		.txtRcptDesc.value = arrRet(14)
		

		Call txtDocCur_OnChange()
		If Trim(.txtDocCur.value) <> "" Then
			Call ggoOper.SetReqAttr(frm1.txtDocCur,   "Q")
		End If
		
	End With		
    lgBlnFlgChgValue = True
End Function


'======================================================================================================
'   Function Name : OpenPopup(Byval strCode, Byval iWhere)
'   Function Desc : 
'======================================================================================================
Function  OpenPopup(Byval strCode, iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function
	
	
	Select Case iWhere
		Case 1
			
			If frm1.txtAcctCd.className = "protected" Then Exit Function
			
			arrParam(0) = "계정코드팝업"								' 팝업 명칭 
			arrParam(1) = "A_Acct, A_ACCT_GP" 								' TABLE 명칭 
			arrParam(2) = Trim(strCode)										' Code Condition
			arrParam(3) = ""												' Name Condition
			arrParam(4) = "A_ACCT.GP_CD=A_ACCT_GP.GP_CD AND A_ACCT.DEL_FG <> " & FilterVar("Y", "''", "S") & " "												' Where Condition
			arrParam(5) = "계정코드"									' 조건필드의 라벨 명칭 

			arrField(0) = "A_ACCT.Acct_CD"									' Field명(0)
			arrField(1) = "A_ACCT.Acct_NM"									' Field명(1)
			arrField(2) = "A_ACCT_GP.GP_CD"									' Field명(2)
			arrField(3) = "A_ACCT_GP.GP_NM"									' Field명(3)
					
			arrHeader(0) = "계정코드"									' Header명(0)
			arrHeader(1) = "계정코드명"									' Header명(1)
			arrHeader(2) = "그룹코드"									' Header명(2)
			arrHeader(3) = "그룹명"										' Header명(3)
    
		Case 2
			
		
		End Select
	IsOpenPop = True	
		arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtAcctCd.focus
		Exit Function
	Else
		Call SetPopUp(arrRet, iWhere)
	End If	
End Function

'======================================================================================================
'   Function Name : SetPopUp(Byval arrRet,byval iWhere)
'   Function Desc : 
'======================================================================================================
Function SetPopUp(Byval arrRet, Byval iWhere)
	With frm1
		Select Case iWhere
			Case 1
					.txtAcctCd.value = arrRet(0)
					.txtAcctNm.value  = arrRet(1)
					Call txtAcctCd_OnChange()
					.txtAcctCd.focus
			Case 2
		End Select
		
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
    Call ggoOper.ClearField(Document, "1")										'⊙: Condition field clear
    Call ggoOper.FormatField(Document, "2", ggStrIntegeralPart, ggStrDeciPointPart, _
							parent.gDateFormat, parent.gComNum1000, parent.gComNumDec)
	Call ggoOper.LockField(Document, "N")										'Lock  Suitable  Field
	Call InitCtrlSpread()
	Call InitVariables()																'Initializes local global variables
    
    Call SetDefaultVal()
    Call SetToolbar("1110100000001111")
	frm1.txtAdJustNo.focus
   

    lgBlnFlgChgValue = False            

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
    Dim  var2
    
    FncQuery = False                                                        
    
    Err.Clear                                                               
	'-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then													'This function check indispensable field
		Exit Function
    End If
    '-----------------------
    'Check previous data area
    '-----------------------
   
    ggoSpread.Source = frm1.vspdData2
    var2 = ggoSpread.SSCheckChange
    
    If lgBlnFlgChgValue = True  Or var2 = True Then		
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO,"X","X")	    
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")										'Clear Contents  Field
    Call InitVariables()																'Initializes local global variables
	
	ggoSpread.Source = frm1.vspdData2
	ggoSpread.ClearSpreadData	
    '-----------------------
    'Query function call area
    '-----------------------
    Call DbQuery()																		'☜: Query db data
    
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
    
    ggoSpread.Source = frm1.vspdData2
    var2 = ggoSpread.SSCheckChange
	'-----------------------
    'Check previous data area
    '----------------------- 
    If lgBlnFlgChgValue = True  Or var2 = True Then
        IntRetCD = DisplayMsgBox("900015", parent.VB_YES_NO,"X","X")               
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
	'-----------------------
    'Erase condition area
    'Erase contents area
    '----------------------- 
    Call ggoOper.ClearField(Document, "1")                                  'Clear Condition Field
    Call ggoOper.ClearField(Document, "2")										'Clear Condition Field
    Call ggoOper.LockField(Document, "N")										'Lock  Suitable  Field
    Call InitVariables()																'Initializes local global variables
    
    Call SetDefaultVal()
    Call txtDocCur_OnChange()
	Call DisableRefPop()    

	ggoSpread.Source = frm1.vspdData2
	ggoSpread.ClearSpreadData
    
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
    
    IntRetCD = DisplayMsgBox("900003", parent.VB_YES_NO,"X","X")					'Will you destory previous data"
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
        
 

    ggoSpread.Source = frm1.vspdData2
    var2 = ggoSpread.SSCheckChange

    If lgBlnFlgChgValue = False  And var2 = False  Then				'⊙: Check If data is chaged
        IntRetCD = DisplayMsgBox("900001","X","X","X")								'⊙: Display Message(There is no changed data.)
		Exit Function
    End If

	If Not chkField(Document, "2") Then												'⊙: Check required field(Single area)
		Exit Function
    End If
   

	If Trim(frm1.txtRcptNo.value)  = "" Then
		IntRetCD = DisplayMsgBox("112700","X","X","X")									'입금정보check
        Exit Function
    End If
    
	
	
    If CheckSpread4 = False Then
	IntRetCD = DisplayMsgBox("110420","X","X","X")									'필수입력 check!!
        Exit Function
    End If
    '-----------------------
    'Save function call area
    '-----------------------
    Call DbSave()  																	'☜: Save db data
    
    FncSave = True      
    	
	Set gActiveElement = document.activeElement    
	                                                 
End Function

'=======================================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================================
Function  FncCopy() 
	Dim  IntRetCD
	
	
End Function

'=======================================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================================
Function  FncCancel() 
	ggoSpread.Source = frm1.vspdData2
	ggoSpread.ClearSpreadData		
     	
	Set gActiveElement = document.activeElement    
	
End Function

'=======================================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================================
Function  FncInsertRow(ByVal pvRowcnt) 
	Dim iCurRowPos
	Dim imRow
    Dim ii
    
	On Error Resume Next															'☜: If process fails
    Err.Clear   
	
    FncInsertRow = False															'☜: Processing is NG

    If IsNumeric(Trim(pvRowcnt)) Then 
		imRow  = Cint(pvRowcnt)
    Else
		imRow = AskSpdSheetAddRowCount()
		If imRow = "" Then
		    Exit Function
		End If
    End If
        
    Call ggoOper.LockField(Document, "I")									'This function lock the suitable field
    
    If Err.number = 0 Then
       FncInsertRow = True                                                          '☜: Processing is OK
    End If   
	
    Set gActiveElement = document.ActiveElement   
End Function

'=======================================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================================
Function  FncDeleteRow() 
   	Dim lDelRows
    Dim DelItemSeq

End Function

'=======================================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================================
Function  FncPrint() 
    On Error Resume Next                                               
    FncPrint()
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
    Call FncFind(parent.C_SINGLEMULTI , True) 
    	
	Set gActiveElement = document.activeElement    
	                         
End Function

'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================
Function  FncExcel() 
	Call FncExport(parent.C_SINGLEMULTI)
	
	Set gActiveElement = document.activeElement    
	
End Function

'========================================================================================
' Function Name : FncSplitColumn
' Function Desc : 
'========================================================================================
Function FncSplitColumn()
    Dim ACol
    Dim ARow
    Dim iRet
    Dim iColumnLimit
    
    iColumnLimit  = 5
    
   
    If gMouseClickStatus = "SP2CRP" Then
		ACol = Frm1.vspdData2.ActiveCol
		ARow = Frm1.vspdData2.ActiveRow

		If ACol > iColumnLimit Then
				Frm1.vspdData2.Col = iColumnLimit : frm1.vspdData2.Row = 0  	 	 	 	 		
				iRet = DisplayMsgBox("900030", "X", Trim(frm1.Vspddata2.text), "X")
				Exit Function  
		End If   
    
		Frm1.vspdData2.ScrollBars = parent.SS_SCROLLBAR_NONE
    
		ggoSpread.Source = Frm1.vspdData2
    
		ggoSpread.SSSetSplit(ACol)    
    
		Frm1.vspdData2.Col = ACol
		Frm1.vspdData2.Row = ARow
    
		Frm1.vspdData2.Action = Parent.SS_ACTION_ACTIVE_CELL     
    
		Frm1.vspdData2.ScrollBars = parent.SS_SCROLLBAR_BOTH
    End If   
	
	Set gActiveElement = document.activeElement    
		
End Function

'=======================================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================================
Function  FncExit()
	Dim IntRetCD
	Dim var1,var2
	
	FncExit = False

    ggoSpread.Source = frm1.vspdData2
    var2 = ggoSpread.SSCheckChange
    
    If lgBlnFlgChgValue = True or var2 = True Then					'⊙: Check If data is chaged
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
    Dim strVal

    DbDelete = False														
    
    strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0003
    strVal = strVal & "&txtRcptNo=" & Trim(frm1.txtRcptNo.value)				'☜: 삭제 조건 데이타    
    strVal = strVal & "&txtAdjustNo=" & Trim(frm1.txtAdjustNo.value)				'☜: 삭제 조건 데이타    

	' 권한관리 추가 
	strVal = strVal & "&lgAuthBizAreaCd="	& lgAuthBizAreaCd			' 사업장 
	strVal = strVal & "&lgInternalCd="		& lgInternalCd				' 내부부서 
	strVal = strVal & "&lgSubInternalCd="	& lgSubInternalCd			' 내부부서(하위포함)
	strVal = strVal & "&lgAuthUsrID="		& lgAuthUsrID				' 개인 
   
    Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 
    
    DbDelete = True                                                         
End Function

'=======================================================================================================
' Function Name : DbDeleteOk
' Function Desc : DbDelete가 성공적일때 수행 
'========================================================================================================
Function DbDeleteOk()												        '삭제 성공후 실행 로직 
	Call ggoOper.ClearField(Document, "1")                                  'Clear Condition Field
	Call ggoOper.ClearField(Document, "2")                                  'Clear Condition Field
    Call ggoOper.LockField(Document, "N")                                   'Lock  Suitable  Field
    Call InitVariables()                                                      'Initializes local global variables
    Call SetDefaultVal()
    Call DisableRefPop()
	ggoSpread.Source = frm1.vspdData2
	ggoSpread.ClearSpreadData	
End Function

'=======================================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'=======================================================================================================
Function DbQueryOk()													'☆: 조회 성공후 실행로직	
	With frm1
        '-----------------------
        'Reset variables area
        '-----------------------
        lgIntFlgMode = parent.OPMD_UMODE												'Indicates that current mode is Update mode
        
        Call ggoOper.LockField(Document, "Q")									'This function lock the suitable field        
        Call SetToolbar("1111100000001111")                                     '버튼 툴바 제어 
         Call DbQuery2()
        
    End With
    
 
	Call txtDocCur_OnChange()
	Call DisableRefPop()
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
			strVal = strVal & "&txtAdjustNo=" & Trim(.txtAdjustNo.value)				'조회 조건 데이타 
			
			
			
		Else
			strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001					'☜: 
			strVal = strVal & "&txtAdjustNo=" & Trim(.txtAdjustNo.value)				'조회 조건 데이타 
		End If
    End With

	' 권한관리 추가 
	strVal = strVal & "&lgAuthBizAreaCd="	& lgAuthBizAreaCd			' 사업장 
	strVal = strVal & "&lgInternalCd="		& lgInternalCd				' 내부부서 
	strVal = strVal & "&lgSubInternalCd="	& lgSubInternalCd			' 내부부서(하위포함)
	strVal = strVal & "&lgAuthUsrID="		& lgAuthUsrID				' 개인 

	Call RunMyBizASP(MyBizASP, strVal)												'☜: 비지니스 ASP 를 가동 

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
        Call SetToolbar("1110100000001111")                                     '버튼 툴바 제어 
        Call DbQuery2()
        
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
    
    GrpCntD = 1: strValD = ""	'관리항목 파라미터 
    
	'=======================================================================
	'2001.06.18 Song,MunGil 관리항목도 입력/수정된 걸로 간주하고 생성함.
	'=======================================================================
	With frm1.vspdData2
	For RowD = 1 To .MaxRows
		.Row = RowD
		.Col = 0
'		If (.Text = ggoSpread.InsertFlag or .Text = ggoSpread.UpdateFlag) then
		If Trim(.Text) <> ggoSpread.DeleteFlag then
			strValD = strValD & "C" & parent.gColSep & RowD & parent.gColSep
			strValD = strValD & "1" & parent.gColSep
			.Col = C_DtlSeq 
			strValD = strValD & Trim(.Text) & parent.gColSep
			.Col = C_CtrlCd
			strValD = strValD & Trim(.Text) & parent.gColSep
			.Col = C_CtrlVal
			strValD = strValD & Trim(.Text) & parent.gRowSep
										
			GrpCntD = GrpCntD + 1
		End If
	Next
	End With				

	
	frm1.txtMaxRows2.value = GrpCntD - 1									'Spread Sheet의 변경된 최대갯수 
	frm1.txtSpread2.value  = strValD				

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
	ggoSpread.Source = frm1.vspdData2
	ggoSpread.ClearSpreadData	

	Call DbQuery()		
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
Function DbQuery2()
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
	  
	   
    	
		Call LayerShowHide(1)
	
		DbQuery2 = False
		strSelect =				" C.DTL_SEQ,  A.CTRL_CD, A.CTRL_NM , C.CTRL_VAL, '',"
		strSelect = strSelect & " CASE  WHEN A.ColM_DATA_TYPE = " & FilterVar("D", "''", "S") & "  THEN " & FilterVar("(Format : YYYY-MM-DD)", "''", "S") & "  END , "
		strSelect = strSelect & " 1 , LTrim(ISNULL(A.TBL_ID,'')), LTrim(ISNULL(A.DATA_COLM_ID,'')), "
		strSelect = strSelect & " LTrim(ISNULL(A.DATA_COLM_NM,'')),  LTrim(ISNULL(A.ColM_DATA_TYPE,'')), LTrim(ISNULL(A.DATA_LEN,'')), "
		strSelect = strSelect & " CASE WHEN B.DR_FG = " & FilterVar("Y", "''", "S") & "  AND  B.CR_FG = " & FilterVar("Y", "''", "S") & "  THEN " & FilterVar("DC", "''", "S") & "  "
		strSelect = strSelect & " WHEN B.DR_FG = " & FilterVar("Y", "''", "S") & "  AND  B.CR_FG = " & FilterVar("N", "''", "S") & "  THEN " & FilterVar("D", "''", "S") & "  "
		strSelect = strSelect & " WHEN B.DR_FG = " & FilterVar("N", "''", "S") & "  AND  B.CR_FG = " & FilterVar("Y", "''", "S") & "  THEN " & FilterVar("C", "''", "S") & "  "
		strSelect = strSelect & " END	, "
		strSelect = strSelect & " LTrim(ISNULL(A.MAJOR_CD,'')),CHAR(8) "
  		
		strFrom = " A_CTRL_ITEM	A (NOLOCK), A_ACCT_CTRL_ASSN B (NOLOCK), A_RCPT_ADJUST_DTL C (NOLOCK), A_RCPT_ADJUST D (NOLOCK) "
		
					
		strWhere =			  " D.ADJUST_NO =  " & FilterVar(UCase(.txtAdjustNO.value), "''", "S") & " "		
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
			Next					
			
		End If 		
	
		Call SetSpread2ColorAR()
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
	Call SetSpread2ColorAR()
    Call txtDocCur_OnChange()
End Function

'==========================================================================================
'   Event Name : txtDocCur_OnChange
'   Event Desc : 
'==========================================================================================
Sub txtDocCur_OnChange()
    Dim arrVal
    lgBlnFlgChgValue = True
    If CommonQueryRs( "CURRENCY_DESC" , "B_CURRENCY" ,  " CURRENCY = " & FilterVar(frm1.txtDocCur.value, "''", "S"), lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then   			        							
		arrVal = Split(lgF0, Chr(11))  
		'frm1.txtDocCurNm.value = arrVal(0)
		Call CurFormatNumericOCX()
		Call CurFormatNumSprSheet()
	End If	  
End Sub

'===================================== DisableRefPop()  =======================================
'	Name : DisableRefPop()
'	Description :
'====================================================================================================
Sub DisableRefPop()
	IF lgIntFlgMode = parent.OPMD_UMODE Then
		RefPop.innerHTML="<font color=""#777777"">가수금정보</font>"
	ELse 
		RefPop.innerHTML="<A href=""vbscript:OpenRcptNo()"">가수금정보</A>"
	End if

END sub
'===================================== CurFormatNumericOCX()  =======================================
'	Name : CurFormatNumericOCX()
'	Description : 화면에서 일괄적으로 Rounding되는 Numeric OCX
'====================================================================================================
Sub CurFormatNumericOCX()
	With frm1
		' 입금액 
		ggoOper.FormatFieldByObjectOfCur .txtRcptAmt, .txtDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec
		' 잔액 
		ggoOper.FormatFieldByObjectOfCur .txtBalAmt, .txtDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec
		' 청산금액 
		ggoOper.FormatFieldByObjectOfCur .txtAdjustAmt, .txtDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec
	End With
End Sub

'===================================== CurFormatNumSprSheet()  ======================================
'	Name : CurFormatNumSprSheet()
'	Description : 화면에서 일괄적으로 Rounding되는 Numeric Spread Sheet
'====================================================================================================
Sub CurFormatNumSprSheet()
	
End Sub

'=======================================================================================================
'   Event Name : InputCtrlVal
'   Event Desc :
'=======================================================================================================  
Sub InputCtrlVal(ByVal Row)
	Dim strAcctCd		
	Dim ii
			
	lgBlnFlgChgValue = True
		
	strAcctCd	= Trim(frm1.txtAcctCd.value)		
		
	Call AutoInputDetail(strAcctCd,Trim(frm1.txtDeptCd.value), frm1.txtAdjustDt.text, Row)

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
Sub PopRestoreSpreadColumnInf()
	Dim indx

	Select Case Trim(UCase(gActiveSpdSheet.Name))
	
		Case "VSPDDATA2"
			Call DeleteHSheet(frm1.hItemSeq.value)
		

			ggoSpread.Source = frm1.vspdData2
			Call ggoSpread.RestoreSpreadInf()
			Call InitCtrlSpread()			'관리항목 그리드 초기화 
			Call ggoSpread.ReOrderingSpreadData()
			Call SetSpread2Color()  
	End Select
End Sub


'=======================================================================================================
'   Event Name : txtAdjustDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub  txtAdjustDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtAdjustDt.Action = 7                
        Call SetFocusToDocument("M")
		Frm1.txtAdjustDt.Focus 
    End If
End Sub
'=======================================================================================================
'   Event Name : txtAdjustDt_Change()
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub  txtAdjustDt_Change() 
    lgBlnFlgChgValue = True
End Sub
'==========================================================================================
'   Event Name : txtAcctCd_OnChange
'   Event Desc : 
'==========================================================================================
Sub txtAcctCd_OnChange
 lgBlnFlgChgValue = True
 If Trim(frm1.txtAcctCd.value) <> "" Then
	Call DbQuery4()
 End If
End Sub
'==========================================================================================
'   Event Name : txtAdjustDt_OnChange
'   Event Desc : 
'==========================================================================================
Sub txtAdjustDt_OnChange
 lgBlnFlgChgValue = True
End Sub

'==========================================================================================
'   Event Name : txtAdjustAmt_OnChange
'   Event Desc : 
'==========================================================================================
Sub txtAdjustAmt_Change
 lgBlnFlgChgValue = True
 frm1.txtAdjustLocAmt.Text = 0
End Sub

'==========================================================================================
'   Event Name : txtAdjustLocAmt_OnChange
'   Event Desc : 
'==========================================================================================
Sub txtAdjustLocAmt_Change
 lgBlnFlgChgValue = True
End Sub

'==========================================================================================
'   Event Name : txtAddesc_OnChange
'   Event Desc : 
'==========================================================================================
Sub txtAdDesc_OnChange
 lgBlnFlgChgValue = True
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
			    <TR HEIGHT=23>
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
					<TD WIDTH=* ALIGN=RIGHT><A HREF="VBSCRIPT:OpenPopupTempGL()">결의전표</A>&nbsp;|&nbsp;<A HREF="VBSCRIPT:OpenPopupGL()">회계전표</A>&nbsp;|&nbsp;<Span id="RefPop"><A HREF="VBSCRIPT:OpenRcptNo()">가수금정보</A></Span></TD>
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
									<TD CLASS="TD5" NOWRAP>청산번호</TD>
									<TD CLASS="TD656" NOWRAP><INPUT NAME="txtAdJustNo" MAXLENGTH=18 STYLE="TEXT-ALIGN: Left" tag ="12XXXU" ALT="청산번호"><IMG align=top name=btnPrpaymNo src="../../image/btnPopup.gif"  TYPE="BUTTON" onclick="vbscript:OpenAdJustNo"></TD>								
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>	
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR HEIGHT=40% >
					<TD WIDTH=100% HEIGHT=* valign=top>
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD CLASS="TD5" NOWRAP>가수금번호</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtRcptNo" SIZE=20 MAXLENGTH=30 STYLE="TEXT-ALIGN: Left" tag="24" ALT="가수금번호"></TD>
								<TD CLASS="TD5" NOWRAP>거래처</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtBpCd" SIZE=10 MAXLENGTH=10 STYLE="TEXT-ALIGN: Left" tag="24" ALT="거래처">&nbsp;<INPUT TYPE=TEXT NAME="txtBpNm" SIZE=25 tag="24" ALT="거래처명"></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>입금일자</TD>
								<TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name=txtRcptDt CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="입금일자" tag="24X1" id=fpDateTime1></OBJECT>');</SCRIPT></TD>
								<TD CLASS="TD5" NOWRAP>부서</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtDeptCd" SIZE=10 MAXLENGTH=10 STYLE="TEXT-ALIGN: Left" tag="24" ALT="회계부서">&nbsp;<INPUT TYPE=TEXT NAME="txtDeptNm" SIZE=25 tag="24" ALT="회계부서명"></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>참조번호</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtRefNo" SIZE=20 MAXLENGTH=30 STYLE="TEXT-ALIGN: Left" tag="24" ALT="참조번호"></TD>
								<TD CLASS=TD5 NOWRAP>거래통화|환율</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtDocCur" SIZE=10 MAXLENGTH=4 tag="24NXXU" STYLE="TEXT-ALIGN: left" ALT="거래통화">&nbsp;
													<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> align ="top" name="txtXchRate" CLASS=FPDS140 title=FPDOUBLESINGLE ALT="환율" tag="24X5Z" id=OBJECT7></OBJECT>');</SCRIPT></TD>
						
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>입금금액|자국</TD>
							    <TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name=txtRcptAmt CLASS=FPDS140 title=FPDOUBLESINGLE ALT="입금금액" tag="24X2" ></OBJECT>');</SCRIPT>&nbsp;
							    <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name=txtRcptLocAmt CLASS=FPDS140 title=FPDOUBLESINGLE ALT="입금금액(자국)" tag="24X2" ></OBJECT>');</SCRIPT></TD>
								<TD CLASS="TD5" NOWRAP>잔액|자국</TD>
							    <TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name=txtBalAmt CLASS=FPDS140 title=FPDOUBLESINGLE ALT="잔액" tag="24X2"></OBJECT>');</SCRIPT>&nbsp;
								<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name=txtBalLocAmt CLASS=FPDS140 title=FPDOUBLESINGLE ALT="잔액(자국)" tag="24X2" ></OBJECT>');</SCRIPT></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>비고</TD>
								<TD CLASS="TD6" NOWRAP COLSPAN=3><INPUT TYPE=TEXT NAME="txtRcptDesc" SIZE=90 MAXLENGTH=128 tag="24" ALT="적요"></TD>
							</TR>						
						</TABLE>
					</TD>
				</TR>
				<TR height="60%">
					<TD WIDTH=100% HEIGHT=* valign=top>
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD CLASS="TD5" NOWRAP>청산일자</TD>
								<TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name=txtAdjustDt CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="청산일자" tag="22X1" id=OBJECT1></OBJECT>');</SCRIPT></TD>
								<TD CLASS="TD5" NOWRAP>계정코드</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtAcctCd" SIZE=10 MAXLENGTH=10 STYLE="TEXT-ALIGN: Left" tag="22XXXU" ALT="계정코드"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDocCur" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript: CALL OpenPopup(frm1.txtAcctCd.value,1)">&nbsp;<INPUT TYPE=TEXT NAME="txtAcctNm" SIZE=25 tag="24" ALT="계정명"></TD>
							</TR>
<!--							<TR>
								<TD CLASS=TD5 NOWRAP>거래통화</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtAdDocCur" SIZE=10 MAXLENGTH=4 tag="24NXXU" STYLE="TEXT-ALIGN: left" ALT="거래통화"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDocCur" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript: CALL OpenPopup(frm1.txtAdDocCur.value,2)">&nbsp;<INPUT TYPE=TEXT NAME="txtAdDocCurNm" SIZE=20 tag="24" ALT="거래통화명"></TD></TD>
								<TD CLASS=TD5 NOWRAP>환율</TD>
								<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name="txtAdXchRate" CLASS=FPDS140 title=FPDOUBLESINGLE ALT="환율" tag="24X5Z" id=OBJECT7></OBJECT>');</SCRIPT></TD>											
							</TR>-->
							<TR>
								
								<TD CLASS="TD5" NOWRAP>청산금액|자국</TD>
							    <TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name=txtAdjustAmt CLASS=FPDS140 title=FPDOUBLESINGLE ALT="청산금액" tag="22X2" id=OBJECT4></OBJECT>');</SCRIPT>&nbsp;
								<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name=txtAdjustLocAmt CLASS=FPDS140 title=FPDOUBLESINGLE ALT="청산금액(자국)" tag="21X2" id=OBJECT5></OBJECT>');</SCRIPT></TD>
								<TD CLASS=TD5 NOWRAP>결의/회계전표번호</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtTEMPGlNo" SIZE=18 MAXLENGTH=20 STYLE="TEXT-ALIGN: Left" tag="24XXXU" ALT="결의전표번호"> /
								<INPUT TYPE=TEXT NAME="txtGlNo" SIZE=18 MAXLENGTH=20 STYLE="TEXT-ALIGN: Left" tag="24XXXU" ALT="회계전표번호"> </TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>비고</TD>
								<TD CLASS="TD6" NOWRAP COLSPAN=3><INPUT TYPE=TEXT NAME="txtAdDesc" SIZE=70 MAXLENGTH=128 tag="21" ALT="비고"></TD>
							</TR>	
							<TR HEIGHT="55%">
								<TD WIDTH="100%" COLSPAN=4>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> HEIGHT="100%" name=vspdData2 width="100%" tag="2" TITLE="SPREAD" id=OBJECT6> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"></OBJECT>');</SCRIPT>
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
			<IFRAME NAME="MyBizASP" WIDTH="100%" HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=no noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA TYPE=hidden Class=hidden name=txtSpread2 tag="24" TABINDEX="-1"></TEXTAREA>
<INPUT TYPE=hidden NAME="txtMode"      tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="hUpdtUserId"  tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="hInsrtUserId" tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="txtFlgMode"   tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="hOrgChangeId" tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="hItemSeq"     tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="hAcctCd"      tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="txtMaxRows2"  tag="24" TABINDEX="-1">
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
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

