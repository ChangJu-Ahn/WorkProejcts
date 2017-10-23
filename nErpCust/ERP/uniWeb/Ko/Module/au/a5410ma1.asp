
<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : 회계 
*  2. Function Name        : 미결관리 
*  3. Program ID           : a5410ma
*  4. Program Name         : 개인별 카드내역 조회 
*  5. Program Desc         : 개인별 카드내역 조회 및 출력 
*  6. Comproxy List        :
*  7. Modified date(First) : 2002/11/11
*  8. Modified date(Last)  : 
*  9. Modifier (First)     : 문희정 
* 10. Modifier (Last)      :
* 11. Comment              :
=======================================================================================================-->
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
<!-- #Include file="../../inc/IncSvrHTML.inc"  -->

<!--
========================================================================================================
=                          3.2 Style Sheet
======================================================================================================== -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		

<!--
========================================================================================================
=                          3.3 Client Side Script
======================================================================================================== -->
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAMain.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAEvent.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliVariables.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAOperation.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliRdsQuery.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incEB.vbs">    </SCRIPT>

<Script Language="VBScript">
Option Explicit                                                        '☜: Turn on the Option Explicit option.
'########################################################################################################
'#                       4.  Data Declaration Part
'########################################################################################################

'========================================================================================================
'=                       4.1 External ASP File
'========================================================================================================
Const BIZ_PGM_ID = "a5410mb1.asp"                                      'Biz Logic ASP 
'========================================================================================================
'=                       4.2 Constant variables For spreadsheet
'========================================================================================================
'⊙: Grid Columns

'@Grid_Column
Dim C_GlDt			'일자 
Dim C_MgntVal1		'카드번호 
Dim C_EmpNM			'사용자 
Dim C_CreditNM		'카드사 
Dim C_GLDesc		'비고 
Dim C_OpenAmt		'금액 

'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim IsOpenPop          

' 권한관리 추가 
Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' 사업장 
Dim lgInternalCd, lgDeptCd, lgDeptNm			' 내부부서 
Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' 내부부서(하위포함)
Dim lgAuthUsrID, lgAuthUsrNm					' 개인 


'========================================================================================================
Sub InitVariables()
	lgIntFlgMode      = Parent.OPMD_CMODE						        '⊙: Indicates that current mode is Create mode
	lgBlnFlgChgValue  = False								    '⊙: Indicates that no value changed
    lgStrPrevKey      = ""                                      '⊙: initializes Previous Key
    lgSortKey         = 1                                       '⊙: initializes sort direction
		
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub


'========================================================================================================
	
Sub SetDefaultVal()
	Dim strSvrDate, strDayCnt
'--------------- 개발자 coding part(실행로직,Start)--------------------------------------------------
	
'		ServerDate	= GetSvrDate

		
		Dim strYear, strMonth, strDay, dtToday, EndDate, StartDate
		EndDate = "<%=GetSvrDate%>"
		Call ExtractDateFrom(EndDate, Parent.gServerDateFormat, Parent.gServerDateType, strYear, strMonth, strDay)
		
		StartDate = UniConvYYYYMMDDToDate(Parent.gDateFormat, strYear, strMonth, "01")
		EndDate = UniConvYYYYMMDDToDate(Parent.gDateFormat, strYear, strMonth, strDay)
	

		frm1.txtDate.Text	= StartDate 
		frm1.txtDate1.Text	= EndDate 	
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub
	

'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 

	<% Call loadInfTB19029A("Q", "A", "NOCOOKIE", "MA") %>

	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub


'========================================================================================================
Sub CookiePage(ByVal Kubun)
   '------ Developer Coding part (Start ) -------------------------------------------------------------- 
   '------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub



'========================================================================================================
Sub MakeKeyStream(pOpt)
   '------ Developer Coding part (Start ) -------------------------------------------------------------- 
   
   Select Case pOpt
	
       Case "MQ"          
			lgKeyStream = UNIConvDateToYYYYMMDD(Trim(frm1.txtDate.Text),parent.gDateFormat,"") & Parent.gColSep       'You Must append one character(Parent.gColSep)
			lgKeyStream = lgKeyStream & UNIConvDateToYYYYMMDD(Trim(frm1.txtDate1.Text),parent.gDateFormat,"") & Parent.gColSep       'You Must append one character(Parent.gColSep)
			
       Case "MR"          
			lgKeyStream = UNIConvDateToYYYYMMDD(Trim(frm1.txtDate.Text),parent.gDateFormat, "") & Parent.gColSep       'You Must append one character(Parent.gColSep)
			lgKeyStream = lgKeyStream & UNIConvDateToYYYYMMDD(Trim(frm1.txtDate1.Text),parent.gDateFormat,"") & Parent.gColSep       'You Must append one character(Parent.gColSep)
            
                  
   End Select 
                   
   '------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub        
	 

	
'========================================================================================================
Sub InitComboBox()
  
End Sub

'========================================================================================================
Sub InitData()

End Sub

'========================================================================================================
Sub initSpreadPosVariables()  
	 C_GlDt			= 1	'일자 
	 C_MgntVal1		= 2	'카드번호 
	 C_EmpNM		= 3	'사용자 
	 C_CreditNM		= 4	'카드사 
	 C_GLDesc		= 5	'비고 
	 C_OpenAmt		= 6	'금액 
End Sub



'========================================================================================================
Sub InitSpreadSheet()

	Call initSpreadPosVariables() 
	With frm1.vspdData
	
       .MaxCols   = C_openAmt + 1                                                  ' ☜:☜: Add 1 to Maxcols
	   .Col       = .MaxCols        : .ColHidden = True

        ggoSpread.Source = frm1.vspdData
        ggoSpread.Spreadinit "V20030204",,parent.gAllowDragDropSpread    
		ggoSpread.ClearSpreadData


	   .ReDraw = false
	   
	   Call GetSpreadColumnPos("A")
	
       Call AppendNumberPlace("6","4","2")

                             'ColumnPosition     Header            Width  Align(0:L,1:R,2:C)  Row   Length  CharCase(0:L,1:N,2:U)

       ggoSpread.SSSetEdit    C_GlDt            ,"일자"			,10     ,0                  ,     ,		,2
       ggoSpread.SSSetEdit    C_MgntVal1        ,"카드번호"     ,15    ,0                  ,     ,		,2
       ggoSpread.SSSetEdit    C_EmpNM			,"성명"			,14	   ,0                  ,     ,		,2       
       ggoSpread.SSSetEdit    C_CreditNM        ,"카드사"		,15	   ,0                  ,     ,		,2
       ggoSpread.SSSetEdit    C_GLDesc          ,"적요"			,45    ,0                  ,     ,		,2
      
                             'Col                Header            Width  Grp            IntegeralPart       DeciPointpart                 Align   Sep    PZ   Min       Max 
       ggoSpread.SSSetFloat   C_OpenAmt			,"금액"			,16    ,"2",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,	,		,"Z"

	   .ReDraw = true
	
       Call SetSpreadLock 
    
    End With
    
   
End Sub


'======================================================================================================

Sub SetSpreadLock()
    ggoSpread.Source = frm1.vspdData
    ggoSpread.SpreadLockWithOddEvenRowColor()
    
End Sub



'======================================================================================================
Sub SetSpreadColor(ByVal pvStartRow,ByVal pvEndRow)
    With frm1
    
    .vspdData.ReDraw = False
                                'Col          Row   Row2
      ggoSpread.SSSetProtected   C_GlDt		,pvStartRow	,pvEndRow
      ggoSpread.SSSetProtected   C_MgntVal1	,pvStartRow	,pvEndRow
      ggoSpread.SSSetProtected   C_EmpNM	,pvStartRow	,pvEndRow     
      ggoSpread.SSSetProtected   C_CreditNM	,pvStartRow	,pvEndRow
      ggoSpread.SSSetProtected   C_GLDesc	,pvStartRow	,pvEndRow
      ggoSpread.SSSetProtected   C_OpenAmt	,pvStartRow	,pvEndRow      
      
    .vspdData.ReDraw = True
    
    End With
End Sub


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
              Frm1.vspdData.Action = 0 ' go to 
              Exit For
           End If
           
       Next
          
    End If   
End Sub


'========================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

			C_GlDt    		= iCurColumnPos(1)
			C_MgntVal1  	= iCurColumnPos(2)
			C_EmpNM    		= iCurColumnPos(3)    
			C_CreditNM   	= iCurColumnPos(4)
			C_GLDesc   		= iCurColumnPos(5)
			C_OpenAmt    	= iCurColumnPos(6)

    End Select    
End Sub

'========================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub


'========================================================================================
Sub PopRestoreSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()      
'   Call InitComboBox
'	Call initData
	Call ggoSpread.ReOrderingSpreadData()
End Sub



'========================================================================================================
Sub Form_Load()
    Err.Clear                                                                        '☜: Clear err status
    
	Call LoadInfTB19029                                                              '☜: Load table , B_numeric_format
		
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
      
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
	Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
	Call ggoOper.FormatField(Document, "3",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
	Call ggoOper.FormatField(Document, "4",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)

	Call ggoOper.LockField(Document, "N")											 '⊙: Lock Field
    
    Call InitSpreadSheet                                                             'Setup the Spread sheet

	Call InitVariables
    Call SetDefaultVal

	frm1.txtDate.focus
	
	
'①메뉴탐색기 ②조회    ③신규    ④삭제        ⑤저장    ⑥행추가       ⑦행삭제 
'⑧취소       ⑨이전    ⑩이후    ⑪레코드복사  ⑫Export  ⑬인쇄         ⑭찾기	
	Call SetToolbar("11000000000111")                                              '☆: Developer must customize
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
	
	Set gActiveElement = document.activeElement		
	
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
	

'========================================================================================================
Sub Form_QueryUnload(Cancel, UnloadMode)
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    '--------- Developer Coding Part (End) ------------------------------------------------------------
End Sub




'========================================================================================================
Function FncQuery()
    Dim IntRetCD 
    
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncQuery = False															  '☜: Processing is NG

    ggoSpread.Source = Frm1.vspdData
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO,"x","x")					  '☜: Data is changed.  Do you want to display it? 
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    

  ' Call ggoOper.ClearField(Document, "2")										  '☜: Clear Contents  Field
	Call InitSpreadSheet     															
    If Not chkField(Document, "1") Then									          '☜: This function check required field
       Exit Function
    End If

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 

    Call InitVariables                                                            '⊙: Initializes local global variables

	If DbQuery("MQ") = False Then                                                 '☜: Query db data
       Exit Function
    End If
    
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    If Err.number = 0 Then	
       FncQuery = True                                                            '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement   

End Function
	

'========================================================================================================
Function FncNew()
    On Error Resume Next                                                          '☜: If process fails
End Function
	

'========================================================================================================
Function FncDelete()
    On Error Resume Next                                                          '☜: If process fails
End Function

'=======================================================================================================
Function FncSave() 
    On Error Resume Next                                                          '☜: If process fails
End Function

'========================================================================================================
Function FncCopy()

End Function


'========================================================================================================
Function FncCancel() 
  
End Function


'========================================================================================================
Function FncInsertRow()


End Function


'========================================================================================================
Function FncDeleteRow()


End Function

'========================================================================================================
Sub fpdtFoundDt_ButtonHit(Button, NewIndex)
	On Error Resume Next
    lgBlnFlgChgValue = True
End Sub


'========================================================================================================
Sub fpdtCloseDt_ButtonHit(Button, NewIndex)
	On Error Resume Next
    lgBlnFlgChgValue = True
End Sub



'========================================================================================================
Function FncPrint() 

Dim StrEbrFile
Dim StrUrl
Dim IntRetCd
	
	If Not chkField(Document, "1") Then									'⊙: This function check indispensable field
	   Exit Function
	End If
		
	Call SetPrintCond(StrEbrFile, StrUrl)
	
    ObjName = AskEBDocumentName(StrEbrFile, "ebr")

	Call FncEBRPrint(EBAction,ObjName,StrUrl)
	
End Function


'========================================================================================================
Function FncPreview()
 
Dim StrEbrFile
Dim StrUrl
Dim IntRetCd
	
	If Not chkField(Document, "1") Then									'⊙: This function check indispensable field
	   Exit Function
	End If
		
	Call SetPrintCond(StrEbrFile, StrUrl)
	
	ObjName = AskEBDocumentName(StrEbrFile, "ebr")
	
	Call FncEBRPreview(ObjName,StrUrl)
			
End Function


'=======================================================================================================
Sub SetPrintCond(StrEbrFile, StrUrl)
	
	Dim	ValGlDate, ValEmpNo, ValFromDate, ValToDate, ValEmpNM
	Dim	strAuthCond

	StrEbrFile = "a5410ma1"
	
	With frm1
		ValFromDate	= (UniConvDateAToB(Trim(frm1.txtDate.Text),Parent.gDateFormat, Parent.gServerDateFormat))
		ValToDate	= (UniConvDateAToB(Trim(frm1.txtDate1.Text),Parent.gDateFormat, Parent.gServerDateFormat))
		ValEmpNo	= UCase(Trim(.txtEmpNo.value))
		ValEmpNM	= Trim(.txtEmpNM.value)
	End With

	If Trim(ValEmpNo) = "" Then		ValEmpNo = "%"

	' 권한관리 추가 
	strAuthCond		= "	"
	
	If lgAuthBizAreaCd <> "" Then			
		strAuthCond		= strAuthCond	& " AND d.BIZ_AREA_CD = " & FilterVar(lgAuthBizAreaCd, "''", "S")  		
	End If			

	If lgInternalCd <> "" Then			
		strAuthCond		= strAuthCond	& " AND d.INTERNAL_CD = " & FilterVar(lgInternalCd, "''", "S")  		
	End If			

	If lgSubInternalCd <> "" Then	
		strAuthCond		= strAuthCond	& " AND d.INTERNAL_CD LIKE " & FilterVar(lgSubInternalCd & "%", "''", "S")  
	End If	

	If lgAuthUsrID <> "" Then	
		strAuthCond		= strAuthCond	& " AND d.INSRT_USER_ID = " & FilterVar(lgAuthUsrID, "''", "S")  
	End If	


	StrUrl = StrUrl & "FromDate|"		& ValFromDate
	StrUrl = StrUrl & "|ToDate1|"		& ValToDate
	StrUrl = StrUrl & "|EMPNo|"			& ValEmpNo
	StrUrl = StrUrl & "|EMPNM|"			& ValEmpNM

	StrUrl = StrUrl & "|strAuthCond|"		& strAuthCond


End Sub
	

'========================================================================================================
Function FncNext() 

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncNext = False                                                               '☜: Processing is NG
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    
    '--------- Developer Coding Part (End) ------------------------------------------------------------
    If Err.number = 0 Then	 
       FncNext = True                                                             '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement   

End Function


'========================================================================================================
Function FncExcel() 

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncExcel = False                                                              '☜: Processing is NG

	Call Parent.FncExport(Parent.C_MULTI)

    If Err.number = 0 Then	 
       FncExcel = True                                                            '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement   

End Function


'========================================================================================================
Function FncFind() 

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncFind = False                                                               '☜: Processing is NG

	Call Parent.FncFind(Parent.C_MULTI, True)

    If Err.number = 0 Then	 
       FncFind = True                                                             '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement   

End Function


'========================================================================================
Sub FncSplitColumn()

    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   

End Sub


'========================================================================================================
Function FncExit()
	Dim IntRetCD

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncExit = False                                                               '☜: Processing is NG
	
    ggoSpread.Source = frm1.vspdData	
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO,"x","x")		              '⊙: Data is changed.  Do you want to exit? 
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

    If Err.number = 0 Then	 
       FncExit = True                                                             '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement   

End Function



'========================================================================================================

Function DbQuery(pDirect)

	Dim strVal
	
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status
    
    DbQuery = False                                                               '☜: Processing is NG
	
    Call DisableToolBar(Parent.TBC_QUERY)                                                '☜: Disable Query Button Of ToolBar
    Call LayerShowHide(1)                                                         '☜: Show Processing Message

	Call MakeKeyStream(pDirect)
  
    With Frm1
		strVal = BIZ_PGM_ID	& "?txtMode="		& Parent.UID_M0001						         
        strVal = strVal		& "&txtKeyStream="	& lgKeyStream         '☜: Query Key
        strVal = strVal		& "&txtMaxRows="	& .vspdData.MaxRows
        strVal = strVal		& "&lgStrPrevKey="	& lgStrPrevKey        '☜: Next key tag
    End With

	If Trim(frm1.txtEmpNo.value) <> "" then strVal = strVal & "&txtEmpNo=" & Trim(frm1.txtEmpNo.Value) 
	
	' 권한관리 추가 
	strVal = strVal	& "&lgAuthBizAreaCd="	& lgAuthBizAreaCd			' 사업장 
	strVal = strVal & "&lgInternalCd="		& lgInternalCd				' 내부부서 
	strVal = strVal & "&lgSubInternalCd="	& lgSubInternalCd			' 내부부서(하위포함)
	strVal = strVal & "&lgAuthUsrID="		& lgAuthUsrID				' 개인 


	if lgStrPrevKey = "" then	frm1.txtSumAmt.text = 0

    Call RunMyBizASP(MyBizASP, strVal)                                           '☜:  Run biz logic

    If Err.number = 0 Then	 
       DbQuery = True                                                             '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement   

End Function

Function DbSave()
    On Error Resume Next                                                          '☜: If process fails
End Function

'========================================================================================
Function DbSaveOk()					'☆: 저장 성공후 실행 로직 
    On Error Resume Next                                                          '☜: If process fails
End Function

'========================================================================================================
Function DbDelete()
    On Error Resume Next                                                          '☜: If process fails
End Function

'========================================================================================================
Sub DbQueryOk()
    On Error Resume Next                                                          '☜: If process fails
End Sub
	
'========================================================================================================
Sub DbDeleteOk()
    On Error Resume Next                                                          '☜: If process fails
End Sub

'========================================================================================================

Sub  txtDate_DblClick(Button)
    If Button = 1 Then
        frm1.txtDate.Action = 7                        
    End If
End Sub

Sub txtDate_KeyDown(KeyCode, Shift)
	If KeyCode = 13 Then Call FncQuery
End Sub

Sub  txtDate1_DblClick(Button)
    If Button = 1 Then
        frm1.txtDate1.Action = 7                        
    End If
End Sub

Sub txtDate1_KeyDown(KeyCode, Shift)
	If KeyCode = 13 Then Call FncQuery
End Sub


Function OpenPopUp(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim strCd, strTempBankCd

	If IsOpenPop = True Then Exit Function

	
	Select Case iWhere
		Case 0		'성명 

	
			arrParam(0) = "성명 팝업"									' 팝업 명칭 
			arrParam(1) = "haa010t b(nolock), b_credit_card a(nolock)" 						' TABLE 명칭 
			arrParam(2) = Trim(strCode)										' Code Condition
			arrParam(3) = ""												' Name Cindition
			arrParam(4) = "b.emp_no=*a.use_user_id and isnull(a.use_user_id,'') <> ''"	' Where Condition
			arrParam(5) = "성명"										' 조건필드의 라벨 명칭 

			arrField(0) = "a.use_user_id"									' Field명(0)
			arrField(1) = "isnull(b.name,a.use_user_id)"					' Field명(1)
    		arrField(2) = "a.rgst_no"										' Field명(2)

			
			arrHeader(0) = "사원번호"										' Header명(0)
			arrHeader(1) = "성명"											' Header명(1)
			arrHeader(2) = "주민등록번호"									' Header명(2)
	
		Case Else
			Exit Function
	End Select
	
	IsOpenPop = True
	
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
	    frm1.txtEmpNo.focus
		Exit Function
	Else
		frm1.txtEmpNo.focus
		frm1.txtEmpNo.value = arrRet(0)
		frm1.txtEmpNM.value = arrRet(1)
	End If	

End Function


'========================================================================================================
Sub vspdData_ButtonClicked(Col, Row, ButtonDown)

End Sub


'========================================================================================================
Sub vspdData_Change(Col , Row)

    Dim iDx
       
   	Frm1.vspdData.Row = Row
   	Frm1.vspdData.Col = Col

        
    Call CheckMinNumSpread(frm1.vspdData, Col, Row)		
	
	ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row
End Sub


'========================================================================================================
Sub vspdData_Click(Col, Row)

	Call SetPopupMenuItemInf("0000111111")
    gMouseClickStatus = "SPC"
    Set gActiveSpdSheet = frm1.vspdData
    
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
    
End Sub


'========================================================================================================
Sub vspdData_GotFocus()
    ggoSpread.Source = Frm1.vspdData


End Sub


'========================================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If

	
End Sub    


'========================================================================================================
sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
	ggoSpread.source = frm1.vspdData
	Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)
End Sub


'========================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
End Sub
  

'========================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft ,ByVal OldTop ,ByVal NewLeft ,ByVal NewTop )

    If OldLeft <> NewLeft Then
       Exit Sub
    End If

	If CheckRunningBizProcess = True Then
	   Exit Sub
	End If
    
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	           
    	If lgStrPrevKey <> "" Then                         
           If DbQuery("MN") = False Then
              Call RestoreToolBar()
              Exit Sub
           End if
        End If
    End if
End Sub


</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>

<!--
'########################################################################################################
'#						6. TAG 부																		#
'######################################################################################################## 
-->
<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR>
		<TD <%=HEIGHT_TYPE_00%>></TD>
	</TR>
	
	<TR HEIGHT=23>
		<TD WIDTH=100%>
			<TABLE  CLASS="BasicTB" CELLSPACING=0 CELLPADDING=0>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 >
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><IMG height=23 src="../../../CShared/image/table/seltab_up_left.gif" width=9></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white><%=Request("strASPMnuMnuNm")%></font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><IMG height=23 src="../../../CShared/image/table/seltab_up_right.gif" width=10></td>
						    </TR>
						</TABLE>
					</TD>

					<TD WIDTH=* ALIGN=RIGHT></TD>
					<TD WIDTH=*>&nbsp;</TD>
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
									<TD CLASS="TD5" NOWRAP>작업일자</TD>
									<TD CLASS="TD6" NOWRAP>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME="txtDate" CLASS=FPDTYYYYMMDD tag="22" Title="FPDATETIME" ALT="시작일" id=fpDateTime1></OBJECT>');</SCRIPT>&nbsp;~&nbsp;
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME="txtDate1" CLASS=FPDTYYYYMMDD tag="22" Title="FPDATETIME" ALT="종료일" id=fpDateTime1></OBJECT>');</SCRIPT>
									</TD>
									<TD CLASS="TD5" NOWRAP>성명</TD>									
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtEmpNo" ALT="사번" MAXLENGTH="30" SIZE=15 STYLE="TEXT-ALIGN: left" tag  ="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnEmpNo" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(frm1.txtEmpNo.Value,0)">
														 <INPUT NAME="txtEmpNM" ALT="성명" MAXLENGTH="30" SIZE=15 STYLE="TEXT-ALIGN: left" tag  ="14">
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
					<TD WIDTH=100% HEIGHT=* valign=top>
						<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD HEIGHT="100%" NOWRAP>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=OBJECT1> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
								</TD>
							</TR>

						</TABLE>
					</TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH=100%  >
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_40%>>		
								<TR>
									<TD <%=HEIGHT_TYPE_03%> WIDTH=80%></TD>

									<TD CLASS=TD5 NOWRAP >합계</TD>
									<TD ><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name=txtSumAmt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 150px" title=FPDOUBLESINGLE ALT="합계" tag="44X2" id=txtSumAmt></OBJECT>');</SCRIPT></TD>
									<TD WIDTH=10></TD>						
								</TR>																													
							</TABLE>
						</FIELDSET>
					</TD>
					
				</TR>				

			</TABLE>
		</TD>
	</TR>
	<TR >
		<TD <%=HEIGHT_TYPE_01%>></TD>
	</TR>
	<TR HEIGHT=20>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_30%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD><BUTTON NAME="btnPreview" CLASS="CLSSBTN" ONCLICK="vbscript:FncPreview()" Flag=1>미리보기</BUTTON>&nbsp;
						<BUTTON NAME="btnPrint" CLASS="CLSSBTN" ONCLICK="vbscript:FncPrint()" Flag=1>인쇄</BUTTON>&nbsp;
					</TD>
					<TD WIDTH=* ALIGN=RIGHT></TD>
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
<TEXTAREA class=hidden name=txtSpread    tag="24" TABINDEX="-1"></TEXTAREA>
<TEXTAREA class=hidden name=txtSpread3   tag="24" TABINDEX="-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMaxRows"      TAG="24">
<INPUT TYPE=HIDDEN NAME="txtMode"         TAG="24">
<INPUT TYPE=HIDDEN NAME="txtKeyStream"    TAG="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode"      TAG="24">
<INPUT TYPE=hidden NAME="htxtTempGlNo"   tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="txtCommandMode" tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="hOrgChangeId"   tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="htxtAcctCd"     TAG="X4">
<INPUT TYPE=HIDDEN NAME="htxtAcctCd1"    TAG="X4">
<INPUT TYPE=HIDDEN NAME="htxtDate"     TAG="X4">
<INPUT TYPE=HIDDEN NAME="htxtDateTo"     TAG="X4">
<INPUT TYPE=HIDDEN NAME="htxtBankAcctNo" TAG="X4">
<INPUT TYPE=HIDDEN NAME="htxtEmpNo" TAG="X4">
<INPUT TYPE=HIDDEN NAME="hCongFg"        tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="hItemSeq"    tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="hAcctCd"     tag="24" TABINDEX="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm"></IFRAME>
</DIV>
<FORM NAME=EBAction TARGET="MyBizASP" METHOD="POST">
</BODY>
</HTML>

