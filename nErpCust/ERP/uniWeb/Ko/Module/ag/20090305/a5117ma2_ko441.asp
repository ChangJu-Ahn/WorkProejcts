<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : Template
*  2. Function Name        : 
*  3. Program ID           : 
*  4. Program Name         : 
*  5. Program Desc         : Multi Sample
*  6. Comproxy List        :
*  7. Modified date(First) : 2001/04/18
*  8. Modified date(Last)  : 2001/04/18
*  9. Modifier (First)     :
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
<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->
<!--
========================================================================================================
=                          3.2 Style Sheet
======================================================================================================== -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		

<!--
========================================================================================================
=                          3.3 Client Side Script
======================================================================================================== -->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMaMain.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMaOperation.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/IncCliRdsQuery.vbs">   </SCRIPT>

<Script Language="VBScript">
Option Explicit                                                        '☜: Turn on the Option Explicit option.
'########################################################################################################
'#                       4.  Data Declaration Part
'########################################################################################################

'========================================================================================================
'=                       4.1 External ASP File
'========================================================================================================
Const BIZ_PGM_ID = "A5117MB1_KO441.asp"                                      'Biz Logic ASP 
'========================================================================================================
'=                       4.2 Constant variables For spreadsheet
'========================================================================================================

Dim C_SEQ
Dim C_TEMP_GL_DT
Dim C_BIZ_AREA_CD
Dim C_BIZ_AREA_NM
Dim C_TEMP_GL_NO
Dim C_GL_NO
Dim C_USER_NM
Dim C_ACCT_CD
Dim C_ACCT_NM
Dim C_DR_ITEM_LOC_AMT
Dim c_CR_ITEM_LOC_AMT
Dim C_ITEM_DESC
Dim C_GL_INPUT_TYPE
Dim C_GL_INPUT_TYPE_NM
Dim C_DEPT_CD
Dim C_DEPT_NM
Dim C_COST_CD
Dim C_COST_NM
Dim C_ITEM_SEQ


'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================

' 권한관리 추가 
Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' 사업장 
Dim lgInternalCd, lgDeptCd, lgDeptNm			' 내부부서		
Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' 내부부서(하위포함)				
Dim lgAuthUsrID, lgAuthUsrNm					' 개인 

Dim IsOpenPop          
<% 
   BaseDate     = GetSvrDate                                                                  'Get DB Server Date
   
Dim lsSvrDate
lsSvrDate = GetSvrDate

%>   
'========================================================================================================
'========================================================================================================
'                        5.1 Common Group-1
'========================================================================================================
'========================================================================================================
' Name : InitSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================
Sub InitSpreadPosVariables()
	C_SEQ				= 1
	C_TEMP_GL_DT		= 2      
	C_BIZ_AREA_CD		= 3     
	C_BIZ_AREA_NM		= 4	     
	C_TEMP_GL_NO        = 5
	C_GL_NO				= 6
	C_USER_NM			= 7
	C_ACCT_CD           = 8
	C_ACCT_NM           = 9
	C_DR_ITEM_LOC_AMT   = 10
	c_CR_ITEM_LOC_AMT   = 11
	C_ITEM_DESC         = 12
	C_GL_INPUT_TYPE     = 13
	C_GL_INPUT_TYPE_NM  = 14
	C_DEPT_CD			= 15
	C_DEPT_NM			= 16
	C_COST_CD			= 17
	C_COST_NM			= 18
	C_ITEM_SEQ			= 19	
End Sub

'========================================================================================================
' Name : InitVariables()	
' Desc : Initialize value
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
' Name : SetDefaultVal()	
' Desc : Set default value
'========================================================================================================
	
Sub SetDefaultVal()

	frm1.txtfromGlDt.text = UniConvDateAToB("<%=lsSvrDate%>", parent.gServerDateFormat, parent.gDateFormat)
    frm1.txttoGlDt.text   = UniConvDateAToB("<%=lsSvrDate%>", parent.gServerDateFormat, parent.gDateFormat)
	frm1.txtfromGlDt.focus 
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	frm1.hOrgChangeId.value = parent.gChangeOrgId
	
	frm1.cboConfFg.value	=	"U"
	Call cboConfFg_OnChange()
	frm1.txtUsr_ID.value = parent.gUsrId
	frm1.txtUsr_NM.value = parent.gUsrNm
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
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	<% Call loadInfTB19029A("I", "*","NOCOOKIE","MA") %>
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub

'========================================================================================================
' Name : CookiePage()
' Desc : Write or Read cookie value 
'========================================================================================================
Sub CookiePage(ByVal Kubun)
   '------ Developer Coding part (Start ) -------------------------------------------------------------- 
   '------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub

'========================================================================================================
' Name : MakeKeyStream
' Desc : This method set focus to pos of err
'========================================================================================================
Sub MakeKeyStream(ByVal pOpt)
   
   '------ Developer Coding part (Start ) -------------------------------------------------------------- 
   'Select Case pOpt
   '    Case "MQ"
   '               'lgKeyStream = frm1.txtPlantCd.Value  & Parent.gColSep       'You Must append one character(Parent.gColSep)
   '    Case "MN"
                  'lgKeyStream = Frm1.htxtPlantCd.Value & Parent.gColSep       'You Must append one character(Parent.gColSep)
   'End Select                 
       
   '------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub        

	
'============================================================================================================
Sub InitComboBox()
	
	Err.clear
	Call CommonQueryRs("minor_cd, minor_nm", "b_minor", "major_cd=" & FilterVar("A1001", "''", "S") & "  order by minor_nm", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
	Call SetCombo2(frm1.cboGlInputType ,lgF0  ,lgF1  ,Chr(11))
	
End Sub

Sub InitComboBox_cond()
	Dim intRetCd,intLoopCnt
	Dim ArrayTemp1
	Dim ArrayTemp2
	IntRetCd = CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("A1007", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) 		    
	
	If IntRetCD=False  Then
	    Call DisplayMsgBox("122300","X","X","X")                         '☜ : Minor코드정보가 없습니다.
	Else
		ArrayTemp1 = Split(lgF0,Chr(11))
		ArrayTemp2 = Split(lgF1,Chr(11))

		For intLoopCnt = 0 To UBound(ArrayTemp1,1) -1
			Call SetCombo(frm1.cboConfFg, ArrayTemp1(intLoopCnt), ArrayTemp2(intLoopCnt))
		Next  

	End If
End Sub

'========================================================================================================
' Name : InitData()
' Desc : Reset ComboBox
'========================================================================================================
Sub InitData()
	Dim intRow
	Dim intIndex 

	With frm1.vspdData
		For intRow = 1 To .MaxRows			
			.Row = intRow
			.Col = 1
			
			select case Trim(.Value)

			Case  "1" 
			    
			    .Col = -1 
			    .Col2 = -1
			    .BackColor = RGB(230, 230, 250)
			Case  "2" 
			    
			    .Col = -1 
			    .Col2 = -1
			    .BackColor = RGB(255, 165, 0)
			    .ForeColor = vbBlue

			End select        
	     next	
	End With
End Sub

'========================================================================================================
' Name : InitSpreadSheet
' Desc : This method initializes spread sheet column property
'========================================================================================================
Sub InitSpreadSheet()
	Call initSpreadPosVariables()
 
	With frm1.vspdData
		
		ggoSpread.Source = frm1.vspdData
		ggoSpread.Spreadinit "V20041105",, parent.gAllowDragDropSpread
		.ReDraw = false
		.MaxCols   = C_ITEM_SEQ + 1                                                  ' ☜:☜: Add 1 to Maxcols
		Call ggoSpread.ClearSpreadData()
		'Call AppendNumberPlace("6","4","2")
		Call GetSpreadColumnPos("A")
			
		ggoSpread.SSSetEdit    C_SEQ		  		  ,"구분" 		,10    ,0                  ,     ,100     ,2	   
		ggoSpread.SSSetDate    C_TEMP_GL_DT		      ,"결의일"     ,13    ,2                  ,parent.gDateFormat   ,-1
		ggoSpread.SSSetEdit    C_BIZ_AREA_CD		  ,"사업장" 	,8    ,0                   ,     ,100     ,2
		ggoSpread.SSSetEdit    C_BIZ_AREA_NM		  ,"사업장명"   ,13    ,0                  ,     ,100     ,2
		ggoSpread.SSSetEdit    C_TEMP_GL_NO           ,"결의번호"	,18    ,0                  ,     ,100     ,2
		ggoSpread.SSSetEdit    C_GL_NO				  ,"회계번호"	,18    ,0                  ,     ,100     ,2
		ggoSpread.SSSetEdit    C_USER_NM              ,"승인자" 	,10    ,0                  ,     ,100     ,2
		ggoSpread.SSSetEdit    C_ACCT_CD              ,"계정코드" 	,8     ,0                  ,     ,100     ,2
		ggoSpread.SSSetEdit    C_ACCT_NM              ,"계정명" 	,15    ,0                  ,     ,100     ,2
		ggoSpread.SSSetFloat   C_DR_ITEM_LOC_AMT      ,"차변" 		,15, parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec
		ggoSpread.SSSetFloat   C_CR_ITEM_LOC_AMT      ,"대변" 		,15, parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec
		ggoSpread.SSSetEdit    C_ITEM_DESC            ,"적요" 		,35    ,0                  ,     ,100     ,2
		ggoSpread.SSSetEdit    C_GL_INPUT_TYPE        ,"입력경로"   ,10    ,0                  ,     ,100     ,1
		ggoSpread.SSSetEdit    C_GL_INPUT_TYPE_NM     ,"입력경로"   ,15    ,0                  ,     ,100     ,2
		ggoSpread.SSSetEdit    C_DEPT_CD              ,"부서" 		,8     ,0                  ,     ,100     ,2
		ggoSpread.SSSetEdit    C_DEPT_NM              ,"부서명" 	,15    ,0                  ,     ,100     ,2						
		ggoSpread.SSSetEdit    C_COST_CD              ,"코스트" 	,8     ,0                  ,     ,100     ,2
		ggoSpread.SSSetEdit    C_COST_NM              ,"코스트명" 	,15    ,0                  ,     ,100     ,2
		ggoSpread.SSSetEdit    C_ITEM_SEQ             ,"순번" 		,8     ,0                  ,     ,100     ,2

        Call ggoSpread.SSSetColHidden(C_SEQ, C_SEQ,True)
		'Call ggoSpread.SSSetColHidden(C_BIZ_AREA_CD, C_BIZ_AREA_CD, True)
		Call ggoSpread.SSSetColHidden(C_GL_INPUT_TYPE, C_GL_INPUT_TYPE, True)
		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
		
		Call ggoSpread.SSSetSplit2(C_TEMP_GL_NO)

	   .ReDraw = true
       Call SetSpreadLock 
    End With    
End Sub

'======================================================================================================
' Name : SetSpreadLock
' Desc : This method set color and protect cell in spread sheet
'======================================================================================================
Sub SetSpreadLock()
    With frm1
    
	    .vspdData.ReDraw = False 
	                                 'Col-1             Row-1       Col-2           Row-2   
	    ggoSpread.SpreadLock       C_SEQ		, -1         , C_ITEM_SEQ        , -1   
	                                   
	    .vspdData.ReDraw = True

    End With
End Sub

'======================================================================================================
' Name : SetSpreadColor
' Desc : This method set color and protect cell by cell in spread sheet
'======================================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
    'With frm1    
       '.vspdData.ReDraw = False
       '                          'Col          Row         Row2
       'ggoSpread.SSSetRequired    C_SID      , pvStartRow, pvEndRow
       'ggoSpread.SSSetRequired    C_SNm      , pvStartRow, pvEndRow
       '                          'Col          Row          Row2
       'ggoSpread.SSSetProtected   C_AddressNm, pvStartRow, pvEndRow
       '.vspdData.ReDraw = True
    'End With
End Sub

'======================================================================================================
' Name : SubSetErrPos
' Desc : This method set focus to position of error
'      : This method is called in MB area
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
' Function Name : GetSpreadColumnPos
' Description   : 
'========================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
	            
			C_SEQ				= iCurColumnPos(1)      
			C_TEMP_GL_DT		= iCurColumnPos(2)     
			C_BIZ_AREA_CD		= iCurColumnPos(3)	     
			C_BIZ_AREA_NM		= iCurColumnPos(4)
			C_TEMP_GL_NO        = iCurColumnPos(5)
			C_GL_NO				= iCurColumnPos(6)
			C_USER_NM			= iCurColumnPos(7)
			C_ACCT_CD           = iCurColumnPos(8)
			C_ACCT_NM           = iCurColumnPos(9)
			C_DR_ITEM_LOC_AMT   = iCurColumnPos(10)
			c_CR_ITEM_LOC_AMT   = iCurColumnPos(11)
			C_ITEM_DESC         = iCurColumnPos(12)
			C_GL_INPUT_TYPE     = iCurColumnPos(13)
			C_GL_INPUT_TYPE_NM  = iCurColumnPos(14)
			C_DEPT_CD			= iCurColumnPos(15)
			C_DEPT_NM			= iCurColumnPos(16)
			C_COST_CD			= iCurColumnPos(17)
			C_COST_NM			= iCurColumnPos(18)			
			C_ITEM_SEQ			= iCurColumnPos(19)					          
			
    End Select		
End Sub

'========================================================================================================
'========================================================================================================
'                        5.2 Common Group-2
'========================================================================================================
'========================================================================================================

'========================================================================================================
' Name : Form_Load
' Desc : This sub is called from window_OnLoad in Common.vbs automatically
'========================================================================================================
Sub Form_Load()
    Err.Clear                                                                        '☜: Clear err status

	Call LoadInfTB19029                                                              '☜: Load table , B_numeric_format
		
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
      
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
	Call ggoOper.LockField(Document, "N")											 '⊙: Lock Field
    Call InitSpreadSheet                                                             'Setup the Spread sheet
	Call InitVariables
    Call SetDefaultVal

    Call SetToolbar("1100000000001111")										
    Call InitComboBox()
    Call InitComboBox_Cond
    Call CookiePage(0)
	frm1.txtAmtFr.Text	= ""
	frm1.txtAmtTo.Text	= ""
    frm1.txtFromGlDt.focus                                         '☆: Developer must customize

	'------ Developer Coding part (End )   -------------------------------------------------------------- 
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
' Name : Form_QueryUnload
' Desc : This sub is called from window_Unonload in Common.vbs automatically
'========================================================================================================
Sub Form_QueryUnload(Cancel, UnloadMode)
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    '--------- Developer Coding Part (End) ------------------------------------------------------------
End Sub



'========================================================================================================
Sub  cboConfFg_OnChange()
    lgBlnFlgChgValue = True
End Sub
'========================================================================================================
Sub txtDeptCD_OnChange()
        
    Dim strSelect, strFrom, strWhere 	
    Dim IntRetCD 
	Dim arrVal1, arrVal2
	Dim ii, jj

	if frm1.txtDeptCd.value = "" then
		frm1.txtDeptNm.value = ""
	end if
	
    lgBlnFlgChgValue = True
	
	If TRim(frm1.txtDeptCd.value) <>"" Then
		'----------------------------------------------------------------------------------------
			strSelect = "dept_cd, ORG_CHANGE_ID"
			strFrom =  " B_ACCT_DEPT "
			strWhere = " ORG_CHANGE_DT >= "
			strWhere = strWhere & " (select max(org_change_dt) from B_ACCT_DEPT where org_change_dt<= " & FilterVar(UNIConvDateToYYYYMMDD(frm1.txtFromGlDt.Text, gDateFormat,Parent.gServerDateType), "''", "S") & ")"
			strWhere = strWhere & " AND ORG_CHANGE_DT <= " 
			strWhere = strWhere & " (select max(org_change_dt) from B_ACCT_DEPT where org_change_dt<= " & FilterVar(UNIConvDateToYYYYMMDD(frm1.txtToGlDt.Text, gDateFormat,Parent.gServerDateType), "''", "S") & ") "
			strWhere =	strWhere & " AND dept_Cd = " & FilterVar(LTrim(RTrim(frm1.txtDeptCd.value)), "''", "S")
		
		If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) = False Then
	
			IntRetCD = DisplayMsgBox("124600","X","X","X")  
			frm1.txtDeptCd.value = ""
			frm1.txtDeptNm.value = ""
			frm1.hOrgChangeId.value = ""
			frm1.txtDeptCd.focus
		Else 
		
			arrVal1 = Split(lgF2By2, Chr(11) & Chr(12))			
			jj = Ubound(arrVal1,1)
					
			For ii = 0 to jj - 1
				arrVal2 = Split(arrVal1(ii), chr(11))			
				frm1.hOrgChangeId.value = Trim(arrVal2(2))
				
			Next	
			
		End If
	End IF
		'----------------------------------------------------------------------------------------

End Sub
'========================================================================================================

Sub txtFromGlDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtFromGlDt.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtFromGlDt.Focus
    End If
End Sub
'========================================================================================================

Sub txtFromGlDt_Change() 
    lgBlnFlgChgValue = True
End Sub
'========================================================================================================

Sub txtToGlDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtToGlDt.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtToGlDt.Focus
    End If
End Sub
'========================================================================================================

Sub txtToGlDt_Change() 
    lgBlnFlgChgValue = True
End Sub

'========================================================================================================
Sub txtFromGlDt_KeyPress(KeyAscii)
	If KeyAscii = 13 Then 
	   Call FncQuery()
	End If   
End Sub

'========================================================================================================
Sub txtFromGlDt_KeyPress(KeyAscii)
	If KeyAscii = 13 Then 
	   Call FncQuery()
	End If   
End Sub

'==========================================================================================
Sub txtAmtFr_Keypress(KeyAscii)
    On Error Resume Next
    If KeyAscii = 27 Then
       Call CancelClick()
    Elseif KeyAscii = 13 Then
       Call fncQuery()
    End if
End Sub
'==========================================================================================
Sub txtAmtTo_Keypress(KeyAscii)
    On Error Resume Next
    If KeyAscii = 27 Then
       Call CancelClick()
    Elseif KeyAscii = 13 Then
       Call fncQuery()
    End if
End Sub




'========================================================================================================
' Name : FncQuery
' Desc : This function is called from MainQuery in Common.vbs
'========================================================================================================
Function FncQuery()
    FncQuery = False                                            
    
    Err.Clear                                                   
    
    Call InitVariables
        
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then									'⊙: This function check indispensable field
       Exit Function
    End If
    
    If CompareDateByFormat(frm1.txtFromGlDt.text,frm1.txtToGlDt.text,frm1.txtFromGlDt.Alt,frm1.txtToGlDt.Alt, _
                        "970025",frm1.txtFromGlDt.UserDefinedFormat,parent.gComDateType,True) = False Then			
		Exit Function
    End If
	If frm1.txtAmtTo.text <> "" Then
		If UNICDbl(frm1.txtAmtTo.text) < UNICDbl(frm1.txtAmtFr.text) Then
			Call DisplayMsgBox("970023","X",frm1.txtAmtTo.Alt,frm1.txtAmtFr.Alt)                         '☜ : 숫자영 
			Exit Function
		End If
	End If
    
	Call ggoOper.ClearField(Document, "2")

    If frm1.txtBizArea.value = "" Then
		frm1.txtBizAreaNm.value = ""
    End If
    
    If frm1.txtCOST_CENTER_CD.value = "" Then
		frm1.txtCOST_CENTER_NM.value = ""
    End If
    
    If frm1.txtdeptcd.value = "" Then
		frm1.txtdeptnm.value = ""
    End If
    
    'Call txtUsr_Id_onChange()
    
    If frm1.txtUsr_Id.value = "" Then
		frm1.txtUsr_Id.value = ""
    End If
    
	IF NOT CheckOrgChangeId Then
		  IntRetCD = DisplayMsgBox("800600","X",frm1.txtFromReqDt.alt,"X")            '⊙: Display Message(There is no changed data.)
		Exit Function
	End if
    '-----------------------
    'Query function call area
    '-----------------------
    IF DbQuery	 = False Then															'☜: Query db data
       Exit Function
    End IF
       
    FncQuery = True	

End Function
	
'========================================================================================================
' Name : FncNew
' Desc : This function is called from MainNew in Common.vbs
'========================================================================================================
Function FncNew()
    Dim IntRetCD 
    
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncNew = False																  '☜: Processing is NG
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'In Multi, You need not to implement this area
    
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    If Err.number = 0 Then	
       FncNew = True                                                              '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement   

End Function
	
'========================================================================================================
' Name : FncDelete
' Desc : This function is called from MainDelete in Common.vbs
'========================================================================================================
Function FncDelete()
End Function

'========================================================================================================
' Name : FncSave
' Desc : This function is called from MainSave in Common.vbs
'========================================================================================================
Function FncSave()
End Function

'========================================================================================================
' Name : FncCopy
' Desc : This function is called from MainCopy in Common.vbs
' Keep : Make sure to clear primary key area
'========================================================================================================
Function FncCopy()
End Function


'========================================================================================================
' Name : FncCancel
' Desc : This function is called by MainCancel in Common.vbs
'========================================================================================================
Function FncCancel() 
End Function


'========================================================================================================
' Name : FncInsertRow
' Desc : This function is called by MainInsertRow in Common.vbs
'========================================================================================================
Function FncInsertRow(ByVal pvRowCnt)
    Dim IntRetCD
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
        .vspdData.ReDraw = False
        .vspdData.focus
        ggoSpread.Source = .vspdData
        ggoSpread.InsertRow ,imRow
        SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1
        .vspdData.ReDraw = True
    End With
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    If Err.number = 0 Then
       FncInsertRow = True                                                          '☜: Processing is OK
    End If   
	
    Set gActiveElement = document.ActiveElement   
End Function


'========================================================================================================
' Name : FncDeleteRow
' Desc : This function is called by MainDeleteRow in Common.vbs
'========================================================================================================
Function FncDeleteRow()
    Dim lDelRows

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncDeleteRow = False                                                          '☜: Processing is NG

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
    If Err.number = 0 Then	
       FncDeleteRow = True                                                            '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement   

End Function

'========================================================================================================
' Name : FncPrint
' Desc : This function is called by MainPrint in Common.vbs
'========================================================================================================
Function FncPrint()

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncPrint = False                                                              '☜: Processing is NG
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    
    '--------- Developer Coding Part (End) ------------------------------------------------------------
	Call Parent.FncPrint()                                                        '☜: Protect system from crashing

    If Err.number = 0 Then	 
       FncPrint = True                                                            '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement   

End Function

'========================================================================================================
' Name : FncPrev
' Desc : This function is called by MainPrev in Common.vbs
'========================================================================================================
Function FncPrev() 

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncPrev = False                                                               '☜: Processing is NG
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    
    '--------- Developer Coding Part (End) ------------------------------------------------------------

    If Err.number = 0 Then	 
       FncPrev = True                                                             '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement   

End Function

'========================================================================================================
' Name : FncNext
' Desc : This function is called by MainPrev in Common.vbs
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
' Name : FncExcel
' Desc : This function is called by MainExcel in Common.vbs
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
' Name : FncFind
' Desc : This function is called by MainFind in Common.vbs
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
' Function Name : FncSplitColumn
' Function Desc : 
'========================================================================================
Sub FncSplitColumn()
    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   
    
End Sub


'========================================================================================
' Function Name : PopSaveSpreadColumnInf
' Description   : 
'========================================================================================
Sub PopSaveSpreadColumnInf()

    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub

'========================================================================================
' Function Name : PopRestoreSpreadColumnInf
' Description   : 
'========================================================================================
Sub PopRestoreSpreadColumnInf()

    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()      
    'Call InitSpreadComboBox()
	Call ggoSpread.ReOrderingSpreadData()
	Call InitData()
End Sub


'========================================================================================================
' Name : FncExit
' Desc : This function is called by MainExit in Common.vbs
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
'========================================================================================================
'                        5.3 Common Group-3
'========================================================================================================
'========================================================================================================

'========================================================================================================
' Name : DbQuery
' Desc : This sub is called by FncQuery
'========================================================================================================
Function DbQuery()
'MsgBox "a"
	Dim strVal
	
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status
    
    DbQuery = False                                                               '☜: Processing is NG
	
    Call DisableToolBar(Parent.TBC_QUERY)                                                '☜: Disable Query Button Of ToolBar
    Call LayerShowHide(1)                                                         '☜: Show Processing Message
    '--------- Developer Coding Part (Start) ----------------------------------------------------------

    With frm1
	
        strVal = BIZ_PGM_ID
        If lgIntFlgMode = parent.OPMD_CMODE Then   ' This means that it is first search
        
			strVal = strVal & "?txtMode=" & parent.UID_M0001	
			strVal = strVal & "&txtFromGlDt=" & Trim(.txtFromGlDt.text)
			strVal = strVal & "&txtToGlDt=" & Trim(.txtToGlDt.text)						'☜: 
			strVal = strVal & "&txtBizArea=" & UCase(Trim(.txtBizArea.value))
			strVal = strVal & "&txtBizArea1=" & UCase(Trim(.txtBizArea1.value))			
			strVal = strVal & "&txtCOST_CENTER_CD=" & Trim(.txtCOST_CENTER_CD.value)
			strVal = strVal & "&txtdeptcd=" & UCase(Trim(.txtdeptcd.value))				'☆: 조회 조건 데이타 
		    strVal = strVal & "&cboGlInputType=" & Trim(.cboGlInputType.value)
			strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
			strVal = strVal & "&OrgChangeId=" & Trim(.hOrgChangeId.Value)
			strVal = strVal & "&txtDesc=" & Trim(.txtDesc.Value)
			strVal = strVal & "&txtRefNo=" & .txtRefNo.value
			strVal = strVal & "&txtAmtFr=" & .txtAmtFr.text
			strVal = strVal & "&txtAmtTo=" & .txtAmtTo.text
			strVal = strVal & "&txtUsr_Id=" & .txtUsr_Id.value
			strVal = strVal & "&cboConfFg=" & Trim(.cboConfFg.value)
			strVal = strVal & "&lgStrPrevKey="  & lgStrPrevKey        '☜: Next key tag
        Else
            strVal = strVal & "?txtMode=" & parent.UID_M0001	
			strVal = strVal & "&txtFromGlDt=" & Trim(.htxtFromGlDt.value)
			strVal = strVal & "&txtToGlDt=" & Trim(.htxtToGlDt.value)						'☜: 
			strVal = strVal & "&txtBizArea=" & UCase(Trim(.htxtBizArea.value))
			strVal = strVal & "&txtBizArea1=" & UCase(Trim(.htxtBizArea1.value))			
			strVal = strVal & "&txtCOST_CENTER_CD=" & Trim(.htxtCOST_CENTER_CD.value)
			strVal = strVal & "&txtdeptcd=" & UCase(Trim(.htxtdeptcd.value))				'☆: 조회 조건 데이타 
		    strVal = strVal & "&cboGlInputType=" & Trim(.hcboGlInputType.value)		
			strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows				
			strVal = strVal & "&OrgChangeId=" & Trim(.hOrgChangeId.Value)
			strVal = strVal & "&txtDesc=" & Trim(.htxtDesc.Value)
			strVal = strVal & "&txtRefNo=" & .htxtRefNo.value
			strVal = strVal & "&txtAmtFr=" & .htxtAmtFr.value
			strVal = strVal & "&txtAmtTo=" & .htxtAmtTo.value
			strVal = strVal & "&txtUsr_Id=" & .htxtUsr_Id.value
			strVal = strVal & "&cboConfFg=" & Trim(.hcboConfFg.value)
			strVal = strVal & "&lgStrPrevKey="  & lgStrPrevKey        '☜: Next key tag			
        End If   
'MsgBox strVal
    '--------- Developer Coding Part (End) ------------------------------------------------------------
    Call RunMyBizASP(MyBizASP, strVal)                                           '☜:  Run biz logic

    End With
    
    
    
    If Err.number = 0 Then	 
       DbQuery = True                                                             '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement   

End Function

'========================================================================================================
' Name : DbSave
' Desc : This sub is called by FncSave
'========================================================================================================
Function DbSave()
End Function

'========================================================================================================
' Name : DbDelete
' Desc : This sub is called by FncDelete
'========================================================================================================
Function DbDelete()

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    DbDelete = False                                                              '☜: Processing is NG
	'------ Developer Coding part (Start)  -------------------------------------------------------------- 
	'In Multi, You need not to implement this area

	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    If Err.number = 0 Then	 
       DbDelete = True                                                             '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement   

End Function

'========================================================================================================
' Name : DbQueryOk
' Desc : Called by MB Area when query operation is successful
'========================================================================================================
Sub DbQueryOk()
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

	lgBlnFlgChgValue = False
    lgIntFlgMode     = parent.OPMD_UMODE												'⊙: Indicates that current mode is Update mode
    lgSaveRow        = 1
    
    Call InitData()
	CALL vspdData_Click(1, 1)
'MsgBox "end"    
End Sub
	
'========================================================================================================
' Name : DbSaveOk
' Desc : Called by MB Area when save operation is successful
'========================================================================================================
Sub DbSaveOk()
End Sub
	
'========================================================================================================
' Function Name : DbDeleteOk
' Function Desc : Called by MB Area when delete operation is successful
'========================================================================================================
Sub DbDeleteOk()
End Sub

'========================================================================================================
'========================================================================================================
'                        5.4 User-defined Method
'========================================================================================================
'========================================================================================================




'========================================================================================================
'========================================================================================================
'                        5.5 Tag Event
'========================================================================================================

'========================================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_ButtonClicked(Col, Row, ButtonDown)
End Sub


'========================================================================================================
'   Event Name : vspdData_Change
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_Change(Col , Row)
End Sub


'========================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : 컬럼을 클릭할 경우 발생
'========================================================================================================
Sub vspdData_Click(ByVal Col , ByVal Row)

	Call SetPopUpMenuItemInf("00000000001")
    gMouseClickStatus = "SPC"	'Split 상태코드 
    Set gActiveSpdSheet = frm1.vspdData

    If frm1.vspdData.MaxRows <= 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
   	
    
    If Row = 0 Then
        ggoSpread.Source = frm1.vspdData
        If lgSortKey = 1 Then
            ggoSpread.SSSort, lgSortKey
            lgSortKey = 2
        Else
            ggoSpread.SSSort, lgSortKey
            lgSortKey = 1
        End If    
    End If
    
    If Col < 1 Then Exit Sub
	'Call SetSpreadColumnValue("A",frm1.vspdData,Col,Row)
    
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
	If frm1.vspdData.MaxRows > 0 Then
		If frm1.vspdData.ActiveRow = Row Or frm1.vspdData.ActiveRow > 0 Then
		End If
	End If
	
End Sub


'========================================================================================================
'   Event Name : vspdData_GotFocus
'   Event Desc : This event is spread sheet data changed
'========================================================================================================
Sub vspdData_GotFocus()
    ggoSpread.Source = Frm1.vspdData

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub

'========================================================================================================
'   Event Name : vspdData_MouseDown
'   Event Desc : Occurs when the user clicks the left mouse button while the pointer is in a cell
'========================================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If
    
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
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

'========================================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
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
           If DbQuery = False Then
              Call RestoreToolBar()
              Exit Sub
           End if
        End If
    End if
End Sub


'========================================================================================================
Sub txtFromGlDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13 Then Call MainQuery
End Sub

'========================================================================================================
Sub txtToGlGlDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13 Then Call MainQuery
End Sub

'========================================================================================================
Sub txtfromamt_KeyDown(KeyCode, Shift)
	If KeyCode = 13 Then Call MainQuery
End Sub

'========================================================================================================
Sub txttoamt_KeyDown(KeyCode, Shift)
	If KeyCode = 13 Then Call MainQuery
End Sub



Function OpenPopUp(Byval strCode, Byval iWhere)

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim strBizAreaCd

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	Select Case iWhere
		Case 0
			arrParam(0) = "사업장 팝업"				' 팝업 명칭 
			arrParam(1) = "B_BIZ_AREA"					' TABLE 명칭 
			arrParam(2) = strCode						' Code Condition
			arrParam(3) = ""							' Name Cindition
			arrParam(4) = ""
			arrParam(5) = "사업장코드"				' 조건필드의 라벨 명칭 
			arrField(0) = "BIZ_AREA_CD"					' Field명(0)
			arrField(1) = "BIZ_AREA_NM"					' Field명(1)

			arrHeader(0) = "사업장코드"				' Header명(0)
			arrHeader(1) = "사업장명"				' Header명(1)

	End Select

   	If iWhere = 0 Then
		arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=400px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	Else
		arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	End If

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetPopUp(arrRet, iWhere)
	End If
	frm1.txtBizArea.focus
	
End Function

'========================================================================================================
Function SetPopUp(Byval arrRet, Byval iWhere)

	With frm1

		Select Case iWhere
			Case 0
				.txtBizArea.value	= arrRet(0)
				.txtBizAreanm.value = arrRet(1)
		End Select

	End With

End Function



'========================================================================================

Function OpenPopupTempGL()

	Dim arrRet
	Dim arrParam(1)	
	Dim iCalledAspName

	If IsOpenPop = True Then Exit Function
	
	With frm1.vspdData
	   if .maxrows > 0 Then	
		.Row = .ActiveRow
		.Col = C_TEMP_GL_NO

	
		arrParam(0) = Trim(.Text)	'결의전표번호 
		arrParam(1) = ""			'Reference번호 
	   End if	
	End With

'	arrParam(0) = Trim(GetKeyPosVal("A", 1))	'전표번호 
'	arrParam(1) = ""			      
	IsOpenPop = True
    
    iCalledAspName = AskPRAspName("a5130ra1")    
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
End Function



'========================================================================================
Function OpenDeptOrgPopup()
	Dim arrRet
	Dim arrParam(8)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = frm1.txtFromGlDt.text								'  Code Condition
   	arrParam(1) = frm1.txtToGlDt.Text
	arrParam(2) = lgUsrIntCd                            ' 자료권한 Condition  
	arrParam(3) = frm1.txtdeptcd.value
	arrParam(4) = "F"									' 결의일자 상태 Condition  
	

	' 권한관리 추가 
	arrParam(5)		= lgAuthBizAreaCd
	arrParam(6)		= lgInternalCd
	arrParam(7)		= lgSubInternalCd
	arrParam(8)		= lgAuthUsrID

	arrRet = window.showModalDialog("../../comasp/DeptPopupOrg.asp", Array(window.parent,arrParam), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtdeptcd.focus
		Exit Function
	Else
		Call SetDept(arrRet)
		frm1.txtdeptcd.focus
	End If	
End Function

'========================================================================================
Function SetDept(Byval arrRet)
		frm1.hOrgChangeId.value=arrRet(2)
		
		frm1.txtDeptCd.value = arrRet(0)
		frm1.txtDeptNm.value = arrRet(1)		
		frm1.txtFromGlDt.text = arrRet(4)
		frm1.txtToGlDt.text = arrRet(5)
End Function


'========================================================================================

Function OpenPopUp(Byval strCode, Byval iWhere)
Dim arrRet
Dim arrParam(5), arrField(6), arrHeader(6)
Dim arrStrRet				'권한관리 추가   							  

dim strgChangeOrgId

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True
	frm1.hOrgChangeId.value = parent.gChangeOrgId

	Select Case iWhere
		'Case 0
'			arrStrRet =  AutorityMakeSql("DEPT",parent.gChangeOrgId, "","","","")	'권한관리 추가   							  
'			
'			arrParam(0) = "부서코드 팝업"								' 팝업 명칭 
'			arrParam(1) = arrstrRet(0)											'권한관리 추가   							  				
'			arrParam(2) = UCase(Trim(frm1.txtDeptCd.Value))	' Code Condition
'			arrParam(3) = ""							' Name Cindition
'			arrParam(4) = arrstrRet(1)											'권한관리 추가   							  
'			
'			arrParam(5) = "부서 코드"			
'	
 '  			arrField(0) = "DEPT_CD"	     									' Field명(0)
'			arrField(1) = "DEPT_NM"			    								' Field명(1)
'		
'			arrHeader(0) = "부서코드"										' Header명(0)
'			arrHeader(1) = "부서코드명"										' Header명(1)
    
		Case 1,3
			arrParam(0) = "사업장 팝업"						' 팝업 명칭 
			arrParam(1) = "B_BIZ_AREA" 							' TABLE 명칭 
			arrParam(2) = strCode								' Code Condition
			arrParam(3) = ""									' Name Cindition

			' 권한관리 추가 
			If lgAuthBizAreaCd <>  "" Then
				arrParam(4) = " BIZ_AREA_CD=" & FilterVar(lgAuthBizAreaCd, "''", "S")			' Where Condition
			Else
				arrParam(4) = ""
			End If

			arrParam(5) = "사업장코드"							' 조건필드의 라벨 명칭 

			arrField(0) = "BIZ_AREA_CD"							' Field명(0)
			arrField(1) = "BIZ_AREA_NM"							' Field명(1)
    
			arrHeader(0) = "사업장코드"				' Header명(0)
			arrHeader(1) = "사업장명"				' Header명(1)
			
		Case 2
			arrParam(0) = "코스트센타 팝업"						' 팝업 명칭 
			arrParam(1) = "B_COST_CENTER"							' TABLE 명칭 
			arrParam(2) = strCode			       				    ' Code Condition
			arrParam(3) = ""										' Name Cindition

			' 권한관리 추가 
			If lgAuthBizAreaCd <>  "" Then
				arrParam(4) = " BIZ_AREA_CD=" & FilterVar(lgAuthBizAreaCd, "''", "S")			' Where Condition
			Else
				arrParam(4) = ""
			End If

			arrParam(5) = "코스트센타"			
	
		    arrField(0) = "COST_CD"									' Field명(0)
			arrField(1) = "COST_NM"									' Field명(1)
    
			arrHeader(0) = "코스트센타코드"					' Header명(0)
			arrHeader(1) = "코스트센타명"						' Header명(1)	
			
		Case 4
			If UCase(frm1.txtUsr_ID.className) = UCase(Parent.UCN_PROTECTED) Then Exit Function
			arrParam(0) = "작성자 팝업"						' 팝업 명칭 
			arrParam(1) = "A_TEMP_GL A, Z_USR_MAST_REC B"							' TABLE 명칭 
			arrParam(2) = strCode			       				    ' Code Condition
			arrParam(3) = ""										' Name Cindition
'hanc::정상적으로 나오지 않음			arrParam(4) = "A.INSRT_USER_ID*=B.USR_ID"										' Where Condition
			arrParam(4) = "A.INSRT_USER_ID=B.USR_ID"										' Where Condition
			arrParam(5) = "작성자"			
	
		    arrField(0) = "A.INSRT_USER_ID"									' Field명(0)
			arrField(1) = "B.USR_NM"									' Field명(1)
    
			arrHeader(0) = "작성자"					' Header명(0)
			arrHeader(1) = "작성자명"						' Header명(1)	
			
			

		
	End Select
    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Call EscPopup(iWhere)
		Exit Function
	Else
		Call SetPopUp(arrRet, iWhere)
	End If	

End Function


'========================================================================================
Function SetPopUp(Byval arrRet, Byval iWhere)
	With frm1
		Select Case iWhere
'			Case 0	     'DEPT
'				.txtdeptcd.value		= UCase(Trim(arrRet(0)))
'				.txtdeptNm.value		= arrRet(1)
'				
'				.txtdeptcd.focus
			Case 1		' Biz area
				.txtBizArea.value		= UCase(Trim(arrRet(0)))
				.txtBizAreaNm.value		= arrRet(1)
				
				.txtBizArea.focus
			Case 2
				.txtCOST_CENTER_CD.value = arrRet(0)
				.txtCOST_CENTER_NM.value = arrRet(1)
				
				.txtCOST_CENTER_CD.focus
			Case 3		' Biz area
				.txtBizArea1.value		= UCase(Trim(arrRet(0)))
				.txtBizAreaNm1.value		= arrRet(1)
				
				.txtBizArea1.focus	
			Case 4		' Biz area
				.txtUsr_ID.value		= UCase(Trim(arrRet(0)))
				.txtUsr_NM.value		= arrRet(1)
				
				.txtUsr_ID.focus					
		End Select

		'lgBlnFlgChgValue = True
	End With
End Function
'=======================================================================================================
Function EscPopup(Byval iWhere)
	With frm1
		Select Case iWhere
			Case 1  
				.txtBizArea.focus
			Case 2 
				.txtCOST_CENTER_CD.focus
		End Select    
	End With
End Function

'========================================================================================================
Sub fpdtFromEnterDt_DblClick(Button)
	If Button = 1 Then
       frm1.fpdtFromEnterDt.Action = 7
       Call SetFocusToDocument("M")	
       frm1.fpdtFromEnterDt.Focus
	End If
End Sub
'========================================================================================================
Sub fpdtToEnterDt_DblClick(Button)
	If Button = 1 Then
       frm1.fpdtToEnterDt.Action = 7
       Call SetFocusToDocument("M")	
       frm1.fpdtToEnterDt.Focus
	End If
End Sub
'========================================================================================================
Sub fpdtFromEnterDt_KeyPress(KeyAscii)
	If KeyAscii = 13 Then 
	   Call MainQuery()
	End If   
End Sub

'========================================================================================================
Sub fpdtToEnterDt_KeyPress(KeyAscii)
	If KeyAscii = 13 Then 
	   Call MainQuery()
	End If   
End Sub

'========================================================================================================
Sub txtUsr_Id_onChange()
	
	If Trim(frm1.txtUsr_Id.value) <> "" Then
		Call CommonQueryRs("USR_NM", "Z_USR_MAST_REC", "USR_ID = " & Filtervar(Trim(frm1.txtUsr_Id.value), "''", "S"), lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
		frm1.txtUsr_NM.value = Replace(lgF0, chr(11), "")
	Else
		frm1.txtUsr_Id.value = ""
		frm1.txtUsr_NM.value = ""
	End If
	
End Sub


'==========================================================================================
Function CheckOrgChangeId()
	Dim strSelect, strFrom, strWhere 	
	Dim IntRetCD 
	Dim ii, jj
	Dim arrVal1, arrVal2
 
	CheckOrgChangeId = True
 
	With frm1
	
		If LTrim(RTrim(.txtDeptCd.value)) <> "" Then
			'----------------------------------------------------------------------------------------
			strSelect = "Distinct ORG_CHANGE_ID"
			strFrom =  " B_ACCT_DEPT "
			strWhere = " ORG_CHANGE_DT >= "
			strWhere = strWhere & " (select max(org_change_dt) from B_ACCT_DEPT where org_change_dt<= " & FilterVar(UNIConvDateToYYYYMMDD(.txtFromGlDt.Text, gDateFormat,Parent.gServerDateType), "''", "S") & ")"
			strWhere = strWhere & " AND ORG_CHANGE_DT <= " 
			strWhere = strWhere & " (select max(org_change_dt) from B_ACCT_DEPT where org_change_dt<= " & FilterVar(UNIConvDateToYYYYMMDD(.txtToGlDt.Text, gDateFormat,Parent.gServerDateType), "''", "S") & ") "
			strWhere = strWhere & " AND ORG_CHANGE_ID =  " & FilterVar(.hOrgChangeId.value , "''", "S") & ""
			strWhere =	strWhere & " AND dept_Cd = " & FilterVar(LTrim(RTrim(frm1.txtDeptCd.value)), "''", "S")

			IntRetCD= CommonQueryRs(strSelect,strFrom,strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
					
			If IntRetCD = False  OR Trim(Replace(lgF0,Chr(11),"")) <> Trim(.hOrgChangeId.value) Then
					.txtDeptCd.value = ""
					.txtDeptNm.value = ""
					.hOrgChangeId.value = ""
					.txtDeptCd.focus
					CheckOrgChangeId = False
			End if
		End If
	End With

End Function



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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>결의전표상세조회</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>					
					<TD WIDTH=*>&nbsp;</TD>
					<TD WIDTH="*" align=right><A HREF="VBSCRIPT:OpenPopupTempGL()">결의전표</A>&nbsp;</td>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>	
	<TR HEIGHT=*>
		<TD WIDTH=100% CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>

				<TR HEIGHT=23 WIDTH=100%>
					<TD>
						<FIELDSET CLASS="CLSFLD">
							<TABLE  <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS=TD5 NOWRAP>결의일자</TD>	                                                  
						            <TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name="txtFromGlDt" CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="시작일자" tag="12" id=fpDateTime1></OBJECT>');</SCRIPT>&nbsp;~&nbsp;
						                                 <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name="txtToGlDt" CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="종료일자" tag="12" id=fpDateTime2></OBJECT>');</SCRIPT></TD>								
									<TD CLASS=TD5 NOWRAP>금액</TD>
									<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name="txtAmtFr" style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 100px" title=FPDOUBLESINGLE tag="11XXXX" ALT="시작금액" id=OBJECT1></OBJECT>');</SCRIPT>&nbsp;~&nbsp;
										 <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name="txtAmtTo" style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 100px" title=FPDOUBLESINGLE tag="11XXXX" ALT="종료금액" id=OBJECT2></OBJECT>');</SCRIPT>
										 </TD>
								</TR>
								 <TR>
									<TD CLASS=TD5 NOWRAP>작성자</TD>
									<TD CLASS=TD6 NOWRAP> <INPUT NAME="txtUsr_ID" MAXLENGTH="12" SIZE=12 ALT ="작성자" tag="14XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btn" ALIGN=top TYPE="BUTTON" ONCLICK="Call OpenPopup(frm1.txtUsr_ID.value, 4)">
														  <INPUT NAME="txtUsr_NM" MAXLENGTH="20" SIZE=24 STYLE="TEXT-ALIGN:left" ALT ="작성자명" tag="14"></TD>
									<TD CLASS=TD5 NOWRAP>참조번호</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtRefNo" ALT="참조번호" MAXLENGTH="30" SIZE="20" tag="11XXXU" ></TD>			
								 </TR>
								<TR>
									<TD CLASS="TD5"NOWRAP>승인상태</TD>
									<TD CLASS="TD6"NOWRAP><SELECT NAME="cboConfFg" tag="1N" STYLE="WIDTH:82px:" Alt="승인상태"><OPTION VALUE="" selected></OPTION></SELECT></TD>
									<TD CLASS=TD5 NOWRAP>비고</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtDesc" ALT="비고" MAXLENGTH="128" SIZE="35" tag="11" ></TD>
								</TR>

								<TR style="display:none">
									<TD CLASS=TD5 NOWRAP></TD>										
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtBizArea1"   ALT="사업장"   Size="12" MAXLENGTH="10" STYLE="TEXT-ALIGN: left" tag   ="1NXXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btn" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtBizArea1.Value, 3)">
														 <INPUT NAME="txtBizAreaNm1" ALT="사업장명" Size="24" MAXLENGTH="20" STYLE="TEXT-ALIGN: left" tag   ="14N"></TD>
								</TR>
								<TD CLASS=TD5 NOWRAP></TD>										
								<TD CLASS=TD6 NOWRAP></TD>										
								<TR style="display:none">
									<TD CLASS=TD5 NOWRAP>코스트센타</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtCOST_CENTER_CD" MAXLENGTH="10" SIZE=12 ALT ="코스트센타 코드" tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btn" ALIGN=top TYPE="BUTTON" ONCLICK="Call OpenPopup(frm1.txtCOST_CENTER_CD.value, 2)">
														 <INPUT NAME="txtCOST_CENTER_NM" MAXLENGTH="20" SIZE=24 STYLE="TEXT-ALIGN:left" ALT ="코스트센타명" tag="14"></TD>
									<TD CLASS=TD5 NOWRAP>전표입력경로</TD>
									<TD CLASS=TD6 NOWRAP><SELECT NAME="cboGlInputType" tag="1N"STYLE="WIDTH:82px:"><OPTION VALUE="" selected></OPTION></SELECT></TD>
								</TR>
								 <TR style="display:none">
									<TD CLASS=TD5 NOWRAP>사업장</TD>										
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtBizArea"   ALT="사업장"   Size="12" MAXLENGTH="10" STYLE="TEXT-ALIGN: left" tag   ="1NXXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btn" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtBizArea.Value, 1)">
														 <INPUT NAME="txtBizAreaNm" ALT="사업장명" Size="24" MAXLENGTH="20" STYLE="TEXT-ALIGN: left" tag   ="14N">&nbsp;~</TD>
						            <TD CLASS=TD5 NOWRAP>부서코드</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtdeptcd" ALT="부서코드" Size="12" MAXLENGTH="10" STYLE="TEXT-ALIGN: left" tag="1NXXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btn" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenDeptOrgPopup()">
														 <INPUT NAME="txtdeptnm" ALT="부서명"   Size="24" MAXLENGTH="20" STYLE="TEXT-ALIGN: left" tag="14N"></TD>
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
							<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD HEIGHT="100%" WIDTH="100%" COLSPAN=6>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=vaSpread> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
								</TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>차변</TD>
								<TD class=TD5 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name="txtNDrAmt" style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 150px" title=FPDOUBLESINGLE ALT="차변" tag="24X2" id=OBJECT22></OBJECT>');</SCRIPT></TD>								
								<TD CLASS=TD5 NOWRAP>대변</TD>
								<TD class=TD5 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name="txtNCrAmt" style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 150px" title=FPDOUBLESINGLE ALT="대변" tag="24X2" id=OBJECT22></OBJECT>');</SCRIPT></TD>								
								<TD CLASS=TD5 NOWRAP>차이</TD>
								<TD class=TD5 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name="txtNSumAmt" style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 150px" title=FPDOUBLESINGLE ALT="차이" tag="24X2" id=OBJECT22></OBJECT>');</SCRIPT></TD>								
							</TR>							
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR >
		<TD <%=HEIGHT_TYPE_01%>></TD>
	</TR>
	
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" TABINDEX="-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="htxtFromGlDt" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="htxtToGlDt" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="htxtBizArea" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="htxtBizArea1" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="htxtCOST_CENTER_CD" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="htxtdeptcd" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hcboGlInputType" tag="24" TABINDEX="-1"> 
<INPUT TYPE=HIDDEN NAME="hOrgChangeId" tag="14" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="htxtglno" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="htxtDesc" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="htxtRefNo" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="htxtAmtFr" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="htxtUsr_Id" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="htxtAmtTo" tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="hcboConfFg"        tag="24" TABINDEX="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

