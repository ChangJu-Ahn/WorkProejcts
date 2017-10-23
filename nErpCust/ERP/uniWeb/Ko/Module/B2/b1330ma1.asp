<%@ LANGUAGE="VBSCRIPT" %>
<!--======================================================================================================
'*  1. Module Name          : Accounting
'*  2. Function Name        : Accounting Basic - Card Company Info.
'*  3. Program ID           : B1330MA1
'*  4. Program Name         : 카드사정보등록 
'*  5. Program Desc         : Register of Card Company
'*  6. Component List       :
'*  7. Modified date(First) : 2002/09/26
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Oh, Soo Min
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
=======================================================================================================-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>


<!-- #Include file="../../inc/IncSvrCcm.inc"  -->
<!-- #Include file="../../inc/IncSvrHTML.inc"  -->	

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliDBAgentA.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliDBAgentVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/Cookie.vbs"></SCRIPT>
<Script Language="VBScript">
Option Explicit                                                        '☜: Turn on the Option Explicit option.
	

'========================================================================================================

Const BIZ_PGM_ID      = "B1330MB1.asp"						           '☆: Biz Logic ASP Name

'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	

'========================================================================================================
Dim IsOpenPop

'========================================================================================================
Sub InitVariables()
	lgIntFlgMode      = parent.OPMD_CMODE						        '⊙: Indicates that current mode is Create mode
	lgBlnFlgChgValue  = False								    '⊙: Indicates that no value changed
    lgStrPrevKey      = ""                                      '⊙: initializes Previous Key
    lgSortKey         = 1                                       '⊙: initializes sort direction
End Sub

'========================================================================================================
Sub SetDefaultVal()
    Dim strYear,strMonth,strDay                                                  '⊙: User-defined variables for date

End Sub
	
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "*","NOCOOKIE","MA") %>
End Sub

'========================================================================================================
Sub CookiePage(ByVal Kubun)
End Sub

'========================================================================================================
Sub MakeKeyStream(ByVal pOpt)
   
   '------ Developer Coding part (Start ) -------------------------------------------------------------- 
   Select Case pOpt
       Case "Q"
			  lgKeyStream = Frm1.txtCardCoCdQ.Value & parent.gColSep       'You Must append one character(parent.gColSep)
       Case "D"
			  lgKeyStream = Frm1.txtCardCoCdQ.Value & parent.gColSep       'You Must append one character(parent.gColSep)
       Case "N"
			  lgKeyStream = Frm1.txtCardCoCdQ.Value & parent.gColSep       'You Must append one character(parent.gColSep)
       Case "P"
			  lgKeyStream = Frm1.txtCardCoCdQ.Value & parent.gColSep       'You Must append one character(parent.gColSep)
   End Select
End Sub        
	
'========================================================================================================
Sub InitComboBox()
    
End Sub

'========================================================================================================
Sub Form_Load()
    Err.Clear
	Call LoadInfTB19029
									'Tag(From "6")   Integeral   Decimal
    Call AppendNumberPlace("6"             ,"8"        ,"2")
	Call AppendNumberPlace("7", "3", "0")
									'Tag(From "0")   Minimal     Maximal
	Call AppendNumberRange("0"             ,"-12x34"   ,"13x440")

	'Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
	'Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
	
	Call InitVariables
    Call SetDefaultVal()
	frm1.txtCardCoCdQ.focus
	Call SetToolbar("1110100000001111")                                              '☆: Developer must customize
	Call ggoOper.LockField(Document, "N")											 '⊙: Lock Field
    Call InitComboBox
'	Call CookiePage (0)   
End Sub	
'========================================================================================
Sub Form_QueryUnload(Cancel, UnloadMode)
End Sub

'========================================================================================
Sub ChkRcptCard_OnClick()
	lgBlnFlgChgValue = True
End Sub

'========================================================================================
Sub ChkPayCard_OnClick()
	lgBlnFlgChgValue = True
End Sub

'========================================================================================
Function OpenPopUp(Byval strCode, Byval iWhere)
	Dim iCalledAspName
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function	

	Select Case iWhere
		Case 0		'캬드사 
			arrParam(0) = "카드사 팝업"				' 팝업 명칭 
			arrParam(1) = "B_CARD_CO"					' TABLE 명칭 
			arrParam(2) = strCode
			arrParam(3) = ""
			arrParam(4) = ""							' Code Condition
			arrParam(5) = "카드사"

			arrField(0) = "CARD_CO_CD"					' Field명(0)
			arrField(1) = "CARD_CO_NM"					' Field명(1)

			arrHeader(0) = "카드사코드"				' Header명(0)
			arrHeader(1) = "카드사명"					' Header명(1)  
		Case 1		'계좌번호		
			If frm1.txtBankAcctNo.className = parent.UCN_PROTECTED Then Exit Function

			arrParam(0) = frm1.txtBankAcctNo.Alt							' 팝업 명칭 
			arrParam(1) = "B_BANK A, B_BANK_ACCT B, F_DPST C"				' TABLE 명칭 
			arrParam(2) = strCode											' Code Condition
			arrParam(3) = ""												' Name Cindition
			arrParam(4) = "A.BANK_CD = B.BANK_CD "							' Where Condition'
			arrParam(4) = arrParam(4) & "AND B.BANK_ACCT_NO = C.BANK_ACCT_NO "
			arrParam(4) = arrParam(4) & "AND A.BANK_CD = C.BANK_CD "
			arrParam(4) = arrParam(4) & "AND (C.DPST_FG = " & FilterVar("SV", "''", "S") & "  OR  C.DPST_FG = " & FilterVar("ET", "''", "S") & " ) "
			arrParam(4) = arrParam(4) & "AND C.DPST_TYPE IN (" & FilterVar("D1", "''", "S") & " ," & FilterVar("D2", "''", "S") & " ," & FilterVar("D3", "''", "S") & " )	 "

			arrParam(5) = frm1.txtBankAcctNo.Alt			' 조건필드의 라벨 명칭 

			arrField(0) = "B.BANK_ACCT_NO"					' Field명(0)
			arrField(1) = "A.BANK_CD"						' Field명(1)
			arrField(2) = "A.BANK_NM"						' Field명(2)

			arrHeader(0) = "계좌번호"						' Header명(0)
			arrHeader(1) = "은행코드"						' Header명(1)
			arrHeader(2) = "은행명"						' Header명(2)
		Case 2		'은행	
			If frm1.txtBankCd.className = parent.UCN_PROTECTED Then Exit Function

			arrParam(0) = frm1.txtBankCd.Alt										' 팝업 명칭 
			arrParam(1) = "B_BANK A, B_BANK_ACCT B, F_DPST C"						' TABLE 명칭 
			arrParam(2) = strCode													' Code Condition
			arrParam(3) = ""														' Name Cindition
			arrParam(4) = "A.BANK_CD = B.BANK_CD "									' Where Condition
			arrParam(4) = arrParam(4) & "AND B.BANK_ACCT_NO = C.BANK_ACCT_NO "
			arrParam(4) = arrParam(4) & "AND A.BANK_CD = C.BANK_CD "
			arrParam(4) = arrParam(4) & "AND (C.DPST_FG = " & FilterVar("SV", "''", "S") & "  OR  C.DPST_FG = " & FilterVar("ET", "''", "S") & " ) "
			arrParam(4) = arrParam(4) & "AND C.DPST_TYPE IN (" & FilterVar("D1", "''", "S") & " ," & FilterVar("D2", "''", "S") & " ," & FilterVar("D3", "''", "S") & " )	 "

			arrParam(5) = frm1.txtBankCd.Alt			' 조건필드의 라벨 명칭 

			arrField(0) = "A.BANK_CD"					' Field명(0)
			arrField(1) = "A.BANK_NM"					' Field명(1)
			arrField(2) = "B.BANK_ACCT_NO"				' Field명(2)

			arrHeader(0) = "은행코드"					' Header명(0)
			arrHeader(1) = "은행명"						' Header명(1)
			arrHeader(2) = "계좌번호"					' Header명(2)

		Case 3		'우편번호 
			arrParam(0) = strCode
			arrParam(1) = ""
			arrParam(2) = parent.gCountry

		Case Else
			Exit Function
	End Select

	IsOpenPop = True

    If iWhere = 3 Then
		iCalledAspName = AskPRAspName("ZipPopup")
		If Trim(iCalledAspName) = "" Then
			IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "ZipPopup", "X")
			IsOpenPop = False
			Exit Function
		End If
		 arrRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
			"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
    ElseIf iWhere = 0 Then
		arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	Else
		arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=760px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	End If

	IsOpenPop = False

	If arrRet(0) = "" Then
		Select Case iWhere

		Case 0	'카드사 
			frm1.txtCardCoCdQ.focus
		Case 1	'은행 
			frm1.txtBankAcctNo.focus
		Case 2	'계좌번호 
			frm1.txtBankCd.focus
		Case 3	'우편번호 
			frm1.txtZipCd.focus
		End Select

		Exit Function
	Else
		Select Case iWhere

		Case 0	'카드사 
			frm1.txtCardCoCdQ.value = arrRet(0)
			frm1.txtCardCoCdQ.focus
		Case 1	'은행 
			frm1.txtBankAcctNo.focus
			frm1.txtBankAcctNo.value  = arrRet(0)
			frm1.txtBankCd.value	= arrRet(1)
			frm1.txtBankNm.value	= arrRet(2)

			lgBlnFlgChgValue = True
		Case 2	'계좌번호 
			frm1.txtBankCd.focus
			frm1.txtBankCd.value	= arrRet(0)
			frm1.txtBankNm.value	= arrRet(1)	
			frm1.txtBankAcctNo.value  = arrRet(2)

			lgBlnFlgChgValue = True
		Case 3	'우편번호 
			frm1.txtZipCd.focus
			frm1.txtZipCd.value = arrRet(0)
			frm1.txtAddr1.value = arrRet(1)
			lgBlnFlgChgValue = True
		End Select
	End If

End Function

'========================================================================================================
' Name : FncQuery
' Desc : This function is called from MainQuery in Common.vbs
'========================================================================================================
Function FncQuery()
    Dim IntRetCD 
    
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status
  
    If lgBlnFlgChgValue = True Then
       IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO,"x","x")					  '☜: Data is changed.  Do you want to display it? 
       If IntRetCD = vbNo Then
          Exit Function
       End If
    End If
    
    Call ggoOper.ClearField(Document, "2")
     
    If Not chkField(Document, "1") Then
       Exit Function
    End If
    
    Call InitVariables

    If DbQuery("Q") = False Then
       Exit Function
    End If
   

    If Err.number = 0 Then
       FncQuery = True
    End If
    Set gActiveElement = document.ActiveElement
End Function
	
'========================================================================================================
' Name : FncNew
' Desc : This function is called from MainNew in Common.vbs
'========================================================================================================
Function FncNew()
    Dim IntRetCD 
    
    On Error Resume Next
    Err.Clear

    FncNew = False

    If lgBlnFlgChgValue = True Then
       IntRetCD = DisplayMsgBox("900015", parent.VB_YES_NO,"x","x")					  '☜: Data is changed.  Do you want to make it new? 
       If IntRetCD = vbNo Then
          Exit Function
       End If
    End If
    
    Call ggoOper.ClearField(Document, "A")                                        '☜: Clear Contents  Field
    Call ggoOper.LockField(Document , "N")                                        '☜: Lock  Field ("N" means New)

    
    Call SetToolbar("11101000000011")
    Call SetDefaultVal    
    Call InitVariables

	If Err.number = 0 Then
       FncNew = True
    End If
    Set gActiveElement = document.ActiveElement
End Function
	
'========================================================================================================
' Name : FncDelete
' Desc : This function is called from MainDelete in Common.vbs
'========================================================================================================
Function FncDelete()
    Dim IntRetCD
    
    On Error Resume Next 
    Err.Clear

    FncDelete = False

    If lgIntFlgMode <> parent.OPMD_UMODE Then
        Call DisplayMsgBox("900002","x","x","x")                                
        Exit Function
    End If

	Call CommonQueryRs(" COUNT(*) ","B_CARD_CO		A, F_CARD		B","A.CARD_CO_CD = B.CARD_CO_CD " & _														
								" AND B.CARD_CO_CD = " & FilterVar(frm1.txtBankAcctNo.value, "''", "S")  _
								,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

	If lgF0 > 0 & chr(11)  Then
		Call DisplayMsgBox("140801","x","x","x")
		Exit Function
	End If
	

    IntRetCD = DisplayMsgBox("900003", parent.VB_YES_NO,"x","x")                         '☜: Do you want to delete? 
    
	If IntRetCD = vbNo Then
       Exit Function
	End If
    

    If DbDelete = False Then 
       Exit Function
    End If
    
    If Err.number = 0 Then    
       FncDelete = True
    End If
       
    Set gActiveElement = document.ActiveElement       
End Function

'==========================================================================================
'   Function Name : chkSaveValue
'   Function Desc : 저장시 거래처구분에 따라 체크박스 Value Change
'==========================================================================================
Function chkSaveValue()

	If frm1.ChkRcptCard.checked = True and frm1.ChkPayCard.checked = True Then
		frm1.txtRcptCard.value = "Y"
		frm1.txtPayCard.value = "Y"	
	ElseIf frm1.ChkRcptCard.checked = True and frm1.ChkPayCard.checked = False Then
		frm1.txtRcptCard.value = "Y"
		frm1.txtPayCard.value = "N"
	ElseIf frm1.ChkRcptCard.checked = False and frm1.ChkPayCard.checked = True Then
		frm1.txtRcptCard.value = "N"
		frm1.txtPayCard.value = "Y"
	ElseIf frm1.ChkRcptCard.checked = False and frm1.ChkPayCard.checked = False Then
		frm1.txtRcptCard.value = "N"
		frm1.txtPayCard.value = "N"
	End If 
	
End Function

'========================================================================================================
Function FncSave()
    Dim IntRetCD 
    
    On Error Resume Next
    Err.Clear

    FncSave = False                                                              '☜: Processing is NG

    If lgBlnFlgChgValue = False Then 
       IntRetCD = DisplayMsgBox("900001","x","x","x")                            '☜:There is no changed data. 
       Exit Function
    End If

    If Not chkField(Document, "2") Then                                          '☜: Check contents area
       Exit Function
    End If
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	If Trim(frm1.txtBankCd.value) <> "" and Trim(frm1.txtBankAcctNo.value) <> "" Then					'계좌, 은행 모두 입력시 
		Call CommonQueryRs(" A.BANK_ACCT_NO, A.BANK_CD, C.DPST_FG , C.DPST_TYPE ", _
									" B_BANK_ACCT A, B_BANK B,  F_DPST C", _
									" A.BANK_CD = B.BANK_CD " & _
									" AND A.BANK_CD = C.BANK_CD " &_
									" AND B.BANK_CD = C.BANK_CD " &_
									" AND A.BANK_ACCT_NO = C.BANK_ACCT_NO " &_
									" AND A.BANK_CD = " & FilterVar(frm1.txtBankCd.value, "''", "S") & _
									" AND A.BANK_ACCT_NO = " & FilterVar(frm1.txtBankAcctNo.value, "''", "S")  _
									,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

		If Trim(lgF0) = "" Then
			Call DisplayMsgBox("120900","x","x","x")
			Exit Function
		End If		

		If Trim(lgF2) = "DP" Then														'예적금 구분 (DP, SV, ET)
			Call DisplayMsgBox("140802","x","x","x")
			Exit Function
		Else
			If Trim(lgF3) > "D3" Then													'예적금 유형( D1 ~ D6)			
				Call DisplayMsgBox("140803","x","x","x")
				Exit Function
			End If
		End If

	ElseIf Trim(frm1.txtBankCd.value) <> "" and Trim(frm1.txtBankAcctNo.value) = "" Then					'은행코드만 입력시 
		Call CommonQueryRs(" COUNT(*) "," B_BANK ","BANK_CD = " & FilterVar(frm1.txtBankCd.value, "''", "S")  _
									,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
		If lgF0 <= 0  & chr(11)  Then
			Call DisplayMsgBox("120800","x","x","x")
			Exit Function
		End If

	ElseIf Trim(frm1.txtBankCd.value) = "" and Trim(frm1.txtBankAcctNo.value) <> "" Then					'은행계좌만 입력시	
		Call CommonQueryRs(" COUNT(*) "," B_BANK_ACCT ", _
									 " BANK_ACCT_NO = " & FilterVar(frm1.txtBankAcctNo.value, "''", "S")  _
									 ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

		If lgF0 <= 0 & chr(11)  Then
			Call DisplayMsgBox("120900","x","x","x")
			Exit Function
		ElseIf lgF0 > 1 Then  
			Call DisplayMsgBox("120911","x","x","x")
			Exit Function
		End If
	End If
	
	Call chkSaveValue()


    If DbSave = False Then                                                       '☜: Query db data
       Exit Function
    End If
    
    If Err.number = 0 Then
       FncSave = True                                                            '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement

End Function

'========================================================================================================
' Name : FncCopy
' Desc : This function is called from MainCopy in Common.vbs
' Keep : Make sure to clear primary key area
'=======================================================================================================
Function FncCopy()
	Dim IntRetCD

    On Error Resume Next
    Err.Clear

    FncCopy = False

    If lgBlnFlgChgValue = True Then
       IntRetCD = DisplayMsgBox("900017", parent.VB_YES_NO,"x","x")
       If IntRetCD = vbNo Then
          Exit Function
       End If
    End If
    
    
    Call ggoOper.ClearField(Document, "1")
    Call ggoOper.LockField(Document, "N")
    Call CancelRestoreToolBar()
    Call SetToolbar("11101000000011")

    frm1.txtCardCoCd.value = ""
    frm1.txtCardCoNm.value = ""


    lgIntFlgMode = parent.OPMD_CMODE

    If Err.number = 0 Then
       FncCopy = True
    End If
    Set gActiveElement = document.ActiveElement
End Function

'========================================================================================
Function FncCancel() 

    On Error Resume Next
    Err.Clear

    FncCancel = False

	If Err.number = 0 Then
       FncCancel = True
    End If
    Set gActiveElement = document.ActiveElement
End Function

'========================================================================================================
Function FncInsertRow()

    On Error Resume Next
    Err.Clear

    FncInsertRow = False

	If Err.number = 0 Then
       FncInsertRow = True                                                       '☜: Processing is OK
    End If
    Set gActiveElement = document.ActiveElement
End Function

'========================================================================================================
Function FncDeleteRow()

    On Error Resume Next
    Err.Clear

    FncDeleteRow = False

	If Err.number = 0 Then
       FncDeleteRow = True
    End If

    Set gActiveElement = document.ActiveElement
End Function

'========================================================================================================
Function FncPrint()

    On Error Resume Next
    Err.Clear

    FncPrint = False
	Call Parent.FncPrint()

    If Err.number = 0 Then
       FncPrint = True
    End If
    Set gActiveElement = document.ActiveElement
End Function

'========================================================================================================
Function FncPrev() 
    Dim IntRetCD

    On Error Resume Next
    Err.Clear

    FncPrev = False

    If lgIntFlgMode <> parent.OPMD_UMODE Then                                           '☜: Please do Display first. 
       Call DisplayMsgBox("900002","x","x","x")
       Exit Function
    End If
	
	If lgBlnFlgChgValue = True Then
       IntRetCD = DisplayMsgBox("900017", parent.VB_YES_NO,"x","x")					 '☜: Will you destory previous data
       If IntRetCD = vbNo Then
          Exit Function
       End If
	End If
	

    Call ggoOper.ClearField(Document, "2")

    Call InitVariables

    If DbQuery("P") = False Then
       Exit Function
    End If

    If Err.number = 0 Then
       FncPrev = True
    End If
    Set gActiveElement = document.ActiveElement
End Function

'========================================================================================================
' Name : FncNext
' Desc : This function is called by MainPrev in Common.vbs
'========================================================================================================
Function FncNext() 
    Dim IntRetCD

    On Error Resume Next
    Err.Clear

    FncNext = False

    If lgIntFlgMode <> parent.OPMD_UMODE Then
        Call DisplayMsgBox("900002","x","x","x")
        Exit Function
    End If
	
	If lgBlnFlgChgValue = True Then
       IntRetCD = DisplayMsgBox("900017", parent.VB_YES_NO,"x","x")
       If IntRetCD = vbNo Then
          Exit Function
       End If
	End If
	

    Call ggoOper.ClearField(Document, "2")

    Call InitVariables

    If DbQuery("N") = False Then
       Exit Function
    End If

    If Err.number = 0 Then
       FncNext = True
    End If
    Set gActiveElement = document.ActiveElement
End Function

'========================================================================================================
' Name : FncExcel
' Desc : This function is called by MainExcel in Common.vbs
'========================================================================================================
Function FncExcel() 

    On Error Resume Next
    Err.Clear

    FncExcel = False

	Call Parent.FncExport(parent.C_SINGLE)

    If Err.number = 0 Then
       FncExcel = True
    End If
    Set gActiveElement = document.ActiveElement
End Function

'========================================================================================================
Function FncFind() 

    On Error Resume Next
    Err.Clear

    FncFind = False
	Call Parent.FncFind(parent.C_SINGLE, True)
    If Err.number = 0 Then
       FncFind = True
    End If
    Set gActiveElement = document.ActiveElement
End Function


'========================================================================================================
Function FncExit()
	Dim IntRetCD

    On Error Resume Next
    Err.Clear

    FncExit = False

	If lgBlnFlgChgValue = True Then
       IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO,"x","x")
       If IntRetCD = vbNo Then
          Exit Function
       End If
    End If

    If Err.number = 0 Then
       FncExit = True
    End If

    Set gActiveElement = document.ActiveElement

End Function


'========================================================================================================
Function DbQuery(pDirect)

	Dim strVal
	
    On Error Resume Next
    Err.Clear

    DbQuery = False

    Call DisableToolBar(parent.TBC_QUERY)
    Call LayerShowHide(1)
    Call MakeKeyStream(pDirect)

    strVal = BIZ_PGM_ID & "?txtMode="          & parent.UID_M0001
    strVal = strVal     & "&txtPrevNext="      & pDirect
    strVal = strVal     & "&txtKeyStream="     & lgKeyStream
    Call RunMyBizASP(MyBizASP, strVal)

    If Err.number = 0 Then
       DbQuery = True
    End If
    Set gActiveElement = document.ActiveElement   
End Function

'========================================================================================================
Function DbSave()
	Dim strVal

    On Error Resume Next
    Err.Clear

	DbSave = False

	Call LayerShowHide(1)
		
	With Frm1
		.txtMode.value        = parent.UID_M0002
		.txtFlgMode.value     = lgIntFlgMode
	End With
	Call ExecMyBizASP(frm1, BIZ_PGM_ID)

	If Err.number = 0 Then
       DbSave  = True
    End If

    Set gActiveElement = document.ActiveElement
End Function
'========================================================================================================
' Name : DbDelete
' Desc : This Sub is called by FncDelete
'========================================================================================================
Function DbDelete()
	Dim strVal
		
    On Error Resume Next
    Err.Clear

	DbDelete = False

	Call LayerShowHide(1)

    Call MakeKeyStream("D")
		
    strVal = BIZ_PGM_ID & "?txtMode="          & parent.UID_M0003
    strVal = strVal     & "&txtKeyStream="     & lgKeyStream 

	Call RunMyBizASP(MyBizASP, strVal)
	
	If Err.number = 0 Then
       DbDelete = True
    End If
       
End Function
'========================================================================================================
' Name : DbQueryOk
' Desc : Called by MB Area when query operation is successful
'========================================================================================================
Sub DbQueryOk()

    On Error Resume Next
    Err.Clear

	lgIntFlgMode      = parent.OPMD_UMODE

    Call CancelRestoreToolBar()
	Call SetToolbar("11111000110111")

    Call ggoOper.LockField(Document, "Q")  

    Set gActiveElement = document.ActiveElement

End Sub
	
'========================================================================================================
' Name : DbSaveOk
' Desc : Called by MB Area when save operation is successful
'========================================================================================================
Sub DbSaveOk()

    On Error Resume Next
    Err.Clear

    Call InitVariables

    Frm1.txtCardCoCdQ.Value     = Frm1.txtCardCoCd.Value

    Call MainQuery()

    Set gActiveElement = document.ActiveElement

End Sub

'========================================================================================================
Sub DbDeleteOk()
    On Error Resume Next
    Err.Clear
	Call MainNew()	
    Set gActiveElement = document.ActiveElement
End Sub

'========================================================================================================


</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->
</HEAD>

<!--
'########################################################################################################
'#						6. TAG 																		#
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>카드사정보등록</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
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
									<TD CLASS=TD5 NOWRAP>카드사</TD>
                                    <TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtCardCoCdQ"  SIZE=20 MAXLENGTH=10   TAG="12XXXU" ALT="카드사코드"><IMG SRC="../../../CShared/image/btnPopup.gif" ALIGN=top TYPE="BUTTON" OnClick="vbscript:Call OpenPopUp(frm1.txtCardCoCdQ.value, 0)"></TD>
                                    <TD CLASS=TDT NOWRAP>&nbsp;</TD>
                                    <TD CLASS=TD6 NOWRAP>&nbsp;</TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=100% VALIGN=TOP>
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD CLASS=TD5 NOWRAP>카드사코드</TD>
                                <TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtCardCoCd"  ALT="카드사코드" SIZE=20 MAXLENGTH=10 TAG="23XXXU"></TD>
                                <TD CLASS=TD5 NOWRAP>카드사명</TD>
                                <TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtCardCoNm"  ALT="카드사명" SIZE=30 MAXLENGTH=30 TAG="22X"></TD>
							</TR>
							<TR>
									<TD CLASS="TD5" NOWRAP>주거래계좌번호</TD>
									<TD CLASS="TD6" NOWRAP><INPUT NAME="txtBankAcctNo" ALT="주거래계좌번호" SIZE="18" MAXLENGTH="30"  tag="21XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBankAcct" align=top TYPE="BUTTON"ONCLICK="vbscript:Call OpenPopUp(frm1.txtBankAcctNo.value, 1)"></TD>
									<TD CLASS="TD5" NOWRAP>주거래은행</TD>
									<TD CLASS="TD6" NOWRAP><INPUT NAME="txtBankCd" ALT="주거래은행" SIZE="10" MAXLENGTH="10"  tag="21XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBankCd" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtBankCd.value, 2)">
																		  <INPUT NAME="txtBankNm" ALT="주거래은행명" STYLE="TEXT-ALIGN: Left" tag="24X"></TD>					
							</TR>
							<TR>
									<TD CLASS=TD5 NOWRAP>수취구매카드</TD>
									<TD CLASS=TD6 NOWRAP><INPUT type="checkbox" CLASS="STYLE CHECK"  NAME=ChkRcptCard ID=ChkRcptCard  value="N" tag="1" ><LABEL FOR=ChkRcptCard>적용</LABEL>&nbsp;
									<TD CLASS=TD5 NOWRAP>지불구매카드</TD>
									<TD CLASS=TD6 NOWRAP><INPUT type="checkbox" CLASS="STYLE CHECK"  NAME=ChkPayCard  ID=ChkPayCard value="N" tag="1" ><LABEL FOR=ChkPayCard>적용</LABEL>&nbsp;</TR>	
							<TR>
                                <TD CLASS=TD5 NOWRAP>우편번호</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtZipCd" ALT="우편번호" MAXLENGTH="12" SIZE="11" STYLE="TEXT-ALIGN:left" tag ="2"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCalType" align=top TYPE="BUTTON" ONCLICK="VBScript:Call OpenPopUp(frm1.txtZipCd.value, 3)"></TD>
								<TD CLASS=TD5 NOWRAP>&nbsp;</TD>
                                <TD CLASS=TD6 NOWRAP>&nbsp;</TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>주소1</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtAddr1" ALT="주소1" MAXLENGTH="50" SIZE=35 STYLE="TEXT-ALIGN:left" tag  ="21"></TD>
								<TD CLASS=TD5 NOWRAP>&nbsp;</TD>
                                <TD CLASS=TD6 NOWRAP>&nbsp;</TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>주소2</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtAddr2" ALT="주소2" MAXLENGTH="50" SIZE=35 STYLE="TEXT-ALIGN:left" tag  ="21"></TD>
								<TD CLASS=TD5 NOWRAP>&nbsp;</TD>
                                <TD CLASS=TD6 NOWRAP>&nbsp;</TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>주소3</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtAddr3" ALT="주소3" MAXLENGTH="50" SIZE=35 STYLE="TEXT-ALIGN:left" tag  ="21"></TD>
								<TD CLASS=TD5 NOWRAP>&nbsp;</TD>
                                <TD CLASS=TD6 NOWRAP>&nbsp;</TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>전화번호1</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtTelNo1" ALT="전화번호" MAXLENGTH="20" STYLE="TEXT-ALIGN:left" tag  ="2"></TD>
								<TD CLASS=TD5 NOWRAP>전화번호2</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtTelNo2" ALT="전화번호" MAXLENGTH="20" STYLE="TEXT-ALIGN:left" tag  ="2"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>FAX번호</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtFaxNo" ALT="FAX번호" MAXLENGTH="20" STYLE="TEXT-ALIGN:left" tag ="2" ></TD>
								<TD CLASS=TD5 NOWRAP>&nbsp;</TD>
                                <TD CLASS=TD6 NOWRAP>&nbsp;</TD>
							</TR>	
							<TR>
								<TD CLASS=TD5 NOWRAP>Homepage URL</TD>
								<TD CLASS=TD6 NOWRAP COLSPAN=3><INPUT NAME="txtUrl" ALT="Homepage URL" MAXLENGTH="50" SIZE="95" STYLE="TEXT-ALIGN:left"  tag="21X" ></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>비고</TD>
								<TD CLASS=TD6 NOWRAP COLSPAN=3><INPUT NAME="txtDesc" ALT="비고" MAXLENGTH="128" SIZE="95" STYLE="TEXT-ALIGN:left"  tag="21X"  ></TD>
							</TR>
                             <% Call SubFillRemBodyTd5656(2) %>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR >
		<TD <%=HEIGHT_TYPE_01%>></TD>
	</TR>
	<TR >
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC= "../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode"        TAG="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtKeyStream"   TAG="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtFlgMode"     TAG="X4" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtPrevNext"    TAG="X4" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="htxtSchoolCdD"  TAG="X4" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtRcptCard" tag="24">
<INPUT TYPE=HIDDEN NAME="txtPayCard" tag="24">

</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm"></IFRAME>
</DIV>
</BODY>
</HTML>

