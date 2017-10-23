<%@ LANGUAGE="VBSCRIPT" %>
<%Option Explicit%>
<!--'======================================================================================================
'*  1. Module Name          : Accounting
'*  2. Function Name        :
'*  3. Program ID           : A5958MA1
'*  4. Program Name         : 유가증권 등록 
'*  5. Program Desc         : 유가증권 등록 
'*  6. Component List       :
'*  7. Modified date(First) : 2002/01/29
'*  8. Modified date(Last)  : 2003/05/30
'*  9. Modifier (First)     : 권기수 
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

<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMaMain.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMaOperation.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliRdsQuery.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/Cookie.vbs">   </SCRIPT>
<SCRIPT LANGUAGE="javascript"	SRC="../../inc/TabScript.js"></SCRIPT>

<Script Language="VBScript">
Option Explicit                                                        '☜: indicates that All variables must be declared in advance

'========================================================================================================

Const BIZ_PGM_ID    = "a5958mb1.asp"
Const BIZ_PGM_JUMP_ID   = "a5959ma1"									'사업부별 손익비교(컴퍼니 메뉴에 등록된 명)
Const COOKIE_SPLIT  =  4877	                                                        'Cookie Split String

'========================================================================================================
'''출금상세내역 

Const TAB1 = 1																		'☜: Tab의 위치 
Const TAB2 = 2

Dim C_Seq
Dim C_RcptType
Dim C_RcptTypePopup
Dim C_RcptTypeNm
Dim C_Amt
Dim C_LocAmt
Dim C_NoteNo
Dim C_NoteNoPopup
Dim C_BankAcct
Dim C_BankAcctPopup
Dim C_BankCd
Dim C_BankNm
Dim irow

Const gIsShowLocal = "Y"
<%
Const gIsShowLocal = "Y"
%>

'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->

'========================================================================================================
Dim lgIsOpenPop
Dim IsOpenPop
Dim lgRecordPage
Dim UserPrevNext
Dim gSelframeFlg
<%
Dim StartDate
StartDate	= GetSvrDate                                               'Get Server DB Date
%>

'========================================================================================================
Sub InitSpreadPosVariables()
	
 C_Seq				= 1
 C_RcptType			= 2									            'Spread Sheet 의 Columns 인덱스 
 C_RcptTypePopup	= 3
 C_RcptTypeNm		= 4									            'Spread Sheet 의 Columns 인덱스 
 C_Amt				= 5
 C_LocAmt			= 6
 C_NoteNo			= 7
 C_NoteNoPopup		= 8
 C_BankAcct			= 9
 C_BankAcctPopup	= 10
 C_BankCd			= 11
 C_BankNm			= 12
 
End Sub

'========================================================================================================
Sub InitVariables()
	lgIntFlgMode      = parent.OPMD_CMODE						        '⊙: Indicates that current mode is Create mode
	lgBlnFlgChgValue  = False								    '⊙: Indicates that no value changed
	lgIntGrpCount     = 0										'⊙: Initializes Group View Size
    lgStrPrevKey      = ""
    lgStrPrevKeyIndex = 0                                       '⊙: initializes Previous Key
    lgSortKey         = 1                                       '⊙: initializes sort direction

	'------ Developer Coding part (Start ) --------------------------------------------------------------

    lgLngCurRows = 0                                            'initializes Deleted Rows Count
	'------ Developer Coding part (End   ) --------------------------------------------------------------
End Sub

'========================================================================================================
Sub SetDefaultVal()
	'------ Developer Coding part (Start ) --------------------------------------------------------------
	Call ggoOper.FormatDate(frm1.txtBillDt, parent.gDateFormat, 1)
	Call ggoOper.FormatDate(frm1.txtPubDt, parent.gDateFormat, 1)
	Call ggoOper.FormatDate(frm1.txtExpireDt, parent.gDateFormat, 1)
	Call ggoOper.FormatDate(frm1.txtInDt, parent.gDateFormat, 1)
	Call ggoOper.FormatDate(frm1.txtOutDt, parent.gDateFormat, 1)
    Call ggoOper.SetReqAttr(frm1.txtDept2,"Q")
    Call ggoOper.SetReqAttr(frm1.txtOutDt,"Q")
    frm1.txtBillDt.text	= UniConvDateAToB("<%=StartDate%>",parent.gServerDateFormat,parent.gDateFormat)
    frm1.txtInDt.text	= UniConvDateAToB("<%=StartDate%>",parent.gServerDateFormat,parent.gDateFormat)

    lgBlnFlgChgValue = False
	'------ Developer Coding part (End )   --------------------------------------------------------------
End Sub

Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "*", "NOCOOKIE", "MA") %>
	<% Call LoadBNumericFormatA("I", "*", "NOCOOKIE", "MA") %>
End Sub

'========================================================================================================
' Name : MakeKeyStream
' Desc : This method set focus to pos of err
'========================================================================================================
Sub MakeKeyStream(pOpt)
    Dim txtBillDt, txtPubDt, txtExpireDt, txtInDt, txtOutDt
    Dim strYear, strMonth, strDay

    '------ Developer Coding part (Start ) --------------------------------------------------------------
    Call ExtractDateFrom(frm1.txtBillDt.Text,frm1.txtBillDt.UserDefinedFormat,parent.gComDateType,strYear,strMonth,strDay)
    txtBillDt = strYear  & strMonth  & strDay

    Call ExtractDateFrom(frm1.txtPubDt.Text,frm1.txtPubDt.UserDefinedFormat,parent.gComDateType,strYear,strMonth,strDay)
    txtPubDt = strYear & strMonth & strDay

    Call ExtractDateFrom(frm1.txtExpireDt.Text,frm1.txtExpireDt.UserDefinedFormat,parent.gComDateType,strYear,strMonth,strDay)
    txtExpireDt = strYear & strMonth & strDay

    Call ExtractDateFrom(frm1.txtInDt.Text,frm1.txtInDt.UserDefinedFormat,parent.gComDateType,strYear,strMonth,strDay)
    txtInDt = strYear & strMonth & strDay

    Call ExtractDateFrom(frm1.txtOutDt.Text,frm1.txtOutDt.UserDefinedFormat,parent.gComDateType,strYear,strMonth,strDay)
    txtOutDt = strYear & strMonth & strDay

Select Case pOpt
       Case "S"
					lgKeyStream = Trim(frm1.txtSecuCode1.value) & parent.gColSep
					lgKeyStream = lgKeyStream & Trim(frm1.txtSecuNm1.value) & parent.gColSep
					lgKeyStream = lgKeyStream & Trim(frm1.txtSecuType.value) & parent.gColSep
					lgKeyStream = lgKeyStream & Trim(frm1.txtDept1Area.value) & parent.gColSep
					lgKeyStream = lgKeyStream & Trim(frm1.txtDept1OrgId.value) & parent.gColSep
					lgKeyStream = lgKeyStream & Trim(frm1.txtDept1.value) & parent.gColSep
					lgKeyStream = lgKeyStream & Trim(frm1.txtDept2.value) & parent.gColSep
					lgKeyStream = lgKeyStream & txtPubDt & parent.gColSep
					lgKeyStream = lgKeyStream & txtExpireDt & parent.gColSep
					lgKeyStream = lgKeyStream & txtInDt & parent.gColSep
					lgKeyStream = lgKeyStream & Trim(frm1.txtCust1.value) & parent.gColSep
					lgKeyStream = lgKeyStream & Trim(frm1.txtCust2.value) & parent.gColSep
					lgKeyStream = lgKeyStream & Trim(frm1.txtTradeCur.value) & parent.gColSep
					lgKeyStream = lgKeyStream & Trim(frm1.txtXchRate.text) & parent.gColSep
					lgKeyStream = lgKeyStream & Trim(frm1.txtBuyAmt.text) & parent.gColSep
					If UNICDbl(frm1.txtLocBuyAmt.text) = 0 or Trim(frm1.txtLocBuyAmt.text) = "" Then
					    lgKeyStream = lgKeyStream & UNICDbl(frm1.txtXchRate.text) * UNICDbl(frm1.txtBuyAmt.text) & parent.gColSep
					Else
					    lgKeyStream = lgKeyStream & Trim(frm1.txtLocBuyAmt.text) & parent.gColSep
					End If
					lgKeyStream = lgKeyStream & Trim(frm1.txtCalRate.text) & parent.gColSep
					lgKeyStream = lgKeyStream & Trim(frm1.selCalYn.value) & parent.gColSep
					lgKeyStream = lgKeyStream & Trim(frm1.selEndYn.value) & parent.gColSep
					lgKeyStream = lgKeyStream & Trim(frm1.txtPriceAmt.text) & parent.gColSep
					If UNICDbl(frm1.txtLocPriceAmt.text) = 0 or Trim(frm1.txtLocPriceAmt.text) = "" Then
					    lgKeyStream = lgKeyStream & UNICDbl(frm1.txtXchRate.text) * UNICDbl(frm1.txtPriceAmt.text) & parent.gColSep
					Else
					    lgKeyStream = lgKeyStream & Trim(frm1.txtLocPriceAmt.text) & parent.gColSep
					End If
					lgKeyStream = lgKeyStream & Trim(frm1.txtCnt.text) & parent.gColSep
					lgKeyStream = lgKeyStream & txtOutDt & parent.gColSep
					lgKeyStream = lgKeyStream & Trim(frm1.selComYn.value) & parent.gColSep
					lgKeyStream = lgKeyStream & Trim(frm1.txtRefNo.value) & parent.gColSep
					lgKeyStream = lgKeyStream & txtBillDt & parent.gColSep
       Case "D" 
	
				  lgKeyStream = Trim(frm1.txtSecuCode1.value)        & parent.gColSep
				  lgKeyStream = lgKeyStream & Trim(frm1.txtGlNo.value)        & parent.gColSep
				  lgKeyStream = lgKeyStream & Trim(frm1.txtTGlNo.value)        & parent.gColSep
				  lgKeyStream = lgKeyStream & Trim(txtBillDt)        & parent.gColSep
	End Select    



    '------ Developer Coding part (End   ) --------------------------------------------------------------
End Sub

'========================================================================================================
Sub InitComboBox()
	Dim iCodeArr
	Dim iNameArr
	Dim i, isize
	Dim IntRetCD1

	on error resume next
	
	i = 0
	'------ Developer Coding part (Start ) --------------------------------------------------------------
    Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("F1080", "''", "S") & "  ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    iCodeArr = Split(lgF0,Chr(11))
    iNameArr = Split(lgF1,Chr(11))

    If isArray(iCodeArr) Then
        Do While Not isNull(iCodeArr(i))
            i = i + 1
            If iCodeArr(i) = "" Then
                Exit Do
            End If
        Loop
        isize = i
        frm1.selEndYn.length = isize

        For i = 0 to isize-1
            frm1.selEndYn.options(i).value = iCodeArr(i)
            frm1.selEndYn.options(i).text	= iNameArr(i)
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
			C_Seq				= iCurColumnPos(1)
			C_RcptType			= iCurColumnPos(2)
			C_RcptTypePopup		= iCurColumnPos(3)    
			C_RcptTypeNm		= iCurColumnPos(4)
			C_Amt				= iCurColumnPos(5)
			C_LocAmt			= iCurColumnPos(6)
			C_NoteNo			= iCurColumnPos(7)
			C_NoteNoPopup		= iCurColumnPos(8)
			C_BankAcct			= iCurColumnPos(9)
			C_BankAcctPopup		= iCurColumnPos(10)
			C_BankCd			= iCurColumnPos(11)
			C_BankNm			= iCurColumnPos(12)
			
    End Select 
    
    		
End Sub    
'========================================================================================================
' Name : InitData()
' Desc : Reset Combox
'========================================================================================================
Sub InitData()
	frm1.txtXchRate.text = 1
	frm1.txtBuyAmt.text = 0
	frm1.txtLocBuyAmt.text = 0
	frm1.txtPriceAmt.text = 0
	frm1.txtLocPriceAmt.text = 0
	frm1.txtCnt.text = 0
	frm1.txtCalRate.text = 0
    frm1.selCalYn.value = "Y"
    frm1.selComYn.value = "N"
    frm1.selEndYn.value = "O"
    frm1.txtSecuCode.focus
    lgBlnFlgChgValue = False
End Sub



Sub vspddata_Change(ByVal Col, ByVal Row)
	Dim intIndex
	Dim varData,varFlag
    Dim loc_amt
    Dim IntRetCD
    Dim RcptType
	
	Frm1.vspdData.Row = Row
	Frm1.vspdData.Col = Col 

	Select Case Col
			Case C_RcptType 
			RcptType = Frm1.vspdData.Text
			frm1.vspdData.ReDraw = False  
          
          If RcptType <> "" Then				
			IntRetCD  =  CommonQueryRs( "REFERENCE" , "B_CONFIGURATION  " , "MAJOR_CD = " & FilterVar("A1006", "''", "S") & "  AND MINOR_CD =  " & FilterVar(UCase(RcptType), "''", "S") & " AND SEQ_NO = 4 ", lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) 
			If IntRetCD = False Then
				Call DisplayMsgBox("141140","X","X","X")
					Frm1.vspdData.Col = C_RcptType
					Frm1.vspdData.Text = ""
					Frm1.vspdData.Col = C_RcptTypeNm
					Frm1.vspdData.Text = ""
				Set gActiveElement = document.activeElement  
			Else
				Select case UCase(Trim(Left(lgF0, Len(lgF0)-1)))
						Case "CS" 
							frm1.vspdData.Col  = C_NoteNo
							frm1.vspdData.Row  = Row
							frm1.vspdData.Text = ""   
										
							frm1.vspddata.Col  = C_BankAcct
							frm1.vspddata.Row  = Row
							frm1.vspddata.Text = ""			
							
						Case "DP" 
							frm1.vspdData.Col  = C_NoteNo
							frm1.vspdData.Row  = Row
							frm1.vspdData.Text = ""   

						Case "NO"
							frm1.vspdData.Col  = C_BankAcct
							frm1.vspdData.Row  = Row
							frm1.vspdData.Text = ""   

						Case Else          
							frm1.vspdData.Col  = C_NoteNo
							frm1.vspdData.Row  = Row
							frm1.vspdData.Text = "" 
							  
							frm1.vspdData.Col  = C_BankAcct
							frm1.vspdData.Row  = Row
							frm1.vspdData.Text = ""   
				End Select
				 
				IntRetCD = CommonQueryRs( "minor_nm" , "B_MINOR a, B_CONFIGURATION b " , "a.minor_cd = b.minor_cd and a.major_cd = " & FilterVar("A1006", "''", "S") & "  and b.seq_no = 2 and b.reference = " & FilterVar("PP", "''", "S") & "  AND a.MINOR_CD =  " & FilterVar(UCase(RcptType), "''", "S") & " ", lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) 
               If IntRetCd = True Then			
				 Frm1.vspdData.Col = C_RcptTypeNm
				 Frm1.vspdData.Text = Trim(Replace(lgF0,Chr(11),""))
			   End If	 
			ENd if
		  End If	
			frm1.vspdData.ReDraw = True 
	
		call subVspdSettingChange(Col,Row,Row, frm1.vspddata.Text)
		
'		Case C_LocAmt
'			If UNICDbl(frm1.vspdData.text) < 0 Then
'				frm1.vspdData.Text  = UNIConvNumPCToCompanyByCurrency(frm1.vspdData.Text * (-1),frm1.txtDocCur.value,parent.ggAmtOfMoneyNo, "X", "X")
'			End if
'			Call DoSum()
		
		Case C_Amt
			
			loc_amt =  UNICDbl(frm1.txtXchRate.text) * UNICDbl(frm1.vspdData.text)
			Frm1.vspdData.col = c_locamt
			frm1.vspdData.text = UNIFormatNumber(loc_amt,ggAmtOfMoney.DecPoint,-2,0,ggAmtOfMoney.RndPolicy,ggAmtOfMoney.RndUnit)
			Call FixDecimalPlaceByCurrency2(frm1.vspdData,Row,Frm1.txtTradeCur.value,C_Amt,  "A" ,"X","X")
			
	End Select		

	ggoSpread.Source = frm1.vspdData
	ggoSpread.UpdateRow Row
	
	lgBlnFlgChgValue = TRUE
End Sub

'==========================================================================================
'   Sub Procedure Name : subVspdSettingChange
'   Sub Procedure Desc : 
'==========================================================================================

Sub subVspdSettingChange(ByVal Col , ByVal Row,  ByVal Row2, Byval varData)	
	Dim intIndex
	Dim strval
	Dim CurRow
	
	For CurRow = Row To Row2
		frm1.vspdData.Col = C_RcptType
		frm1.vspdData.Row = CurRow
		strval = UCase(TRim(frm1.vspdData.Text))

		If CommonQueryRs( "REFERENCE" , "B_CONFIGURATION  " , "MAJOR_CD = " & FilterVar("A1006", "''", "S") & "  AND MINOR_CD =  " & FilterVar(strval , "''", "S") & " AND SEQ_NO = 4 ", lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then 
			Select Case UCase(lgF0)
				Case "DP" & Chr(11)   ' 예적금 
					ggoSpread.SpreadUnLock C_BankAcct,    CurRow,   CurRow 
					ggoSpread.SSSetRequired C_BankAcct,    CurRow,   CurRow 
					ggoSpread.SpreadUnLock C_BankAcctPopUp, CurRow, C_BankAcctPopUp,  CurRow
					ggoSpread.SSSetEdit	C_BankAcct, "예적금코드", 25, 0, CurRow, 30
					    
					ggoSpread.SpreadLock C_NoteNo,    CurRow, C_NoteNo,       CurRow
					ggoSpread.SpreadLock C_NoteNoPopup,     CurRow, C_NoteNoPopup,    CurRow  
				Case "NO" & Chr(11)
					ggoSpread.SpreadLock C_BankAcct,   CurRow, C_BankAcct,     CurRow 
					ggoSpread.SpreadLock C_BankAcctPopup,   CurRow, C_BankAcctPopup,  CurRow 
					ggoSpread.SSSetProtected C_BankAcct,      CurRow, CurRow
													
					ggoSpread.SpreadUnLock C_NoteNo,   CurRow, C_NoteNo,       CurRow
					ggoSpread.SSSetEdit      C_NoteNo, "어음번호", 25, 0, CurRow, 30	
					ggoSpread.SSSetRequired C_NoteNo,   CurRow, CurRow
					ggoSpread.SpreadUnLock C_NoteNoPopup,   CurRow, C_NoteNoPopup,    CurRow     
				Case Else
					ggoSpread.SpreadLock     C_BankAcct,      CurRow, C_BankAcct,     CurRow   			
					ggoSpread.SpreadLock     C_BankAcctPopup, CurRow, C_BankAcctPopup,CurRow
					ggoSpread.SSSetProtected C_BankAcct,      CurRow, CurRow							
		
					ggoSpread.SpreadLock     C_NoteNo,        CurRow, C_NoteNo,     CurRow
					ggoSpread.SpreadLock     C_NoteNoPopup,   CurRow, C_NoteNoPopup,CurRow		
					ggoSpread.SSSetProtected C_NoteNo,        CurRow, CurRow													
			End Select
		End If
	 Next 

End Sub

'===================================== CurFormatNumericOCX()  =======================================
'	Name : CurFormatNumericOCX()
'	Description : 화면에서 일괄적으로 Rounding되는 Numeric OCX 
'====================================================================================================
Sub CurFormatNumericOCX()

	With frm1
	'해당되는 금액이 있는 Data 필드에 대하여 각각 처리 
		'취득금액 
		ggoOper.FormatFieldByObjectOfCur .txtBuyAmt, .txtTradeCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec
		'액면금액 
		ggoOper.FormatFieldByObjectOfCur .txtPriceAmt, .txtTradeCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec
	End With
End Sub

'===================================== CurFormatNumSprSheet()  ======================================
'	Name : CurFormatNumSprSheet()
'	Description : 화면에서 일괄적으로 Rounding되는 Numeric Spread Sheet
'====================================================================================================
Sub CurFormatNumSprSheet()
	Dim ii
	With frm1
		'해당되는 금액이 있는 Grid에 대하여 각각 처리 
		ggoSpread.Source = frm1.vspdData
		For ii = 1 To .vspdData.MaxRows 
			Call FixDecimalPlaceByCurrency2(frm1.vspdData,ii,.txtTradeCur.value,C_Amt,"A" ,"X","X")
      	Next
       Call ReFormatSpreadCellByCellByCurrency2(Frm1.vspdData,1,-1,.txtTradeCur.value,C_Amt,"A" ,"I","X","X")

	End With
End Sub  

'==========================================================================================
'   Event Name : txtDocCur_OnChangeASP
'   Event Desc : 
'==========================================================================================
Sub txtDocCur_OnChangeASP()
 
    IF CommonQueryRs( "CURRENCY_DESC" , "B_CURRENCY" ,  " CURRENCY =  " & FilterVar(frm1.txtTradeCur.value , "''", "S") & "" , lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then   			        							

		Call CurFormatNumericOCX()
		Call CurFormatNumSprSheet()

	END IF	    
End Sub
'==========================================================================================
Sub vspddata_Click(ByVal Col, ByVal Row)

	
	Call SetPopupMenuItemInf("1101111111")

    gMouseClickStatus = "SPC"	'Split 상태코드 
      

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
    
    	frm1.vspdData.Row = Row
End Sub


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
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub

'========================================================================================================
Sub vspdData_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    
    Call GetSpreadColumnPos("A")
End Sub
'==========================================================================================
'   Event Desc : Spread Split 상태코드 
'==========================================================================================

Sub vspddata_MouseDown(Button, Shift, X, Y)
	If Button = 2 And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If
End Sub

'==========================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : 
'==========================================================================================

Sub vspddata_ButtonClicked(ByVal Col, ByVal Row, ByVal ButtonDown)
	Dim strTemp
	Dim intPos1
	Dim strCard

	With frm1.vspddata 
		If Row > 0 then
			if  Col = C_BankAcctPopup Then
				.Col = C_BankAcct
				.Row = Row
				strTemp = Trim(.text)

				.col = C_RcptType
				strCard = .text
				Call OpenBankAcct(strTemp, strCard)
			elseif Col = C_NoteNoPopup Then
				.Col = C_NoteNo
				.Row = Row
				strTemp = Trim(.text)

				.col = C_RcptTypeNm
				strCard = .text
				Call OpenNoteNo(strTemp, strCard)
			elseif Col = C_RcptTypePopup Then
  				.Col = C_RcptType
				Call OpenRcptType(frm1.vspdData.Text, 1, Row)
			    
			end if
		End If
	End With	
End Sub


Function OpenRcptType(strCode, iWhere, Row)

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	Select Case iWhere
	    Case 1
	    	arrParam(1) = "B_MINOR a, B_CONFIGURATION b "	' TABLE 명칭 
	    	arrParam(2) = Trim(strCode)	                        ' Code Condition
	    	arrParam(3) = ""									' Name Cindition
   	        arrParam(4) = "a.minor_cd = b.minor_cd and a.major_cd = " & FilterVar("A1006", "''", "S") & "  and b.seq_no = 2 and b.reference = " & FilterVar("PP", "''", "S") & "  "          <%' Where Condition%>	       
	    	arrParam(5) = "출금유형"		   				    ' TextBox 명칭 
	
	    	arrField(0) = "a.minor_cd"		                ' Field명(0)
	    	arrField(1) = "a.minor_nm"    						' Field명(1)%>
    
	    	arrHeader(0) = "출금유형"		        		' Header명(0)%>
	    	arrHeader(1) = "출금유형명"	        					' Header명(1)%>
			arrParam(0) = arrParam(5)								  ' 팝업 명칭 

	End Select


	arrRet = window.showModalDialog("../../comasp/AdoCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Call SetActiveCell(frm1.vspdData,C_Rcpttype,frm1.vspdData.ActiveRow ,"M","X","X")
		Exit Function
	Else
		Call SetRcpt(arrRet, iWhere, Row)
	End If	
	
End Function



'------------------------------------------  SetSItemDC()  --------------------------------------------------
'	Name : SetRcpt()
'	Description : OpenSItemDC Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetRcpt(arrRet, iWhere, Row)

	With frm1
        .vspdData.Row = Row
		Select Case iWhere
		    Case 1
		        .vspdData.Col = C_RcptType
		    	.vspdData.text = arrRet(0) 
		    	.vspdData.Col = C_RcptTypeNm
		    	.vspdData.text = arrRet(1)
				Call subVspdSettingChange(C_RcptType, frm1.vspdData.ActiveRow ,frm1.vspdData.ActiveRow, arrRet(0) )
				Call vspdData_Change(C_Rcpttype, .vspdData.Row)
				Call SetActiveCell(frm1.vspdData,C_RcptType,frm1.vspdData.ActiveRow ,"M","X","X")

		End Select

		lgBlnFlgChgValue = True

	End With
	ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row
    
End Function



Sub vspddata_KeyPress(index , KeyAscii )
    lgBlnFlgChgValue = True                                                 '⊙: Indicates that value changed
End Sub

'========================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If

	If CheckRunningBizProcess = True Then
	   Exit Sub
	End If
    
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	           
		If lgStrPrevKeyIndex <> 0 Then                         
           
           If DbQuery = False Then
           
              Exit Sub
           End if
    	End If
    End if
End Sub
'=============================================== 2.2.3 InitSpreadSheet() ========================================
Sub InitSpreadSheet()

	Call InitSpreadPosVariables()

	ggoSpread.Source = frm1.vspddata
	ggoSpread.SpreadInit "V20030102",,parent.gAllowDragDropSpread

    With frm1.vspdData
		.Redraw = False
		.MaxCols = C_BankNm + 1												'☜: 최대 Columns의 항상 1개 증가시킴 
		.Col = .MaxCols															'공통콘트롤 사용 Hidden Column
		.ColHidden = True    
		.MaxRows = 0

		Call GetSpreadColumnPos("A")
   

    ggoSpread.SSSetEdit	  C_Seq,       "순번",        5, 2, -1, 5
	ggoSpread.SSSetEdit   C_RcptType,  "출금유형",    10,,,10,2 
	ggoSpread.SSSetButton   C_RcptTypePopup
	ggoSpread.SSSetEdit   C_RcptTypeNm,  "출금유형명",  15,,,20,2 
	ggoSpread.SSSetFloat  C_Amt,       "금액",       19, "A", ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
	ggoSpread.SSSetFloat  C_LocAmt,    "금액(자국)", 19, Parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
    
      
	If gIsShowLocal = "N" Then
		.Col		= C_LocAmt
		.ColHidden = True
		.ColHidden = True
	End If

	ggoSpread.SSSetEdit   C_NoteNo,    "어음번호",     25, 0, -1, 30,2
	ggoSpread.SSSetButton C_NoteNoPopup		    
	ggoSpread.SSSetEdit	  C_BankAcct,  "예적금코드",   25, 0, -1, 30,2
	ggoSpread.SSSetButton C_BankAcctPopup
	ggoSpread.SSSetEdit	  C_BankCd,  "은행코드",   20, 0, -1, 30,2
	ggoSpread.SSSetEdit	  C_BankNm,  "은행명",   20, 0, -1, 30,2

	Call ggoSpread.SSSetColHidden(C_BankCd,C_BankCd,True)
	Call ggoSpread.SSSetColHidden(C_BankNm,C_BankNm,True)

	call ggoSpread.MakePairsColumn(C_RcptType,C_RcptTypePopup)
	call ggoSpread.MakePairsColumn(C_Amt,C_LocAmt)
	call ggoSpread.MakePairsColumn(C_NoteNo,C_NoteNoPopup)
	call ggoSpread.MakePairsColumn(C_BankAcct,C_BankAcctPopup)


	.ReDraw = true

	Call SetSpreadLock 
    
    End With
    
End Sub


'================================== 2.2.4 SetSpreadLock() ==================================================
Sub SetSpreadLock() 'byVal gird_fg, byVal lock_fg, byVal iRow)
    With frm1
		
		'if Ucase(gird_fg) = "I" then
			ggoSpread.Source = .vspddata		
			.vspddata.ReDraw = False
		
			ggoSpread.SpreadLock		C_Seq,			-1, C_Seq			, -1
			ggoSpread.SpreadLock		C_RcptTypePopup,-1, C_RcptTypePopup	, -1
			ggoSpread.SpreadLock		C_RcptTypeNm,	-1, C_RcptTypeNm	, -1
			ggoSpread.SpreadLock		C_NoteNo,		-1, C_NoteNo		, -1
			ggoSpread.SpreadLock		C_NoteNoPopup,	-1, C_NoteNoPopup	, -1
			ggoSpread.SpreadLock		C_BankAcct,		-1, C_BankAcct		, -1
			ggoSpread.SpreadLock		C_BankAcctPopup,-1, C_BankAcctPopup	, -1
			
			ggoSpread.SSSetRequired  C_RcptType, -1, -1 
			ggoSpread.SSSetRequired  C_Amt, -1, -1 
			
			ggoSpread.SSSetProtected	.vspdData.MaxCols,-1,-1

			.vspddata.ReDraw = True
		
		'end if
   End With    
End Sub

'================================== 2.2.5 SetSpreadColor() ==================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================

Sub SetSpreadColor_Item(ByVal iwhere, ByVal pvStartRow, ByVal pvEndRow)

    With frm1.vspddata 
    
    ggoSpread.Source = frm1.vspddata
    
    .ReDraw = False
	Select Case iwhere
	Case "Q"
		ggoSpread.SSSetProtected C_Seq			, pvStartRow, pvEndRow  
		ggoSpread.SSSetRequired	 C_Amt			, pvStartRow, pvEndRow  
		ggoSpread.SSSetRequired	 C_RcptType		, pvStartRow, pvEndRow  
		ggoSpread.SSSetProtected C_RcptTypeNm	, pvStartRow, pvEndRow   
		ggoSpread.SpreadLock	 C_NoteNo		, pvStartRow, pvEndRow   
		ggoSpread.SpreadLock	 C_NoteNoPopup	, pvStartRow, pvEndRow   
		ggoSpread.SpreadLock	 C_BankAcct		, pvStartRow, pvEndRow   
		ggoSpread.SpreadLock	 C_BankAcctPopup, pvStartRow, pvEndRow  
	CASE "I"
		ggoSpread.SSSetProtected C_Seq			, pvStartRow, pvEndRow  
		ggoSpread.SSSetRequired	 C_Amt			, pvStartRow, pvEndRow  
		ggoSpread.SSSetRequired	 C_RcptType		, pvStartRow, pvEndRow  
		ggoSpread.SSSetProtected C_RcptTypeNm	, pvStartRow, pvEndRow   
	End Select
	.ReDraw = True	        
   
    End With
End Sub



'========================================================================================================
Sub Form_Load()
    Err.Clear                                                                        '☜: Clear err status
	Call LoadInfTB19029                                                              '☜: Load table , B_numeric_format

	'------ Developer Coding part (Start ) --------------------------------------------------------------

	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
	Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
	Call ggoOper.LockField(Document, "N")                                            '⊙: Lock  Suitable  Field
	'Call ggoOper.FormatNumber(frm1.txtCnt, "9999999999999", "0", true)	
    
    Call InitVariables                                                               '⊙: Setup the Spread sheet
	Call InitData
	Call SetDefaultVal
	Call InitSpreadSheet()                                                               '⊙: Setup the Spread sheet
	Call InitComboBox
	Call SetToolbar("1111100000111111")                                                     '☆: Developer must customize

	lgBlnFlgChgValue = false
	'------ Developer Coding part (End )   --------------------------------------------------------------
	gSelframeFlg = TAB1

End Sub

'========================================================================================================
Sub Form_QueryUnload(Cancel, UnloadMode)
End Sub

'========================================================================================================
Function FncQuery()
    Dim IntRetCD
    Dim var_m
    FncQuery = False															 '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status


	If Not chkField(Document, "1") Then									         '☜: This function check required field
       Exit Function
    End If

    ggoSpread.Source = frm1.vspdData
    var_m = ggoSpread.SSCheckChange
 
	 
   	If lgBlnFlgChgValue = True Or var_m = True    Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO,"X","X") '☜ "데이타가 변경되었습니다. 조회하시겠습니까?"
	
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If	


    Call ggoOper.ClearField(Document, "2")										 '☜: Clear Contents  Field
    Call ggoOper.LockField(Document , "N")                                        '☜: Lock  Field
    
	call ClickTab1()
 
	ggoSpread.Source= frm1.vspdData
	ggoSpread.ClearSpreadData
	Call InitData

 	Call SetDefaultVal()
   
    Call InitVariables
		

	Call InitComboBox
	
    If DbQuery = False Then
       Exit Function
    End If                                                                 '☜: Query db data
	
    Set gActiveElement = document.ActiveElement
    FncQuery = True                                                              '☜: Processing is OK
End Function

'========================================================================================================
Function FncNew()
    Dim IntRetCD

    FncNew = False																  '☜: Processing is NG
    Err.Clear                                                                     '☜: Clear err status

    If lgBlnFlgChgValue = True Then
       IntRetCD = DisplayMsgBox("900015", parent.VB_YES_NO,"x","x")					  '☜: Data is changed.  Do you want to make it new?
       If IntRetCD = vbNo Then
          Exit Function
       End If
    End If

    Call ggoOper.ClearField(Document, "A")                                        '☜: Clear Condition Field
    Call ggoOper.LockField(Document , "N")                                        '☜: Lock  Field
'    Call ggoOper.SetReqAttr(frm1.txtSecuCode1,"N")
    '--------- Developer Coding Part (Start) ----------------------------------------------------------

'    frm1.btnDept2.disabled = 1

	Call InitVariables                                                               '⊙: Setup the Spread sheet

    ggoSpread.Source = frm1.vspdData        				
    ggoSpread.ClearSpreadData						
	Call InitData
	Call InitComboBox
	Call SetDefaultVal
    Call ClickTab1()
    

	Call SetToolbar("1111100000101111")
    call txtDocCur_OnChangeASP()
	

	'------ Developer Coding part (End )   --------------------------------------------------------------
    Set gActiveElement = document.ActiveElement
    FncNew = True
End Function

'========================================================================================================
Function FncDelete()
    Dim intRetCD

    FncDelete = False                                                             '☜: Processing is NG
    Err.Clear                                                                     '☜: Clear err status

    If lgIntFlgMode <> parent.OPMD_UMODE Then                                            '☜: Please do Display first.
        Call DisplayMsgBox("900002","x","x","x")
        Exit Function
    End If

    IntRetCD = DisplayMsgBox("900003", parent.VB_YES_NO,"x","x")                         '☜: Do you want to delete?
	If IntRetCD = vbNo Then
		Exit Function
	End If

    If Not chkField(Document, "1") Then									        '⊙: This function check indispensable field
       Exit Function
    End If

	'------ Developer Coding part (Start ) --------------------------------------------------------------
	'------ Developer Coding part (End )   --------------------------------------------------------------
    Call MakeKeyStream("D")
    If DbDelete = False Then                                                      '☜: Query db data
       Exit Function
    End If

    Set gActiveElement = document.ActiveElement
    FncDelete = True                                                              '☜: Processing is OK
End Function

'========================================================================================================
Function FncSave()
	Dim IntRetCD
	Dim FrDt
	Dim strSelect, strFrom, strWhere
	Dim var1, Pos
    FncSave = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
    If lgBlnFlgChgValue = False Then 
        IntRetCD = DisplayMsgBox("900001","x","x","x")                           '☜:There is no changed data. 
        Exit Function
    End If

    Call deptCheck()
    
    If Not chkField(Document, "2") Then
       Exit Function
    End If
	
'	IntRetCD= CommonQueryRs(" TEMP_GL_FG "," B_CALENDAR "," CALENDAR_DT = '" & UNIConvDate(frm1.txtBillDt.Text) & "'" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) 
'		If IntRetCD=False Or Trim(Replace(lgF0,Chr(11),"")) = "" Then
'		Else
'			If Trim(Replace(lgF0,Chr(11),"")) = "C" Then
'				Call DisplayMsgBox("121291","X","X","X")                         '☜ : 결의전표가 마감되었습니다.(수정요망)
'				Exit Function
'		    End IF 
'		End If
	
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    If frm1.selCalYn.value = "Y" Then
         If CompareDateByFormat(frm1.txtBillDt.text,frm1.txtExpireDt.text,frm1.txtBillDt.Alt,frm1.txtExpireDt.Alt,"970023",frm1.txtBillDt.UserDefinedFormat,parent.gComDateType,True) = False Then
		    Exit Function
        End If
    End If
	
	If CompareDateByFormat(frm1.txtInDt.Text, frm1.txtBillDt.Text,frm1.txtInDt.Alt,frm1.txtBillDt.Alt, "970025", frm1.txtInDt.UserDefinedFormat, parent.gComDateType, True)=False Then  '☜ : 시장월은 종료월보다 작아야합니다. 
'	        frm1.txtBillDt.Text = ""
            frm1.txtBillDt.focus
            Set gActiveElement = document.ActiveElement
            Exit Function
    End if 
	
'	If CompareDateByFormat(frm1.txtBillDt.text,frm1.txtInDt.text,frm1.txtBillDt.Alt,frm1.txtInDt.Alt,"970023",frm1.txtBillDt.UserDefinedFormat,parent.gComDateType,True) = False Then
'		    Exit Function
'    End If
	
    If frm1.selComYn.value = "Y" Then
         If CompareDateByFormat(frm1.txtBillDt.text,frm1.txtOutDt.text,frm1.txtBillDt.Alt,frm1.txtOutDt.Alt,"970023",frm1.txtBillDt.UserDefinedFormat,parent.gComDateType,True) = False Then
		    Exit Function
        End If
    End If

	
	
	If UNICDbl(frm1.txtXchRate.text) = 0 Then
		Call DisplayMsgBox("800443","X",frm1.txtXchRate.alt,"0")                         '☜ : 숫자영 '//환율 
		frm1.txtXchRate.focus
		Set gActiveElement = document.activeElement  	
		Exit Function
	End If	
	
	If UNICDbl(frm1.txtBuyAmt.text) = 0 Then
		Call DisplayMsgBox("800443","X",frm1.txtBuyAmt.alt,"0")                         '☜ : 숫자영'//취득 
		frm1.txtBuyAmt.focus
		Set gActiveElement = document.activeElement  	
		Exit Function
	End If	
	
	If UNICDbl(frm1.txtPriceAmt.text) = 0 Then
		Call DisplayMsgBox("800443","X",frm1.txtPriceAmt.alt,"0")                         '☜ : 숫자영'//액면 
		frm1.txtPriceAmt.focus
		Set gActiveElement = document.activeElement  	
		Exit Function
	End If	
	
	If UNICDbl(frm1.txtCnt.text) = 0 Then
		Call DisplayMsgBox("800443","X",frm1.txtCnt.alt,"0")                         '☜ : 숫자영'//매수 
		frm1.txtCnt.focus
		Set gActiveElement = document.activeElement  	
		Exit Function
	End If	
	
	If 	frm1.txtSecuTypeNm.value <> "" Then
	    var1 =  frm1.txtSecuTypeNm.value
		Pos =  instr(1,var1,"주식")
	   If Pos = 0 Then	
			If UNICDbl(frm1.txtCalRate.text) = 0 Then
				Call DisplayMsgBox("141157","X","X","X")                         '☜ : 숫자영'//이자율 
				frm1.txtCalRate.focus
				Set gActiveElement = document.activeElement  	
				Exit Function
			End If	
	   Else	
	   End If			
	End If	
	
	FrDt = UniConvDateToYYYYMMDD(frm1.txtInDt.Text,parent.gDateFormat,"")   '//parent.UNIConvDate(frm1.txtInDt.Text)
	strSelect = strSelect & " isnull(case t.loc_cur when " & FilterVar(UCase(frm1.txtTradeCur.value), "''", "S") & " Then 1 "
	strSelect = strSelect & "    Else  Case t.xch_rate_fg " 
	strSelect = strSelect & "        When " & FilterVar("M", "''", "S") & "  Then ( SELECT isnull(STD_RATE,0) "
	strSelect = strSelect & "               FROM    b_monthly_exchange_rate (nolock) "
	strSelect = strSelect & "               WHERE apprl_yrmnth  = CONVERT (varchar(06), " & FilterVar(FrDt,"" & FilterVar("29991231", "''", "S") & " ", "S") & ", 112) "
	strSelect = strSelect & "               and from_currency   = " & FilterVar(UCase(frm1.txtTradeCur.value), "''", "S") 
	strSelect = strSelect & "               and to_currency     = t.loc_cur ) "
	strSelect = strSelect & "		  Else (	SELECT  isnull(STD_RATE,0) "
	strSelect = strSelect & "               FROM   b_daily_exchange_rate (nolock) "
	strSelect = strSelect & "               WHERE   apprl_dt    = " & FilterVar(FrDt,"" & FilterVar("29991231", "''", "S") & " ", "S")
	strSelect = strSelect & "               and from_currency   = " & FilterVar(UCase(frm1.txtTradeCur.value), "''", "S") 
	strSelect = strSelect & "               and to_currency     = t.loc_cur ) "
	strSelect = strSelect & "         End "
	strSelect = strSelect & " End,0) as xch_rate "
	strFrom  = " (SELECT isnull(XCH_RATE_FG,'') as xch_rate_fg, loc_cur from b_company) t "  
	strWhere = ""
	IntRetCD= CommonQueryRs(strSelect,strFrom,strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) 			
	If IntRetCD = False or Trim(Replace(lgF0,Chr(11),""))="0" or Trim(Replace(lgF0,Chr(11),"")) = ""  Then
		Call DisplayMsgBox("am0023","X","X","X")         
		Exit Function
	End If	
	
	Call MakeKeyStream("S")
    '--------- Developer Coding Part (End) ------------------------------------------------------------

    If DbSave = False Then                                                       '☜: Query db data
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
     '--------- Developer Coding Part (Start) ----------------------------------------------------------

    
    If gSelframeFlg = TAB1 Then	 
		If lgBlnFlgChgValue = True Then
			IntRetCD = DisplayMsgBox("900017", Parent.VB_YES_NO,"x","x")				     '☜: Data is changed.  Do you want to continue? 
			If IntRetCD = vbNo Then
				Exit Function
			End If
		End If
		Call ggoOper.ClearField(Document, "1")                                       '⊙: Clear Condition Field
		Call ggoOper.LockField(Document, "N")		     
		lgIntFlgMode = Parent.OPMD_CMODE												     '⊙: Indicates that current mode is Crate mode
		
		frm1.vspdData.maxRows = 0
	    ggoSpread.Source = frm1.vspdData 
		ggoSpread.Spreadinit 

     
		frm1.txtSecuCode.value = ""
		frm1.txtSecuCode1.value = ""
		frm1.txtSecuNm.value = ""
		frm1.txtSecuNm1.value = ""
		frm1.txtTGlNo.value = ""
		frm1.txtGlNo.value = ""

		frm1.vspdData.ReDraw = True
		
	Elseif  gSelframeFlg = TAB2 Then	 
	
		If lgIntFlgMode <> Parent.OPMD_UMODE Then
			lgIntFlgMode = Parent.OPMD_CMODE
		End If
		frm1.vspddata.ReDraw = False

		if frm1.vspddata.MaxRows < 1 then Exit Function
	
		ggoSpread.Source = frm1.vspddata	
		ggoSpread.CopyRow
		Call ReFormatSpreadCellByCellByCurrency2(Frm1.vspdData,frm1.vspdData.ActiveRow,frm1.vspdData.ActiveRow,frm1.txtTradeCur.value,C_Amt,   "A" ,"I","X","X")

		Call vspddata_Change(C_RcptType, frm1.vspddata.ActiveRow)
		Call SetSpreadColor_Item("I", frm1.vspddata.ActiveRow, frm1.vspddata.ActiveRow)
 
		MaxSpreadVal frm1.vspddata.ActiveRow				
    
    	frm1.vspddata.Col = C_RcptType
    	
		
		frm1.vspddata.ReDraw = True	
		
	End if    
	lgBlnFlgChgValue = True
    '--------- Developer Coding Part (End) ------------------------------------------------------------
    FncCopy = True                                                               '☜: Processing is OK
End Function



'========================================================================================================
Function FncCancel()
    Dim lRow
	Dim strVal
	Dim varFlag
    FncCancel = False                                                            '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
	If gSelframeFlg = TAB2 Then  'Master단 
		ggoSpread.Source = frm1.vspdData
		ggoSpread.EditUndo


'		Call subVspdSettingChange(Frm1.vspdData.ActiveRow,varFlag )
        Call ReFormatSpreadCellByCellByCurrency2(Frm1.vspdData,Frm1.vspdData.ActiveRow,Frm1.vspdData.ActiveRow,frm1.txtTradeCur.value,C_Amt,   "A" ,"I","X","X")

	End If	

    '--------- Developer Coding Part (End) ------------------------------------------------------------
    Set gActiveElement = document.ActiveElement   
    FncCancel = True                                                             '☜: Processing is OK
End Function

'========================================================================================================
Function  FncInsertRow(ByVal pvRowCnt) 

    FncInsertRow = False														 '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
	Dim varMaxRow,iCurRowPos
	Dim strDoc
	Dim varXrate

	Dim IntRetCD
	Dim imRow
	Dim imRow2


    if IsNumeric(Trim(pvRowCnt)) Then
			imRow = CInt(pvRowCnt)
		else
			imRow = AskSpdSheetAddRowcount()

			If ImRow="" then
			Exit Function
			End If
	End If

    
        
    If   gSelframeFlg = TAB2 Then        '''' Acq Item
     
		with frm1
		
			.vspddata.focus
			
			varMaxRow = .vspddata.MaxRows		
			iCurRowPos = .vspddata.ActiveRow		
		
			ggoSpread.Source = .vspddata
			.vspddata.ReDraw = False
		
			For imRow2=1 to imRow
			ggoSpread.InsertRow ,1
			.vspddata.ReDraw = True
			Call SetSpreadColor_Item ("Q", .vspdData.ActiveRow, .vspdData.ActiveRow)
		
			'SetSpreadColor_Item .vspddata.ActiveRow			
			MaxSpreadVal .vspddata.ActiveRow				

			
			Next
        Call ReFormatSpreadCellByCellByCurrency2(Frm1.vspdData,iCurRowPos + 1,iCurRowPos + imRow,frm1.txtTradeCur.value,C_Amt,"A" ,"I","X","X")

		end with
		
	END if
	
	Set gActiveElement = document.ActiveElement 
    '--------- Developer Coding Part (End) ------------------------------------------------------------
    FncInsertRow = True   
    
    
                                                           '☜: Processing is OK
End Function

'========================================================================================================
Function FncDeleteRow()
    Dim lDelRows

    FncDeleteRow = False                                                         '☜: Processing is NG
	Err.Clear
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
	Dim varMaxRow
	Dim strDoc
	Dim varXrate
	 
        
	If gSelframeFlg = TAB2 Then	
		frm1.vspdData.focus
    	ggoSpread.Source = frm1.vspdData
		if frm1.vspdData.MaxRows < 1 then Exit Function
	
		ggoSpread.DeleteRow
	End If
		
	lgBlnFlgChgValue = True
    '--------- Developer Coding Part (End) ------------------------------------------------------------
        Set gActiveElement = document.ActiveElement   
    FncDeleteRow = True                                                          '☜: Processing is OK
End Function

'========================================================================================================
Function FncPrint()
    FncPrint = False
	Err.Clear
	Call Parent.FncPrint()
    FncPrint = True
End Function

'========================================================================================================
' Name : FncExcel
' Desc : This function is called by MainExcel in Common.vbs
'========================================================================================================
Function FncExcel()
    FncExcel = False
    Err.Clear
	Call Parent.FncExport(parent.C_SINGLE)
    FncExcel = True
End Function

'========================================================================================================
' Name : FncFind
' Desc : This function is called by MainFind in Common.vbs
'========================================================================================================
Function FncFind()
    FncFind = False
    Err.Clear
	Call Parent.FncFind(parent.C_SINGLE, True)
    FncFind = True
End Function

'========================================================================================================
Function FncExit()
	Dim IntRetCD

    FncExit = False
    Err.Clear

	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO,"x","x")
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If
    FncExit = True
End Function
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
	Dim iRow
	Dim intIndex
	Dim	varData
    
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()
    Call InitComboBox
    Call ggoSpread.ReOrderingSpreadData()
	Call ReFormatSpreadCellByCellByCurrency2(Frm1.vspdData,-1 , -1 ,frm1.txtTradeCur.value ,C_AMT ,   "A" ,"I","X","X")
   	
End Sub

'========================================================================================
Sub FncSplitColumn()

    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   

End Sub

'========================================================================================================
Function DbQuery()
    Err.Clear                                                                    '☜: Clear err status
    DbQuery = False                                                              '☜: Processing is NG

	if LayerShowHide(1) = false then
	    Exit Function
	end if

	Dim strVal
	'------ Developer Coding part (Start)  --------------------------------------------------------------

    With frm1
    strVal = BIZ_PGM_ID & "?txtMode="          & parent.UID_M0001                       '☜: Query
    strVal = strVal     & "&txtPrevNext="      & ""	                         '☜: Direction
    strVal = strVal     & "&txtSecuCode="      & Trim(frm1.txtSecuCode.value)                       '☜: Query Key
    strVal = strVal     & "&lgStrPrevKeyIndex=" & lgStrPrevKeyIndex
    strVal = strVal     & "&txtMaxRows="         & Frm1.vspdData.MaxRows         '☜: Max fetched data
    strVal = strVal		& "&lgCurrency="		& frm1.txtTradeCur.value
    End With
	'------ Developer Coding part (End )   --------------------------------------------------------------
	
	Call RunMyBizASP(MyBizASP, strVal)                                               '☜: Run Biz Logic

    DbQuery = True
     Set gActiveElement = document.ActiveElement   
End Function

'========================================================================================================
Function DbSave()
	Dim lGrpcnt 
	Dim strVal, strDel
	Dim IntRows
	
    Err.Clear                                                                    '☜: Clear err status
    DbSave = False                                                               '☜: Processing is NG
   	lGrpCnt =0

	IF	ValidCheck = False Then
		Exit Function
	End If
	 
   if LayerShowHide(1) = false then                                                        '☜: Show Processing Message
		exit function
	end if
	'------ Developer Coding part (Start)  --------------------------------------------------------------

  	With frm1
		.txtMode.value        = parent.UID_M0002                                        '☜: Delete
        .txtFlgMode.value     = lgIntFlgMode
        .txtKeyStream.Value   = lgKeyStream                                      '☜: Save Key
        .txtUpdtUserId.value  = parent.gUsrID
        .txtInsrtUserId.value = parent.gUsrID
        .txtSpread.value	  = ""
	End With

	With frm1.vspdData
	    
    
    For IntRows = 1 To .MaxRows
		
		.Row = IntRows
		.Col = 0		
		
		If .Text = ggoSpread.DeleteFlag Then
			strDel = strDel & "Sheet1" & parent.gColSep  & "D" & parent.gColSep & IntRows & parent.gColSep				'D=Delete
		ElseIf .Text = ggoSpread.UpdateFlag Then
			strVal = strVal & "Sheet1" & parent.gColSep  & "U" & parent.gColSep & IntRows & parent.gColSep				'U=Update
		ElseIf .Text = ggoSpread.InsertFlag Then
			strVal = strVal & "Sheet1" & parent.gColSep  & "C" & parent.gColSep & IntRows & parent.gColSep				'C=Create
		End If		
	
		Select Case .Text		    
		        
		    Case ggoSpread.DeleteFlag

				.Col = C_Seq
				strDel = strDel & Trim(.Text) & parent.gRowSep				    '마지막 데이타는 Row 분리기호를 넣는다 
					
				lGrpcnt = lGrpcnt + 1            
		    
		    Case ggoSpread.UpdateFlag
		        .Col = C_Seq	'3
		        strVal = strVal & Trim(.Text) & parent.gColSep
		             
		        .Col = C_RcptType   '4
		        strVal = strVal & Trim(.Text) & parent.gColSep

				.Col = C_Amt		'5

		        strVal = strVal & Trim(.Text) & parent.gColSep
		        
		        .Col = C_LocAmt		'6
		        strVal = strVal & Trim(.Text) & parent.gColSep
   		        
 		        .Col = C_NoteNo		'7
		        strVal = strVal & Trim(.Text) & parent.gColSep			        

   		        .Col = C_BankCd		'8				
		        strVal = strVal & Trim(.Text) & parent.gColSep		        

		        .Col = C_BankAcct	'9				
		        strVal = strVal & Trim(.Text) & parent.gRowSep	        
		           		        
		        		           		        
		        lGrpCnt = lGrpCnt + 1
			
			Case ggoSpread.InsertFlag
		        .Col = C_Seq	'3
		        strVal = strVal & Trim(.Text) & parent.gColSep
		             
		        .Col = C_RcptType   '4
		        strVal = strVal & Trim(.Text) & parent.gColSep

				.Col = C_Amt		'5

		        strVal = strVal & Trim(.Text) & parent.gColSep
		        
		        .Col = C_LocAmt		'6
		        strVal = strVal & Trim(.Text) & parent.gColSep
		        

		        .Col = C_NoteNo		'7
		        strVal = strVal & Trim(.Text) & parent.gColSep			        
 
 		        .Col = C_BankCd		'8				
		        strVal = strVal & Trim(.Text) & parent.gColSep		        
  		        
		        .Col = C_BankAcct	'9				
		        strVal = strVal & Trim(.Text) & parent.gRowSep		        
		           		        
		        		           		        
		        lGrpCnt = lGrpCnt + 1

		End Select

    Next

	End With
	
    frm1.txtMaxRows.value  = lGrpCnt										'Spread Sheet의 변경된 최대갯수 
	frm1.txtSpread.value   = strDel & strVal
	
	'------ Developer Coding part (End )   --------------------------------------------------------------
	Call ExecMyBizASP(frm1, BIZ_PGM_ID)

    DbSave = True                                                               '☜: Processing is OK

    Set gActiveElement = document.ActiveElement
End Function

'========================================================================================================
Function DbDelete()
	Dim strVal
    Err.Clear                                                                    '☜: Clear err status

	DbDelete = False			                                                 '☜: Processing is NG

	if LayerShowHide(1) = False then
	   Exit Function
	end if

    strVal = BIZ_PGM_ID & "?txtMode="          & parent.UID_M0003                       '☜: Query
    strVal = strVal     & "&txtKeyStream="     & lgKeyStream                     '☜: Key
    
	Call RunMyBizASP(MyBizASP, strVal)                                           '☜: Run Biz logic

	DbDelete = True                                                              '⊙: Processing is NG
End Function

'========================================================================================================
Sub DbQueryOk()

	dim iRow
	Dim varData
	Dim intIndex
	
	lgIntFlgMode      = Parent.OPMD_UMODE
	'------ Developer Coding part (Start)  --------------------------------------------------------------
    If Trim(frm1.txtSecuCode1.value) <> "" Then
        Call ggoOper.SetReqAttr(frm1.txtSecuCode1,"Q")
        Call ggoOper.SetReqAttr(frm1.txtDept2,"N")
    End If
   
	If frm1.txtGlNo.value = "" Then
		
	    Call SetToolbar("1111100011011111")                                              '☆: Developer must customize
	    Call ggoOper.LockField(Document, "N")
	    Call ggoOper.SetReqAttr(frm1.txtSecuCode1,"Q")
	    selCalYn_Change()
		selComYn_Change()
	Else
	    Call SetToolbar("1110100011011111")                                              '☆: Developer must customize
	    Call ggoOper.LockField(Document, "Q")
	    '-- 승인코드가 있으므로 수정불가 
	    Call ggoOper.SetReqAttr(frm1.selComYn,"N")
	    Call ggoOper.SetReqAttr(frm1.txtDept1,"N")
	    Call ggoOper.SetReqAttr(frm1.txtDept2,"N")
    End If

    Set gActiveElement = document.ActiveElement  
	Call SetSpreadColor_Item("Q",-1, -1)
	
	If gSelframeFlg = TAB1 Then
		Call SetToolbar("111110000011111")                                                     '☆: Developer must customize
	Else
		Call SetToolbar("1100111100100111")              
	END IF	
	With frm1				
	
		.vspdData.Redraw = False
	
		For iRow = 1 To frm1.vspdData.MaxRows
	
			.vspdData.Col = C_RcptType		
			.vspdData.Row = iRow
			
			varData = frm1.vspdData.text
			
			Call vspddata_Change(.vspdData.Col,iRow)
			Call subVspdSettingChange(C_RcptType,1,frm1.vspdData.Maxrows, varData)
				
			ggoSpread.Source = frm1.vspdData
			ggoSpread.EditUndo iRow				
		Next
		
		.vspdData.Redraw = True			
	End With

	Call txtDocCur_OnChangeASP()

   	'Call txtDocCur_OnChangeASP()
    lgBlnFlgChgValue = False       
	'------ Developer Coding part (End )   --------------------------------------------------------------
End Sub

'========================================================================================================
Sub DbSaveOk()
	'------ Developer Coding part (Start)  --------------------------------------------------------------
	On error Resume next
	
    ggoSpread.Source= frm1.vspdData
	ggoSpread.ClearSpreadData
	
	Set gActiveElement = document.ActiveElement   
	frm1.txtSecuCode.value = frm1.txtSecuCode1.value
	Call DbQuery()
End Sub

'========================================================================================================
Sub DbDeleteOk()
'	Call SetToolbar("11101000000011")
	Call InitVariables()
	Call FncNew()
End Sub

'=======================================================================================================
'Description : 결의전표 생성내역 팝업 
'=======================================================================================================
Function OpenPopupTempGL()

	Dim arrRet
	Dim arrParam(1)	
	Dim iCalledAspName
	Dim IntRetCD

	If IsOpenPop = True Then Exit Function
	
	iCalledAspName = AskPRAspName("A5130RA1")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "A5130RA1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	arrParam(0) = Trim(frm1.txtTGlNo.value)	'결의전표번호 
	arrParam(1) = ""							'Reference번호 
	
	If arrParam(0) = "" Then
		IntRetCD = DisplayMsgBox("970000","X" , "결의전표", "X") 	
		IsOpenPop = False
		Exit Function
	End If
	

	IsOpenPop = True
    arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		     
	'arrRet = window.showModalDialog("../../ComAsp/a5130ra1.asp", Array(window.parent, arrParam), _
	'	     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	
End Function
'=======================================================================================================
'Description : 회계전표 생성내역 팝업 
'=======================================================================================================
Function OpenPopupGL()

	Dim arrRet
	Dim arrParam(1)	
	Dim iCalledAspName
	Dim IntRetCD
	
	If IsOpenPop = True Then Exit Function
	
	iCalledAspName = AskPRAspName("A5120RA1")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "A5120RA1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	arrParam(0) = Trim(frm1.txtGlNo.value)	'회계전표번호 
	arrParam(1) = ""						'Reference번호 
	
	If arrParam(0) = "" Then
		IntRetCD = DisplayMsgBox("970000","X" , "회계전표", "X") 	
		IsOpenPop = False
		Exit Function
	End If

	IsOpenPop = True
    arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		     
'	arrRet = window.showModalDialog("../../ComAsp/a5120ra1.asp", Array(window.parent, arrParam), _
'		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	IsOpenPop = False
	
End Function
'======================================================================================================
Function ValidCheck()
	Dim iRow,tempRow
	Dim	amt_sum,locamt_sum
	Dim diff_locamt
	Dim max_amt
	Dim intRetCd

	ValidCheck = False
	

	amt_sum		= 0
	locamt_sum	= 0
	max_amt		= 0
	tempRow	= 0
	
	With Frm1
	
		' 취득금액와 출금내역의 합계금액이 일치하는지 확인	
		For iRow = 1 To frm1.vspdData.MaxRows
			.vspdData.Row = iRow
			.vspdData.Col = 0		

			If .vspdData.Text <> ggoSpread.DeleteFlag Then

				.vspdData.Col = C_Amt
				.vspdData.Row = iRow
				amt_sum = amt_sum + UNICDbl(.vspdData.text)
				
				IF UNICDbl(.vspdData.text) > max_amt Then
					tempRow = iRow
					max_amt = UNICDbl(.vspdData.text)
				End If
				
				.vspdData.Col = C_LocAmt
				.vspdData.Row = iRow
				locamt_sum = locamt_sum + UNICDbl(.vspdData.text)
			End If
		
		Next		 
	
		diff_locamt = locamt_sum - UNICDbl(.txtLocBuyAmt.text)
		
		IF UNICDbl(.txtBuyAmt.text) = amt_sum Then
			If diff_locamt <> 0 then
				.vspdData.Col = C_LocAmt		
				.vspdData.Row = tempRow
				.vspdData.text = UNIFormatNumber(UNICDbl(.vspdData.text)-diff_locamt,ggAmtOfMoney.Decpoint,-2,0,ggAmtOfMoney.RndPolicy,ggAmtOfMoney.RndUnit) 
			End If
		Else
			Call DisplayMsgBox("am0029","X","X","X")                  '☜ :유가증권 취득금액과 출금내역의 금액의 합이 일치하지 않습니다.
			exit function
		End If
		
		' 어음정보와 계좌정보가 유효한지 Check	
		For iRow = 1 To frm1.vspdData.MaxRows
		
			.vspdData.Col = C_NoteNo		
			.vspdData.Row = iRow

			IF .vspdData.text <> "" Then
				IntRetCD= CommonQueryRs(" Note_No "," F_Note "," NOTE_NO = " & FilterVar(.vspdData.text, "''", "S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) 
				If IntRetCD = False Then
					Call DisplayMsgBox("800054","X","X","X")		'☜ : 등록되지  않은 코드입니다 
					exit function
				End If
			End If
			
			.vspdData.Col = C_BankAcct		
			.vspdData.Row = iRow

			IF .vspdData.text <> "" Then
				IntRetCD= CommonQueryRs(" Bank_Acct_No "," B_BANK_ACCT "," Bank_Acct_NO = " & FilterVar(.vspdData.text, "''", "S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) 
				If IntRetCD = False Then
					Call DisplayMsgBox("800054","X","X","X")		'☜ : 등록되지 않은 코드입니다 
					exit function
				End If
			End If
		Next		

		
	End With

	ValidCheck = True
	
End Function


'======================================================================================================
Function btnSecuCodeOnClick()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function
	
'	If BtnPopupDisabled(Inobj) = False Then Exit Function

	IsOpenPop = True

	arrParam(0) = "유가증권"		    						' 팝업 명칭 
	arrParam(1) = "A_SECURITY"									' TABLE 명칭 
	arrParam(2) = frm1.txtSecuCode.value						' Code Condition
	arrParam(3) = "" 		            						' Name Cindition
	arrParam(4) = "(ISNULL(TEMP_GL_NO,'') <> '' OR ISNULL(GL_NO, '') <> '')"					
	arrParam(5) = "유가증권"

    arrField(0) = "SECURITY_CD"	     							' Field명(1)
    arrField(1) = "SECURITY_NM"									' Field명(0)


    arrHeader(0) = "유가증권코드"			    				' Header명(0)
    arrHeader(1) = "유가증권명"								' Header명(1)

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=470px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
        frm1.txtSecuCode.focus
		Exit Function
	Else
		Call SetSecuCode(arrRet)
	End If

End Function

'======================================================================================================
Function SetSecuCode(Byval arrRet)
    With frm1
        .txtSecuCode.focus
        .txtSecuCode.value  = arrRet(0)
        .txtSecuNm.value    = arrRet(1)
    End With
End Function

'======================================================================================================
'	Name : btnSecuTypeOnClick()
'	Description : Major PopUp
'======================================================================================================
Function btnSecuTypeOnClick(Inobj)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function
	
	If BtnPopupDisabled(Inobj) = False Then Exit Function	

	IsOpenPop = True

	arrParam(0) = "유가증권종류"		    			' 팝업 명칭	
	arrParam(1) = "B_MINOR"								' TABLE 명칭 
	arrParam(2) = frm1.txtSecuType.value				' Code Condition	
	arrParam(3) = "" 		            				' Name Cindition
	arrParam(4) = " MAJOR_CD = " & FilterVar("A1031", "''", "S") & "  "				' Where Condition	
	arrParam(5) = "유가증권종류"

    arrField(0) = "MINOR_CD"	     					' Field명(1)
    arrField(1) = "MINOR_NM"							' Field명(0)	


    arrHeader(0) = "유가증권종류코드"			    	' Header명(0)	
    arrHeader(1) = "유가증권종류명"					' Header명(1)	

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=470px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
        frm1.txtSecuType.focus
		Exit Function
	Else
		Call SetSecuType(arrRet)
	End If

End Function

'======================================================================================================
'	Name : SetSecuType()
'	Description : Item Popup에서 Return되는 값 setting
'======================================================================================================
Function SetSecuType(Byval arrRet)
	Dim  var1
	Dim  Pos
	
    With frm1
        .txtSecuType.focus
        .txtSecuType.value  = arrRet(0)
        .txtSecuTypeNm.value    = arrRet(1)

        var1 = .txtSecuTypeNm.value 
		Pos =  instr(1,var1,"주식")

		If Pos = 0 Then
			Call ggoOper.SetReqAttr(frm1.txtExpireDt,"N")
			Call ggoOper.SetReqAttr(frm1.txtCalRate,"N")
		Else
			Call ggoOper.SetReqAttr(frm1.txtExpireDt,"Q")
			Call ggoOper.SetReqAttr(frm1.txtCalRate,"Q")
		End If	
    End With

    
    lgBlnFlgChgValue = True 
End Function

'======================================================================================================
'	Name : btnTradeCurOnClick()
'	Description : Major PopUp
'======================================================================================================
Function btnTradeCurOnClick(Inobj)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function
	
	If BtnPopupDisabled(Inobj) = False Then Exit Function	

	IsOpenPop = True

	arrParam(0) = "거래통화"		    					'팝업 명칭 
	arrParam(1) = "B_CURRENCY"								' TABLE 명칭 
	arrParam(2) = frm1.txtTradeCur.value					' Code Condition
	arrParam(3) = "" 		            					' Name Cindition
	arrParam(4) = ""										' Where Condition
	arrParam(5) = "거래통화"

    arrField(0) = "CURRENCY"	     						' Field명(1)
    arrField(1) = "CURRENCY_DESC"							' Field명(0)


    arrHeader(0) = "통화코드"			    				' Header명(0)
    arrHeader(1) = "통화명"								' Header명(1)

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=470px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
        frm1.txtTradeCur.focus
		Exit Function
	Else
		Call SetTradeCur(arrRet)
	End If

End Function

'======================================================================================================
'	Name : SetTradeCur()
'	Description : Item Popup에서 Return되는 값 setting
'======================================================================================================
Function SetTradeCur(Byval arrRet)
    With frm1
        .txtTradeCur.focus
        .txtTradeCur.value  = arrRet(0)
        .txtTradeCurNm.value    = arrRet(1)
    End With
	Call txtDocCur_OnChangeASP
    Call txtTradeCur_OnChange()

End Function

'======================================================================================================
'	Name : OpenDept
'	Description : 
'======================================================================================================


Function OpenDept(Byval strCode, iWhere)
	Dim arrRet
	Dim arrParam(5)
	Dim iCalledAspName
	Dim IntRetCD
	
	If IsOpenPop = True Then Exit Function

	iCalledAspName = AskPRAspName("DEPTPOPUPDT3")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "DEPTPOPUPDT3", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	IsOpenPop = True

	arrParam(0) = strCode		            '  Code Condition
   	arrParam(1) = frm1.txtBillDt.Text
	arrParam(2) = lgUsrIntCd                            ' 자료권한 Condition  

	' T : protected F: 필수 
	If lgIntFlgMode = Parent.OPMD_UMODE then
		arrParam(3) = "T"									' 결의일자 상태 Condition  
	Else
		arrParam(3) = "F"									' 결의일자 상태 Condition  
	End If
	
	arrParam(4) = iWhere
	arrParam(5) = Trim(frm1.txtDept1Area.value)
	
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam), _
		     "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		     
		
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Select Case iWhere
		     Case 1
               frm1.txtDept1.focus
		     Case 2
               frm1.txtDept2.focus
	    End Select
		Exit Function
	Else
		Call SetDept(arrRet, iWhere)
	End If

End Function

'------------------------------------------  SetDept()  ------------------------------------------------
'	Name : SetDept()
'	Description : Dept Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------
Function SetDept(Byval arrRet, Byval iWhere)

	With frm1
		Select Case iWhere
		Case 1
			.txtDept1.focus
			.txtDept1.value = arrRet(0)
			.txtDept1Nm.value = arrRet(1)
			.txtDept1Area.value = arrRet(2)
			.txtDept1OrgId.value = arrRet(3)
			.txtInternalCd1.value = arrRet(4)
			.txtBillDt.text = arrRet(5)
			.txtDept2.value = ""
			.txtDept2Nm.value = ""
			call txtDept1_OnChange()  
			.txtDept1.focus
		Case 2
			.txtDept2.focus
			.txtDept2.value = arrRet(0)
			.txtDept2Nm.value = arrRet(1)
			.txtInternalCd2.value = arrRet(4)
			.txtBillDt.text = arrRet(5)
			call txtDept2_OnChange()  
		End Select
	End With
End Function       
'------------------------------------------  OpenBp()  ---------------------------------------
'	Name : OpenBp()
'	Description : OpenAp Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function OpenBp(Byval strCode, byval iWhere)
	Dim arrRet
	Dim arrParam(5)

	If IsOpenPop = True Then Exit Function
	
	if UCase(frm1.txtCust1.className) = "PROTECTED" Then Exit Function
	IsOpenPop = True

	arrParam(0) = strCode								'  Code Condition
   	arrParam(1) = ""							' 채권과 연계(거래처 유무)
	arrParam(2) = ""								'FrDt
	arrParam(3) = ""								'ToDt
	arrParam(4) = "T"							'B :매출 S: 매입 T: 전체 
	arrParam(5) = ""									'SUP :공급처 PAYTO: 지급처 SOL:주문처 PAYER :수금처 INV:세금계산 	
	
	arrRet = window.showModalDialog("../../Module/ar/BpPopup.asp", Array(window.Parent,arrParam), _
		"dialogWidth=780px; dialogHeight=450px; : Yes; help: No; resizable: No; status: No;")
		
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
        frm1.txtCust1.focus
		Exit Function
	Else
		Call SetCust1(arrRet)
		lgBlnFlgChgValue = True
	End If

End Function
'======================================================================================================
'	Name : btnCust1OnClick()
'	Description : Major PopUp
'======================================================================================================
Function btnCust1OnClick(Inobj)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function
	
	If BtnPopupDisabled(Inobj) = False Then Exit Function	

	IsOpenPop = True

	arrParam(0) = "발행거래처"		    	' 팝업 명칭 

    arrParam(1) = "B_BIZ_PARTNER"
    arrParam(2) = Trim(frm1.txtCust1.value)
    arrParam(3) = ""
    arrParam(4) = ""
    arrParam(5) = "발행거래처코드"

    arrField(0) = "BP_CD"
    arrField(1) = "BP_NM"

    arrHeader(0) = "거래처코드"
    arrHeader(1) = "거래처명"

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=470px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
        frm1.txtCust1.focus
		Exit Function
	Else
		Call SetCust1(arrRet)
	End If

End Function

'======================================================================================================
'	Name : SetCust1()
'	Description : Item Popup에서 Return되는 값 setting
'======================================================================================================
Function SetCust1(Byval arrRet)
    With frm1
        .txtCust1.focus
        .txtCust1.value  = arrRet(0)
        .txtCust1Nm.value    = arrRet(1)
    End With
    Call txtCust1_Change()
End Function

'======================================================================================================
'	Name : btnCust2OnClick()
'	Description : Major PopUp
'======================================================================================================
Function btnCust2OnClick(Inobj)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function
	
	If BtnPopupDisabled(Inobj) = False Then Exit Function	

	IsOpenPop = True

	arrParam(0) = "주식보관회사"		    	' 팝업 명칭 

    arrParam(1) = "B_BIZ_PARTNER"
    arrParam(2) = Trim(frm1.txtCust2.value)
    arrParam(3) = ""
    arrParam(4) = ""
    arrParam(5) = "주식보관회사코드"

    arrField(0) = "BP_CD"
    arrField(1) = "BP_NM"

    arrHeader(0) = "주식보관회사코드"
    arrHeader(1) = "주식보관회사명"

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=470px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
        frm1.txtCust2.focus
		Exit Function
	Else
		Call SetCust2(arrRet)
	End If

End Function

'======================================================================================================
'	Name : SetCust1()
'	Description : Item Popup에서 Return되는 값 setting
'======================================================================================================
Function SetCust2(Byval arrRet)
    With frm1
        .txtCust2.focus
        .txtCust2.value  = arrRet(0)
        .txtCust2Nm.value    = arrRet(1)
    End With
    lgBlnFlgChgValue = True 
End Function

Function BtnPopupDisabled(Inobj) 

	If UCase(Inobj.className) = UCase("protected") Then 
		IsOpenPop = False
		BtnPopupDisabled = False
	Else
		BtnPopupDisabled = True
	End If

End Function

'======================================================================================================
'	Name : OpenAcctPopup()
'	Description : Major PopUp
'======================================================================================================
Function OpenAcctPopup(iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function
	
	Select Case iWhere
		Case 0
			If BtnPopupDisabled(frm1.txtAcct1) = False Then Exit Function
		Case 1
			If BtnPopupDisabled(frm1.txtAcct2) = False Then Exit Function
	End select		
	


	IsOpenPop = True
	
	Select Case iwhere

	Case 0			
		arrParam(0) = "이자수익계정"															
		arrParam(1) = "A_ACCT"							
		arrParam(2) = Trim(frm1.txtAcct1.value)
		arrParam(3) = ""
		arrParam(4) = "DEL_FG <> " & FilterVar("Y", "''", "S") & " "
		arrParam(5) = "미수수익계정"								
	
		arrField(0) = "ACCT_CD"									
		arrField(1) = "ACCT_NM"									
		
		arrHeader(0) = "미수수익계정코드"								
		arrHeader(1) = "미수수익계정명"	
	Case 1
		arrParam(0) = "이자수익계정"										
		arrParam(1) = "A_ACCT"							
		arrParam(2) = Trim(frm1.txtAcct2.value)
		arrParam(3) = ""
		arrParam(4) = "DEL_FG <> " & FilterVar("Y", "''", "S") & " "
		arrParam(5) = "이자수익계정"								
	
		arrField(0) = "ACCT_CD"									
		arrField(1) = "ACCT_NM"									
		
		arrHeader(0) = "이자수익계정코드"								
		arrHeader(1) = "이자수익계정명"					
	End Select				
	
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Select Case iwhere
			Case 0
				frm1.txtAcct1.focus
			Case 1
				frm1.txtAcct2.focus
		End Select
		Exit Function
	Else
		Call SetAcctPopup(arrRet, iWhere)
	End If	
End Function

'======================================================================================================
'	Name : SetAcctPopup()
'	Description : Item Popup에서 Return되는 값 setting
'======================================================================================================
Function SetAcctPopup(Byval arrRet, Byval iwhere)

	With frm1
		Select Case iwhere
			Case 0
				.txtAcct1.focus
				.txtAcct1.value = arrRet(0)
				.txtAcctNm1.value = arrRet(1)
			Case 1
				.txtAcct2.focus
				.txtAcct2.value = arrRet(0)
				.txtAcctNm2.value = arrRet(1)
		End Select
	End With
	lgBlnFlgChgValue = True	

End Function

'========================================================================================================
' Name : txtInsureAcct_Onchange
' Desc : developer describe this line
'========================================================================================================
Function txtAcct1_Onchange()
    Dim IntRetCd
	
		If  frm1.txtAcct1.value = "" Then
			frm1.txtAcct1.value = ""
			frm1.txtAcctNM1.value=""
			frm1.txtAcct1.focus
		Else
		    IntRetCD= CommonQueryRs(" ACCT_CD,ACCT_NM "," A_ACCT "," ACCT_CD = " & FilterVar(frm1.txtAcct1.value, "''", "S") & " and DEL_FG <> " & FilterVar("Y", "''", "S") & " " ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) 		    
				If IntRetCD=False And Trim(frm1.txtAcct1.value)<>"" Then
				    Call DisplayMsgBox("800054","X","X","X")                         '☜ : 등록되지 않은 코드입니다.
				    frm1.txtAcct1.value=""
				    frm1.txttxtAcctNm1.value=""
				    frm1.txtAcct1.focus
				    Set gActiveElement = document.activeElement  
				Else
				    frm1.txtAcct1.value=Trim(Replace(lgF0,Chr(11),""))
				    frm1.txtAcctNm1.value=Trim(Replace(lgF1,Chr(11),""))
				End If
		End if
	lgBlnFlgChgValue = True   
End Function 

'========================================================================================================
' Name : txtInsureAcct_Onchange
' Desc : developer describe this line
'========================================================================================================
Function txtAcct2_Onchange()
    Dim IntRetCd
	
		If  frm1.txtAcct2.value = "" Then
			frm1.txtAcct2.value = ""
			frm1.txtAcctNm2.value=""
			frm1.txtAcct2.focus
		Else
		    IntRetCD= CommonQueryRs(" ACCT_CD,ACCT_NM "," A_ACCT "," ACCT_CD = " & FilterVar(frm1.txtAcct1.value, "''", "S") & " and DEL_FG <> " & FilterVar("Y", "''", "S") & " " ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) 		    
				If IntRetCD=False And Trim(frm1.txtAcct2.value)<>"" Then
				    Call DisplayMsgBox("800054","X","X","X")                         '☜ : 등록되지 않은 코드입니다.
				    frm1.txtAcct2.value=""
				    frm1.txttxtAcctNm2.value=""
				    frm1.txtAcct2.focus
				    Set gActiveElement = document.activeElement  
				Else
				    frm1.txtAcct2.value=Trim(Replace(lgF0,Chr(11),""))
				    frm1.txtAcctNm2.value=Trim(Replace(lgF1,Chr(11),""))
				End If
		End if
	lgBlnFlgChgValue = True   
End Function 
'=======================================================================================================
'	Name : OpenBankAcct()
'	Description : Bank Account No PopUp
'=======================================================================================================
Function OpenBankAcct(byVal strCode,byVal strCard)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim iWhere
	
	IF strCard = "" Then Exit Function
	If IsOpenPop = True  Then Exit Function

	IsOpenPop = True
	
	arrParam(0) = "예적금코드 팝업"	' 팝업 명칭 
	arrParam(1) = "B_BANK A, B_BANK_ACCT B, F_DPST C"			' TABLE 명칭 
	arrParam(2) = ""						' Code Condition
	arrParam(3) = strCode							' Name Cindition
	arrParam(4) = "A.BANK_CD = B.BANK_CD "	
	arrParam(4) = arrParam(4) & "AND B.BANK_CD = C.BANK_CD " 					' Where Condition 
	arrParam(4) = arrParam(4) & "AND B.BANK_ACCT_NO = C.BANK_ACCT_NO " 
	arrParam(4) = arrParam(4) & "AND (C.DPST_FG = " & FilterVar("SV", "''", "S") & "  OR C.DPST_FG = " & FilterVar("ET", "''", "S") & " ) " 
	arrParam(4) = arrParam(4) & "AND C.DPST_TYPE IN (" & FilterVar("D1", "''", "S") & " ," & FilterVar("D2", "''", "S") & " ," & FilterVar("D3", "''", "S") & " ) " 

	arrParam(5) = "은행코드"				' 조건필드의 라벨 명칭 
	
	arrField(0) = "A.BANK_CD"					' Field명(1)	
	'arrField(1) = "A.BANK_NM"					' Field명(1)
	arrField(1) = "B.BANK_ACCT_NO"				' Field명(2)
   
   	arrHeader(0) = "은행코드"						' Header명(1)
'	arrHeader(1) = "은행명"						' Header명(1)
	arrHeader(1) = "예적금코드"

    
            
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	iWhere = "BankAcct"
	If arrRet(0) = "" Then
		Call GridsetFocus(iWhere)
		Exit Function
	Else
		Call SetReturnVal(arrRet,iWhere)
	End If	

End Function

'=======================================================================================================
'	Name : OpenNoteNo()
'	Description : Note No PopUp
'=======================================================================================================
Function OpenNoteNo(byVal strCode, byVal strCard)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim iWhere
	IF strCard = "" Then Exit Function
	If IsOpenPop = True  Then Exit Function

	IsOpenPop = True	

	IF UCase(strCard) = "CP"	Then
		arrParam(0) = "지불구매카드 팝업"	
		arrParam(1) = "F_NOTE A,B_BANK B,B_BIZ_PARTNER C"				
		arrParam(2) = strCode
		arrParam(3) = ""
		
		arrParam(4) = "A.NOTE_STS = " & FilterVar("BG", "''", "S") & "  AND A.NOTE_FG = " & FilterVar("CP", "''", "S") & "  AND A.BP_CD = C.BP_CD AND A.BANK_CD = B.BANK_CD"				
		arrParam(5) = "지불구매카드번호"			
		
	    arrField(0) = "A.NOTE_NO"		
	    arrField(1) = "F2" & parent.gColSep & "A.NOTE_AMT"
	    arrField(2) = "C.BP_NM"	    
	    arrField(3) = "DD" & parent.gColSep & "A.ISSUE_DT"
	    arrField(4) = "DD" & parent.gColSep & "A.DUE_DT"	
	    arrField(5) = "B.BANK_NM"	        
	    
	    arrHeader(0) = "지불구매카드번호"
	    arrHeader(1) = "금액"        		
		arrHeader(2) = "거래처"        		        	
		arrHeader(3) = "발행일"        		        
		arrHeader(4) = "만기일"        		        
		arrHeader(5) = "은행"       

	
	Elseif  UCase(strCard) = "NP"	Then
		arrParam(0) = "지급어음번호 팝업"	
		arrParam(1) = "F_NOTE A,B_BANK B,B_BIZ_PARTNER C"				
		arrParam(2) = strCode
		arrParam(3) = ""
		
		arrParam(4) = "A.NOTE_STS = " & FilterVar("BG", "''", "S") & "  AND A.NOTE_FG = " & FilterVar("D3", "''", "S") & "  AND A.BP_CD = C.BP_CD AND A.BANK_CD = B.BANK_CD"				
		arrParam(5) = "지급어음번호"			
		
	    arrField(0) = "A.NOTE_NO"		
	    arrField(1) = "F2" & parent.gColSep & "A.NOTE_AMT"
	    arrField(2) = "C.BP_NM"	    
	    arrField(3) = "DD" & parent.gColSep & "A.ISSUE_DT"
	    arrField(4) = "DD" & parent.gColSep & "A.DUE_DT"	
	    arrField(5) = "B.BANK_NM"	        
	    
	    arrHeader(0) = "지급어음번호"
	    arrHeader(1) = "어음금액"        		
		arrHeader(2) = "거래처"        		        	
		arrHeader(3) = "발행일"        		        
		arrHeader(4) = "만기일"        		        
		arrHeader(5) = "은행" 

	Else
		arrParam(0) = "배서어음번호 팝업"	
		arrParam(1) = "F_NOTE A,B_BANK B,B_BIZ_PARTNER C"				
		arrParam(2) = strCode
		arrParam(3) = ""
	
		arrParam(4) = "A.NOTE_STS = " & FilterVar("ED", "''", "S") & "  AND A.NOTE_FG = " & FilterVar("D1", "''", "S") & "  AND A.BP_CD = C.BP_CD AND A.BANK_CD = B.BANK_CD"				
		arrParam(5) = "배서어음번호"			
	
		arrField(0) = "A.NOTE_NO"		
		arrField(1) = "F2" & parent.gColSep & "A.NOTE_AMT"
		arrField(2) = "C.BP_NM"	    
		arrField(3) = "DD" & parent.gColSep & "A.ISSUE_DT"
		arrField(4) = "DD" & parent.gColSep & "A.DUE_DT"	
		arrField(5) = "B.BANK_NM"	        
    
		arrHeader(0) = "배서어음번호"
		arrHeader(1) = "어음금액"        		
		arrHeader(2) = "거래처"        		        	
		arrHeader(3) = "발행일"        		        
		arrHeader(4) = "만기일"        		        
		arrHeader(5) = "은행"    

		      
 	End If	
	
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
	"dialogWidth=700px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False

	iWhere = "NoteNo"
	If arrRet(0) = "" Then
		Call GridSetFocus(iWhere)
		Exit Function
	Else
		Call SetReturnVal(arrRet,iWhere)
	End If
End Function
'=======================================================================================================
Function GridSetFocus(Byval iWhere)
		With frm1
			Select Case iWhere
			Case "BankAcct"
				Call SetActiveCell(.vspdData,C_BankAcct,.vspdData.ActiveRow ,"M","X","X")
			Case "NoteNo"
				Call SetActiveCell(.vspdData,C_NoteNo,.vspdData.ActiveRow ,"M","X","X")
			End Select
		End With
End Function

 '------------------------------------------  SetReturnVal()  --------------------------------------------------
'	Name : SetReturnVal()
'	Description : Plant Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetReturnVal(byval arrRet,byval iWhere)
	With frm1	
		Select case iWhere
			case "BankAcct"
				.vspddata.Col			= C_BankCd
				.vspddata.Text			= arrRet(0)
				.vspddata.Col			= C_BankNm
				.vspddata.Text			= ""
				.vspddata.Col			= C_BankAcct
				.vspddata.Text			= arrRet(1)
			case "NoteNo"
				.vspddata.Col			= C_NoteNo
				.vspddata.Text			= arrRet(0)
				.vspddata.Col			= C_Amt
				.vspddata.Text			= arrRet(1)
				.vspddata.Col			= C_LocAmt
				.vspddata.Text			= arrRet(1)
		End select
		lgBlnFlgChgValue = True
		Call GridsetFocus(iWhere)
	End With
End Function

'========================================================================================
' Function Name : MaxSpreadVal
' Function Desc : 
'========================================================================================

Function MaxSpreadVal(byval Row)
  Dim iRows
  Dim MaxValue  
  Dim tmpVal

	MAxValue = 0

		with frm1
			For iRows = 1 to  .vspddata.MaxRows
				.vspddata.row = iRows
		        .vspddata.col = C_Seq

				if .vspddata.Text = "" then 
					tmpVal = 0
				else
  					tmpVal = UNICDbl(.vspddata.text)
				end if

				if tmpval > MaxValue   then
					MaxValue = UNICDbl(tmpVal)
				end if
			Next

			MaxValue = MaxValue + 1

			.vspddata.row = row
			.vspddata.col = C_Seq
			.vspddata.text = MaxValue
		end with
		
end Function
 '==========================================  2.3.1 Tab Click 처리  =================================================
'	기능: Tab Click
'	설명: Tab Click시 필요한 기능을 수행한다.
'=================================================================================================================== 
 '----------------  ClickTab1(): Header Tab처리 부분 (Header Tab이 있는 경우만 사용)  ---------------------------- 
Function ClickTab1()

	If gSelframeFlg = TAB1 Then Exit Function
	Call changeTabs(TAB1)	 '~~~ 첫번째 Tab 
	gSelframeFlg = TAB1
	
	Call SetToolbar("1111100000111111")                                                     '☆: Developer must customize

	
End Function

Function ClickTab2()
	Dim IntRetCD
	
	If gSelframeFlg = TAB2 Then Exit Function

	Call changeTabs(TAB2)	 '~~~ 첫번째 Tab 
	gSelframeFlg = TAB2

	Call SetToolbar("1100111100111111")        

End Function




'========================================================================================================
Sub txtBillDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtBillDt.Action = 7
		Call deptCheck()
		Call SetFocusToDocument("M")
		Frm1.txtBillDt.Focus
		lgBlnFlgChgValue = True 
	End If
End Sub

'========================================================================================================
Sub txtPubDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtPubDt.Action = 7
		Call SetFocusToDocument("M")
		Frm1.txtPubDt.Focus
		lgBlnFlgChgValue = True 
	End If
End Sub

'========================================================================================================
Sub txtExpireDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtExpireDt.Action = 7
		Call SetFocusToDocument("M")
		Frm1.txtExpireDt.Focus
		lgBlnFlgChgValue = True 
	End If
End Sub

'========================================================================================================
Sub txtInDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtInDt.Action = 7
		Call txtTradeCur_OnChange()
		Call SetFocusToDocument("M")
		Frm1.txtInDt.Focus
		lgBlnFlgChgValue = True 
	End If
End Sub

Sub txtInDt_Change()
    If frm1.txtInDt.text <> "" AND frm1.txtCust1.value <> "" Then
        Call txtTradeCur_OnChange()
    End If
End Sub

'========================================================================================================
Sub txtOutDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtOutDt.Action = 7
		Call SetFocusToDocument("M")
		Frm1.txtOutDt.Focus
		lgBlnFlgChgValue = True 
	End If
End Sub

Sub txtXchRate_Change()
	Dim iRows
	Dim iVal
	Dim iVal2
	Dim iVal3

  	
	ggoSpread.Source = frm1.vspdData
	
	with frm1
		For iRows = 1 to  .vspddata.MaxRows
			.vspddata.row = iRows
	        .vspddata.col = C_Amt

				iVal = UNICDbl(frm1.txtXchRate.text) * UNICDbl(.vspddata.text)
				
				.vspddata.col = C_LocAmt
				.vspddata.text = UNIFormatNumber(iVal,ggAmtOfMoney.DecPoint,-2,0,ggAmtOfMoney.RndPolicy,ggAmtOfMoney.RndUnit)
				
				ggoSpread.UpdateRow iRows
			Next
		end with

		iVal2 = UNICDbl(frm1.txtXchRate.text) * UNICDbl(frm1.txtBuyAmt.text)
		iVal3 = UNICDbl(frm1.txtXchRate.text) * UNICDbl(frm1.txtPriceAmt.text)
		frm1.txtLocBuyAmt.text = UNIFormatNumber(iVal2,ggAmtOfMoney.DecPoint,-2,0,ggAmtOfMoney.RndPolicy,ggAmtOfMoney.RndUnit)
		frm1.txtLocPriceAmt.text = UNIFormatNumber(iVal3,ggAmtOfMoney.DecPoint,-2,0,ggAmtOfMoney.RndPolicy,ggAmtOfMoney.RndUnit)

		lgBlnFlgChgValue = True   
End Sub


Sub txtBuyAmt_Change()

	Dim iVal1
	iVal1 = UNICDbl(frm1.txtXchRate.text) * UNICDbl(frm1.txtBuyAmt.text)
	frm1.txtLocBuyAmt.text = UNIFormatNumber(iVal1,ggAmtOfMoney.DecPoint,-2,0,ggAmtOfMoney.RndPolicy,ggAmtOfMoney.RndUnit)
'	frm1.txtLocBuyAmt.text = 0
	lgBlnFlgChgValue = True   
End Sub

Sub txtPriceAmt_Change()
	Dim iVal1
	iVal1 = UNICDbl(frm1.txtXchRate.text) * UNICDbl(frm1.txtPriceAmt.text)
	frm1.txtLocPriceAmt.text = UNIFormatNumber(iVal1,ggAmtOfMoney.DecPoint,-2,0,ggAmtOfMoney.RndPolicy,ggAmtOfMoney.RndUnit)
'    frm1.txtLocPriceAmt.text = 0
    lgBlnFlgChgValue = True   
End Sub

Sub txtLocBuyAmt_Change()
    lgBlnFlgChgValue = True   
End Sub
Sub txtLocPriceAmt_Change()
    lgBlnFlgChgValue = True   
End Sub

Sub txtCnt_Change()
    lgBlnFlgChgValue = True   
End Sub

Sub selCalYn_Change()
    If frm1.selCalYn.value = "Y" Then
        Call ggoOper.SetReqAttr(frm1.txtCalRate,"N")
        Call ggoOper.SetReqAttr(frm1.selEndYn,"N")
        Call ggoOper.SetReqAttr(frm1.txtExpireDt,"N")
    Else
        Call ggoOper.SetReqAttr(frm1.txtCalRate,"Q")
        Call ggoOper.SetReqAttr(frm1.selEndYn,"Q")
        Call ggoOper.SetReqAttr(frm1.txtExpireDt,"Q")
    End If
        lgBlnFlgChgValue = True 
End Sub

Sub selComYn_Change()
    If frm1.selComYn.value = "Y" Then
        Call ggoOper.SetReqAttr(frm1.txtOutDt,"N")
    Else
        Call ggoOper.SetReqAttr(frm1.txtOutDt,"Q")
    End If
        lgBlnFlgChgValue = True 
End Sub

Sub txtSecuType_Change()
    Dim var1
    Dim Pos

    If Trim(frm1.txtSecuType.value) <> "" Then
		Call CommonQueryRs(" MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("A1031", "''", "S") & "  AND MINOR_CD =  " & FilterVar(frm1.txtSecuType.value , "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
		var1 = Replace(lgF0, Chr(11), "")

		If var1 = "" Then
		    Call DisplayMsgBox("970000","X",frm1.txtSecuType.alt,"X")
		    frm1.txtSecuType.value = ""
		    frm1.txtSecuTypeNm.value = ""
		    frm1.txtSecuType.focus
		    Set gActiveElement = document.activeElement
		Else
		    frm1.txtSecuTypeNm.value = var1
			Pos =  instr(1,var1,"주식")

			If Pos = 0 Then

				Call ggoOper.SetReqAttr(frm1.txtExpireDt,"N")
				Call ggoOper.SetReqAttr(frm1.txtCalRate,"N") 
			Else
				Call ggoOper.SetReqAttr(frm1.txtExpireDt,"Q")
				Call ggoOper.SetReqAttr(frm1.txtCalRate,"Q") 
			End If	
		End If
	Else
		frm1.txtSecuType.value = ""
		frm1.txtSecuTypeNm.value = ""
	End If	
        lgBlnFlgChgValue = True 
End Sub


'========================================================================================================
' Name : txtDept1_Onchange
' Desc : developer describe this line
'========================================================================================================
Function txtDept1_Onchange()
    Dim strSelect
	Dim strFrom
	Dim strWhere 	
    Dim IntRetCD 
	Dim arrVal1
	Dim arrVal2
	Dim ii
	Dim jj
	
	With Frm1
	
    If  .txtDept1.value = "" Then
		.txtDept1.value = ""
		.txtDept1Nm.value=""
		.txtDept2.value=""
		.txtDept2Nm.value=""
		.txtDept1Area.value=""
		.txtDept1OrgId.value = ""
		Call ggoOper.SetReqAttr(.txtDept2, "Q")
		.txtDept1.focus
		Set gActiveElement = document.activeElement
		lgBlnFlgChgValue = True 
		Exit Function
    End If
    
    
    If Trim(.txtBillDt.Text = "") Then    
		Exit Function
    End If
    lgBlnFlgChgValue = True

	'----------------------------------------------------------------------------------------
	strSelect	=			 " a.dept_cd,a.dept_nm, a.org_change_id, a.internal_cd, b.biz_area_cd "    		
	strFrom		=			 " b_acct_dept a, b_cost_center b "		
	strWhere	= " a.cost_cd = b.cost_cd " 	 
	strWhere	= strWhere & " and a.dept_Cd = " & FilterVar(LTrim(RTrim(.txtDept1.value)), "''", "S")
	strWhere	= strWhere & " and a.org_change_id = (select distinct org_change_id "			
	strWhere	= strWhere & " from b_acct_dept where org_change_dt = ( select max(org_change_dt)"
	strWhere	= strWhere & " from b_acct_dept where org_change_dt <= " & FilterVar(UniConvDateToYYYYMMDD(.txtBillDt.Text, parent.gDateFormat,""), "''", "S") & "))"			
		
	
		
	If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) = False Then
	
		IntRetCD = DisplayMsgBox("124600","X","X","X")  
		.txtDept1.value = ""
		.txtDept1Nm.value = ""
		.txtDept1Area.value = ""
		.txtDept1OrgId.value = ""
		.txtInternalCd1.value = ""
		.txtInternalCd2.value = ""
		.txtDept2.value = ""
		.txtDept2Nm.value = ""
		Call ggoOper.SetReqAttr(.txtDept2, "Q")
		.txtDept1.focus
		Set gActiveElement = document.activeElement  
							    
	Else 
		
		arrVal1 = Split(lgF2By2, Chr(11) & Chr(12))			
		jj = Ubound(arrVal1,1)
					
		For ii = 0 to jj - 1
			arrVal2 = Split(arrVal1(ii), chr(11))			
			frm1.txtDept1Nm.value=Trim(arrVal2(2))
		    frm1.txtDept1OrgId.value =Trim(arrVal2(3))
		    frm1.txtInternalCd1.value =Trim(arrVal2(4))
		    frm1.txtDept1Area.value =Trim(arrVal2(5))
		    frm1.txtDept2.value=""
		    frm1.txtDept2Nm.value=""
		    frm1.txtDept2.focus
		    Call ggoOper.SetReqAttr(frm1.txtDept2, "N")
		Next	
	End If
	
	End With
	lgBlnFlgChgValue = True   
End Function 
'========================================================================================================
' Desc : developer describe this line
'========================================================================================================
Function txtDept2_Onchange()
   Dim strSelect
	Dim strFrom
	Dim strWhere 	
    Dim IntRetCD 
	Dim arrVal1
	Dim arrVal2
	Dim ii
	Dim jj
	
	With Frm1
	
    If  .txtDept2.value = "" Then
		.txtDept2.value=""
		.txtDept2Nm.value=""
		.txtDept2.focus
		lgBlnFlgChgValue = True 
		Exit Function
    End If
    
    
    If Trim(.txtBillDt.Text = "") Then    
		Exit Function
    End If
    lgBlnFlgChgValue = True

	strSelect	=			 " a.dept_cd,a.dept_nm "    		
	strFrom		=			 " b_acct_dept a, b_cost_center b "		
	strWhere	= " a.cost_cd = b.cost_cd " 	 
	strWhere	= strWhere & " and a.dept_Cd = " & FilterVar(LTrim(RTrim(.txtDept2.value)), "''", "S")
	strWhere	= strWhere & " and b.biz_area_cd = " & FilterVar(LTrim(RTrim(.txtDept1Area.value)), "''", "S")
	strWhere	= strWhere & " and a.org_change_id = (select distinct org_change_id "			
	strWhere	= strWhere & " from b_acct_dept where org_change_dt = ( select max(org_change_dt)"
	strWhere	= strWhere & " from b_acct_dept where org_change_dt <= " & FilterVar(UniConvDateToYYYYMMDD(.txtBillDt.Text, parent.gDateFormat,""), "''", "S") & "))"			
		
	If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) = False Then
	
		IntRetCD = DisplayMsgBox("124600","X","X","X")  
		.txtDept2.value = ""
		.txtDept2Nm.value = ""
		.txtDept2.focus
		Set gActiveElement = document.activeElement  
							    
	Else 
		
		arrVal1 = Split(lgF2By2, Chr(11) & Chr(12))			
		jj = Ubound(arrVal1,1)
					
		For ii = 0 to jj - 1
			arrVal2 = Split(arrVal1(ii), chr(11))			
			frm1.txtDept2.value=Trim(arrVal2(1))
		    frm1.txtDept2Nm.value=Trim(arrVal2(2))
		    frm1.txtCust2.focus
		Next	
	End If
	
	End With
	lgBlnFlgChgValue = True   
End Function 



Sub txtCust1_Change()
    Dim var1, var2
    If Trim(frm1.txtCust1.value) <> "" Then
		Call CommonQueryRs(" BP_NM, CURRENCY "," B_BIZ_PARTNER "," BP_CD =  " & FilterVar(frm1.txtCust1.value , "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
		var1 = Replace(lgF0, Chr(11), "")
		var2 = Replace(lgF1, Chr(11), "")

		If var1 = "" Then
		    Call DisplayMsgBox("126100","X","X","X")
		    frm1.txtCust1.value = ""
		    frm1.txtCust1Nm.value = ""
		    frm1.txtCust1.focus
		    Set gActiveElement = document.activeElement
		Else
		    frm1.txtCust1Nm.value = var1
		    'If Trim(var2) = "" Then
		       ' Call CommonQueryRs(" LOC_CUR "," B_COMPANY ","",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
		        'var2 = Replace(lgF0, Chr(11), "")
		    'End If

		    frm1.txtTradeCur.value = Trim(var2)
		    Call txtTradeCur_OnChange()
		End If
	Else
		frm1.txtCust1.value = ""
		frm1.txtCust1Nm.value = ""	
	End If	
        lgBlnFlgChgValue = True 
End Sub

Sub txtCust2_Change()
    Dim var1
    If Trim(frm1.txtCust2.value) <> "" Then
		Call CommonQueryRs(" BP_NM "," B_BIZ_PARTNER "," BP_CD =  " & FilterVar(frm1.txtCust2.value , "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
		var1 = Replace(lgF0, Chr(11), "")

		If var1 = "" Then
		    Call DisplayMsgBox("126100","X","X","X")
		    frm1.txtCust2.value = ""
		    frm1.txtCust2Nm.value = ""
		    frm1.txtCust2.focus
		    Set gActiveElement = document.activeElement
		Else
		    frm1.txtCust2Nm.value = var1
		End If
	Else
		frm1.txtCust2.value = ""
		frm1.txtCust2Nm.value = ""	
	End If	
        lgBlnFlgChgValue = True 
End Sub

Sub txtTradeCur_OnChange()
    Dim var1, var2
	Dim FrDt
	Dim IntRetCD, strSelect, strFrom, strWhere
	
	
	If Trim(frm1.txtInDt.Text) = "" Then
		frm1.txtInDt.Text = UniConvDateAToB("<%=StartDate%>",parent.gServerDateFormat,parent.gDateFormat)
		Exit Sub
	End IF

	FrDt = UniConvDateToYYYYMMDD(frm1.txtInDt.Text,parent.gDateFormat,"")   '//parent.UNIConvDate(frm1.txtInDt.Text)
	Call txtDocCur_OnChangeASP()
	If Trim(frm1.txtTradeCur.value) <> "" Then
	    Call CommonQueryRs(" CURRENCY_DESC "," B_CURRENCY "," CURRENCY =  " & FilterVar(frm1.txtTradeCur.value, "''", "S") & "  ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

	    var1 = Replace(lgF0, Chr(11), "")

	    If var1 = "" Then
	'        Call DisplayMsgBox("110100","X","X","X")
	'        frm1.txtTradeCur.value = ""
	'        frm1.txtTradeCurNm.value = ""
	'        frm1.txtTradeCur.focus
	        frm1.txtXchRate.text = 1
	        
	        Set gActiveElement = document.activeElement
	    Else
	        
	        strSelect = strSelect & "isnull(case t.loc_cur when " & FilterVar(UCase(frm1.txtTradeCur.value), "''", "S") & " Then 1 "
			strSelect = strSelect & " 	Else  Case t.xch_rate_fg " 
			strSelect = strSelect & "        When " & FilterVar("M", "''", "S") & "  Then ( SELECT isnull(STD_RATE,0) "
			strSelect = strSelect & "				FROM    b_monthly_exchange_rate (nolock) "
			strSelect = strSelect & "				WHERE  apprl_yrmnth 	= CONVERT (varchar(06), " & FilterVar(FrDt,"" & FilterVar("29991231", "''", "S") & " ", "S") & ", 112) "
			strSelect = strSelect & "               and from_currency	= " & FilterVar(UCase(frm1.txtTradeCur.value), "''", "S") 
			strSelect = strSelect & "               and to_currency	= t.loc_cur ) "
			strSelect = strSelect & "         Else (	SELECT  isnull(STD_RATE,0) "
			strSelect = strSelect & "               FROM  b_daily_exchange_rate (nolock) "
			strSelect = strSelect & "               WHERE apprl_dt      = " & FilterVar(FrDt,"" & FilterVar("29991231", "''", "S") & " ", "S")
			strSelect = strSelect & "               and from_currency   = " & FilterVar(UCase(frm1.txtTradeCur.value), "''", "S") 
			strSelect = strSelect & "               and to_currency     = t.loc_cur ) "
			strSelect = strSelect & "         End  "
			strSelect = strSelect & "End,0) as xch_rate "
			strFrom  = " (SELECT isnull(XCH_RATE_FG,'') as xch_rate_fg, loc_cur  from  b_company) t "  
	       	strWhere = ""
			IntRetCD= CommonQueryRs(strSelect,strFrom,strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) 			
	       
	       If IntRetCD = True Then
				var2 = Trim(Replace(lgF0,Chr(11),""))
		    End If
	        frm1.txtTradeCurNm.value = Trim(var1)
	        If Trim(var2) = "" Then
	            var2 = 1
	        End If
	        frm1.txtXchRate.text = Trim(var2)
	    End If

	Else
		frm1.txtTradeCur.value = ""
		frm1.txtTradeCurNm.value = ""  
	End If	
    lgBlnFlgChgValue = True 
End Sub

Function txtBillDt_onblur()
	Call deptCheck()
End Function


Function txtInDt_onblur()
	Call txtTradeCur_OnChange()
End Function



'//////////////////////////check 작업중/////////////////////////////
Function deptCheck()

	Dim strSelect
	Dim strFrom
	Dim strWhere 	
    Dim IntRetCD 
	
	With Frm1
		If  .txtDept1.value = "" Then
			.txtDept1.value = ""
			.txtDept1Nm.value = ""
			.txtDept1Area.value = ""
			.txtDept1OrgId.value = ""
			.txtInternalCd1.value = ""
			.txtInternalCd2.value = ""
			.txtDept2.value = ""
			.txtDept2Nm.value = ""
			Call ggoOper.SetReqAttr(.txtDept2, "Q")
			.txtDept1.focus
			Set gActiveElement = document.activeElement
			lgBlnFlgChgValue = True 
			Exit Function
		End If
    
    
		If Trim(.txtBillDt.Text = "") Then    
			Exit Function
		End If
		

		'----------------------------------------------------------------------------------------
		strSelect	=			 " distinct org_change_id "    		
		strFrom		=			 " b_acct_dept "		
		strWhere	=			 " org_change_id = (select distinct org_change_id "			
		strWhere	= strWhere & " from b_acct_dept where org_change_dt = ( select max(org_change_dt)"
		strWhere	= strWhere & " from b_acct_dept where org_change_dt <= " & FilterVar(UniConvDateToYYYYMMDD(.txtBillDt.Text, parent.gDateFormat,""), "''", "S") & "))"			
			
	
		IntRetCD= CommonQueryRs(strSelect,strFrom,strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) 			
	
		If IntRetCD = False or Trim(Replace(lgF0,Chr(11),"")) <> .txtDept1OrgId.value  Then
			.txtDept1.value = ""
			.txtDept1Nm.value = ""
			.txtDept1Area.value = ""
			.txtDept1OrgId.value = ""
			.txtInternalCd1.value = ""
			.txtInternalCd2.value = ""
			.txtDept2.value = ""
			.txtDept2Nm.value = ""
			Call ggoOper.SetReqAttr(.txtDept2, "Q")
			lgBlnFlgChgValue = True
			Exit Function
		End If	
	End With
	
End Function



</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->
</HEAD>



<BODY SCROLL="No" TABINDEX="-1">
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
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="ClickTab1()">
							<TR>
								<td background="../../image/table/seltab_up_bg.gif"><img src="../../image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>유가증권정보등록</font></td>
								<td background="../../image/table/seltab_up_bg.gif" align="right"><img src="../../image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="ClickTab2()">
							<TR>
								<td background="../../image/table/tab_up_bg.gif"><img src="../../image/table/tab_up_left.gif" width="9" height="23"></td>
								<td background="../../image/table/tab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>출금내역</font></td>
								<td background="../../image/table/tab_up_bg.gif" align="right"><img src="../../image/table/tab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* ALIGN=RIGHT><A HREF="VBSCRIPT:OpenPopupTempGL()">결의전표</A>&nbsp;|&nbsp;<A HREF="VBSCRIPT:OpenPopupGL()">회계전표</A></TD>					
					<TD WIDTH=*>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=*>
		<TD WIDTH=100% CLASS="Tab11">
			<DIV ID="TabDiv" STYLE="FlOAT: left; HEIGHT:100%; OVERFLOW:auto; WIDTH:100%;" SCROLL=no>
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS=TD5 NOWRAP>유가증권코드</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtSecuCode" SIZE="20" MAXLENGTH="20" TAG="12XXXU" ALT="유가증권 코드" ><IMG SRC="../../image/btnPopup.gif" NAME="btnSecuCode" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:btnSecuCodeOnClick()">&nbsp;<INPUT NAME="txtSecuNm" TYPE=TEXT SIZE="30" MAXLENGTH="30"   TAG="24XXXU" ALT="증권명칭"></TD>
                                    <TD CLASS=TDT NOWRAP></TD>
									<TD CLASS=TD6 NOWRAP></TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD WIDTH=100% VALIGN=TOP>
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD CLASS=TD5 NOWRAP>유가증권코드</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtSecuCode1" TYPE=TEXT SIZE="20" MAXLENGTH="20"   TAG="23XXXU" ALT="증권코드"></TD>
								<TD CLASS=TD5 NOWRAP>증권명칭</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtSecuNm1" TYPE=TEXT SIZE="30" MAXLENGTH="30"   TAG="23XXXU" ALT="증권명칭"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>증권종류</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtSecuType" TYPE=TEXT SIZE=10  MAXLENGTH="20" TAG="23XXXU" ALT="증권종류" OnChange="txtSecuType_Change()"><IMG SRC="../../image/btnPopup.gif" NAME="btnSecuType" align=top TYPE="BUTTON" ONCLICK ="vbscript:btnSecuTypeOnClick(frm1.txtSecuType)">&nbsp;<INPUT TYPE=TEXT NAME="txtSecuTypeNm"  SIZE="20" MAXLENGTH="50" TAG="24"></TD>
								<TD CLASS=TD5 NOWRAP>전표일자</TD>
								<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/a5958ma1_txtBillDt_txtBillDt.js'></script></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>발행거래처</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtCust1" TYPE=TEXT SIZE=10  MAXLENGTH="10" TAG="23XXXU" ALT="발행거래처" OnChange="txtCust1_Change()"><IMG SRC="../../image/btnPopup.gif" NAME="btnCust1" align=top TYPE="BUTTON" ONCLICK ="vbscript:Call OpenBp(frm1.txtCust1.value, 1)">&nbsp;<INPUT TYPE=TEXT NAME="txtCust1Nm"  SIZE="20" MAXLENGTH="50" TAG="24"></TD>
								<TD CLASS=TD5 NOWRAP>매수</TD>
								<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/a5958ma1_txtCnt_txtCnt.js'></script></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>거래통화</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtTradeCur" SIZE="10"  MAXLENGTH="3" TAG="23XXXU" ALT="거래통화" ><IMG SRC="../../image/btnPopup.gif" NAME="btnTradeCur" align=top TYPE="BUTTON" ONCLICK ="vbscript:btnTradeCurOnClick(frm1.txtTradeCur)">&nbsp;<INPUT TYPE=TEXT NAME="txtTradeCurNm" SIZE="20" MAXLENGTH="50" TAG="24"></TD>
								<TD CLASS=TD5 NOWRAP>환율</TD>
								<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/a5958ma1_txtXchRate_txtXchRate.js'></script></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>취득금액</TD>
								<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/a5958ma1_txtBuyAmt_txtBuyAmt.js'></script></TD>
								<TD CLASS=TD5 NOWRAP>취득금액(자국)</TD>
								<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/a5958ma1_txtLocBuyAmt_txtLocBuyAmt.js'></script></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>액면금액</TD>
								<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/a5958ma1_txtPriceAmt_txtPriceAmt.js'></script></TD>
								<TD CLASS=TD5 NOWRAP>액면금액(자국)</TD>
								<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/a5958ma1_txtLocPriceAmt_txtLocPriceAmt.js'></script></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>발의부서</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtDept1" TYPE=TEXT SIZE=10  MAXLENGTH="10" TAG="23XXXU" ALT="발의부서"><IMG SRC="../../image/btnPopup.gif" NAME="btnDept1" align=top TYPE="BUTTON" ONCLICK ="vbscript:Call OpenDept(frm1.txtDept1.value,1)">&nbsp;<INPUT TYPE=TEXT NAME="txtDept1Nm"  SIZE="20" MAXLENGTH="50" TAG="24"></TD>
								<TD CLASS=TD5 NOWRAP>귀속부서</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtDept2" TYPE=TEXT SIZE=10  MAXLENGTH="10" TAG="23XXXU" ALT="귀속부서"><IMG SRC="../../image/btnPopup.gif" NAME="btnDept2" align=top TYPE="BUTTON" ONCLICK ="vbscript:Call OpenDept(frm1.txtDept2.value,2)">&nbsp;<INPUT TYPE=TEXT NAME="txtDept2Nm"  SIZE="20" MAXLENGTH="50" TAG="24"></TD>
								
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>주식보관회사</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtCust2" TYPE=TEXT SIZE=10  MAXLENGTH="20" TAG="25XXXU" ALT="주식보관회사" OnChange="txtCust2_Change()"><IMG SRC="../../image/btnPopup.gif" NAME="btnCust2" align=top TYPE="BUTTON" ONCLICK ="vbscript:btnCust2OnClick(frm1.txtCust2)" >&nbsp;<INPUT TYPE=TEXT NAME="txtCust2Nm"  SIZE="20" MAXLENGTH="50" TAG="24"></TD>
								<TD CLASS=TD5 NOWRAP>발행일자</TD>
								<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/a5958ma1_txtPubDt_txtPubDt.js'></script></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>미수수익계정</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtAcct1" SIZE="10"  MAXLENGTH="20" TAG="23XXXU" ALT="미수수익계정"><IMG SRC="../../image/btnPopup.gif" NAME="btnAcct1" align=top TYPE="BUTTON" ONCLICK ="vbscript:Call OpenAcctPopup(0)">&nbsp;<INPUT TYPE=TEXT NAME="txtAcctNm1" SIZE="20" MAXLENGTH="50" TAG="24"></TD>
								<TD CLASS=TD5 NOWRAP>이자수익계정</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtAcct2" SIZE="10"  MAXLENGTH="20" TAG="23XXXU" ALT="이자수익계정"><IMG SRC="../../image/btnPopup.gif" NAME="btnAcct2" align=top TYPE="BUTTON" ONCLICK ="vbscript:Call OpenAcctPopup(1)">&nbsp;<INPUT TYPE=TEXT NAME="txtAcctNm2" SIZE="20" MAXLENGTH="50" TAG="24"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>만기여부</TD>
								<TD CLASS=TD6 NOWRAP><SELECT NAME="selComYn" TAG="23XXXU" ALT="완료여부" OnChange="selComYn_Change()"><OPTION VALUE="Y">Y</OPTION><OPTION VALUE="N">N</OPTION></SELECT></TD>
								<TD CLASS=TD5 NOWRAP>이자율</TD>
								<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/a5958ma1_txtCalRate_txtCalRate.js'></script></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>이자계산</TD>
								<TD CLASS=TD6 NOWRAP><SELECT NAME="selCalYn" TAG="23XXXU" ALT="이자계산" OnChange="selCalYn_Change()"><OPTION VALUE="Y">계산</OPTION><OPTION VALUE="N">미계산</OPTION></SELECT></TD>
								<TD CLASS=TD5 NOWRAP>만기일자</TD>
								<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/a5958ma1_txtExpireDt_txtExpireDt.js'></script></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>양편구분</TD>
								<TD CLASS=TD6 NOWRAP><SELECT NAME="selEndYn" TAG="23XXXU" ALT="양편구분"></SELECT></TD>
								<TD CLASS=TD5 NOWRAP>취득일자</TD>
								<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/a5958ma1_txtInDt_txtInDt.js'></script></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>관리번호</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtRefNo" TYPE=TEXT SIZE="20" MAXLENGTH="20"   TAG="25XXXU" ALT="관리번호"></TD>
								<TD CLASS=TD5 NOWRAP>처분일자</TD>
								<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/a5958ma1_txtOutDt_txtOutDt.js'></script></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>결의전표번호</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtTGlNo" TYPE=TEXT SIZE="20" MAXLENGTH="20"   TAG="24XXXU" ALT="결의전표번호"></TD>
								<TD CLASS=TD5 NOWRAP>회계전표번호</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtGlNo" TYPE=TEXT SIZE="20" MAXLENGTH="20"   TAG="24XXXU" ALT="회계전표번호"></TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</DIV>
		<!-- 두번째 탭 내용  -->
		<DIV ID="TabDiv" STYLE="DISPLAY: none;" SCROLL=no>
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD HEIGHT="100%" NOWRAP>
						<script language =javascript src='./js/a5958ma1_fpSpread2_vspdData.js'></script>
					</TD>
				</TR>
			</TABLE>
		</DIV>	
		</TD>
	</TR>
	<TR >
		<TD <%=HEIGHT_TYPE_01%>></TD>
	</TR>
	
	<TR >
		<TD WIDTH=100% HEIGHT=<%=BIZSIZE%>> <IFRAME NAME="MyBizASP" SRC= "../../blank.htm" WIDTH=100% HEIGHT=<%=BIZSIZE%> FRAMEBORDER=0 SCROLLING=No noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" TABINDEX="-1"></TEXTAREA><% '업무처리ASP로 넘기기 위한 변수를 담고 있는 Tag들 %>

<INPUT TYPE=HIDDEN NAME="txtMode" TAG="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtKeyStream"    TAG="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtPrevNext" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtDept1Area" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtDept1OrgId" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtInternalCd1" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtInternalCd2" tag="24" TABINDEX="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm" TABINDEX="-1"></IFRAME>
</DIV>
</BODY>
</HTML>
