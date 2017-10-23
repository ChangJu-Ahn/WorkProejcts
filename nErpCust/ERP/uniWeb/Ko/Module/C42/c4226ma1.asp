<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% session.CodePage=949 %>

<!-- '======================================================================================================
'*  1. Module Name          : Cost
'*  2. Function Name        :평균단가/재고평가내역 
'*  3. Program ID           : c4226ma1.asp
'*  4. Program Name         : 평균단가/재고평가내역 
'*  5. Program Desc         : 평균단가/재고평가내역 
'*  6. Modified date(First) : 2005-11-25
'*  7. Modified date(Last)  : 2005-11-25
'*  8. Modifier (First)     : HJO
'*  9. Modifier (Last)      : 
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'======================================================================================================= -->

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
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<!--
========================================================================================================
=                          3.2 Style Sheet
======================================================================================================== -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		

<!--
========================================================================================================
=                          3.3 Client Side Script
======================================================================================================== -->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliRdsQuery.vbs">	</SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>
Option Explicit

'########################################################################################################
'#                       4.  Data Declaration Part
'########################################################################################################

'========================================================================================================
'=                       4.1 External ASP File
'========================================================================================================

Const BIZ_PGM_ID = "c4226mb1.asp"                               'Biz Logic ASP

Dim iDBSYSDate
Dim iStrFromDt
Dim iStrToDt

iDBSYSDate = "<%=GetSvrDate%>"
iStrFromDt = UNIConvDateAtoB(iDBSYSDate, parent.gServerDateFormat, parent.gDateFormat)	
iStrToDt= UNIDateAdd("m", -1,iStrFromDt, parent.gServerDateFormat)
'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	


'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim lgQueryFlag
Dim IsOpenPop          
Dim lgCurrGrid
Dim lgCopyVersion
Dim lgErrRow, lgErrCol
'Dim lgStrPrevKey2
Dim lgSTime		' -- 디버깅 타임체크 
'========================================================================================================
'                        5.1 Common Method-1
'========================================================================================================
'========================================================================================================
'--spread A
Dim C_PlantCD	
Dim C_AcctNM		
Dim C_ItemCd
Dim C_ItemNM
Dim C_TrackingNO
Dim C_BasQTY
Dim C_BasAMT
Dim C_BasPrice
Dim C_RcptQTY
Dim C_RcptStdAmt
Dim C_RcptDiffAmt
Dim C_TransQtyPrc
Dim C_TransStdAmtPrc

Dim C_TransDebitDiffAmtPrc
Dim C_AvgPrc
Dim C_InvQty
Dim C_invAmt
Dim C_RcptDiffAmt2
Dim C_TransDiffAmt
Dim C_AInvAmt
Dim C_DiffAmt

 


'==========================================  2.2.7 InitSpreadPosVariables()  =============================
' Function Name : InitSpreadPosVariables
' Function Desc : This method Assigns Sequential Number to spread sheet column 
'=========================================================================================================
Sub InitSpreadPosVariables(byVal pvSpd)

		If pvSpd="" or pvSpd="A" Then
			C_PlantCD		 =1
			C_AcctNM		 =2
			C_ItemCd		 =3
			C_ItemNM	 =4
			C_TrackingNO =5
			C_BasQTY=6
			C_BasAMT	     =7
			C_BasPrice      =8
			C_RcptQTY     =9
			C_RcptStdAmt		     =10		
			C_RcptDiffAmt            =11
			C_TransQtyPrc            =12
			C_TransStdAmtPrc            =13
			
			C_TransDebitDiffAmtPrc            =14
			C_AvgPrc            =15
			C_InvQty            =16
			C_invAmt            =17
			C_RcptDiffAmt2            =18
			C_TransDiffAmt            =19
			C_AInvAmt            =20
			C_DiffAmt            =21
		End If
	
	
End Sub

'========================================================================================================
' Name : InitVariables()	
' Desc : Initialize value
'======================================================================================================== 
Sub InitVariables()

    lgIntFlgMode = Parent.OPMD_CMODE
    lgBlnFlgChgValue = False 
    
    lgStrPrevKey = ""
    'lgStrPrevKey2 = ""		

End Sub

'========================================================================================================
' Name : SetDefaultVal()	
' Desc : Set default value
'========================================================================================================
Sub SetDefaultVal()
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 

	frm1.txtYYYYMM.Text =UniConvDateAToB(iStrFromDt, parent.gServerDateFormat, parent.gDateFormat)
	
	Call ggoOper.FormatDate(frm1.txtYYYYMM, parent.gDateFormat, 2)

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
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
	<% Call loadInfTB19029A("Q", "C", "NOCOOKIE", "MA") %>
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
   
   '------ Developer Coding part (Start ) -------------------------------------------------------------- 
       
   '------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub        


'========================================================================================================
' Name : InitSpreadSheet
' Desc : This method initializes spread sheet column property
'========================================================================================================
Sub InitSpreadSheet(byVal pvSpd)
	Dim i, ret
	
	Call InitSpreadPosVariables(pvSpd)
    'Call AppendNumberPlace("6","3","0")
    
    If pvSpd = "" or pvSpd ="A" Then 
		With frm1.vspdData
		
		ggoSpread.Source = frm1.vspdData
		'ggoSpread.Spreadinit "V20021106", , Parent.gAllowDragDropSpread
		ggoSpread.Spreadinit "V20021106", , ""
			
		.MaxCols = C_DiffAmt+1
				'헤더를 2줄로    
		.ColHeaderRows = 2
		
		Call GetSpreadColumnPos("A")
		.ReDraw = False
		
		ggoSpread.SSSetEdit		C_PlantCD,	"공장"	, 5,,,,1	
		ggoSpread.SSSetEdit		C_AcctNM,	"품목계정"	, 5,,,,1	
		ggoSpread.SSSetEdit		C_ItemCd,	"품목"	, 7,,,,1
		ggoSpread.SSSetEdit		C_ItemNM,	"품목명",7
		ggoSpread.SSSetEdit		C_TrackingNO,	"Tracking No.",7,,,,1
		ggoSpread.SSSetFloat	C_BasQTY,	"기초수량"	, 10,		Parent.ggQtyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	C_BasAMT,	"기초금액"	, 10,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	C_BasPrice,	"기초단가"	, 10,		Parent.ggUnitCostNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		
		ggoSpread.SSSetFloat	C_RcptQTY,	"평균단가"	, 10,		Parent.ggUnitCostNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	C_RcptStdAmt,	""	, 10,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec		
		ggoSpread.SSSetFloat	C_RcptDiffAmt,	""	, 10,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	C_TransQtyPrc,	""	, 10,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	C_TransStdAmtPrc,	""	, 10,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	C_TransDebitDiffAmtPrc,	""	, 10,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	C_AvgPrc,	""	, 10,		Parent.ggUnitCostNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		
		ggoSpread.SSSetFloat	C_InvQty,	"재고평가"	, 10,		Parent.ggQtyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	C_invAmt,	""	, 10,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	C_RcptDiffAmt2,	""	, 10,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	C_TransDiffAmt,	""	, 10,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	C_AInvAmt,	""	, 10,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	C_DiffAmt,	""	, 10,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		
				

		Call ggoSpread.SSSetColHidden(.MaxCols,.MaxCols,True)
		'ggoSpread.SSSetSplit2(C_ItemCd) 
		

		'header 1
		for i = C_PlantCD to C_BasPrice
			ret = .AddCellSpan(i,-1000,1,2)
		next
		ret = .AddCellSpan(C_RcptQTY,-1000,7,1)
		ret = .AddCellSpan(C_InvQty,-1000,6,1)
		
		.Row=-999
		.Col = C_RcptQTY :.Text ="입고수량"
		.Col = C_RcptStdAmt :.Text ="입고금액"
		.Col = C_RcptDiffAmt:.Text ="입고차이금액"
		.Col = C_TransQtyPrc :.Text ="이동입고수량"
		.Col = C_TransStdAmtPrc :.Text ="이동입고금액"
		.Col = C_TransDebitDiffAmtPrc :.Text ="이동입고차이금액"
		.Col = C_AvgPrc:.Text ="평균단가"
		
		.Col = C_InvQty :.Text ="기말재고수량"
		.Col = C_invAmt:.Text ="평가전기말재고금액"
		.Col = C_RcptDiffAmt2:.Text ="입고차이금액"
		.Col = C_TransDiffAmt:.Text ="이동차이금액"
		.Col = C_AInvAmt:.Text ="평가후기말재고금액"
		.Col = C_DiffAmt:.Text ="재고차이금액"
		
		
		.rowheight(-1000) =15	' 높이 재지정 
		.rowheight(-999) = 20	' 높이 재지정 
		
		
		.ReDraw = True		
		End With
		Call SetSpreadLock("A")
	End If
	
	

	'ggoSpread.SpreadLockWithOddEvenRowColor()		

End Sub


'======================================================================================================
' Name : SetSpreadLock
' Desc : This method set color and protect cell in spread sheet
'======================================================================================================
Sub SetSpreadLock(byVal pvSpd)
	If pvSpd="A" Then
		ggoSpread.Source = frm1.vspdData    
		ggoSpread.SpreadLockWithOddEvenRowColor()
	End If
	
End Sub

'======================================================================================================
' Name : SetSpreadColor
' Desc : This method set color and protect cell by cell in spread sheet
'======================================================================================================
Sub SetSpreadColor(ByVal pvStartRow,ByVal pvEndRow)
    With frm1
    
    ggoSpread.Source = frm1.vspdData
    .vspdData.ReDraw = False

	
    .vspdData.ReDraw = True
    
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
			C_PlantCD		     = iCurColumnPos(1)	
			C_AcctNM		 = iCurColumnPos(2)	
			C_ItemCd		 = iCurColumnPos(3)
			C_ItemNM = iCurColumnPos(4)
			C_TrackingNO =iCurColumnPos(5)
			C_BasQTY = iCurColumnPos(6)
			C_BasAMT = iCurColumnPos(7)
			C_BasPrice = iCurColumnPos(8)
			C_RcptQTY = iCurColumnPos(9)
			C_RcptStdAmt = iCurColumnPos(10)
			
			C_RcptDiffAmt = iCurColumnPos(11)
			C_TransQtyPrc = iCurColumnPos(12)
			C_TransStdAmtPrc = iCurColumnPos(13)
			
			C_TransDebitDiffAmtPrc= iCurColumnPos(14)
			C_AvgPrc= iCurColumnPos(15)
			C_InvQty= iCurColumnPos(16)
			C_invAmt= iCurColumnPos(17)
			C_RcptDiffAmt2= iCurColumnPos(18)
			C_TransDiffAmt= iCurColumnPos(19)
			C_AInvAmt= iCurColumnPos(20)
			C_DiffAmt= iCurColumnPos(21)
		
 		End Select  	
 	
End Sub
'******************************************  2.4 POP-UP 처리함수  ****************************************
'	기능: 기준 POP-UP
'   설명: 기준 POP-UP에 관한 Open은 Include한다. 
'	      하나의 ASP에서 Popup이 중복되면 하나는 링크시켜 사용하고 나머지는 재정의하여 사용한다.
'********************************************************************************************************* 
Function OpenPopUp(Byval iWhere)
	Dim arrRet, sTmp
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	With frm1
	
	Select Case iWhere
		Case 0
			arrParam(0) = "공장 팝업"
			arrParam(1) = "dbo.B_PLANT"	
			arrParam(2) = Trim(.txtPLANT_CD.value)
			arrParam(3) = ""
			arrParam(4) = ""
			arrParam(5) = "공장" 

			arrField(0) = "PLANT_CD"	
			arrField(1) = "PLANT_NM"
			arrField(2) = ""		
    
			arrHeader(0) = "공장"	
			arrHeader(1) = "공장명"
			arrHeader(2) = ""
			
		Case 1
			arrParam(0) = "품목계정 팝업"
			arrParam(1) = "dbo.B_MINOR"	
			arrParam(2) = Trim(.txtITEM_ACCT.value)
			arrParam(3) = ""
			arrParam(4) = "MAJOR_CD =" & FilterVar("P1001", "''", "S")
			arrParam(5) = "품목계정" 

			arrField(0) ="ED10" & Parent.gColSep & "MINOR_CD"
			arrField(1) ="ED30" & Parent.gColSep & "MINOR_NM"		
			arrField(2) = ""	
    
			arrHeader(0) = "품목계정"
			arrHeader(1) = "품목계정명"
			arrHeader(2) = "C/C Level"	

		Case 2
			arrParam(0) = "품목 팝업"
			arrParam(1) = "dbo.B_ITEM"	
			arrParam(2) = Trim(.txtITEM_CD.value)
			arrParam(3) = ""	
			arrParam(4) = ""
			arrParam(5) = "품목" 

			arrField(0) = "ED20" & Parent.gColSep &"ITEM_CD"	
			arrField(1) = "ED30" & Parent.gColSep &"ITEM_NM"
			arrField(2) = ""		
    
			arrHeader(0) = "품목"	
			arrHeader(1) = "품목명"
			arrHeader(2) = ""
		

	End Select
    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetPopUp(arrRet, iWhere)
	End If	

	End With
End Function


Function SetPopUp(Byval arrRet, Byval iWhere)
	Dim sTmp
	
	With frm1
		Select Case iWhere		
			Case 0
				.txtPLANT_CD.value		= arrRet(0)
				.txtPLANT_NM.value		= arrRet(1)
				
			Case 1
				.txtITEM_ACCT.value		= arrRet(0)
				.txtITEM_ACCT_NM.value	= arrRet(1)
				
			Case 2
				.txtITEM_CD.value		= arrRet(0)
				.txtITEM_NM.value		= arrRet(1)
						
		End Select
		lgBlnFlgChgValue = True
	End With
	
End Function

'------------------------------------------  OpenCalType()  ----------------------------------------------
'	Name :InitComboBox
'	Description : 
'--------------------------------------------------------------------------------------------------------- 
Sub InitComboBox
    
End Sub

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
	
    Call LoadInfTB19029
    Call ggoOper.LockField(Document, "N") 
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
	Call ggoOper.FormatDate(frm1.txtYYYYMM, parent.gDateFormat,2)

    Call InitSpreadSheet("")

    Call InitVariables

    Call SetDefaultVal
    Call SetToolbar("110000000001111")	
   
    If parent.gPlant <> "" Then
		frm1.txtPlant_Cd.value = UCase(parent.gPlant)
		frm1.txtPlant_Nm.value = parent.gPlantNm
		frm1.txtITem_Acct.focus 		
	Else
		frm1.txtPlant_Cd.focus 		
	End If


    
   	Set gActiveElement = document.activeElement			    
    
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
'========================================================================================================
'                        5.4 User-defined Method
'========================================================================================================


'========================================================================================================
'========================================================================================================
'                        5.5 Tag Event
'========================================================================================================

Sub txtYYYYMM_KeyPress(KeyAscii)
	If KeyAscii = 13 Then
		Call FncQuery
	End If
End Sub

Sub txtYYYYMM_DblClick(Button)
    If Button = 1 Then
        frm1.txtYYYYMM.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtYYYYMM.Focus
    End If
End Sub


Sub txtPlant_cd_onKeyPress()
	If window.event.keyCode = 13 Then
		Call FncQuery
	End If
End Sub


Sub txtPlant_cd_onChange()
	If frm1.txtPlant_cd.value ="" then frm1.txtPlant_nm.value=""
End Sub


'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    'ggoSpread.Source = frm1.vspdData
    'Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

End Sub

'========================================================================================================
'   Event Name : vspdData_DblClick
'   Event Desc : 
'========================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)				
    
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
    	
End Sub


'==========================================================================================
'   Event Desc : Spread Split 상태코드 
'==========================================================================================
Sub vspdData_MouseDown(Button, Shift, X, Y)
End Sub
'========================================================================================================
'   Event Name : vspdData_ScriptLeaveCell
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_ScriptLeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
    If Row <> NewRow And NewRow > 0 Then
        With frm1
            .vspdData.Row = NewRow
            .vspdData.Col = C_AcctNM

        End With

        frm1.vspddata.Col = 0
		'lgStrPrevKey2=""

		'Call DbDtlQuery(NewRow)
    End If
End Sub
'========================================================================================================
'   Event Name : vspdData_Change
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
'==========================================================================================
'   Event Name : vspdData_Click
'   Event Desc :
'==========================================================================================

Sub vspdData_Click(ByVal Col , ByVal Row )
	
	Call SetPopupMenuItemInf("0000111111")         '화면별 설정		
	
    gMouseClickStatus = "SPC"	'Split 상태코드 
	Set gActiveSpdSheet = frm1.vspdData         

	'lgStrPrevKey2=""
End Sub

'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
End Sub


Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
	If OldLeft <> NewLeft Then	Exit Sub
	    
	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) And lgStrPrevKey <> "" Then
		If CheckRunningBizProcess = True Then Exit Sub
	    
		Call DisableToolBar(Parent.TBC_QUERY)
		If DBQuery = False Then
			Call RestoreToolBar()
			Exit Sub
		End If
	End if    

End Sub



'========================================================================================================
' Name : FncQuery
' Desc : This function is called from MainQuery in Common.vbs
'========================================================================================================
Function FncQuery() 
    Dim IntRetCD , sStartDt, sEndDt
    
    FncQuery = False
    
    Err.Clear
    
    If Not chkField(Document, "1") Then
       Exit Function
    End If
    
    sStartDt= Replace(frm1.txtYYYYMM.text, parent.gComDateType, "")

    IF ChkKeyField()=False Then Exit Function 
    
    Call ggoOper.ClearField(Document, "2")
    ggoSpread.Source = frm1.vspdData
	ggoSpread.ClearSpreadData

    Call InitVariables 	

    IF DbQuery = False Then
		Exit Function
	END IF
       
    FncQuery = True		
    
End Function

'========================================================================================================
' Name : FncNew
' Desc : This function is called from MainNew in Common.vbs
'========================================================================================================
Function FncNew() 
    Dim IntRetCD 
    
  
End Function

'========================================================================================================
' Name : FncSave
' Desc : This function is called from MainSave in Common.vbs
'========================================================================================================
Function FncSave() 
    
    FncSave = True      
    
End Function


'========================================================================================================
' Name : FncCopy
' Desc : This function is called from MainCopy in Common.vbs
' Keep : Make sure to clear primary key area
'========================================================================================================
Function FncCopy() 


End Function


Function FncCancel() 
    Dim lDelRows

	lgBlnFlgChgValue = True
End Function


'========================================================================================================
' Name : FncInsertRow
' Desc : This function is called by MainInsertRow in Common.vbs
'========================================================================================================
Function FncInsertRow(ByVal pvRowCnt)
    Dim IntRetCD, iSeqNo, iSubSeqNo
    Dim imRow
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

End Function


Function FncDeleteRow() 
    Dim lDelRows
	
	lgBlnFlgChgValue = True
End Function
Function FncPrint()
    Call parent.FncPrint() 
End Function

Function FncPrev() 
End Function

Function FncNext() 
End Function

Function FncExcel() 
    Call parent.FncExport(Parent.C_MULTI)
End Function

Function FncFind() 
    Call parent.FncFind(Parent.C_MULTI, False) 
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
    Call InitSpreadSheet(gActiveSpdSheet.id)      
    'Call InitComboBox
	Call ggoSpread.ReOrderingSpreadData()
	Call InitData()
End Sub

'========================================================================================================
' Name : FncExit
' Desc : This function is called by MainExit in Common.vbs
'========================================================================================================
Function FncExit()
Dim IntRetCD
	
	FncExit = False
	
    ggoSpread.Source = frm1.vspdData	
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO,"x","x")
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    FncExit = True
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
	Dim strVal

    DbQuery = False
    IF LayerShowHide(1) = False Then
		Exit Function
	END IF

    Err.Clear	
    
    Dim sStartDt, sEndDt, sYear, sMon, sDay
    
    With frm1
		Call parent.ExtractDateFromSuper(.txtYYYYMM.Text, parent.gDateFormat,sYear,sMon,sDay)
		
		sStartDt= sYear&sMon
		

		If lgIntFlgMode = parent.OPMD_UMODE Then
			strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
			strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
			strVal = strVal & "&txtYYYYMM=" & Trim(.hYYYYMM.value)		
			strVal = strVal & "&txtPlant_cd=" & Trim(.hPlant_cd.value)
			strVal = strVal & "&txtItem_Acct=" & Trim(.hItem_Acct.value)
			strVal = strVal & "&txtItem_cd=" & Trim(.hItem_cd.value)

		Else
			strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
			strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
			strVal = strVal & "&txtYYYYMM=" & sStartDt		
			strVal = strVal & "&txtPlant_cd=" & Trim(.txtPlant_cd.value)
			strVal = strVal & "&txtItem_Acct=" & Trim(.txtItem_Acct.value)
			strVal = strVal & "&txtItem_cd=" & Trim(.txtItem_cd.value)

		End If


		Call RunMyBizASP(MyBizASP, strVal)
   
    End With
    
    DbQuery = True

End Function

'========================================================================================================
' Name : DbQueryOk
' Desc : Called by MB Area when query operation is successful
'========================================================================================================
Function DbQueryOk()	
	lgIntFlgMode = parent.OPMD_UMODE	
	Frm1.vspdData.Focus
    Set gActiveElement = document.ActiveElement   

End Function

'========================================================================================
' Function Name : SetQuerySpreadColor
' Function Desc : 소계 및 총계 색상변경 
'========================================================================================
Sub SetQuerySpreadColor(byVal arrStr)

	Dim arrRow, arrCol, iRow
	Dim iLoopCnt, i
	Dim ret, iCnt, strRowI

	With frm1.vspdData
	.ReDraw = False
	
	arrRow = Split(arrStr, Parent.gRowSep)
	
	iLoopCnt = UBound(arrRow, 1)

	For i = 0 to iLoopCnt -1
		arrCol = Split(arrRow(i), Parent.gColSep)
	
		.Col = -1
		.Row = CDbl(arrCol(2))	' -- 행 
	
		Select Case arrCol(0)
			Case "%1"
				iRow = .Row	: .Row2=.Row
				.Col = arrCol(1) +1 : .Col2=.MaxCols
				.BlockMode = True
			   'ret = .AddCellSpan(C_PlantCd, 1 ,5, iRow)   '시작컬럼, 시작로, 길이컬럼, 길이행 
			   'ret = .AddCellSpan(1, iRow ,5, 1)
				.BackColor = RGB(250,250,210) 
				.ForeColor = vbBlack
				.BlockMode = False
			Case "%2"
				iRow = .Row :.Row2=.Row
				.Col = arrCol(1)+1 :.Col2=.MaxCols
				.BlockMode = True
				'ret = .AddCellSpan(2, iRow , 4, 1)
				'ret = .AddCellSpan(C_CCCd, 1 , 1, iRow)
				.BackColor = RGB(204,255,153) 
				.ForeColor = vbBlack
				.BlockMode = False
			Case "%3"
				iRow = .Row : .Row2=.Row
				.Col = arrCol(1) +2 : .Col2 =.MaxCols
				.BlockMode = True
				'ret = .AddCellSpan(C_ItemAcctNm-2, 1 , 1, iRow)
				'ret = .AddCellSpan(5, iRow , 2, 1)
				.BackColor = RGB(204,255,255) 
				.ForeColor = vbBlack
				.BlockMode = False
			Case "%4"  
				iRow = .Row
				.Col =arrCol(1)+1
				.Col2 = .MaxCols
				.Row2=.Row
				.BlockMode = True
				'ret = .AddCellSpan(C_AcctNM-1,1, 1,.maxRows)
				'ret = .AddCellSpan(C_ItemAcctNm-2,1, 1,.maxRows)
				.BackColor = RGB(255,228,181) 
				.ForeColor = vbBlack
				.BlockMode =False
		End Select
		.Col = 1: .Row = -1: .ColMerge = 1
		.Col = 2: .Row = -1: .ColMerge = 1
'		.Col = 3: .Row = -1: .ColMerge = 1
'		.Col = 4: .Row = -1: .ColMerge = 1
'		.Col = 5: .Row = -1: .ColMerge = 1
		strRowI = strRowI & CDbl(arrCol(2)) & Parent.gColSep
	Next

	frm1.txtTmp.value=frm1.txtTmp.value & strRowI
	.ReDraw = True
	End With

End Sub


'========================================================================================================
' Name : DbSave
' Desc : This sub is called by FncSave
'========================================================================================================
Function DbSave() 
    DbSave = True        
End Function

'========================================================================================================
' Name : DbSaveOk
' Desc : Called by MB Area when save operation is successful
'========================================================================================================
Function DbSaveOk()	
   
End Function

'========================================================================================================
' Function Name : DbDeleteOk
' Function Desc : Called by MB Area when delete operation is successful
'========================================================================================================
Sub DbDeleteOk()
	'------ Developer Coding part (Start)  -------------------------------------------------------------- 

	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub
'=================================================================================
'	Name : ChkKeyField()
'	Description : check the valid data
'=========================================================================================================
Function ChkKeyField()
	Dim strDataCd, strDataNm
    Dim strWhere , strFrom 
    Dim lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6
    
    Err.Clear                                       

	ChkKeyField = true		
'check plant
	If Trim(frm1.txtPlant_cd.value) <> "" Then
		strWhere = " plant_cd= " & FilterVar(frm1.txtPlant_Cd.value, "''", "S") & "  "

		Call CommonQueryRs(" plant_nm ","	 b_plant ", strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
		IF Len(lgF0) < 1 Then 
			Call DisplayMsgBox("970000","X",frm1.txtPlant_Cd.alt,"X")			
			frm1.txtPlant_nm.value = ""
			ChkKeyField = False
			frm1.txtPlant_Cd.focus 
			Exit Function
		End If
	
		strDataNm = split(lgF0,chr(11))
		frm1.txtPlant_NM.value = strDataNm(0)
	Else
		frm1.txtPLANT_NM.value=""
	End If
'check
	If Trim(frm1.txtItem_Acct.value) <> "" Then
		strWhere = " minor_cd = " & FilterVar(frm1.txtItem_Acct.value, "''", "S") & " "			
		strWhere = strWhere & " and major_cd=" & FilterVar("P1001", "''", "S") & " "			
		Call CommonQueryRs(" minor_nm ","	 b_minor  ", strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
		IF Len(lgF0) < 1 Then 
			Call DisplayMsgBox("970000","X",frm1.txtItem_Acct.alt,"X")			
			frm1.txtITEM_ACCT_NM.value = ""
			ChkKeyField = False
			frm1.txtITEM_ACCT.focus
			Exit Function
		End If
	
		strDataNm = split(lgF0,chr(11))
		frm1.txtITEM_ACCT_NM.value = strDataNm(0)
	Else
		frm1.txtITEM_ACCT_NM.value=""
	End If
'check 
	If Trim(frm1.txtItem_cd.value) <> "" Then
		If  trim(frm1.txtPlant_cd.value)="" Then 
			Call DisplayMsgBox("970000","X",frm1.txtPlant_cd.alt,"X")
			Exit function 
		End If
		
		strFrom = " b_item "
		strWhere = " item_cd = " & FilterVar(frm1.txtItem_cd.value, "''", "S") & " "		
		
		Call CommonQueryRs(" item_nm ",strFrom , strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
		IF Len(lgF0) < 1 Then 
			Call DisplayMsgBox("970000","X",frm1.txtItem_cd.alt,"X")			
			frm1.txtItem_nm.value = ""
			ChkKeyField = False
			frm1.txtItem_cd.focus 
			Exit Function
		End If	
		strDataNm = split(lgF0,chr(11))
		frm1.txtItem_nm.value = strDataNm(0)
	Else
		frm1.txtITEM_NM.value=""
	End If
End Function
</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->	
</HEAD>

<BODY TABINDEX="-1" SCROLL="no" oncontextmenu="javascript:return false">
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white><%=Request("strASPMnuMnuNm")%></font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* align=right>&nbsp;&nbsp;</TD>
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
									<TD CLASS="TD5">작업년월</TD>
									<TD CLASS="TD6" valign=top><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name=txtYYYYMM CLASS=FPDTYYYYMM title=FPDATETIME ALT="작업년월" tag="12" id=txtYYYYMM></OBJECT>');</SCRIPT>&nbsp;
									
									</TD>
									<TD CLASS="TD5" NOWRAP>공장</TD>
									<TD CLASS="TD6" NOWRAP><input NAME="txtPLANT_CD" TYPE="Text" MAXLENGTH="4" tag="15XXXU" size="10" ALT="공장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCalType" align=top TYPE="BUTTON" onclick="vbscript:Call OpenPopup(0)">
									<input NAME="txtPLANT_NM" TYPE="TEXT"  tag="14XXX" size="25">
									</TD>
								</TR>    
								<TR>									
									<TD CLASS="TD5">품목계정</TD>
									<TD CLASS="TD6" NOWRAP><input NAME="txtITEM_ACCT" TYPE="Text" MAXLENGTH="2" tag="15XXXU" size="15" ALT="품목계정"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCalType" align=top TYPE="BUTTON" onclick="vbscript:Call OpenPopup(1)">
									<input NAME="txtITEM_ACCT_NM" TYPE="TEXT"  tag="14XXX" size="25">
									</TD>
								
									<TD CLASS="TD5" NOWRAP>품목</TD>
									<TD CLASS="TD6" NOWRAP><input NAME="txtITEM_CD" TYPE="Text" MAXLENGTH="18" tag="15XXXU" size="20" ALT="품목"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCalType" align=top TYPE="BUTTON" onclick="vbscript:Call OpenPopup(2)">
									<input NAME="txtITEM_NM" TYPE="TEXT"  tag="14XXX" size="20">
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
					<TD WIDTH=100% HEIGHT=100% valign=top>
						<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD HEIGHT="60%" NOWRAP>
								<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> ID=vspdData NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" > <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
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
		<TD WIDTH=100% HEIGHT=<%=BizSize%>>
			<IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=no  noresize framespacing=0 TABINDEX= "-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="hYYYYMM" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="hPlant_cd" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="hItem_cd" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="hItem_Acct" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="txtTmp" tag="24" TABINDEX= "-1">

</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

