<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% session.CodePage=949 %>

<!-- '======================================================================================================
'*  1. Module Name          : Cost
'*  2. Function Name        :공통재료비투입현황 
'*  3. Program ID           : c4237ma1.asp
'*  4. Program Name         : 공통재료비투입현황 
'*  5. Program Desc         : 공통재료비투입현황 
'*  6. Modified date(First) : 2005-12-30
'*  7. Modified date(Last)  : 2005-12-22
'*  8. Modifier (First)     : HJO
'*  9. Modifier (Last)      : HJO
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

Const BIZ_PGM_ID = "c4237mb1.asp"                               'Biz Logic ASP

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
Dim C_CostCd
Dim C_CostNm

Dim C_OrderNo
Dim C_OrderSeq		
Dim C_WcCd		
Dim C_WcNm
Dim C_ItemAcct		
Dim C_AcctNm
Dim C_ItemCd		
Dim C_ItemNm
Dim C_TypeCd		
Dim C_TypeNm
Dim C_WipQty
Dim C_WipAmt

'==========================================  2.2.7 InitSpreadPosVariables()  =============================
' Function Name : InitSpreadPosVariables
' Function Desc : This method Assigns Sequential Number to spread sheet column 
'=========================================================================================================
Sub InitSpreadPosVariables(byVal pvSpd)

		If pvSpd="" or pvSpd="A" Then
			C_PlantCD=1
			C_CostCd=2
			C_CostNm=3

			C_WcCd		=4
			C_WcNm=5
			C_ItemAcct=6		
			C_AcctNm=7
			C_ItemCd		=8
			C_ItemNm=9
			C_TypeCd		=10
			C_TypeNm=11
			C_WipQty=12
			C_WipAmt=13
		End If
	
	
End Sub

'========================================================================================================
' Name : InitVariables()	
' Desc : Initialize value
'======================================================================================================== 
Sub InitVariables()

    lgIntFlgMode = Parent.OPMD_CMODE
    lgBlnFlgChgValue = False 
    lgSortKey         = 1                                       '⊙: initializes sort direction
    lgStrPrevKey = ""
    

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
		ggoSpread.Spreadinit "V20021106", , parent.gForbidDragDropSpread
			
		.MaxCols = C_WipAmt + 1
				'헤더를 2줄로    
		.ColHeaderRows = 1
		
		Call GetSpreadColumnPos("A")
		.ReDraw = False
		ggoSpread.SSSetEdit		C_PlantCD,	"공장"	, 10,,,,1	
		ggoSpread.SSSetEdit		C_CostCd,	"C/C",12,,,,1
		ggoSpread.SSSetEdit		C_CostNm,	"C/C명"	,15

		ggoSpread.SSSetEdit		C_WcCd,	"작업장",12,,,,1
		ggoSpread.SSSetEdit		C_WcNm,	"작업장명"	,15
					
		ggoSpread.SSSetEdit		C_ItemAcct,	"품목계정"	, 15,,,,1	
		ggoSpread.SSSetEdit		C_AcctNm,	"품목계정명"	, 20
		ggoSpread.SSSetEdit		C_ItemCd,	"자품목"	, 15,,,,1	
		ggoSpread.SSSetEdit		C_ItemNm,	"자품목명"	, 20
		ggoSpread.SSSetEdit		C_TypeCd,	"수불유형"	, 15,,,,1	
		ggoSpread.SSSetEdit		C_TypeNm,	"수불유형명"	, 20
		ggoSpread.SSSetFloat	C_WipQty,	"투입수량"	, 12,		Parent.ggQtyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	C_WipAmt,	"투입금액"	, 12,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
				
		Call ggoSpread.SSSetColHidden(.MaxCols,.MaxCols,True)
		Call ggoSpread.SSSetColHidden(C_ItemAcct,C_ItemAcct,True)
	
		.ReDraw = True		
		End With
		Call SetSpreadLock("A")
	End If
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
 			
			C_PlantCD  = iCurColumnPos(1)	
			C_CostCd  = iCurColumnPos(2)	
			C_CostNm  = iCurColumnPos(3)	

			C_WcCd		  = iCurColumnPos(4)	
			C_WcNm  = iCurColumnPos(5)	
			C_ItemAcct		  = iCurColumnPos(6)	
			C_AcctNm  = iCurColumnPos(7)	
			C_ItemCd		  = iCurColumnPos(8)	
			C_ItemNm  = iCurColumnPos(9)	
			C_TypeCd		  = iCurColumnPos(10)	
			C_TypeNm  = iCurColumnPos(11)	
			C_WipQty  = iCurColumnPos(12)	
			C_WipAmt  = iCurColumnPos(13)	
			
 		End Select  	
 	
End Sub
'******************************************  2.4 POP-UP 처리함수  ****************************************
'	기능: 기준 POP-UP
'   설명: 기준 POP-UP에 관한 Open은 Include한다. 
'	      하나의 ASP에서 Popup이 중복되면 하나는 링크시켜 사용하고 나머지는 재정의하여 사용한다.
'********************************************************************************************************* 
Function OpenPopUp(Byval iWhere)
	Dim arrRet, sTmp,strYYYYMM
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim sYear,sMon,sDay

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True
	
	If  trim(frm1.txtYYYYMM.TExt)="" Then 
			Call DisplayMsgBox("970000","X",frm1.txtYYYYMM.alt,"X")			
			IsOpenPop=false
			frm1.txtYYYYMM.Focus
			Exit function 
	End If

	With frm1

	Call parent.ExtractDateFromSuper(.txtYYYYMM.Text, parent.gDateFormat,sYear,sMon,sDay)
		
	strYYYYMM= (sYear&sMon)

	
	Select Case iWhere
		Case 0
			arrParam(0) = "공장 팝업"
			arrParam(1) = " B_PLANT "
			arrParam(2) = Trim(.txtPlant_cd.value)
			arrParam(3) = ""
			arrParam(4) = ""
			arrParam(5) = "공장" 

			arrField(0) ="ED10" & parent.gColsep &  "PLANT_CD"	
			arrField(1) ="ED20" & parent.gColsep &  "PLANT_NM"    
			arrHeader(0) = "공장"	
			arrHeader(1) = "공장명"
		Case 1
			arrParam(0) = "C/C 팝업"
			arrParam(1) = " b_cost_center(nolock) "
			
			arrParam(2) = Trim(.txtCost_cd.value)
			arrParam(3) = ""
			arrParam(4) = " COST_TYPE='M' "
			arrParam(5) = "C/C" 

			arrField(0) = "ED10" & Parent.gColSep & "COST_CD"	
			arrField(1) ="ED30" & Parent.gColSep & "COST_NM"
			arrField(2) = ""		
    
			arrHeader(0) = "C/C"	
			arrHeader(1) = "C/C명"
			arrHeader(2) = ""
				
		Case 2
			arrParam(0) = "자품목계정 팝업"
			arrParam(1) = "dbo.B_MINOR A(NOLOCK) INNER JOIN B_ITEM_ACCT_INF B(NOLOCK) ON A.MINOR_CD=B.ITEM_ACCT "	
			arrParam(2) = Trim(.txtITEM_ACCT.value)
			arrParam(3) = ""
			arrParam(4) = "MAJOR_CD =" & FilterVar("P1001", "''", "S") & " AND B.ITEM_ACCT_GROUP <> '6MRO' "
			arrParam(5) = "자품목계정" 

			arrField(0) = "MINOR_CD"
			arrField(1) = "MINOR_NM"		
			arrField(2) = ""	
    
			arrHeader(0) = "자품목계정"
			arrHeader(1) = "자품목계정명"
			arrHeader(2) = ""	
		Case 3
			arrParam(0) = "자품목 팝업"
			arrParam(1) ="  b_item "
			arrParam(2) = Trim(.txtITEM_CD.value)
			arrParam(3) = ""	
			arrParam(4) = ""
			arrParam(5) = "자품목" 

			arrField(0) = "ED20" & Parent.gColSep &"ITEM_CD"	
			arrField(1) = "ED30" & Parent.gColSep &"ITEM_NM"
			arrField(2) = ""		
    
			arrHeader(0) = "자품목"	
			arrHeader(1) = "자품목명"
			arrHeader(2) = ""
		Case 4
			arrParam(0) = "작업장 팝업"
			arrParam(1) =" P_WORK_CENTER "
			arrParam(2) = Trim(.txtWC_Cd.value)
			arrParam(3) = ""	
			arrParam(4) = ""
			arrParam(5) = "작업장" 

			arrField(0) = "ED20" & Parent.gColSep &"WC_CD"	
			arrField(1) = "ED30" & Parent.gColSep &"WC_NM"
			
    
			arrHeader(0) = "작업장"	
			arrHeader(1) = "작업장명"
			
		Case 5
			arrParam(0) = "수불유형 팝업"
			arrParam(1) =" (select distinct b.mov_type, a.minor_nm "
			arrParam(1) = arrParam(1) & "	from b_minor a(nolock) join c_common_material_s b(nolock) "
			arrParam(1) = arrParam(1) & "	on a.major_cd = 'i0001' and a.minor_cd = b.mov_type "
			arrParam(1) = arrParam(1) & "	where b.yyyymm = " & filtervar ( strYYYYMM,"","S") &  ") z"
			arrParam(2) = Trim(.txtMovType.value)
			arrParam(3) = ""	
			arrParam(4) = "  "
			arrParam(5) = "수불유형" 

			arrField(0) = "ED20" & Parent.gColSep &"mov_type"	
			arrField(1) = "ED30" & Parent.gColSep &"minor_nm"
			
    
			arrHeader(0) = "수불유형"	
			arrHeader(1) = "수불유형명"

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
				.txtPlant_CD.value= arrRet(0)
				.txtPlant_NM.value= arrRet(1)
				.txtPlant_CD.focus
					
			Case 1
				.txtCost_cd.value		= arrRet(0)
				.txtCost_nm.value		= arrRet(1)
				.txtCost_cd.focus
			Case 2
				.txtITEM_ACCT.value		= arrRet(0)
				.txtITEM_ACCT_NM.value	= arrRet(1)	
				.txtITEM_ACCT.focus					
			Case 3
				.txtITEM_CD.value		= arrRet(0)
				.txtITEM_NM.value		= arrRet(1)
				.txtITEM_CD.focus
			Case 4
				.txtWC_Cd.value= arrRet(0)
				.txtWC_NM.value=arrRet(1)
				.txtWC_Cd.focus
			Case 5
				.txtMovType.value= arrRet(0)
				.txtMovNm.value=arrRet(1)
				.txtMovType.focus
			
			
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
		frm1.txtPlant_CD.value = UCase(parent.gPlant)
		frm1.txtPlant_Nm.value = parent.gPlantNm
		frm1.txtCOST_Cd.focus 		
	Else
		frm1.txtPlant_CD.focus 		
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


Sub txtCost_cd_onKeyPress()
	If window.event.keyCode = 13 Then
		Call FncQuery
	End If
End Sub


Sub txtCost_cd_onChange()
	If frm1.txtCost_cd.value ="" then frm1.txtCost_nm.value=""
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

'========================================================================================================
'   Event Name : vspdData_MouseDown
'   Event Desc :
'========================================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If
End Sub 

'========================================================================================================
'   Event Name : vspdData_ScriptLeaveCell
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_ScriptLeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
    If Row <> NewRow And NewRow > 0 Then
        With frm1
            .vspdData.Row = NewRow
            .vspdData.Col = C_ItemCd

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
	
'	Call SetPopupMenuItemInf("0000111111")         '화면별 설정		
	
    gMouseClickStatus = "SPC"

    Set gActiveSpdSheet = frm1.vspdData

	
	If frm1.vspdData.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If

'    If Row <= 0 Then
 '       ggoSpread.Source = frm1.vspdData
  '      If lgSortKey = 1 Then
   '         ggoSpread.SSSort	Col			'Sort in ascending
    '        lgSortKey = 2
     '   Else
      '      ggoSpread.SSSort Col	,lgSortKey	'Sort in descending
       '     lgSortKey = 1
        'End If
    'Else
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    	frm1.vspdData.Row = Row
	'	frm1.vspdData.Col = C_MajorCd
	
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
    'End If
    
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
    Call InitSpreadSheet(gActiveSpdSheet.alt)      
    
	Call ggoSpread.ReOrderingSpreadData()
	
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
		
		sStartDt= (sYear&sMon)
		

		If lgIntFlgMode = parent.OPMD_UMODE Then
			strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
			strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
			strVal = strVal & "&txtYYYYMM=" & Trim(.hYYYYMM.value)		
			strVal = strVal & "&txtCost_cd=" & Trim(.hCost_cd.value)
			
			strVal = strVal & "&txtItem_cd=" & Trim(.hItem_cd.value)
			strVal = strVal & "&txtWc_Cd=" & Trim(.hWc_cd.value)
			strVal = strVal & "&txtItemAcct=" & Trim(.hItemAcct.value)
			strVal = strVal & "&txtPlant_cd=" & Trim(.hPlant_cd.value)
			strVal = strVal & "&txtMovType=" & Trim(.hMovType.value)
			

		Else
			strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
			strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
			strVal = strVal & "&txtYYYYMM=" & sStartDt		
			strVal = strVal & "&txtCost_cd=" & Trim(.txtCost_cd.value)
			
			strVal = strVal & "&txtItem_cd=" & Trim(.txtItem_cd.value)
			strVal = strVal & "&txtWc_Cd=" & Trim(.txtWc_cd.value)
			strVal = strVal & "&txtItemAcct=" & Trim(.txtITEM_ACCT.value)
			strVal = strVal & "&txtPlant_cd=" & Trim(.txtPlant_cd.value)
			strVal = strVal & "&txtMovType=" & Trim(.txtMovType.value)

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
			   'ret = .AddCellSpan(C_OrderNo, 1 ,5, iRow)   '시작컬럼, 시작로, 길이컬럼, 길이행 
				ret = .AddCellSpan(.Col,iRow, 11,1)
				.BackColor = RGB(250,250,210) 
				.ForeColor = vbBlack
				.BlockMode = False
			Case "%2"
				iRow = .Row :.Row2=.Row
				.Col = arrCol(1)+1 :.Col2=.MaxCols
				.BlockMode = True
				ret = .AddCellSpan(.Col,iRow, 10,1)
				'ret = .AddCellSpan(C_CCCd, 1 , 1, iRow)
				.BackColor = RGB(204,255,153) 
				.ForeColor = vbBlack
				.BlockMode = False
			Case "%3"
				iRow = .Row : .Row2=.Row
				.Col = arrCol(1) +1 : .Col2 =.MaxCols
				.BlockMode = True
				ret = .AddCellSpan(.Col,iRow, 8,1)
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
				ret = .AddCellSpan(.Col,iRow, 6,1)
				'ret = .AddCellSpan(C_ItemAcctNm-2,1, 1,.maxRows)
				.BackColor = RGB(255,228,181) 
				.ForeColor = vbBlack
				.BlockMode =False
			Case "%5"  
				iRow = .Row
				.Col =arrCol(1)+1
				.Col2 = .MaxCols
				.Row2=.Row
				.BlockMode = True
				ret = .AddCellSpan(.Col,iRow, 4,1)
				'ret = .AddCellSpan(C_ItemAcctNm-2,1, 1,.maxRows)
				.BackColor = RGB(255,200,181) 
				.ForeColor = vbBlack
				.BlockMode =False
			Case "%6"  
				iRow = .Row
				.Col =arrCol(1)+1
				.Col2 = .MaxCols
				.Row2=.Row
				.BlockMode = True
				ret = .AddCellSpan(.Col,iRow, 2,1)
				'ret = .AddCellSpan(C_ItemAcctNm-2,1, 1,.maxRows)
				.BackColor = RGB(255,228,0) 
				.ForeColor = vbBlack
				.BlockMode =False
		End Select
		.Col = 1: .Row = -1: .ColMerge = 1
		.Col = 2: .Row = -1: .ColMerge = 1
		.Col = 3: .Row = -1: .ColMerge = 1
		.Col = 8: .Row = -1: .ColMerge = 1
		.Col = 9: .Row = -1: .ColMerge = 1
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
    
    Dim sStartDt,sYear,sMon,sDay
    
    On Error Resume Next
    
    Err.Clear       

	ChkKeyField = true		
	If  trim(frm1.txtYYYYMM.TExt)="" Then 
			Call DisplayMsgBox("970000","X",frm1.txtYYYYMM.alt,"X")			
			ChkKeyField=false
			frm1.txtYYYYMM.Focus
			Exit function 
	End If
	
	Call parent.ExtractDateFromSuper(frm1.txtYYYYMM.Text, parent.gDateFormat,sYear,sMon,sDay)
		
	sStartDt= (sYear&sMon)                  
'check plant
	If Trim(frm1.txtPlant_cd.value) <> "" Then		
		strFrom ="	 b_plant "
		strWhere = " plant_cd  = " & FilterVar(frm1.txtPlant_cd.value, "''", "S") & "  "
		
		Call CommonQueryRs(" distinct  plant_nm  ",strFrom , strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
		IF Len(lgF0) < 1 Then 
			Call DisplayMsgBox("970000","X",frm1.txtPlant_cd.alt,"X")			
			frm1.txtPlant_nm.value = ""
			ChkKeyField = False
			frm1.txtPlant_cd.focus 
			Exit Function
		End If
	
		strDataNm = split(lgF0,chr(11))
		frm1.txtPlant_nm.value = strDataNm(0)
	Else
		frm1.txtPlant_nm.value=""
	End If
'check cost
	If Trim(frm1.txtCost_cd.value) <> "" Then		
		strFrom ="	b_cost_center(nolock)  "
		strWhere = " cost_cd = " & FilterVar(frm1.txtCost_cd.value, "''", "S") & "  "
		strWhere = strwhere & "	and cost_type='M' "
		Call CommonQueryRs(" distinct  cost_nm ",strFrom , strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
		IF Len(lgF0) < 1 Then 
			Call DisplayMsgBox("970000","X",frm1.txtCost_cd.alt,"X")			
			frm1.txtCost_nm.value = ""
			ChkKeyField = False
			frm1.txtCost_cd.focus 
			Exit Function
		End If
	
		strDataNm = split(lgF0,chr(11))
		frm1.txtCost_nm.value = strDataNm(0)
	Else
		frm1.txtCost_nm.value=""
	End If
'check wc cd
	If Trim(frm1.txtWC_Cd.value) <> "" Then
		strFrom = "p_work_center(nolock)"
		
		strWhere = " wc_cd  = " & FilterVar(frm1.txtWC_Cd.value, "''", "S") & " "			
		
		Call CommonQueryRs(" distinct wc_nm ", strFrom , strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
		IF Len(lgF0) < 1 Then 
			Call DisplayMsgBox("970000","X",frm1.txtWC_Cd.alt,"X")	
			frm1.txtWC_nm.value=""		
			ChkKeyField = False
			frm1.txtWC_Cd.focus
			Exit Function
		End If	
		strDataNm = split(lgF0,chr(11))
		frm1.txtWC_nm.value = strDataNm(0)
	Else
		frm1.txtWC_nm.value=""
	End If
	
'check item acct	
	If Trim(frm1.txtITEM_ACCT.value) <> "" Then
	
		strFrom = " B_MINOR a(nolock) inner join b_item_acct_inf b(nolock) on a.minor_cd=b.item_acct  "
		strWhere = " MAJOR_CD =" & FilterVar("P1001", "''", "S") & " and b.item_acct_group <>'6MRO' "
		strWhere = strWhere & "  and minor_cd= " & FilterVar(frm1.txtITEM_ACCT.value, "''", "S") & " "
		
		Call CommonQueryRs(" distinct minor_nm  ",strFrom , strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
		IF Len(lgF0) < 1 Then 
			Call DisplayMsgBox("970000","X",frm1.txtITEM_ACCT.alt,"X")			
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
	'check item cd
	If Trim(frm1.txtItem_cd.value) <> "" Then
	
		strFrom = "   b_item(nolock)  "
		strWhere = " item_cd = " & FilterVar(frm1.txtItem_cd.value, "''", "S") & " "	
		
		Call CommonQueryRs(" distinct item_nm ",strFrom , strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
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
	

	If Trim(frm1.txtMovType.value) <> "" Then
	
		strFrom = " b_minor(nolock)  "
		strWhere = " major_cd = 'I0001' and minor_cd = " & FilterVar(frm1.txtMovType.value, "''", "S") & " "	
		
		Call CommonQueryRs(" minor_nm ",strFrom , strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
		IF Len(lgF0) < 1 Then 
			Call DisplayMsgBox("970000","X",frm1.txtMovType.alt,"X")			
			frm1.txtMovNm.value = ""
			ChkKeyField = False
			frm1.txtMovType.focus 
			Exit Function
		End If	
		strDataNm = split(lgF0,chr(11))
		frm1.txtMovNm.value = strDataNm(0)
	Else
		frm1.txtMovNm.value=""
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
									<TD CLASS="TD6" NOWRAP><input NAME="txtPlant_CD" TYPE="Text" MAXLENGTH="4" tag="15XXXU" size="10" ALT="공장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCalType" align=top TYPE="BUTTON" onclick="vbscript:Call OpenPopup(0)">
									<input NAME="txtPlant_Nm" TYPE="TEXT"  tag="14XXX" size="25">
									</TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>C/C</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtCost_cd" SIZE=10 MAXLENGTH=10 tag="11xxxU" ALT="C/C"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnOprCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenPopUp(1)">&nbsp;<INPUT TYPE=TEXT NAME="txtCost_nm" SIZE=25 tag="14"></TD>
									<TD CLASS="TD5">작업장</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtWC_Cd" SIZE=10 MAXLENGTH=7 tag="11xxxU" ALT="작업장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnOprCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenPopup(4)">
									<INPUT TYPE=TEXT NAME="txtWC_NM" SIZE=25 tag="14">
									</TD>									
								</TR>
								<TR>								
									<TD CLASS="TD5">자품목계정</TD>
									<TD CLASS="TD6" NOWRAP><input NAME="txtITEM_ACCT" TYPE="Text" MAXLENGTH="10" tag="11XXXU" size="10" ALT="자품목계정"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCalType" align=top TYPE="BUTTON" onclick="vbscript:Call OpenPopup(2)">
									<input NAME="txtITEM_ACCT_NM" TYPE="TEXT"  tag="14XXX" size="20">
									</TD>
									<TD CLASS="TD5" NOWRAP>자품목</TD>
									<TD CLASS="TD6" NOWRAP><input NAME="txtITEM_CD" TYPE="Text" MAXLENGTH="18" tag="15XXXU" size="20" ALT="자품목"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCalType" align=top TYPE="BUTTON" onclick="vbscript:Call OpenPopup(3)">
									<input NAME="txtITEM_NM" TYPE="TEXT"  tag="14XXX" size="20">
									</TD>									
								</TR>								
								<TR>								
									<TD CLASS="TD5">수불유형</TD>
									<TD CLASS="TD6" NOWRAP><input NAME="txtMovType" TYPE="Text" MAXLENGTH="3" tag="11XXXU" size="10" ALT="수불유형"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnMovType" align=top TYPE="BUTTON" onclick="vbscript:Call OpenPopup(5)">
									<input NAME="txtMovNm" TYPE="TEXT"  tag="14XXX" size="20">
									</TD>
									<TD CLASS="TD5" NOWRAP>&nbsp;</TD>
									<TD CLASS="TD6" NOWRAP>&nbsp;</TD>									
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
								<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> ID=vspdData NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" alt="A"> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
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
<INPUT TYPE=HIDDEN NAME="hCost_cd" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="hItem_cd" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="hOrder_no" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="txtTmp" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="hWc_Cd" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="hPlant_Cd" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="hItemAcct" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="hMovType" tag="24" TABINDEX= "-1">

</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

