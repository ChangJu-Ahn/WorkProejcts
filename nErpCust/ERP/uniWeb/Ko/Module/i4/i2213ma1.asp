<%@ LANGUAGE="VBSCRIPT" %>
<!--'**********************************************************************************************
'*  1. Module Name          : Inventory
'*  2. Function Name        : Stock Requirement 조회 
'*  3. Program ID           : I2213ma1.asp
'*  4. Program Name         : Stock Requirement정보 
'*  5. Program Desc         :
'*  6. Comproxy List        :
'
'*  7. Modified date(First) : 2000/05/3
'*  8. Modified date(Last)  : 2003/05/15
'*  9. Modifier (First)     : Nam hoon kim
'* 10. Modifier (Last)      : Lee Seung Wook
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*                            -2000/04/18 : ..........
'**********************************************************************************************-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!--'#########################################################################################################
'            1. 선 언 부 
'##########################################################################################################-->
<!--'******************************************  1.1 Inc 선언   **********************************************
' 기능: Inc. Include
'********************************************************************************************************* -->
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->
<!--'==========================================  1.1.1 Style Sheet  ======================================
'==========================================================================================================-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'==========================================  1.1.2 공통 Include   ======================================
'==========================================================================================================-->

<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/Cookie.vbs"></SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit                 

Const BIZ_PGM_ID = "i2213mb1.asp"

Dim	C_Date        
Dim	C_MvmtFlag 
Dim	C_TrackingNo 
Dim	C_PlanQty 
Dim	C_RemainQty
Dim	C_AvalQty
Dim IsOpenPop

<!-- #Include file="../../inc/lgvariables.inc" -->
'==========================================  1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
'=======================================================================================================
Sub InitVariables()
	lgIntGrpCount = 0                          
	lgBlnFlgChgValue = False
	lgStrPrevKey = ""
	lgLngCurRows = 0                           
End Sub

'========================================  2. SetDefaultVal()  ======================================
' Name : SetDefaultVal()
' Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
'===================================================================================================  
Sub SetDefaultVal()
	frm1.txtYyyyMmDd.text = UniConvDateAToB("<%=GetSvrDate%>",Parent.gServerDateFormat,Parent.gDateFormat) 
	lgBlnFlgChgValue = False
	if frm1.txtPlantCd.value = "" Then
	   frm1.txtPlantNm.value = ""
	End if
	if frm1.txtItemCd.value = "" Then
	   frm1.txtItemNm.value = ""
	End if
 
	If Parent.gPlant <> "" Then
	 frm1.txtPlantCd.value = UCase(Parent.gPlant)
	 frm1.txtPlantNm.value = Parent.gPlantNm
	 frm1.txtItemCd.focus   
	End If     
End Sub

'========================================= 3. LoadInfTB19029() ==================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'===================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "I","NOCOOKIE","MA") %>
End Sub

'========================================= 4. InitSpreadSheet() =========================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'=========================================================================================================
Sub InitSpreadSheet()
	Call InitSpreadPosVariables()
	ggoSpread.Source = frm1.vspdData
	ggoSpread.Spreadinit "V20021106", , Parent.gAllowDragDropSpread
	
	With frm1.vspdData
		.ReDraw = false
		.MaxCols = C_AvalQty+1         
		.MaxRows = 0
		  
		Call GetSpreadColumnPos("A")

		ggoSpread.SSSetDate C_Date, "계획일", 10,2, Parent.gDateFormat
		ggoSpread.SSSetEdit C_MvmtFlag, "오더구분", 15,2  
		ggoSpread.SSSetEdit C_TrackingNo, "Tracking No", 20
		ggoSpread.SSSetFloat C_PlanQty, "계획량", 20, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec   '수량'
		ggoSpread.SSSetFloat C_RemainQty, "잔량", 20, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec    '수량'
		ggoSpread.SSSetFloat C_AvalQty, "가용재고", 20, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec  '수량'

		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
		
		.ReDraw = true

		Call SetSpreadLock
		ggoSpread.SSSetSplit2(1)
	End With
End Sub

'========================================= 5. InitSpreadPosVariables() =========================================
' Function Name : InitSpreadPosVariables
' Function Desc : This method initializes spread sheet column property
'=========================================================================================================
Sub InitSpreadPosVariables()
	C_Date			= 1
	C_MvmtFlag		= 2
	C_TrackingNo	= 3
	C_PlanQty		= 4
	C_RemainQty		= 5
	C_AvalQty		= 6
End Sub


'========================================================================================
' Function Name : GetSpreadColumnPos
' Function Desc : 
'========================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
	Dim iCurColumnPos
	
	Select Case UCase(pvSpdNo)
	Case "A"
		ggoSpread.Source = frm1.vspdData 
		
		Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

		C_Date       = iCurColumnPos(1)
		C_MvmtFlag   = iCurColumnPos(2)
		C_TrackingNo = iCurColumnPos(3)
		C_PlanQty    = iCurColumnPos(4)
		C_RemainQty  = iCurColumnPos(5)
		C_AvalQty    = iCurColumnPos(6)
	End Select
End Sub

'========================================= 2.7 SetSpreadLock() ===========================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'=========================================================================================================
Sub SetSpreadLock()
 With frm1
  .vspdData.ReDraw = False
  ggoSpread.SpreadLockWithOddEvenRowColor()
  .vspdData.ReDraw = True
 End With
End Sub


Sub SetSpreadColor(ByVal lRow)
	With frm1
		.vspdData.ReDraw = False
		ggoSpread.SSSetProtected C_Date, lRow, lRow
		ggoSpread.SSSetProtected C_MvmtFlag, lRow, lRow
		ggoSpread.SSSetProtected C_TrackingNo, lRow, lRow
		ggoSpread.SSSetProtected C_PlantQty, lRow, lRow
		ggoSpread.SSSetProtected C_RemainQty, lRow, lRow
		ggoSpread.SSSetProtected C_AvalQty, lRow, lRow
		ggoSpread.SSSetProtected .vspdData.MaxCols, lRow, lRow
		.vspdData.ReDraw = True
	End With
End Sub

'------------------------------------------ OpenPlant()  --------------------------------------------------
'	Name : OpenPlant()
'	Description : Plant Popup
'---------------------------------------------------------------------------------------------------------
Function OpenPlant()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function
	IsOpenPop = True

	arrParam(0) = "공장팝업" 
	arrParam(1) = "B_PLANT"    
	arrParam(2) = Trim(frm1.txtPlantCd.Value)
	arrParam(3) = ""
	arrParam(4) = ""   
	arrParam(5) = "공장"   
	 
	arrField(0) = "PLANT_CD" 
	arrField(1) = "PLANT_NM" 
	 
	arrHeader(0) = "공장"  
	arrHeader(1) = "공장명"  

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
	 "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	 
	IsOpenPop = False
	 
	If arrRet(0) = "" Then
		frm1.txtPlantCd.focus
		Exit Function
	Else
		Call SetPlant(arrRet)
	End If 
	 
End Function

'------------------------------------------ OpenItem()  --------------------------------------------------
'	Name : OpenItem()
'	Description : Item Popup
'---------------------------------------------------------------------------------------------------------
Function OpenItem()
	Dim iCalledAspName
	Dim IntRetCD

	Dim arrRet
	Dim arrParam(5), arrField(6)
	 
	If Trim(frm1.txtPlantCd.Value) = "" then
		Call DisplayMsgBox("169901","X", "X", "X")   
		frm1.txtPlantCd.focus
		Exit Function
	End If

	If  CommonQueryRs(" PLANT_NM "," B_PLANT ", " PLANT_CD = " & FilterVar(frm1.txtPlantCd.Value, "''", "S"), _
					  lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
	   
		Call DisplayMsgBox("125000","X","X","X")
		frm1.txtPlantCd.focus
		Exit function
	End If
	 
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True
	 
	arrParam(0) = Trim(frm1.txtPlantCd.value) 
	arrParam(1) = Trim(frm1.txtItemCd.Value)  
	arrParam(2) = ""      
	arrParam(3) = ""      
	 
	arrField(0) = 1  
	arrField(1) = 2  
	arrField(2) = 9  
	arrField(3) = 6  

	iCalledAspName = AskPRAspName("B1B11PA3")
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040",parent.VB_INFORMATION,"B1B11PA3","x")
		IsOpenPop = False
		Exit Function
	End If
	 
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam, arrField), _
	"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
	 
	IsOpenPop = False
	 
	If arrRet(0) = "" Then
		frm1.txtItemCd.focus
		Exit Function
	Else
		Call SetItem(arrRet)
	End If 
End Function

'------------------------------------------ OpenItemAcct()  --------------------------------------------------
'	Name : OpenItemAcct()
'	Description : ItemAcct Popup
'---------------------------------------------------------------------------------------------------------
Function OpenItemAcct()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function
 
	IsOpenPop = True

	arrParam(0) = "품목계정 팝업"    
	arrParam(1) = "B_MINOR"      
	arrParam(2) = Trim(frm1.txtItemAcct.Value)
	arrParam(3) = ""       
	arrParam(4) = "MAJOR_CD = " & FilterVar("P1001", "''", "S") & ""
	arrParam(5) = "품목계정"   
 
	arrField(0) = "MINOR_CD"      
	arrField(1) = "MINOR_NM"      
 
	arrHeader(0) = "품목계정" 
	arrHeader(1) = "품목계정명"
 
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
	 "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
 
	IsOpenPop = False
 
	If arrRet(0) = "" Then
		frm1.txtItemAcct.focus
		Exit Function
	Else
	 Call SetItemAcct(arrRet)
	End If 
End Function

'------------------------------------------  SetPlant()  --------------------------------------------------
'	Name : SetPlant()
'	Description : OpenPlant Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetPlant(byRef arrRet)
	frm1.txtPlantCd.Value    = arrRet(0)  
	frm1.txtPlantNm.Value    = arrRet(1)  
	frm1.txtPlantCd.focus
End Function
'------------------------------------------  SetItem()  --------------------------------------------------
'	Name : SetItem()
'	Description : OpenItem Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetItem(ByRef arrRet)
	frm1.txtItemCd.value = arrRet(0) 
	frm1.txtItemNm.value = arrRet(1)
	frm1.txtItemCd.focus
End Function


Function CookiePage()

	Dim strItemCd, strBaseDt, strPlantCd

	strItemCd  = ReadCookie("PoNo")
	strBaseDt  = ReadCookie("BaseDt")
	strPlantCd = ReadCookie("PlantCd")
 
	If strItemCd = ""  or strBaseDt = ""  then Exit Function

	frm1.txtItemCd.value =  strItemCd
	frm1.txtYyyyMmDd.Text = strBaseDt
	frm1.txtPlantCd.Value = strPlantCd

	Call dbquery()
     
	WriteCookie "PoNo" ,   ""
	WriteCookie "BaseDt" , ""
	WriteCookie "PlantCd", ""

End Function

'==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분 
'=====================================================================================================
Sub Form_Load()

    Call LoadInfTB19029           
    Call ggoOper.LockField(Document, "N")                                  
    Call ggoOper.FormatField(Document, "A", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec)
                        
    Call InitSpreadSheet          
    Call InitVariables            
    If Parent.gPlant <> "" Then
  frm1.txtPlantCd.value = UCase(Parent.gPlant)
  frm1.txtPlantNm.value = Parent.gPlantNm
  frm1.txtItemCd.focus    
 Else
  frm1.txtPlantCd.focus 
 End If  
    
    Call SetDefaultVal
 Call SetToolbar("11000000000011")         
 
 Call CookiePage

End Sub


'==========================================================================================
'   Event Name : txtYyyyMmDd_DblClick
'   Event Desc :
'==========================================================================================
Sub txtYyyyMmDd_DblClick(Button) 
    If Button = 1 Then
        frm1.txtYyyyMmDd.Action = 7
        Call SetFocusToDocument("M")        
        frm1.txtYyyyMmDd.Focus
    End If
End Sub

'==========================================================================================
'   Event Name : txtYyyyMmDd_KeyPress
'   Event Desc :
'==========================================================================================
Sub txtYyyyMmDd_KeyPress(KeyAscii)
	If KeyAscii = 13 Then
		Call MainQuery()
	End If
End Sub

'==========================================================================================
'   Event Name : vspdData_Click
'   Event Desc :
'==========================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
	Call SetPopupMenuItemInf("0000111111")	
	gMouseClickStatus = "SPC"   
	Set gActiveSpdSheet = frm1.vspdData
	
	If frm1.vspdData.MaxRows = 0 Then
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
End Sub

Sub vspdData_DblClick(ByVal Col, ByVal Row)
	Dim iColumnName
   
	If Row <= 0 Then
		Exit Sub
	End If
	If frm1.vspdData.MaxRows = 0 Then
		Exit Sub
	End If
End Sub

Sub vspdData_MouseDown(Button , Shift , x , y)

   If Button = 2 And gMouseClickStatus = "SPC" Then
      gMouseClickStatus = "SPCR"
   End If
End Sub 

Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)

   ggoSpread.Source = frm1.vspdData
   Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub 

Sub vspdData_ScriptDragDropBlock( Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
   ggoSpread.Source = frm1.vspdData
   Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
   Call GetSpreadColumnPos("A")
End Sub

Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )

    If OldLeft <> NewLeft Then
        Exit Sub
    End If
     
	If CheckRunningBizProcess = True Then
		Exit Sub
	End If
    
	if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData, NewTop) Then 
		If lgStrPrevKey <> "" Then
			Call DisableToolBar(Parent.TBC_QUERY)
			If DbQuery = False Then
				Call RestoreToolBar()
				Exit Sub
			End if
		End IF
	End if
	 
End Sub 

Sub PopSaveSpreadColumnInf()

   ggoSpread.Source = gActiveSpdSheet
   Call ggoSpread.SaveSpreadColumnInf()
End Sub 

Sub PopRestoreSpreadColumnInf()

   ggoSpread.Source = gActiveSpdSheet
   Call ggoSpread.RestoreSpreadInf()
   Call InitSpreadSheet
   Call ggoSpread.ReOrderingSpreadData
End Sub 

Sub FncSplitColumn()

    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   

End Sub

Function FncQuery()

 FncQuery = False                                                       
 
 Err.Clear                                                              
 
 If Not chkField(Document, "1") Then        
	Exit Function
 End If

 Call ggoOper.ClearField(Document, "2")        
 Call InitVariables           

 If  CommonQueryRs(" PLANT_NM "," B_PLANT ", " PLANT_CD = " & FilterVar(frm1.txtPlantCd.Value, "''", "S"), _
	lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
	 
	Call DisplayMsgBox("125000","X","X","X")
	frm1.txtPlantNm.value = ""
	frm1.txtPlantCd.focus
	Exit function
 End If
 lgF0 = Split(lgF0,Chr(11))
 frm1.txtPlantNm.value = lgF0(0)
 
 If  CommonQueryRs(" ITEM_NM "," B_ITEM ", " ITEM_CD= " & FilterVar(frm1.txtItemCd.value, "''", "S"), _
  lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
	Call DisplayMsgBox("122600","X","X","X")
	frm1.txtItemNm.value = ""
	frm1.txtItemCd.focus
	Exit function
 End If
 lgF0 = Split(lgF0,Chr(11))
 frm1.txtItemNm.value = lgF0(0)
   
 If  CommonQueryRs(" ITEM_CD "," B_ITEM_BY_PLANT ", " PLANT_CD = " & FilterVar(frm1.txtPlantCd.Value, "''", "S") & " AND ITEM_CD= " & FilterVar(frm1.txtItemCd.value, "''", "S"), _
  lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
   
  Call DisplayMsgBox("122700","X","X","X")
  frm1.txtItemNm.value = ""
  frm1.txtItemCd.focus
  Exit function
 End If 

 Call SetToolbar("11000000000111") 
 
 If DbQuery = False Then
	frm1.txtItemCd.focus
	Exit Function
 End if
 
 FncQuery = True
 Set gActiveElement = document.activeElement            
End Function


Function FncPrint()
	Call parent.FncPrint()
End Function


Function FncExcel()
	Call parent.FncExport(Parent.C_MULTI)
End Function

Function FncFind() 
    Call parent.FncFind(Parent.C_MULTI , True)                                                    
End Function

Function FncExit() 
    FncExit = True
End Function


Function DbQuery()
	Dim LngLastRow
	Dim LngMaxRow
	Dim StrNextKey

	Call LayerShowHide(1)

	DbQuery = False
 
	Err.Clear                                                        
 
	Dim strVal
 
	With frm1
     
	strVal = BIZ_PGM_ID &	"?txtPlantCd="    & Trim(.txtPlantCd.value)		& _
							"&txtYyyyMmDd="   & Trim(.txtYyyyMmDd.text)		& _  
							"&txtItemCd="     & Trim(.txtItemCd.value)		& _
							"&txtInsrtUsrId=" & Parent.gUsrID				& _
							"&lgStrPrevKey=" & lgStrPrevKey				& _
							"&htxtLoginDt="   & Trim(.htxtLoginDt.value)	& _
							"&htxtMvmtFlag=" & Trim(.htxtMvmtFlag.value)	& _
							"&htxtRefNo="     & Trim(.htxtRefNo.value)		& _
							"&htxtQty="       & Trim(.htxtQty.value)		& _
							"&txtMaxRows="    & .vspdData.MaxRows
	Call RunMyBizASP(MyBizASP,strVal)         
 
	End With
 
	DbQuery = True

End Function

Function DbQueryOk()             
	lgBlnFlgChgValue = False
	frm1.vspdData.focus
End Function


</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" --> 
</HEAD>
<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR>
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
								<TD background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></TD>
								<TD background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>Stock Requirement</font></TD>
								<TD background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></TD>
							</TR>
						</TABLE>
					</TD>
					<TD WIDTH=* align=right>&nbsp;</TD>     
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
					<TD HEIGHT=20 >
					<FIELDSET CLASS="CLSFLD">
						<TABLE WIDTH=100% <%=LR_SPACE_TYPE_40%>>
							<TR>
								<TD CLASS="TD5" NOWRAP>공장</TD>      
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT SIZE=6 NAME="txtPlantCd" MAXLENGTH="4" tag="13XXXU" ALT = "공장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd"  align="top" TYPE="BUTTON" ONCLICK="vbscript:OpenPlant()">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=29 MAXLENGTH=40 tag="14"></TD>          
								<TD CLASS="TD5" NOWRAP>기준일자</TD>
								<TD CLASS="TD6" NOWRAP>
								<script language =javascript src='./js/i2213ma1_fpDateTime1_txtYyyyMmDd.js'></script>
								</TD>      
							</TR>
							<TR>      
								<TD CLASS="TD5" NOWRAP>품목</TD>      
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd" SIZE=15 MAXLENGTH=18 tag="13XXXU" ALT = "품목" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" align="top" TYPE="BUTTON" ONCLICK="vbscript:OpenItem()">&nbsp;<INPUT TYPE=TEXT NAME="txtItemNm" SIZE=20 MAXLENGTH=40 tag="14"></TD>
								<TD CLASS="TD5" NOWRAP></TD>
								<TD CLASS="TD6" NOWRAP></TD>
							</TR>
						</TABLE>
					</FIELDSET>
				</TD>
			</TR>
			<TR>
				<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
			</TR>    
			<TR>
				<TD HEIGHT=* WIDTH=100% valign=top>
					<TABLE WIDTH="100%" <%=LR_SPACE_TYPE_60%>>
						<TR>
							<TD CLASS="TD5" NOWRAP>품목</TD>
							<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd2" SIZE=15 MAXLENGTH=18 tag="24" ALT="품목">&nbsp;<INPUT TYPE=TEXT NAME="txtItemNm2" SIZE=20 tag="24"></TD>
							<TD CLASS="TD5" NOWRAP>단위</TD>
							<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtBaseUnit" SIZE=5 MAXLENGTH=3 tag="24" ALT="단위"></TD>           
						</TR>
						<TR>
							<TD CLASS="TD5" NOWRAP>규격</TD>
							<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtItemSpec" SIZE=40 MAXLENGTH=40 tag="24" ALT="규격"></TD>
							<TD CLASS="TD5" NOWRAP>양품재고/안전재고</TD>
							<TD CLASS="TD6" NOWRAP>
							<script language =javascript src='./js/i2213ma1_I780201339_txtOnhandQty.js'></script>&nbsp;
							<script language =javascript src='./js/i2213ma1_fpDoubleSingle1_txtSsQty.js'></script>
							</TD>           
						</TR>
						<TR>
							<TD HEIGHT="100%" WIDTH=100% COLSPAN=4>
							<script language =javascript src='./js/i2213ma1_I867044991_vspdData.js'></script></TD>
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
        </TABLE>
    </TD>
 </TR> 
 <TR>
	<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
	</TD>
 </TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" TABINDEX="-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hPlantCd" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hItemCd" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="htxtLoginDt" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="htxtQty" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="htxtMvmtFlag" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="htxtRefNo" tag="24" TABINDEX="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

