<%@ LANGUAGE="VBSCRIPT" %>
<!--'********************************************************************************************************
'*  1. Module Name          : Inventory       
'*  2. Function Name        : LOT Popup
'*  3. Program ID           : i2511pa1.asp     
'*  4. Program Name         :       
'*  5. Program Desc         :      
'*  7. Modified date(First) : 2000/10/09     
'*  8. Modified date(Last)  : 2000/10/09     
'*  9. Modifier (First)     : Kim Nam Hoon
'* 10. Modifier (Last)      : Lee Seung Wook
'* 11. Comment              :         
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change" 
'*                            this mark(⊙) Means that "may  change" 
'*                            this mark(☆) Means that "must change"   
'* 13. History              :       
'*                            2000/10/09 : 4th Iteration
'********************************************************************************************************-->
<HTML>
<HEAD>
<!--'########################################################################################################
'#      1. 선 언 부                  #
'########################################################################################################
'********************************************  1.1 Inc 선언  ********************************************
'* Description : Inc. Include        
'********************************************************************************************************-->
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!--'============================================  1.1.1 Style Sheet  =======================================
'========================================================================================================-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">
<!--'============================================  1.1.2 공통 Include  ======================================
'========================================================================================================-->
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliPAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliPAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliPAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript" SRC="../../inc/incImage.js"></SCRIPT>
<Script Language="VBS">
Option Explicit
Const BIZ_PGM_ID = "i2511pb1.asp"       

Dim C_LotNo     
Dim C_LotSubNo  
Dim C_LotGenDt  
Dim C_OrderType 
Dim C_TrackingNo

'*********************************************  1.3 변 수 선 언  ****************************************
'* 설명: Constant는 반드시 대문자 표기.                *
'********************************************************************************************************
<!-- #Include file="../../inc/lgvariables.inc" -->
Dim lgStrPrevSubKey

Dim arrParam     
Dim arrReturn    

arrParam		= window.dialogArguments
Set PopupParent = arrParam(0)

top.document.title = PopupParent.gActivePRAspName

'==========================================  2.2.1 SetDefaultVal()  =====================================
'= Name : SetDefaultVal()                    =
'= Description : 화면 초기화(수량 Field나 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)  =
'========================================================================================================
Sub SetDefaultVal()
	txthPlantCd.value	= arrParam(1)
	txtItemCd.Value		= arrParam(2)
	txtLotNo.value		= arrParam(3)
	txtLotSubNo.Value	= arrParam(4) 
	txtItemNm.Value		= arrParam(5)
 
	Self.Returnvalue = Array("")
End Sub

'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'========================================================================================
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029A("Q", "I","NOCOOKIE","PA") %>
End Sub

'==========================================  2.2.2 InitSpreadSheet()  ===================================
'= Name : InitSpreadSheet()                   =
'= Description : This method initializes spread sheet column property         =
'========================================================================================================
Sub InitSpreadSheet()

	Call InitSpreadPosVariables()
	ggoSpread.Source = vspdData
	ggoSpread.Spreadinit "V20021106", , PopupParent.gAllowDragDropSpread

	With vspdData
	 
		.ReDraw = false
		.MaxCols = C_TrackingNo+1        
		.MaxRows = 0
		Call GetSpreadColumnPos("A")
				  
		ggoSpread.SSSetEdit C_LotNo, "Lot No.", 20, 12
		ggoSpread.SSSetEdit C_LotSubNo, "지번", 12, 3 
		ggoSpread.SSSetDate C_LotGenDt, "입고일", 14, 20, PopupParent.gDateFormat
		ggoSpread.SSSetEdit C_OrderType, "오더구분", 11, 3
		ggoSpread.SSSetEdit C_TrackingNo, "Tracking No", 20, 16 

		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
		ggoSpread.SSSetSplit2(2)
		.ReDraw = true
		  
		Call SetSpreadLock
	End With

End Sub

'========================================================================================
' Function Name : InitSpreadPosVariables
' Function Desc : 
'======================================================================================== 
Sub InitSpreadPosVariables()
	C_LotNo      = 1
	C_LotSubNo   = 2
	C_LotGenDt   = 3
	C_OrderType  = 4
	C_TrackingNo = 5
End Sub

'========================================================================================
' Function Name : GetSpreadColumnPos
' Function Desc : 
'======================================================================================== 
Sub GetSpreadColumnPos(ByVal pvSpdNo)
	Dim iCurColumnPos
	
	Select Case UCase(pvSpdNo)
	Case "A"
		ggoSpread.Source = vspdData 
		
		Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

		C_LotNo      = iCurColumnPos(1)
		C_LotSubNo   = iCurColumnPos(2)
		C_LotGenDt   = iCurColumnPos(3)
		C_OrderType  = iCurColumnPos(4)
		C_TrackingNo = iCurColumnPos(5)
	End Select

End Sub

'================================== 2.2.4 SetSpreadLock() ==================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadLock()
     vspdData.ReDraw = False    
     ggoSpread.SpreadLockWithOddEvenRowColor()
 vspdData.ReDraw = True
End Sub

'===========================================  2.3.1 OkClick()  ==========================================
' Name : OkClick()                     
'= Description : Return Array to Opener Window when OK button click        
'========================================================================================================
Function OKClick()
 Dim intColCnt
 
 If vspdData.ActiveRow > 0 Then 
  Redim arrReturn(vspdData.MaxCols - 1)
 
  vspdData.Row = vspdData.ActiveRow
    
	vspdData.Col = C_LotNo
	arrReturn(0) = vspdData.Text
	vspdData.Col = C_LotSubNo
	arrReturn(1) = vspdData.Text
	vspdData.Col = C_LotGenDt
	arrReturn(2) = vspdData.Text
	vspdData.Col = C_OrderType
	arrReturn(3) = vspdData.Text
	vspdData.Col = C_TrackingNo
	arrReturn(4) = vspdData.Text

   
  Self.Returnvalue = arrReturn
 End If
 
 Self.Close()
End Function

'=========================================  2.3.2 CancelClick()  ========================================
'= Name : CancelClick()                    =
'= Description : Return Array to Opener Window for Cancel button click         =
'========================================================================================================
Function CancelClick()
 Self.Close()
End Function

'=========================================  3.1.1 Form_Load()  ==========================================
'= Name : Form_Load()                     =
'= Description : Window Load시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분    =
'========================================================================================================
Sub Form_Load() 

 Call MM_preloadImages("../../../CShared/image/Query.gif","../../../CShared/image/OK.gif","../../../CShared/image/Cancel.gif")
 Call ggoOper.LockField(Document, "N")                                 
 
 Call SetDefaultVal()
 Call InitSpreadSheet()
 Call FncQuery()
End Sub

Function FncQuery()
    ggoSpread.Source = vspdData
	ggoSpread.ClearSpreadData
	
	lgStrPrevKey = ""
	 
	If Not chkField(Document, "1") Then
		Exit Function
	End If
	 
	Call DbQuery()
End Function

'========================================================================================
' Function Name : vspdData_Click
' Function Desc : 그리드 헤더 클릭시 정렬 
'========================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
	Call SetPopupMenuItemInf("0000111111")
	gMouseClickStatus = "SPC"   
   
	Set gActiveSpdSheet = vspdData
   
	If vspdData.MaxRows = 0 Then
		Exit Sub
	End If
	
	If Row <= 0 Then
		ggoSpread.Source = vspdData 
		If lgSortKey = 1 Then
			ggoSpread.SSSort Col					
			lgSortKey = 2
		Else
			ggoSpread.SSSort Col, lgSortKey		
			lgSortKey = 1
		End If
	End If
End Sub

'========================================================================================
' Function Name : vspdData_DblClick
' Function Desc : 그리드 해더 더블클릭시 네임 변경 
'========================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)
	Dim iColumnName
   
	If Row <= 0 Then
		Exit Sub
	End If
	If vspdData.MaxRows = 0 Then
		Exit Sub
	End If
	'------ Developer Coding part (Start)
	If vspdData.MaxRows > 0 Then
		If vspdData.ActiveRow = Row Or vspdData.ActiveRow > 0 Then
			Call OKClick()
		End If
	End If
	'------ Developer Coding part (End)
End Sub

'========================================================================================
' Function Name : vspdData_MouseDown
' Function Desc : 
'========================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)

   If Button = 2 And gMouseClickStatus = "SPC" Then
      gMouseClickStatus = "SPCR"
   End If
End Sub 

'========================================================================================
' Function Name : vspdData_ColWidthChange
' Function Desc : 그리드 폭조정 
'========================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)

   ggoSpread.Source = vspdData
   Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub 

'========================================================================================
' Function Name : vspdData_ScriptDragDropBlock
' Function Desc : 그리드 위치 변경 
'========================================================================================
Sub vspdData_ScriptDragDropBlock( Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)

   ggoSpread.Source = vspdData
   Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
   Call GetSpreadColumnPos("A")
End Sub 

'========================================================================================
' Function Name : PopSaveSpreadColumnInf
' Function Desc : 그리드 현상태를 저장한다.
'========================================================================================
Sub PopSaveSpreadColumnInf()

   ggoSpread.Source = gActiveSpdSheet
   Call ggoSpread.SaveSpreadColumnInf()
End Sub 

'========================================================================================
' Function Name : PopRestoreSpreadColumnInf
' Function Desc : 그리드 현상태를 저장한다.
'========================================================================================
Sub PopRestoreSpreadColumnInf()

   ggoSpread.Source = gActiveSpdSheet
   Call ggoSpread.RestoreSpreadInf()
   Call InitSpreadSheet
   Call ggoSpread.ReOrderingSpreadData
End Sub 



Function vspdData_KeyPress(KeyAscii)
	On Error Resume Next 
	If KeyAscii = 13 And vspdData.ActiveRow > 0 Then
		Call OKClick()
	ElseIf KeyAscii = 27 Then
		Call CancelClick()
	End If
End Function

Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
	If OldLeft <> NewLeft Then
		Exit Sub
	End If
	If vspdData.MaxRows < NewTop + VisibleRowCnt(vspdData, NewTop) And lgStrPrevKey <> "" Then
		DbQuery
	End if
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
 
'********************************************  5.1 DbQuery()  *******************************************
' Function Name : DbQuery                    *
' Function Desc : This function is data query and display            *
'********************************************************************************************************
Function DbQuery()

 Dim strVal
 Dim txtMaxRows
 

    Call LayerShowHide(1)  

 DbQuery = False 
 txtMaxRows = vspdData.MaxRows
 
 strVal = BIZ_PGM_ID &	"?txtLotNo="		& Trim(txtLotNo.Value)		& _
						"&txtLotSubNo="		& Trim(txtLotSubNo.Value)	& _
						"&txtItemCd="		& Trim(txtItemCd.Value)		& _
						"&txtPlantCd="		& Trim(txthPlantCd.Value)	& _
						"&lgStrPrevKey="	& lgStrPrevKey				& _
						"&lgStrPrevSubKey=" & lgStrPrevSubKey			& _
						"&txtMaxRows="       & txtMaxRows
 Call RunMyBizASP(MyBizASP, strVal)         
 DbQuery = True                                                          
 
End Function

Function DbQueryOk()       
	vspdData.focus 
End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" --> 
</HEAD>
<BODY SCROLL=NO TABINDEX="-1">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR HEIGHT=*>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS="TD5" NOWRAP>LOT NO</TD>
									<TD CLASS="TD6" NOWRAP><input NAME="txtLotNo" TYPE="Text" SIZE="12" MAXLENGTH="12" STYLE="Text-Transform: uppercase" tag="11" ALT = "LOT NO">&nbsp;<input NAME="txtLotSubNo" TYPE="Text" SIZE="5" MAXLENGTH="3" STYLE="Text-Transform: uppercase" tag="11x3" ALT = "지번"></TD>
									<TD CLASS="TD5" NOWRAP>품목</TD>
									<TD CLASS="TD6" NOWRAP>
									<input TYPE=TEXT NAME="txtItemCd" SIZE="15" MAXLENGTH="18" ALT="품목" tag="14">&nbsp;<input TYPE=TEXT NAME="txtItemNm" SIZE="20" tag="14" ></TD>     
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
								<TD WIDTH=100% HEIGHT=100%>
								<script language =javascript src='./js/i2511pa1_I581524109_vspdData.js'></script>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_01%>></TD>
				</TR>
				<TR HEIGHT=20>
					<TD WIDTH=100%>
						<TABLE <%=LR_SPACE_TYPE_30%>>
							<TR>
								<TD WIDTH=10>&nbsp;</TD>
								<TD WIDTH=70% NOWRAP>     <IMG SRC="../../../CShared/image/query_d.gif"  Style="CURSOR: hand" ALT="Search" NAME="Search" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Query.gif',1)"  ONCLICK="FncQuery()"   ></IMG></TD>
								<TD WIDTH=30% ALIGN=RIGHT><IMG SRC="../../../CShared/image/ok_d.gif"     Style="CURSOR: hand" ALT="OK"     NAME="pop1"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/OK.gif',1)"     ONCLICK="OkClick()"    ></IMG>
								                    <IMG SRC="../../../CShared/image/cancel_d.gif" Style="CURSOR: hand" ALT="CANCEL" NAME="pop2"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)" ONCLICK="CancelClick()"></IMG>     </TD>
								<TD WIDTH=10>&nbsp;</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
					</TD>
				</TR>
			</TABLE>
			<INPUT TYPE=HIDDEN NAME="txthPlantCd" tag="24" TABINDEX="-1">
			<DIV ID="MousePT" NAME="MousePT">
			<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
		</DIV>
	</BODY>
</HTML>

