
<%@ LANGUAGE="VBScript" CODEPAGE=949%>
<% session.CodePage=949 %>
<%'======================================================================================================
'*  1. Module Name          : 법인세 
'*  2. Function Name        : 외국세액납부 
'*  3. Program ID           :
'*  4. Program Name         : 
'*  5. Program Desc         : 
'*  6. Modified date(First) : 
'*  7. Modified date(Last)  : 
'*  8. Modifier (First)     : 
'*  9. Modifier (Last)      : 
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../wcm/inc_SvrOperation.asp" -->
<%
'=========================  로긴중인 유저의 법인코드를 출력하기 위해  ======================
    Call LoadBasisGlobalInf()
%>
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"	SRC="../WCM/inc_CliOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../WCM/incEB.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit

'============================================  상수/변수 선언  ====================================

<!-- #Include file="../../inc/lgvariables.inc" -->
Const BIZ_MNU_ID = "w6129MA1"											 '☆: 비지니스 로직 ASP명 
Const BIZ_PGM_ID = "w6129MB1.asp"											 '☆: 비지니스 로직 ASP명 
Const EBR_RPT_ID = "W6115OA1"


Dim IsOpenPop          
Dim lgStrPrevKey2
Dim lgRefMode	' 레퍼런스 참조상태 : ERP 추출사용유무 



'============================================  초기화 함수  ====================================
Sub InitSpreadPosVariables()

    

    
End Sub

Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
    
    lgStrPrevKey = ""                           'initializes Previous Key
    lgStrPrevKey2 = ""                          'initializes Previous Key
    lgLngCurRows = 0                            'initializes Deleted Rows Count
    lgSortKey = 1
    lgRefMode = False
End Sub

Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029A("I", "*","NOCOOKIE","MA") %>
End Sub

'========================================================================================================
' Function Name : MakeKeyStream
' Function Desc : This method set focus to pos of err
'========================================================================================================
Sub MakeKeyStream(pOpt)
	Dim strYear
	Dim strMonth
	Dim strInsurDt
	Dim stReturnrInsurDt

   lgKeyStream = UCASE(Frm1.txtCo_Cd.Value) &  parent.gColSep  
   lgKeyStream = lgKeyStream & (Frm1.txtFISC_YEAR.text) &  parent.gColSep ' 
   lgKeyStream = lgKeyStream & UCASE(Frm1.cboREP_TYPE.Value ) &  parent.gColSep '   
   lgKeyStream = lgKeyStream & Trim(frm1.txtw1_1A.Value ) &  parent.gColSep '  
   lgKeyStream = lgKeyStream & unicdbl(frm1.txtw1_1B.Value ) &  parent.gColSep '
   lgKeyStream = lgKeyStream & Trim(frm1.txtw1_2A.Value ) &  parent.gColSep '  
   lgKeyStream = lgKeyStream & unicdbl(frm1.txtw1_2B.Value ) &  parent.gColSep '
   lgKeyStream = lgKeyStream & Trim(frm1.txtw1_3A.Value ) &  parent.gColSep ' 
   lgKeyStream = lgKeyStream & unicdbl(frm1.txtw1_3B.Value ) &  parent.gColSep ' 
   lgKeyStream = lgKeyStream & Trim(frm1.txtw1_Sum.Value ) &  parent.gColSep ' 
    
  if  frm1.chkW2_4.checked = true then
       lgKeyStream = lgKeyStream & "1" &  parent.gColSep '  
   else
       lgKeyStream = lgKeyStream & "0" &  parent.gColSep '  
   end if
   
   if  frm1.chkW2_5.checked = true then
       lgKeyStream = lgKeyStream & "1" &  parent.gColSep '  
   else
       lgKeyStream = lgKeyStream & "0" &  parent.gColSep '  
   end if
  
   if  frm1.chkW2_6.checked = true then
       lgKeyStream = lgKeyStream & "1" &  parent.gColSep '  
   else
       lgKeyStream = lgKeyStream & "0" &  parent.gColSep '  
   end if
   
   if  frm1.chkW2_7.checked = true then
       lgKeyStream = lgKeyStream & "1" &  parent.gColSep '  
   else
       lgKeyStream = lgKeyStream & "0" &  parent.gColSep '  
   end if
   
   lgKeyStream = lgKeyStream & unicdbl(frm1.txtw3_8.Value ) &  parent.gColSep '
   lgKeyStream = lgKeyStream & unicdbl(frm1.txtw3_9.Value ) &  parent.gColSep '  
   lgKeyStream = lgKeyStream & unicdbl(frm1.txtw3_10.Value ) &  parent.gColSep ' 
   lgKeyStream = lgKeyStream & unicdbl(frm1.txtw3_11.Value ) &  parent.gColSep ' 
   lgKeyStream = lgKeyStream & unicdbl(frm1.txtw3_12.Value ) &  parent.gColSep ' 
   lgKeyStream = lgKeyStream & unicdbl(frm1.txtw3_Sum.Value ) &  parent.gColSep ' 
    
   

End Sub 

'============================================  신고구분 콤보 박스 채우기  ====================================

Sub InitComboBox()
	Dim IntRetCD1
	call CommonQueryRs("MINOR_CD,MINOR_NM"," B_MINOR "," MAJOR_CD = 'W1018' ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboREP_TYPE ,lgF0  ,lgF1  ,Chr(11))
End Sub

Sub InitSpreadSheet()
	
    
End Sub


'============================================  그리드 함수  ====================================





'============================================  조회조건 함수  ====================================

'============================================  폼 함수  ====================================
Sub Form_Load()

    Call LoadInfTB19029     
                                             <%'Load table , B_numeric_format%>
    Call ggoOper.LockField(Document, "N")                                   <%'Lock  Suitable  Field%>                         
   
    Call InitSpreadSheet()                                                    <%'Setup the Spread sheet%>
    Call InitVariables                                                      <%'Initializes local global variables%>
    
    Call SetToolbar("1110100000001111")										<%'버튼 툴바 제어 %>

	' 변경한곳 
    Call InitComboBox
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.FormatField(Document, "2", ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
    Call ggoOper.FormatDate(frm1.txtFISC_YEAR, parent.gDateFormat,3)
    Call SetDefaultVal()

     
    ' 세무조정 체크호출 
	
  
End Sub




Sub SetDefaultVal()
dim strWhere 
DIM strW1

    frm1.txtFISC_YEAR.text 	= "<%=wgFISC_YEAR%>"
	frm1.txtCO_CD.value 	= "<%=wgCO_CD%>"
	frm1.txtCO_NM.value 	= "<%=wgCO_NM%>"
	frm1.cboREP_TYPE.value 	= "<%=wgREP_TYPE%>"

End Sub



Function GetRef()	' 금액가져오기 링크 클릭시 
	

   
End Function



'============================================  이벤트 함수  ====================================
Sub txtFISC_YEAR_DblClick(Button)
    If Button = 1 Then
        frm1.txtFISC_YEAR.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtFISC_YEAR.Focus
    End If
End Sub


Function CalSum()
	   frm1.txtw1_sum.value = unicdbl(frm1.txtw1_1b.value) + unicdbl(frm1.txtw1_2b.value) + unicdbl(frm1.txtw1_3b.value) 
	   frm1.txtw3_sum.value = unicdbl(frm1.txtw3_8.value) + unicdbl(frm1.txtw3_9.value) + unicdbl(frm1.txtw3_10.value) + unicdbl(frm1.txtw3_11.value) + unicdbl(frm1.txtw3_12.value) 
       

end function








Sub txtw1_1b_Change()  
     
    lgBlnFlgChgValue  = True 
   
    Call CalSum() 
    
End Sub

Sub txtw1_2b_Change()  
     
    lgBlnFlgChgValue  = True 
   
    Call CalSum() 
    
End Sub


Sub txtw1_3b_Change()  
     
    lgBlnFlgChgValue  = True 
   
    Call CalSum() 
    
End Sub

Sub txtw1_3b_Change()  
     
    lgBlnFlgChgValue  = True 
   
    Call CalSum() 
    
End Sub

Sub txtw3_8_Change()  
     
    lgBlnFlgChgValue  = True 
   
    Call CalSum() 
    
End Sub

Sub txtw3_9_Change()  
     
    lgBlnFlgChgValue  = True 
   
    Call CalSum() 
    
End Sub

Sub txtw3_10_Change()  
     
    lgBlnFlgChgValue  = True 
   
    Call CalSum() 
    
End Sub

Sub txtw3_11_Change()  
     
    lgBlnFlgChgValue  = True 
   
    Call CalSum() 
    
End Sub

Sub txtw3_12_Change()  
     
    lgBlnFlgChgValue  = True 
   
    Call CalSum() 
    
End Sub

Sub chkW2_4_OnChange()  
     
    lgBlnFlgChgValue  = True 
  
    
End Sub


Sub chkW2_5_OnChange()  
     
    lgBlnFlgChgValue  = True 
   
    
End Sub


Sub chkW2_6_OnChange()  
     
    lgBlnFlgChgValue  = True 
   
  
End Sub


Sub chkW2_7_OnChange()  
     
    lgBlnFlgChgValue  = True 
 
    
End Sub
'==========================================================================================
Sub InitData()
	Dim intRow
	Dim intIndex 
	

End Sub


Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub

Sub PopRestoreSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()      
	Call ggoSpread.ReOrderingSpreadData()
End Sub




'============================================  툴바지원 함수  ====================================
Function FncNew() 
    Dim IntRetCD 
    
    FncNew = False                                                          
    

	If lgBlnFlgChgValue = true Then
		IntRetCD = DisplayMsgBox("900015", parent.VB_YES_NO, "X", "X") 
		If IntRetCD = vbNo Then	Exit Function
    End If

     Call ggoOper.ClearField(Document, "2")
     Call ggoOper.LockField(Document, "N")
     Call InitVariables               

     Call SetToolbar("1110110000001111")          '⊙: 버튼 툴바 제어 
    FncNew = True                

End Function
Function FncQuery() 
    Dim IntRetCD 

    
    FncQuery = False															 '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO,"x","x")					 '☜: Data is changed.  Do you want to display it? 
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    Call ggoOper.LockField(Document, "Q")
    Call  ggoOper.ClearField(Document, "2")										 '☜: Clear Contents  Field
    
    If Not chkField(Document, "1") Then									         '☜: This function check required field
       Exit Function
    End If

    Call InitVariables                                                           '⊙: Initializes local global variables
    Call MakeKeyStream("Q")
    
	Call  DisableToolBar( parent.TBC_QUERY)
    If DbQuery = False Then
		Call  RestoreToolBar()
        Exit Function
    End If
              
    FncQuery = True  
    
End Function

Function FncSave() 
        
    FncSave = False                                                         
    dim IntRetCD
    
    
    

    Err.Clear                                                               <%'☜: Protect system from crashing%>
    'On Error Resume Next                                                   <%'☜: Protect system from crashing%>    
    
<%  '-----------------------
    'Precheck area
    '----------------------- %>
    If lgBlnFlgChgValue = False Then
		IntRetCD =  DisplayMsgBox("900001","x","x","x")  					 '☜: Data is changed.  Do you want to display it? 

			Exit Function

    End If
<%  '-----------------------
    'Save function call area
    '----------------------- %>
    Call ggoOper.LockField(Document, "N")
    Call MakeKeyStream("Q")
    If DbSave = False Then Exit Function                                        '☜: Save db data
  
    FncSave = True                                                          
    
End Function

Function FncCopy() 
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncCopy = False                                                               '☜: Processing is NG


	
    Set gActiveElement = document.ActiveElement   
	
End Function
'========================================================================================================
' Name : FncDelete
' Desc : developer describe this line Called by MainDelete in Common.vbs
'========================================================================================================
Function FncDelete()
    Dim IntRetCd

    FncDelete = False                                                             '☜: Processing is NG
    
    
    <%  '-----------------------
    'Check previous data area
    '----------------------- %>

    If lgBlnFlgChgValue = true Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "X", "X")			    <%'데이타가 변경되었습니다. 조회하시겠습니까?%>
    	If IntRetCD = vbNo Then
      	Exit Function
    	End If
    End If
    
    
    
    If lgIntFlgMode <>  parent.OPMD_UMODE  Then                                            'Check if there is retrived data
        Call  DisplayMsgBox("900002","X","X","X")                                  '☜: Please do Display first.
        Exit Function
    End If

    IntRetCD =  DisplayMsgBox("900003",  parent.VB_YES_NO,"X","X")		                  '☜: Do you want to delete?
	If IntRetCD = vbNo Then
		Exit Function
	End If
    Call MakeKeyStream("Q")
    If DbDelete= False Then
       Exit Function
    End If												                  '☜: Delete db data

    FncDelete=  True                                                              '☜: Processing is OK
End Function


Function FncCancel() 
                                           '☜: Protect system from crashing
End Function

Function FncInsertRow(ByVal pvRowCnt) 
    Dim IntRetCD
    Dim imRow
    Dim iRow, iSeqNo

    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
    Dim uCountID

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status
    
	
    Set gActiveElement = document.ActiveElement   
    
End Function

Function FncDeleteRow() 
    Dim lDelRows

    
End Function

Function FncPrint()
    Call parent.FncPrint()                                                   '☜: Protect system from crashing
End Function

Function FncExcel() 
    Call parent.FncExport(parent.C_MULTI)											 <%'☜: 화면 유형 %>
End Function

Function FncFind() 
    Call parent.FncFind(parent.C_MULTI, False)                                      <%'☜:화면 유형, Tab 유무 %>
End Function

Function FncExit()
Dim IntRetCD

	FncExit = False
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO, "X", "X")

		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    FncExit = True
End Function


'============================================  DB 억세스 함수  ====================================
Function DbQuery() 

    Dim strVal
    Err.Clear                                                                    '☜: Clear err status

    DbQuery = False                                                              '☜: Processing is NG

	If   LayerShowHide(1) = False Then
	     Exit Function
	End If

    strVal = BIZ_PGM_ID & "?txtMode="          &  parent.UID_M0001                       '☜: Query
    strVal = strVal     & "&txtKeyStream="     & lgKeyStream                     '☜: Query Key
    strVal = strVal     & "&txtPrevNext="      & ""	                             '☜: Direction
    Call RunMyBizASP(MyBizASP, strVal)                                           '☜:  Run biz logic

    DbQuery = True                                       
  
End Function

Function DbQueryOk()													<%'조회 성공후 실행로직 %>
	
    Dim iNameArr
    Dim iDx    
    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
    dim IntRetCD

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	 Call ggoOper.LockField(Document, "N")
    '-----------------------
    'Reset variables area
    '-----------------------
    lgIntFlgMode = parent.OPMD_UMODE

    lgBlnFlgChgValue = false
 										<%'버튼 툴바 제어 %>
        Call CheckTaxDoc(frm1.txtCO_CD.value, frm1.txtFISC_YEAR.text, frm1.cboREP_TYPE.value, BIZ_MNU_ID)

	'1 컨펌체크 
	If wgConfirmFlg = "Y" Then
	    Call SetToolbar("1110100000011111")										<%'버튼 툴바 제어 %>
	    
		
	Else
	   '2 디비환경값 , 로드시환경값 비교 
		 Call SetToolbar("1111100000011111")										<%'버튼 툴바 제어 %>
		
	End If
  
		
End Function


'========================================================================================================
' Name : DbSave
' Desc : This function is data query and display
'========================================================================================================
Function DbSave() 


 
    DbSave = False														         '☜: Processing is NG
		
	If   LayerShowHide(1) = False Then
	     Exit Function
	End If

	With Frm1
		.txtMode.value        =  parent.UID_M0002                                        '☜: Delete
		.txtFlgMode.value     = lgIntFlgMode
        .txtKeyStream.Value   = lgKeyStream                                      '☜: Save Key
	End With

	Call ExecMyBizASP(frm1, BIZ_PGM_ID)
		
    DbSave  = True                                            
End Function


Function DbSaveOk()													        <%' 저장 성공후 실행 로직 %>
	Call InitVariables
	

    Call MainQuery()
End Function

'========================================================================================================
' Name : DbDelete
' Desc : This function is called by FncDelete
'========================================================================================================
Function DbDelete()
	Dim strVal
    Err.Clear                                                                    '☜: Clear err status

	DbDelete = False			                                                 '☜: Processing is NG

    If LayerShowHide(1) = false Then
        Exit Function
    End If

	With Frm1
		.txtMode.value        =  parent.UID_M0003                                        '☜: Delete
		.txtFlgMode.value     = lgIntFlgMode
        .txtKeyStream.Value   = lgKeyStream                                      '☜: Save Key
	End With

		Call ExecMyBizASP(frm1, BIZ_PGM_ID)
	DbDelete = True                                                              '⊙: Processing is NG

End Function
'========================================================================================================
' Function Name : DbDeleteOk
' Function Desc : Called by MB Area when delete operation is successful
'========================================================================================================
Function DbDeleteOk()
	Call InitVariables()
	Call MainNew()
End Function
'======================================================================================================
' Function Name : SubSetErrPos
' Function Desc : This method set focus to pos of err
'======================================================================================================
Sub SubSetErrPos(iPosArr)
    Dim iDx
    Dim iRow
    iPosArr = Split(iPosArr, parent.gColSep)
    If IsNumeric(iPosArr(0)) Then
       iRow = CInt(iPosArr(0))
       For iDx = 1 To  frm1.vspdData.MaxCols - 1
           Frm1.vspdData.Col = iDx
           Frm1.vspdData.Row = iRow
           If Frm1.vspdData.ColHidden <> True And Frm1.vspdData.BackColor <>  parent.UC_PROTECTED Then
              Frm1.vspdData.Col = iDx
              Frm1.vspdData.Row = iRow
              Frm1.vspdData.Action = 0 ' go to
              Exit For
           End If

       Next

    End If
End Sub

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->	
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
					<TD CLASS="CLSMTABP_BAK">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" NOWRAP><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB_BAK"><font color=white><%=Request("strASPMnuMnuNm")%></font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=*>&nbsp;</TD>
					<TD WIDTH=* ALIGN=RIGHT>&nbsp;
						<a href="vbscript:GetRef">금액 불러오기</A>  
					</TD>
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
									<TD CLASS="TD5">사업연도</TD>
									<TD CLASS="TD6"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name=txtFISC_YEAR CLASS=FPDTYYYY title=FPDATETIME ALT="사업연도" tag="14X1" id=txtFISC_YEAR></OBJECT>');</SCRIPT>
									<TD CLASS="TD5">법인명</TD>
									<TD CLASS="TD6"><INPUT TYPE=TEXT NAME="txtCO_CD" Size=10 tag="14">
										<INPUT TYPE=TEXT NAME="txtCO_NM" Size=20 tag="14">
									</TD>
									</TD>
								</TR>
								<TR>
									<TD CLASS="TD5">신고구분</TD>
									<TD CLASS="TD6"><SELECT NAME="cboREP_TYPE" ALT="신고구분" STYLE="WIDTH: 50%" tag="14X1"></SELECT>
									</TD>
									<TD CLASS="TD5"></TD>
									<TD CLASS="TD6"></TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=*> </TD>
					
				</TR>
				<TR>
					<TD WIDTH=620 valign=top  >
				          <TABLE   border = 0 cellpadding = 1 cellspacing = 1 ID="Table2" WIDTH =  100%>
				                     <TR>
										<TD WIDTH=620 valign=top  >
										   
										       <FIELDSET CLASS="CLSFLD"><LEGEND ALIGN=LEFT>1.외국남부세액공제 등 대상</LEGEND>
														<TABLE bgcolor = #696969  border = 0 cellpadding = 1 cellspacing = 1 ID="Table1" WIDTH =  100%>
														   
														
															<TR>
																 <TD CLASS="TD51" align =CENTER  >
																	관련근거법 
																</TD>
																
															    <TD bgcolor =  #d1e8f9 align =CENTER >
																	(101)외국납부세목 
																</TD>
																
																<TD bgcolor =  #d1e8f9 align =CENTER>
																	(102)금액(원화)
																</TD>
																
															</TR>
															
															
															<TR>
																 <TD CLASS="TD51" align =LEFT>
																	(1)법 제 57조 제 1항에 의한 외국납부세액 
																</TD>
																
																 <TD bgcolor =#eeeeec align =CENTER  >
																	 <INPUT TYPE=TEXT NAME="txtw1_1A"   tag="25"  maxlength=100 size=25>
																</TD>
																
																 <TD bgcolor =#eeeeec align =CENTER  >
																	<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtw1_1B" name=txtw1_1B CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="25X26" width = 100% ></OBJECT>');</SCRIPT>
																</TD>
															</TR>
															
															<TR>
																 <TD CLASS="TD51" align =LEFT>
																	(2)법 제 57조 제 3항에 의한 외국납부세액 
																</TD>
																
																 <TD bgcolor =#eeeeec align =CENTER  >
																	 <INPUT TYPE=TEXT NAME="txtw1_2A"   tag="25"  maxlength=100  size=25>
																</TD>
																
																 <TD bgcolor =#eeeeec align =CENTER  >
																	<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtw1_2B" name=txtw1_2B CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="25X26" width = 100% ></OBJECT>');</SCRIPT>
																</TD>
															</TR>
															<TR>
																 <TD CLASS="TD51" align =LEFT>
																	(3)법 제 57조 제 4항에 의한 외국납부세액 
																</TD>
																
																 <TD bgcolor =#eeeeec align =CENTER  >
																	 <INPUT TYPE=TEXT NAME="txtw1_3A"   tag="25"  maxlength=100 size=25 >
																</TD>
																
																 <TD bgcolor =#eeeeec align =CENTER  >
																	<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtw1_3B" name=txtw1_3B CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="25X26" width = 100% ></OBJECT>');</SCRIPT>
																</TD>
															</TR>
															<TR>
																 <TD CLASS="TD51" align =CENTER>
																	합계 
																</TD>
																
																 <TD bgcolor =#eeeeec align =CENTER  >
																	 
																</TD>
																
																 <TD CLASS="TD51" align =CENTER  >
																	<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW1_Sum" name=txtW1_Sum CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="24X26" width = 100% ></OBJECT>');</SCRIPT>
																</TD>
															</TR>
										
																
														</TABLE>
											   </FIELDSET>	
											   			
										</TD>
									</TR>
								   <TR>
										<TD WIDTH=620 valign=top  >
										   
										       <FIELDSET CLASS="CLSFLD"><LEGEND ALIGN=LEFT>2.선택한 외국납부세액공제 등 방법(해당란에 O표시)</LEGEND>
														<TABLE bgcolor = #696969  border = 0 cellpadding = 1 cellspacing = 1 ID="Table1" WIDTH =  100%>
														    
														
															<TR>
															      <TD CLASS="TD51" align =LEFT WIDTH =  50%>
																	(4)세액공제(일괄하여 계산하는 방법)
															     </TD>
																 <TD bgcolor = #eeeeec   align =CENTER  >
																	<INPUT TYPE=CHECKBOX NAME="chkW2_4" ID="chkW2_4" tag="25" Class="Check"  value=0>
																 </TD>
																
																
															</TR>
															
															
															<TR>
															      <TD CLASS="TD51" align =LEFT WIDTH =  50%>
																	(5)세액공제(국가별로 계산하는 방법)
															     </TD>
																 <TD bgcolor = #eeeeec   align =CENTER  >
																	<INPUT TYPE=CHECKBOX NAME="chkW2_5" ID="chkW2_5" tag="25" Class="Check"  value=0>
																 </TD>
																
																
															</TR>
															
															
															
															<TR>
															      <TD CLASS="TD51" align =LEFT WIDTH =  50%>
																	(6)손금산입(신고조정)
															     </TD>
																 <TD bgcolor = #eeeeec   align =CENTER  >
																	<INPUT TYPE=CHECKBOX NAME="chkW2_6" ID="chkW2_6" tag="25" Class="Check"  value=0>
																 </TD>
																
																
															</TR>
															
																<TR>
															      <TD CLASS="TD51" align =LEFT WIDTH =  50%>
																	(7)손금산입(결산조정)
															     </TD>
																 <TD bgcolor = #eeeeec   align =CENTER  >
																	<INPUT TYPE=CHECKBOX NAME="chkW2_7" ID="chkW2_7" tag="25" Class="Check" value=0>
																 </TD>
																
																
															</TR>
														</TABLE>
											   </FIELDSET>	
											   			
										</TD>
									</TR>					
															
														 <TR>
															<TD WIDTH=620 valign=top  >
																   
															       <FIELDSET CLASS="CLSFLD"><LEGEND ALIGN=LEFT>3.손금 산입명세</LEGEND>
																			<TABLE bgcolor = #696969  border = 0 cellpadding = 1 cellspacing = 1 ID="Table3" WIDTH =  100%>
																				   
																				
																				<TR>
																					 <TD CLASS="TD51" align =CENTER  WIDTH =  50% >
																						구분 
																					</TD>
																						
																				    <TD bgcolor =  #d1e8f9 align =CENTER >
																						금액 
																					</TD>
																						
																				</TR>
																					
																					
																				<TR>
																					 <TD CLASS="TD51" align =LEFT>
																						(8)신고조정에 의한 손금산입 
																					</TD>
																						
																						
																					 <TD bgcolor =#eeeeec align =CENTER  >
																						<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW3_8" name=txtW3_8 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="25X26" width = 100% ></OBJECT>');</SCRIPT>
																					</TD>
																				</TR>
																					
																				<TR>
																					 <TD CLASS="TD51" align =LEFT>
																						(9)매출원가 계상 
																					</TD>
																						
																						 
																					 <TD bgcolor =#eeeeec align =CENTER  >
																						<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW3_9" name=txtW3_9 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="25X26" width = 100% ></OBJECT>');</SCRIPT>
																					</TD>
																				</TR>
																				<TR>
																					 <TD CLASS="TD51" align =LEFT>
																						(10)판매비와 일반관리비에 계상 
																					</TD>
																						
																						
																					 <TD bgcolor =#eeeeec align =CENTER  >
																						<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW3_10" name=txtW3_10 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="25X26" width = 100% ></OBJECT>');</SCRIPT>
																					</TD>
																				</TR>
																				<TR>
																					 <TD CLASS="TD51" align =LEFT>
																						(11)제조원가에 계상 
																					</TD>
																						
																					 <TD CLASS="TD51" align =CENTER  >
																						<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW3_11" name=txtW3_11 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="25X26" width = 100% ></OBJECT>');</SCRIPT>
																					</TD>
																				</TR>
																				<TR>
																					 <TD CLASS="TD51" align =LEFT>
																						(12)기타계정과목에 계상 
																					</TD>
																						
																					 <TD CLASS="TD51" align =CENTER  >
																						<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW3_12" name=txtW3_12 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="25X26" width = 100% ></OBJECT>');</SCRIPT>
																					</TD>
																				</TR>
																				<TR>
																					 <TD CLASS="TD51" align =CENTER>
																						합계 
																					</TD>
																						
																					 <TD CLASS="TD51" align =CENTER  >
																						<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtWSUM" name=txtW3_SUM CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="24X26" width = 100% ></OBJECT>');</SCRIPT>
																					</TD>
																				</TR>
																
																						
																			</TABLE>
																   </FIELDSET>	
																	   			
															</TD>
														</TR>															
															
										
																
									
						  </TABLE>
					</TD>
				</tr>			
		        
			    	
			</TABLE>
		</TD>
	</TR>
	
	<TR>
		<TD <%=HEIGHT_TYPE_01%>></TD>
	</TR>
		<TR HEIGHT=20>
		<TD WIDTH=100%>
			<TABLE CLASS="TB3" CELLSPACING=0>
	
		
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD><BUTTON NAME="bttnPreview"  CLASS="CLSSBTN" ONCLICK="vbscript:FncBtnPrint('VIEW')" Flag=1>미리보기</BUTTON>&nbsp;
						<BUTTON NAME="bttnPrint"	CLASS="CLSSBTN" ONCLICK="vbscript:FncBtnPrint('PRINT')"   Flag=1>인쇄</BUTTON></TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">

<INPUT TYPE=HIDDEN NAME="txtKeyStream" tag="24" >
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24" >
<INPUT TYPE=HIDDEN NAME="txtW154" tag="24" ALT="">
<INPUT TYPE=HIDDEN NAME="txtW134" tag="24" ALT="">
<INPUT TYPE=HIDDEN NAME="txtW150" tag="24" ALT="">


</FORM>
<DIV ID="MousePT" NAME="MousePT">

<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

