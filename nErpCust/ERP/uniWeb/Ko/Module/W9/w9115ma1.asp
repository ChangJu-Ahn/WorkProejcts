
<%@ LANGUAGE="VBSCRIPT" %>
<%'======================================================================================================

'*  1. Module Name          : 법인세 
'*  2. Function Name        : 제63호감가상각방법신고서 
'*  3. Program ID           : W9115MA1
'*  4. Program Name         : W9115MA1.asp
'*  5. Program Desc         : 제63호감가상각방법신고서 
'*  6. Modified date(First) : 2005/03/18
'*  7. Modified date(Last)  : 2006/02/02
'*  8. Modifier (First)     : 홍지영 
'*  9. Modifier (Last)      : hjo 
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
<SCRIPT LANGUAGE="VBScript"	SRC="../WCM/inc_CliGrid.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"	SRC="../WCM/inc_CliOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../WCM/incEB.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit

'============================================  상수/변수 선언  ====================================

<!-- #Include file="../../inc/lgvariables.inc" -->
Const BIZ_MNU_ID = "w9115ma1"	
Const BIZ_PGM_ID = "w9115mb1.asp"											 '☆: 비지니스 로직 ASP명 
Const EBR_RPT_ID = "W9115OA1"



Dim C_SEQ_NO
Dim C_W8
Dim C_W9_Fr
Dim C_W9
Dim C_W9_To
Dim C_W10
Dim C_W11
dim C_W12

dim strMode

Dim IsOpenPop          
Dim lgStrPrevKey2
Dim lgRefMode	' 레퍼런스 참조상태 : ERP 추출사용유무 

'============================================  초기화 함수  ====================================
Sub InitSpreadPosVariables()
	 C_SEQ_NO	= 1
     C_W8		= 2
     C_W9_Fr	= 3
	 C_W9		= 4
	 C_W9_To	= 5
	 C_W10		= 6
	 C_W11		= 7
	 C_W12		= 8
	
	

    
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



'============================================  신고구분 콤보 박스 채우기  ====================================

Sub InitComboBox()
	Dim IntRetCD
	call CommonQueryRs("MINOR_CD,MINOR_NM"," B_MINOR "," MAJOR_CD = 'W1018' ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboREP_TYPE ,"" & Chr(11) & lgF0  ,"" & Chr(11) &  lgF1  ,Chr(11))
    
    
 
    
    call CommonQueryRs("MINOR_CD,MINOR_NM","   dbo.ufn_TB_MINOR('W1087', '" & C_REVISION_YM & "')  ","  ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboW13_A  ,"" & Chr(11) & lgF0  ,"" & Chr(11) &  lgF1  ,Chr(11))
    
    call CommonQueryRs("MINOR_CD,MINOR_NM"," dbo.ufn_TB_MINOR('W1087', '" & C_REVISION_YM & "')  "," ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboW13_B  ,"" & Chr(11) & lgF0  ,"" & Chr(11) &  lgF1  ,Chr(11))
    
    call CommonQueryRs("MINOR_CD,MINOR_NM"," dbo.ufn_TB_MINOR('W1088', '" & C_REVISION_YM & "') ","  ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboW14_A  ,"" & Chr(11) & lgF0  ,"" & Chr(11) &  lgF1  ,Chr(11))
    
    call CommonQueryRs("MINOR_CD,MINOR_NM"," dbo.ufn_TB_MINOR('W1088', '" & C_REVISION_YM & "') "," ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboW14_B  ,"" & Chr(11) & lgF0  ,"" & Chr(11) &  lgF1  ,Chr(11))
    
    
    call CommonQueryRs("MINOR_CD,MINOR_NM"," dbo.ufn_TB_MINOR('W1042', '" & C_REVISION_YM & "') ","  ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboW15_A  ,"" & Chr(11) & lgF0  ,"" & Chr(11) &  lgF1  ,Chr(11))
     call CommonQueryRs("MINOR_CD,MINOR_NM"," dbo.ufn_TB_MINOR('W1042', '" & C_REVISION_YM & "') ","  ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboW15_B  ,"" & Chr(11) & lgF0  ,"" & Chr(11) &  lgF1  ,Chr(11))
End Sub




Sub SetDefaultVal()

	frm1.txtFISC_YEAR.text 	= "<%=wgFISC_YEAR%>"
	frm1.txtCO_CD.value 	= "<%=wgCO_CD%>"
	frm1.txtCO_NM.value 	= "<%=wgCO_NM%>"
	frm1.cboREP_TYPE.value 	= "<%=wgREP_TYPE%>"

     
    Call GetRef()
    wgRefDoc = GetDocRef(frm1.txtCO_CD.value, frm1.txtFISC_YEAR.text, frm1.cboREP_TYPE.value, BIZ_MNU_ID)



End Sub
Sub InitSpreadSheet()

	Call initSpreadPosVariables()  

					      	
	With frm1.vspdData
		
		ggoSpread.Source = frm1.vspdData	
		'patch version
		ggoSpread.Spreadinit "V20041222",,parent.gForbidDragDropSpread    
								 
		.ReDraw = false

		.MaxCols = C_W12 + 1							'☜: 최대 Columns의 항상 1개 증가시킴 
		.Col = .MaxCols														'☆: 사용자 별 Hidden Column
		.ColHidden = True    
									       
		.MaxRows = 0
		ggoSpread.ClearSpreadData

		Call AppendNumberPlace("6","3","2")
		.RowHeight(0) = 30 
								 
		' Call GetSpreadColumnPos("A")   					 

		ggoSpread.SSSetEdit     C_SEQ_NO, "순번", 10,,,100,1
		ggoSpread.SSSetEdit     C_W8, "(8)자산 및 업종명", 15,,,100,1
		ggoSpread.SSSetMask     C_W9_Fr,"내용연수" & vbCrLf & "(From)", 10, 2 ,"99년"
		ggoSpread.SSSetEdit     C_W9, "~", 1,2,,1,1
		ggoSpread.SSSetMask     C_W9_to,"내용연수" & vbCrLf & "(To)", 10, 2 ,"99년"
		ggoSpread.SSSetMask     C_W10, "(10)신고내용" & vbCrLf & "연수",10, 2 ,"99년"
		ggoSpread.SSSetMask     C_W11, "(11)변경내용" & vbCrLf & "연수",10, 2 ,"99년"
		ggoSpread.SSSetEdit     C_W12, "(12)변경사유", 20,,,50
								
		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
		Call ggoSpread.SSSetColHidden(C_SEQ_NO,C_SEQ_NO,True)						
								
		.ReDraw = true				 
		Call SetSpreadLock					 
	End With    
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
	Dim strChkW1 , strChkW2 , strChkW3  , strChkW4, strChkW5
			
	if 	pOpt = "S" then
	         if frm1.chkW1.checked = true  then
	            strChkW1 = 1
	         else
	            strChkW1 =0 
	         end if     
	         
	          if frm1.chkW2.checked = true   then
	            strChkW2 = 1
	         else
	            strChkW2 =0 
	         end if     
	           
	         if frm1.chkW3.checked = true   then
	            strChkW3 = 1
	         else
	            strChkW3 =0 
	         end if    
	         
	         
	         if frm1.chkW4.checked = true   then
	            strChkW4 = 1
	         else
	            strChkW4 =0 
	         end if 
	         
	         if frm1.chkW5.checked = true   then
	            strChkW5 = 1
	         else
	            strChkW5 =0 
	         end if        

		    lgKeyStream = UCASE(Frm1.txtCo_Cd.Value) &  parent.gColSep  
			lgKeyStream = lgKeyStream & (Frm1.txtFISC_YEAR.text) &  parent.gColSep ' 
			lgKeyStream = lgKeyStream & UCASE(Frm1.cboREP_TYPE.Value ) &  parent.gColSep '
          
			lgKeyStream = lgKeyStream & strChkW1 &  parent.gColSep		' 3   감가상각방법신고서 
			lgKeyStream = lgKeyStream & strChkW2 &  parent.gColSep		' 4   감가상각방법변경신청서               
			lgKeyStream = lgKeyStream & strChkW3 &  parent.gColSep		' 5   내용연수신고서 
			lgKeyStream = lgKeyStream & strChkW4 &  parent.gColSep		' 6   내용연수승인(변경신청서)
			lgKeyStream = lgKeyStream & strChkW5 &  parent.gColSep		' 7   내용연수승인변경신고서 
			lgKeyStream = lgKeyStream &  Trim(Frm1.txtChange_DT.text )  &  parent.gColSep			' 8   변경일자 
			lgKeyStream = lgKeyStream & Trim(Frm1.cboW13_A.Value ) &  parent.gColSep		' 9 13_신고상각방법 
			lgKeyStream = lgKeyStream & Trim(Frm1.cboW13_B.Value ) &  parent.gColSep		' 10 13_변경신고방법 
			lgKeyStream = lgKeyStream & Trim(Frm1.txtW13_C.Value ) &  parent.gColSep		' 11 13_내용 
			lgKeyStream = lgKeyStream & Trim(Frm1.cboW14_A.Value ) &  parent.gColSep		' 12 14_신고상각방법 
			lgKeyStream = lgKeyStream & Trim(Frm1.cboW14_B.Value ) &  parent.gColSep		' 13 14_변경신고방법 
			lgKeyStream = lgKeyStream & Trim(Frm1.txtW14_C.Value ) &  parent.gColSep		' 14 14_내용 
			lgKeyStream = lgKeyStream & Trim(Frm1.cboW15_A.Value ) &  parent.gColSep		' 15 15_신고상각방법 
			lgKeyStream = lgKeyStream & Trim(Frm1.cboW15_B.Value ) &  parent.gColSep		' 16 15_변경신고방법 
			lgKeyStream = lgKeyStream & Trim(Frm1.txtW15_C.Value ) &  parent.gColSep		' 17 15_내용 
			
	
	Else
	        lgKeyStream = UCASE(Frm1.txtCo_Cd.Value) &  parent.gColSep  
			lgKeyStream = lgKeyStream & (Frm1.txtFISC_YEAR.text) &  parent.gColSep ' 
			lgKeyStream = lgKeyStream & UCASE(Frm1.cboREP_TYPE.Value ) &  parent.gColSep '
	End if		
    
End Sub 


'============================================  그리드 함수  ====================================

Sub InitSpreadComboBox()
    Dim iCodeArr , IntRetCD1
    Dim iNameArr
    Dim iDx



End Sub


Sub SetSpreadLock()


           With frm1
    
				.vspdData.ReDraw = False
				  ggoSpread.SpreadLock C_SEQ_NO, -1, C_SEQ_NO
                  ggoSpread.SSSetRequired C_W8,	  -1, C_W8


				.vspdData.ReDraw = True

			End With
 
   
End Sub

Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
dim sumRow
    ggoSpread.Source = frm1.vspdData
    With frm1

    .vspdData.ReDraw = False
        ggoSpread.SSSetRequired C_W8,	  -1, C_W8
        ggoSpread.SSSetProtected C_w9 , pvStartRow, pvEndRow
        ggoSpread.SSSetProtected C_SEQ_NO , pvStartRow, pvEndRow  
        
    .vspdData.ReDraw = True
    
    End With
End Sub


Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos

    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

    End Select    
End Sub

'============================================  조회조건 함수  ====================================
'======================================================================================================
'   Function Name : OpenPopUp()
'   Function Desc : 
'=======================================================================================================
Function  OpenPopUp(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim strCd
	
	If IsOpenPop = True Then Exit Function
	
	Select Case iWhere
		Case 0

		Case 5
		
	
		Case Else
			Exit Function
	End Select

	IsOpenPop = True
			
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=720px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")			
		
	IsOpenPop = False
	
	If arrRet(0) = "" Then	    
		Exit Function
	Else
		Call SetPopup(arrRet, iWhere)
	End If
End Function
'======================================================================================================
'   Function Name : SetPopup(Byval arrRet)
'   Function Desc : 
'=======================================================================================================

Function SetPopup(Byval arrRet,Byval iWhere)
	With frm1
		Select Case iWhere
			Case 0
			Case 5
			
		End Select
	End With
	If iwhere  <> 0 Then
		lgBlnFlgChgValue = True
	End If
End Function

'============================================  폼 함수  ====================================
Sub Form_Load()

    Call LoadInfTB19029     
                                                    <%'Load table , B_numeric_format%>
    Call ggoOper.LockField(Document, "N")                                   <%'Lock  Suitable  Field%>                         
   
    Call InitSpreadSheet()                                                    <%'Setup the Spread sheet%>
                                            <%'Initializes local global variables%>

    Call SetToolbar("1100110100000111")										<%'버튼 툴바 제어 %>

	' 변경한곳 
	Call  AppendNumberPlace("7", "2", "0")
    Call InitComboBox

    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
        
    Call ggoOper.FormatDate(frm1.txtFISC_YEAR, parent.gDateFormat,3)
    Call ggoOper.FormatDate(frm1.txtFOUNDATION_DT, parent.gDateFormat,1)
            
	Call SetDefaultVal
   Call InitVariables 
    Call FncQuery
End Sub
'============================================  사용자 함수  ====================================


Function MaxSpreadVal_S(Byref objSpread, ByVal intCol, byval Row)

	Dim iRows
	Dim MaxValue
	Dim tmpVal

	MAxValue = 0

	For iRows = 1 to  objSpread.MaxRows -2
		objSpread.row = iRows
	    objSpread.col = intCol

		If objSpread.Text = "" Then
		   tmpVal = 0
		Else
  		   tmpVal = cdbl(objSpread.value)
		End If

		If tmpval > MaxValue And tmpval < SUM_SEQ_NO Then
		   MaxValue = cdbl(tmpVal)
		End If
	Next

	MaxValue = MaxValue + 1

	objSpread.row	= row
	objSpread.col	= intCol
	objSpread.text	= MaxValue
	MaxSpreadVal_S = MaxValue
end Function


'============================================  이벤트 함수  ====================================

Function GetRef()	
    Dim IntRetCD , i
    Dim sMesg
    Dim sFiscYear, sRepType, sCoCd
    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6,BackColor
    
	sCoCd		= frm1.txtCO_CD.value
	sFiscYear	= frm1.txtFISC_YEAR.text
	sRepType	= frm1.cboREP_TYPE.value
	
	
	if wgConfirmFlg = "Y" then		  
	    Exit function
	end if

	call CommonQueryRs(" CO_NM, OWN_RGST_NO, CO_ADDR,REPRE_NM,FOUNDATION_DT,REPRE_RGST_NO ","dbo.ufn_TB_COMPANY_HISTORY_GetRef('" & sCoCd & "', '" & sFiscYear & "', '" & sRepType & "')", "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

	If lgF0 = "" Then	 Exit Function

	frm1.txtCO_NM2.value       = REPLACE(lgF0, chr(11),"")
	frm1.txtOwn_Rgst_No.value  = REPLACE(lgF1, chr(11),"")
	frm1.txtaddr.value		   = REPLACE(lgF2, chr(11),"")
	frm1.txtREPRE_NM.value	   = REPLACE(lgF3, chr(11),"")
	frm1.txtFOUNDATION_DT.text =REPLACE(lgF4, chr(11),"")
	frm1.txtReg_No.value=replace(lgF5,chr(11),"")
	    
End Function




Sub txtFISC_YEAR_DblClick(Button)
    If Button = 1 Then
        frm1.txtFISC_YEAR.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtFISC_YEAR.Focus
    End If
End Sub

Sub cboW13_A_OnChange()
    lgBlnFlgChgValue = True
End Sub


Sub cboW13_B_OnChange()
    lgBlnFlgChgValue = True
End Sub

Sub txtW13_C_OnChange()
    lgBlnFlgChgValue = True
End Sub

Sub cboW14_A_OnChange()
    lgBlnFlgChgValue = True
End Sub


Sub cboW14_B_OnChange()
    lgBlnFlgChgValue = True
End Sub

Sub txtW14_C_OnChange()
    lgBlnFlgChgValue = True
End Sub

Sub cboW15_A_OnChange()
    lgBlnFlgChgValue = True
End Sub


Sub cboW15_B_OnChange()
    lgBlnFlgChgValue = True
End Sub

Sub txtW15_C_OnChange()
    lgBlnFlgChgValue = True
End Sub

Sub chkW1_OnChange()
    lgBlnFlgChgValue = True
End Sub
Sub chkW2_OnChange()
    lgBlnFlgChgValue = True
End Sub
Sub chkW3_OnChange()
    lgBlnFlgChgValue = True
End Sub
Sub chkW4_OnChange()
    lgBlnFlgChgValue = True
End Sub
Sub chkW5_OnChange()
    lgBlnFlgChgValue = True
End Sub

'======

Sub txtFOUNDATION_DT_Change()
    lgBlnFlgChgValue = True
End Sub
'=======================================================================================================
Sub txtFOUNDATION_DT_DblClick(Button)
    If Button = 1 Then
        frm1.txtFOUNDATION_DT.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtFOUNDATION_DT.Focus
    End If
End Sub



Sub txtCHANGE_DT_Change()
    lgBlnFlgChgValue = True
End Sub
'=======================================================================================================
Sub  txtCHANGE_DT_DblClick(Button)
    If Button = 1 Then
        frm1.txtCHANGE_DT.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtCHANGE_DT.Focus
    End If
End Sub



'========================================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : 
'========================================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)

End Sub

Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
	Dim intIndex

	
End Sub

'==========================================================================================
Sub InitData()
	Dim intRow
	Dim intIndex 
	
	With frm1.vspdData
		
	End With
End Sub

Sub vspdData_Change(ByVal Col , ByVal Row )
     Dim iDx
    Dim IntRetCD
    Dim i , j,txtW9 ,txtW9_Row1 , txtW9_Row2, dblW10, dblW11,dblW13,dblW14
   
 
    Frm1.vspdData.Row = Row
    Frm1.vspdData.Col = Col
  '------ Developer Coding part (Start ) -------------------------------------------------------------- 

  '--------------------'그리드에 입력된 내역이 기존데이터에 있을때 체크----------------------------------
    Select Case Col
        Case C_W9
             

        
           
        
    End Select
    
 '------ Developer Coding part (End   ) -------------------------------------------------------------- 
    If Frm1.vspdData.CellType = parent.SS_CELL_TYPE_FLOAT Then
      If uniCDbl(Frm1.vspdData.text) < uniCDbl(Frm1.vspdData.TypeFloatMin) Then
         Frm1.vspdData.text = Frm1.vspdData.TypeFloatMin
      End If
    End If
	lgBlnFlgChgValue = True
    ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row

End Sub




Sub vspdData_Click(ByVal Col, ByVal Row)
    Call SetPopupMenuItemInf("1101011111") 

    gMouseClickStatus = "SPC"   

    Set gActiveSpdSheet = frm1.vspdData
   
    If frm1.vspdData.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
   	    
   ' If Row <= 0 Then
   '    ggoSpread.Source = frm1.vspdData
       
   '    If lgSortKey = 1 Then
   '        ggoSpread.SSSort Col               'Sort in ascending
   '        lgSortKey = 2
   '    Else
   '        ggoSpread.SSSort Col, lgSortKey    'Sort in descending 
   '        lgSortKey = 1
   '    End If
       
   '    Exit Sub
   ' End If

	frm1.vspdData.Row = Row
End Sub

Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
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

Sub vspdData_GotFocus()
    ggoSpread.Source = Frm1.vspdData

End Sub

Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If

End Sub    

Sub vspdData_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
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
    
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	           
    	If lgStrPrevKeyIndex <> "" Then                         
      	   Call DisableToolBar(Parent.TBC_QUERY)
			If DbQuery = False Then
				Call RestoreTooBar()
			    Exit Sub
			End If  				
    	End If
    End if
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


Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If CheckRunningBizProcess = True Then					'☜: 조회중이면 다음 조회 안하도록 체크 
        Exit Sub
	End If
	
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	    <%'☜: 재쿼리 체크 %>
      
    	If lgStrPrevKey <> "" And lgStrPrevKey2 <> "" Then                  <%'다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 %>
      		Call DisableToolBar(parent.TBC_QUERY)					'☜ : Query 버튼을 disable 시킴.
			If DBQuery = False Then 
			   Call RestoreToolBar()
			   Exit Sub 
			End If 
    	End If

    End If
End Sub







'============================================  툴바지원 함수  ====================================
'=====================================================
Function FncNew() 
    Dim IntRetCD 
    
    FncNew = False                                                          
    
	ggoSpread.Source = frm1.vspdData 
	If ggoSpread.SSCheckChange = True or lgBlnFlgChgValue = true Then
		IntRetCD = DisplayMsgBox("900015", parent.VB_YES_NO, "X", "X") 
		If IntRetCD = vbNo Then	Exit Function
    End If

     Call ggoOper.ClearField(Document, "2")
    Call SetDefaultVal
    Call InitVariables               

    Call SetToolbar("1100110100000111")          '⊙: 버튼 툴바 제어 
    FncNew = True                

End Function

'=====================================================
Function FncQuery() 
    Dim IntRetCD 
    
    FncQuery = False                                                        
    
    Err.Clear                                                               <%'Protect system from crashing%>

<%  '-----------------------
    'Check previous data area
    '----------------------- %>
    ggoSpread.Source = frm1.vspdData
    If lgBlnFlgChgValue = true Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "X", "X")			    <%'데이타가 변경되었습니다. 조회하시겠습니까?%>
    	If IntRetCD = vbNo Then
      	Exit Function
    	End If
    End If
    
<%  '-----------------------
    'Erase contents area
    '----------------------- %>
    Call ggoOper.ClearField(Document, "2")									<%'Clear Contents  Field%>
    ggoSpread.ClearSpreadData
    Call InitVariables                                                      <%'Initializes local global variables%>
    															
<%  '-----------------------
    'Check condition area
    '----------------------- %>
    If Not chkField(Document, "1") Then								<%'This function check indispensable field%>
       Exit Function
    End If    
	
	Call MakeKeyStream("X")
     
    CALL DBQuery()
    
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
    ggoSpread.Source = frm1.vspdData
    'If ggoSpread.SSCheckChange = True or lgBlnFlgChgValue = true Then
	'	IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "X", "X")			    <%'데이타가 변경되었습니다. 조회하시겠습니까?%>
    '	If IntRetCD = vbNo Then
     ' 	Exit Function
    '	End If
    'End If
    
    
    
    If lgIntFlgMode <>  parent.OPMD_UMODE  Then                                            'Check if there is retrived data
        Call  DisplayMsgBox("900002","X","X","X")                                  '☜: Please do Display first.
        Exit Function
    End If

    IntRetCD =  DisplayMsgBox("900003",  parent.VB_YES_NO,"X","X")		                  '☜: Do you want to delete?
	If IntRetCD = vbNo Then
		Exit Function
	End If

    Call MakeKeyStream("X")

    
    If DbDelete= False Then
       Exit Function
    End If												                  '☜: Delete db data

    FncDelete=  True                                                              '☜: Processing is OK
End Function





Function FncSave() 
   dim IntRetCD
    FncSave = False                                                         
    
    
    
    

    Err.Clear                                                               <%'☜: Protect system from crashing%>
    'On Error Resume Next                                                   <%'☜: Protect system from crashing%>    
    
<%  '-----------------------
    'Precheck area
    '----------------------- %>
    ggoSpread.Source = frm1.vspdData
    if frm1.vspdData.maxrows = 0 then 
       IntRetCD =  DisplayMsgBox("WC0002","x","x","x")                           '☜:There is no changed data. 
	    Exit Function
    end if 
    
    If lgBlnFlgChgValue = False Then
                 ggoSpread.Source = frm1.vspdData
			If  ggoSpread.SSCheckChange = False Then
			    IntRetCD =  DisplayMsgBox("900001","x","x","x")                           '☜:There is no changed data. 
			    Exit Function
			End If
       
    End If
    
    
    
    if frm1.chkW1.checked = false and frm1.chkW2.checked = false and frm1.chkW3.checked = false and frm1.chkW4.checked = false and frm1.chkW5.checked = false then
        IntRetCD =  DisplayMsgBox("X","x","신고서 종류를 선택하여 주십시오.","x")                           '☜:There is no changed data. 
	        Exit Function
    end if
    
    
    
    ggoSpread.Source = frm1.vspdData
	If Not  ggoSpread.SSDefaultCheck Then                                         '☜: Check contents area
	      Exit Function
	End If  
    
    Call MakeKeyStream("S")
<%  '-----------------------
    'Save function call area
    '----------------------- %>
    If DbSave = False Then Exit Function                                        '☜: Save db data
    
    FncSave = True                                                          
    
End Function

Function FncCopy() 
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncCopy = False                                                               '☜: Processing is NG

    If Frm1.vspdData.MaxRows < 1 Then
       Exit Function
    End If
 
    ggoSpread.Source = Frm1.vspdData

	With frm1
		If .vspdData.ActiveRow > 0 and .vspdData.ActiveRow <> .vspdData.maxrows Then
			.vspdData.focus
			.vspdData.ReDraw = False
		
			ggoSpread.CopyRow
			SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow

			
			.vspdData.ReDraw = True
		End If
	End With
	
    If Err.number = 0 Then	
       FncCopy = True                                                            '☜: Processing is OK
    End If
	
    Set gActiveElement = document.ActiveElement   
	
End Function

Function FncCancel() 
    ggoSpread.Source = frm1.vspdData	
    ggoSpread.EditUndo                 '재계산 
    
End Function

Function FncInsertRow(ByVal pvRowCnt) 
    Dim IntRetCD
    Dim imRow
    Dim iRow, iSeqNo

    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
    Dim uCountID

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
 
	With frm1.vspdData	' 포커스된 그리드 
			
		ggoSpread.Source = frm1.vspdData
			
		iRow = .ActiveRow
		.ReDraw = False
					
		If .MaxRows = 0 Then	' 첫 InsertRow는 1줄+합계줄 

			iRow = 1
			ggoSpread.InsertRow , 1
			Call SetSpreadColor( iRow, iRow+1) 
			.Row = iRow		
	       Call SetSeqNo(iRow+1, imRow)

		Else
				
			If iRow = .MaxRows Then	' -- 마지막 합계줄에서 InsertRow를 하면 상위에 추가한다.
				ggoSpread.InsertRow , imRow 
				SetSpreadColor iRow+1, iRow + imRow

				Call SetSeqNo(iRow + 1, imRow)
			Else
				ggoSpread.InsertRow ,imRow
				SetSpreadColor iRow+1, iRow + imRow

				Call SetSeqNo(iRow+1, imRow)
			End If   
		End If
		
		.col=C_W9
		.text = "~"
	End With

    If Err.number = 0 Then
       FncInsertRow = True                                                          '☜: Processing is OK
    End If   
	
    Set gActiveElement = document.ActiveElement   
    
    
End Function

' 그리드에 SEQ_NO, TYPE 넣는 로직 
Function SetSeqNo( iRow, iAddRows)
	
	Dim i, iSeqNo
	
	With frm1.vspdData	' 포커스된 그리드 

	ggoSpread.Source = frm1.vspdData
	
	If iAddRows = 1 Then ' 1줄만 넣는경우 
		.Row = iRow
		MaxSpreadVal frm1.vspdData, C_SEQ_NO, iRow
	Else
		iSeqNo = MaxSpreadVal(frm1.vspdData, C_SEQ_NO, iRow)	' 현재의 최대SeqNo를 구한다 
		
		For i = iRow to iRow + iAddRows -1
			.Row = i
			.Col = C_SEQ_NO : .Value = iSeqNo : iSeqNo = iSeqNo + 1
		Next
	End If
	End With
End Function

Function FncDeleteRow() 
   If frm1.vspdData.MaxRows < 1 Then Exit Function

    Dim lDelRows
    Dim iDelRowCnt, i
    
    With frm1  

    .vspdData.focus
    ggoSpread.Source = .vspdData 

       lDelRows = ggoSpread.DeleteRow                                              '☜: Protect system from crashing
     
	

	
 
    lgBlnFlgChgValue = True
    
    End With
    
    
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
	
    ggoSpread.Source = frm1.vspdData	
    If lgBlnFlgChgValue = true Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO, "X", "X")                <%'데이타가 변경되었습니다. 종료 하시겠습니까?%>
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    FncExit = True
End Function


'============================================  DB 억세스 함수  ====================================
Function DbQuery() 

    DbQuery = False
    
    Err.Clear                                                               <%'☜: Protect system from crashing%>
	
	Call LayerShowHide(1)
	
	Dim strVal
    
    With Frm1
    
       
        
        
        
        	strVal = BIZ_PGM_ID & "?txtMode="            & parent.UID_M0001						         
            strVal = strVal     & "&txtKeyStream="       & lgKeyStream                       '☜: Query Key
		    strVal = strVal     & "&txtMaxRows="         & .vspdData.MaxRows 
    


		    Call RunMyBizASP(MyBizASP, strVal)   
    End With                                           '☜:  Run biz logic

    DbQuery = True  
  
End Function


Function DbQueryOk()													<%'조회 성공후 실행로직 %>
	
    Dim iNameArr
    Dim iDx    
    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
    
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	lgBlnFlgChgValue = false
    '-----------------------
    'Reset variables area
    '-----------------------
    Call CheckTaxDoc(frm1.txtCO_CD.value, frm1.txtFISC_YEAR.text, frm1.cboREP_TYPE.value, BIZ_MNU_ID)
    Call InitData
	'1 컨펌체크 
	If wgConfirmFlg = "Y" Then

	    Call SetToolbar("1100000000000111")	
		
	Else
	   '2 디비환경값 , 로드시환경값 비교 
		  Call SetToolbar("1101111100000111")									<%'버튼 툴바 제어 %>
	
	End If
	
    lgIntFlgMode = parent.OPMD_UMODE
    Call SetSpreadColor(-1,-1)  


	frm1.vspdData.focus			
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
	strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0003							'☜: 비지니스 처리 ASP의 상태 
    strVal = strVal     & "&txtKeyStream="       & lgKeyStream                       '☜: Query Key	
	Call RunMyBizASP(MyBizASP, strVal)                                           '☜: Run Biz logic
	DbDelete = True                                                              '⊙: Processing is NG

End Function
'========================================================================================================
' Function Name : DbDeleteOk
' Function Desc : Called by MB Area when delete operation is successful
'========================================================================================================
Function DbDeleteOk()
	Call InitVariables()
	Call MainNew()
	Call SetToolbar("1100110100000111")	
End Function
'========================================================================================================
' Name : DbSave
' Desc : This function is data query and display
'========================================================================================================
Function DbSave() 
    Dim pP21011
    Dim lRow        
    Dim lGrpCnt     
    Dim retVal      
    Dim boolCheck   
    Dim lStartRow, lEndRow     
    Dim lRestGrpCnt 
    Dim strVal, strDel
 
    DbSave = False                                                          
    
	if LayerShowHide(1) = false then
		exit Function
	end if

    strVal = ""
    strDel = ""
    lGrpCnt = 1
    ggoSpread.Source = frm1.vspdData 
	With Frm1
	

	     .txtFlgMode.value = lgIntFlgMode	
		 strMode	   = .txtFlgMode.value
		For lRow = 1 To .vspdData.MaxRows
    
           .vspdData.Row = lRow
           .vspdData.Col = 0
        
           Select Case .vspdData.Text
        

        
               Case  ggoSpread.InsertFlag                                      '☜: Insert
                                                  strVal = strVal & "C"  &  Parent.gColSep
                                                  'strVal = strVal & lRow &  Parent.gColSep
                    .vspdData.Col = C_SEQ_NO          : strVal = strVal & Trim(.vspdData.Text) &  Parent.gColSep
                    .vspdData.Col = C_W8           : strVal = strVal & Trim(.vspdData.Text) &  Parent.gColSep
                    .vspdData.Col = C_W9_Fr          : strVal = strVal & replace(Trim(.vspdData.Text),"년","") &  Parent.gColSep
                    .vspdData.Col = C_W9_To          : strVal = strVal & replace(Trim(.vspdData.Text),"년","") &  Parent.gColSep
                    .vspdData.Col = C_W10          : strVal = strVal & replace(Trim(.vspdData.Text),"년","") &  Parent.gColSep
                    .vspdData.Col = C_W11          : strVal = strVal & replace(Trim(.vspdData.Text),"년","") &  Parent.gColSep
                    .vspdData.Col = C_W12          : strVal = strVal & Trim(.vspdData.Text) &  Parent.gRowSep
                    
    

                    lGrpCnt = lGrpCnt + 1
                    
               Case  ggoSpread.UpdateFlag                                      '☜: Update
                                                  strVal = strVal & "U"  &  Parent.gColSep

                   .vspdData.Col = C_SEQ_NO          : strVal = strVal & Trim(.vspdData.Text) &  Parent.gColSep
                    .vspdData.Col = C_W8           : strVal = strVal & Trim(.vspdData.Text) &  Parent.gColSep
                    .vspdData.Col = C_W9_Fr          : strVal = strVal & replace(Trim(.vspdData.Text),"년","") &  Parent.gColSep
                    .vspdData.Col = C_W9_To          : strVal = strVal & replace(Trim(.vspdData.Text),"년","") &  Parent.gColSep
                    .vspdData.Col = C_W10          : strVal = strVal & replace(Trim(.vspdData.Text),"년","") &  Parent.gColSep
                    .vspdData.Col = C_W11          : strVal = strVal & replace(Trim(.vspdData.Text),"년","") &  Parent.gColSep
                    .vspdData.Col = C_W12          : strVal = strVal & Trim(.vspdData.Text) &  Parent.gRowSep
                    
                    
    
                    
                    lGrpCnt = lGrpCnt + 1
               Case  ggoSpread.DeleteFlag                                      '☜: Delete
                                                  strDel = strDel & "D"  &  Parent.gColSep
                                                  'strVal = strVal & lRow &  Parent.gColSep
                    .vspdData.Col = C_SEQ_NO          : strVal = strVal & Trim(.vspdData.Text) &  Parent.gRowSep
   
                    lGrpCnt = lGrpCnt + 1
           End Select
       Next

	Frm1.txtSpread.value      = strDel & strVal

    Frm1.txtMaxRows.value  =     lGrpCnt - 1

	Frm1.txtMode.value        =  Parent.UID_M0002
	frm1.txtFlgMode.value	  =  lgIntFlgMode
	frm1.txtKeyStream.value      =  lgKeyStream
		
	End With	



	Call ExecMyBizASP(frm1, BIZ_PGM_ID) 
    DbSave = True                                                           
End Function


Function DbSaveOk()													        <%' 저장 성공후 실행 로직 %>
	Call InitVariables
	frm1.vspdData.MaxRows = 0
    ggoSpread.Source = frm1.vspdData
    ggoSpread.ClearSpreadData

    Call MainQuery()
End Function



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
		<TD WIDTH="100%">
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP_BAK">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><IMG src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23" ></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP_BAK"><font color=white><%=Request("strASPMnuMnuNm")%></font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><IMG src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23" ></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* align=right></TD>
					<TD WIDTH=10>&nbsp;</TD>
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
									<TD CLASS="TD6"><script language =javascript src='./js/w9115ma1_txtFISC_YEAR_txtFISC_YEAR.js'></script>
									<TD CLASS="TD5">법인명</TD>
									<TD CLASS="TD6"><INPUT TYPE=TEXT NAME="txtCO_CD" Size=10 tag="14">
										<INPUT TYPE=TEXT NAME="txtCO_NM" Size=20 tag="14">
									</TD>
									</TD>
								</TR>
								<TR>
									<TD CLASS="TD5">신고구분</TD>
									<TD CLASS="TD6"><SELECT NAME="cboREP_TYPE" ALT="신고구분" STYLE="WIDTH: 50%" tag="14X"></SELECT>
									</TD>
									<TD CLASS="TD5"></TD>
									<TD CLASS="TD6"></TD>
								</TR>
							</TABLE>
						</FIELDSET>							
					</TD>				
				</TR>		
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%> </TD>
				</TR>
				<TR>
					<TD valign=top HEIGHT="*" >
					    <DIV ID="TabDiv" STYLE="FLOAT: left; HEIGHT: 100%; WIDTH: 100%; overflow=auto">
					   
							<TABLE width = 100%  border = 0 cellpadding = 1 cellspacing = 1 ID="Table1">
									   <TR>
									            <TD>
													<TABLE width = 100% bgcolor = #696969  border = 0 cellpadding = 1 cellspacing = 1 ID="Table1">
														     <TR>
																	<TD CLASS="TD51" align =center width = 20%  >
																		신고서종류 
																	</TD>
																	<TD bgcolor = #eeeeec   align =CENTER  > 
																	    <TABLE width = 100% bgcolor = #696969  border = 0 cellpadding = 0 cellspacing = 0 ID="Table2"> 
																	          <TR>
																					<TD CLASS="TD61" align =right width = 5% >
																					    <INPUT TYPE=CHECKBOX NAME="chkW1" ID="chkW1" tag="25" Class="Check"  value=0> <br>
																					</TD>
																					
																					<TD CLASS="TD61" align =left width = 20% >
																					    감가상각방법신고서 
																					</TD>
																					<TD CLASS="TD61" align =right width = 5%  >
																					    <INPUT TYPE=CHECKBOX NAME="chkW2" ID="chkW2" tag="25" Class="Check"  value=0> 
																					</TD>
																					<TD CLASS="TD61" align =left width = 20% >
																					    감가상각방법변경신청서 
																					</TD>																	
																				</TR>  																				
																				 <TR>
																					<TD CLASS="TD61" align =right width = 5%  >
																					    <INPUT TYPE=CHECKBOX NAME="chkW3" ID="chkW3" tag="25" Class="Check"  value=0> 
																					</TD>
																					<TD CLASS="TD61" align =left width = 20% >
																					    내용연수신고서<br>
																					</TD>
																					<TD CLASS="TD61" align =right width = 5%  >
																					    <INPUT TYPE=CHECKBOX NAME="chkW4" ID="chkW4" tag="25" Class="Check"  value=0> 
																					</TD>
																					<TD CLASS="TD61" align =left width = 20% >
																					     내용연수승인(변경승인)신청서 
																					</TD>																	
																				</TR> 
																				<TR>
																					<TD CLASS="TD61" align =right width = 5%  >
																					    <INPUT TYPE=CHECKBOX NAME="chkW5" ID="chkW5" tag="25" Class="Check"  value=0> 
																					</TD>
																					<TD CLASS="TD61" align =left width = 20% >
																					    내용연수변경신고서 
																					</TD>
																					<TD CLASS="TD61" align =right width = 5%  >																					  
																					</TD>
																					<TD CLASS="TD61" align =left width = 20% >																					  
																					</TD>																	
																				</TR>  
																		</TABLE >       
													                </TD>													
												              </TR>										
								                  	</TABLE>								
					                         </TD>
				                         </TR>
										<TR>
											<TD WIDTH=800 valign=top HEIGHT="100" >
											   <FIELDSET CLASS="CLSFLD"><LEGEND ALIGN=LEFT>1.신고(청)인 인적사항</LEGEND>
											   
															<TABLE width = 100% bgcolor = #696969  border = 0 cellpadding = 1 cellspacing = 1 ID="Table1">
															       <TR>
																			<TD CLASS="TD51" align =left width = 20%  >
																				(1)법인명 
																			</TD>
																			<TD CLASS="TD61" align =left width = 20%  >
																				<INPUT TYPE=TEXT NAME="txtCO_NM2"   tag="14"  maxlength=100 size=25 >
																			</TD>
																			<TD CLASS="TD51" align =left width = 20%  >
																				(2)사업자등록번호 
																			</TD>
																			
																			<TD CLASS="TD61" align =left width = 20%  >
																				<INPUT TYPE=TEXT NAME="txtOwn_Rgst_No"   tag="14"  maxlength=100 size=25 >
																			</TD>	
																		</TR>
																		<TR>
																			<TD CLASS="TD51" align =left width = 20%  >
																				(3)본점소재지 
																			</TD>
																			<TD CLASS="TD61" align =left width = 20%  colspan =3 >
																				<INPUT TYPE=TEXT NAME="txtaddr"   tag="14"  maxlength=200 size=95 >
																			</TD>
																		</TR>
																		<TR>
																			<TD CLASS="TD51" align =left width = 20%  >
																				(4)대표자성명 
																			</TD>
																			<TD CLASS="TD61" align =left width = 20%  >
																				<INPUT TYPE=TEXT NAME="txtREPRE_NM"   tag="14"  maxlength=100 size=25 >
																			</TD>
																			<TD CLASS="TD51" align =left width = 20%   >
																			(5)주민등록번호																				
																			</TD>
																			<TD CLASS="TD61" align =center width = 20% >		
																			<INPUT TYPE=TEXT NAME="txtReg_No"   tag="14"  maxlength=100 size=25 >																		
																			</TD>
																		</TR>
																		<TR>
																			<TD CLASS="TD51" align =left width = 20%  >
																				(6)사업개시일 
																			</TD>
																			<TD CLASS="TD61" align =center width = 20%  >
																				<script language =javascript src='./js/w9115ma1_txtFOUNDATION_DT_txtFOUNDATION_DT.js'></script>
																			</TD>
																			<TD CLASS="TD51" align =left width = 20%   >
																				(7)변경방법적용사업연도 
																			</TD>
																			<TD CLASS="TD61" align =center width = 20%  >
																				<script language =javascript src='./js/w9115ma1_txtChange_DT_txtChange_DT.js'></script>
																			</TD>	
																		</TR>	
															</TABLE>														
												 	   </FIELDSET>	  			
											</TD>
										</TR>
										<TR>										    
												<TD WIDTH=100%  valign=top HEIGHT="160">												   
																<TABLE <%=LR_SPACE_TYPE_20%>>
																            <TR>
																				<TD COLSPAN=3>
																					 2.내용연수 신고(청) 및 변경 
																				</TD>																				
																			</TR>																	       
																			<TR>
																				<TD HEIGHT="100%" COLSPAN=3>
																					<script language =javascript src='./js/w9115ma1_vaSpread1_vspdData.js'></script>
																				</TD>																				
																			</TR>																
																</TABLE>												
												</TD>
										</TR>	
										<TR>	
												<TD WIDTH=800 valign=top  HEIGHT="*">														
														   <FIELDSET CLASS="CLSFLD"><LEGEND ALIGN=LEFT>3.감가상각방법 신고(청) 및 변경</LEGEND>
														   					<TABLE width = 100% bgcolor = #696969  border = 0 cellpadding = 1 cellspacing = 1 ID="Table1">
																		   	        <TR>
																						<TD CLASS="TD51" align =center width = 20%  >
																							자산명 
																						</TD>
																						<TD CLASS="TD51" align =center width = 20%  >
																							신고상각방법 
																						</TD>
																						<TD CLASS="TD51" align =center width = 20%  >
																							변경상각방법 
																						</TD>																						
																						<TD CLASS="TD51" align =center width = 20%  >
																							변경사유 
																						</TD>																		                 
																					</TR>
																					<TR>
																						<TD CLASS="TD51" align =left width = 20%  >
																							(13)유형고정자산(건축물외)
																						</TD>
																						<TD CLASS="TD61" align =center width = 20%   >
																							<SELECT NAME="cboW13_A" ALT="신고상각방법" STYLE="WIDTH: 50%" tag="25X1"></SELECT>
																						</TD>
																						<TD CLASS="TD61" align =center width = 20%   >
																							<SELECT NAME="cboW13_B" ALT="변경상각방법" STYLE="WIDTH: 50%" tag="25X1"></SELECT>
																						</TD>
																						<TD CLASS="TD61" align =center width = 20%   >
																							<INPUT TYPE=TEXT NAME="txtW13_C"   tag="25X1"  maxlength=100 size=25 >
																						</TD>
																					</TR>
																					<TR>
																						<TD CLASS="TD51" align =left width = 20%  >
																							(14)광업권 
																						</TD>
																							<TD CLASS="TD61" align =center width = 20%   >
																							<SELECT NAME="cboW14_A" ALT="신고상각방법" STYLE="WIDTH: 50%" tag="25X1"></SELECT>
																						</TD>
																						<TD CLASS="TD61" align =center width = 20%   >
																							<SELECT NAME="cboW14_B" ALT="변경상각방법" STYLE="WIDTH: 50%" tag="25X1"></SELECT>
																						</TD>
																						<TD CLASS="TD61" align =center width = 20%   >
																							<INPUT TYPE=TEXT NAME="txtW14_C"   tag="25X1"  maxlength=100 size=25 >
																						</TD>
																						
																					</TR>
																					<TR>
																						<TD CLASS="TD51" align =left width = 20%  >
																							(15)광업용고정자산(건축물외)
																						</TD>
																						<TD CLASS="TD61" align =center width = 20%   >
																							<SELECT NAME="cboW15_A" ALT="신고상각방법" STYLE="WIDTH: 50%" tag="25X1"></SELECT>
																						</TD>
																						<TD CLASS="TD61" align =center width = 20%   >
																							<SELECT NAME="cboW15_B" ALT="변경상각방법" STYLE="WIDTH: 50%" tag="25X1"></SELECT>
																						</TD>
																						<TD CLASS="TD61" align =center width = 20%   >
																							<INPUT TYPE=TEXT NAME="txtW15_C"   tag="25X1"  maxlength=100 size=25 >
																						</TD>
																					</TR>									
																		</TABLE>																	
															 	   </FIELDSET>				
															   
														</TD>														
											</TR>  											
							</Table>
						</DIV>
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
					<TD><BUTTON NAME="bttnPreview"  CLASS="CLSSBTN" ONCLICK="vbscript:FNCBtnPrint('VIEW')" Flag=1>미리보기</BUTTON>&nbsp;
						<BUTTON NAME="bttnPrint"	CLASS="CLSSBTN" ONCLICK="vbscript:FNCBtnPrint('PRINT')"   Flag=1>인쇄</BUTTON></TD>
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
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode"     TAG="24">
<INPUT TYPE=HIDDEN NAME="txtW3value"     TAG="24">

<INPUT TYPE=HIDDEN NAME="txtKeyStream" tag="24">
<TEXTAREA CLASS=hidden NAME=txtSpread tag="24" tabindex="-1"></TEXTAREA>
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
<FORM NAME=EBAction TARGET="MyBizASP" METHOD="POST">
	<INPUT TYPE="HIDDEN" NAME="uname"    TABINDEX="-1">
	<INPUT TYPE="HIDDEN" NAME="dbname"   TABINDEX="-1">
	<INPUT TYPE="HIDDEN" NAME="filename" TABINDEX="-1">
	<INPUT TYPE="HIDDEN" NAME="condvar"  TABINDEX="-1">
	<INPUT TYPE="HIDDEN" NAME="date"     TABINDEX="-1">	
</FORM>
</BODY>
</HTML>

