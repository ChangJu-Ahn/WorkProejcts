
<%@ LANGUAGE="VBSCRIPT" %>
<%'======================================================================================================

'*  1. Module Name          : 법인세 
'*  2. Function Name        : 32호 퇴직급여충당금조정명세서 
'*  3. Program ID           : W3101MA1
'*  4. Program Name         : W3101MA1.asp
'*  5. Program Desc         : 32호 퇴직급여충당금조정명세서 
'*  6. Modified date(First) : 2005/03/18
'*  7. Modified date(Last)  : 2005/03/18
'*  8. Modifier (First)     : 홍지영 
'*  9. Modifier (Last)      : 홍지영 
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
Const BIZ_MNU_ID  = "W3101MA1"	
Const BIZ_PGM_ID  = "w3101mb1.asp"											 '☆: 비지니스 로직 ASP명 
Const EBR_RPT_ID  = "w3101oa1"
Const BIZ_REF_PGM_ID = "w3101mb2.asp"


Dim C_SEQ_NO		'	순번 
Dim C_ACCT			'	계정명 
Dim C_W14_1			'	총급여액 
Dim C_W14_2			'	금액 
dim C_W15_1			'	1년미만근로한  사용인에 대한 급여액 
dim C_W15_2			'	금액 
dim C_W16_1			'	1년간 계속근로한 임원사용인에 대한 급여액 
dim C_W16_2			'	금액 
dim C_W17_1			'	기말현재전사용인퇴직시퇴직급여추계액 
dim C_W17_2			'   금액 
dim strMode
					 
DIM IsRunEvents
Const C_SHEETMAXROWS = 13


Dim IsOpenPop          
Dim lgStrPrevKey2
Dim lgRefMode	' 레퍼런스 참조상태 : ERP 추출사용유무 

IsRunEvents = FALSE

'============================================  초기화 함수  ====================================
Sub InitSpreadPosVariables()
	 C_SEQ_NO	= 1      ' 
	 C_ACCT		= 2
	 C_W14_1	= 3	
	 C_W14_2	= 4
	 C_W15_1	= 5
	 C_W15_2	= 6
	 C_W16_1	= 7
	 C_W16_2	= 8
	 C_W17_1	= 9
	 C_W17_2	= 10
    
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
    Call SetCombo2(frm1.cboREP_TYPE ,lgF0  ,lgF1  ,Chr(11))
End Sub

Sub SetDefaultVal()

    frm1.txtFISC_YEAR.text 	= "<%=wgFISC_YEAR%>"
	frm1.txtCO_CD.value 	= "<%=wgCO_CD%>"
	frm1.txtCO_NM.value 	= "<%=wgCO_NM%>"
	frm1.cboREP_TYPE.value 	= "<%=wgREP_TYPE%>"
    frm1.txtW2.value = "5/100"

End Sub

Sub InitSpreadSheet()

    Call initSpreadPosVariables()  

			      	
	With frm1.vspdData
	
			ggoSpread.Source = frm1.vspdData	
			'patch version
			 ggoSpread.Spreadinit "V20041222",,parent.gAllowDragDropSpread    
						 
				.ReDraw = false

			    .MaxCols = C_W17_2 + 1							'☜: 최대 Columns의 항상 1개 증가시킴 
				.Col = .MaxCols														'☆: 사용자 별 Hidden Column
				.ColHidden = True    
							       
			 .MaxRows = 0
			 ggoSpread.ClearSpreadData
			 .ColHeaderRows(2)
			 Call AppendNumberPlace("6","3","2")
			 .RowHeight(0) = 30 
						 
			 Call GetSpreadColumnPos("A")    
					
			 ggoSpread.SSSetEdit     C_SEQ_NO, "순번", 10,,,100,1
			 ggoSpread.SSSetEdit     C_ACCT, "계정명", 18,,,50
			 ggoSpread.SSSetFloat     C_W14_1,"(14)총급여액",   10,	    Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec,,,,"0"
			 ggoSpread.SSSetFloat     C_W14_2,   "금액", 10,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec  ,,,,"0"
			 ggoSpread.SSSetFloat     C_W15_1,"(15)퇴직급여 지급대상이 " & vbCr & "아닌 임원 또는 사용인에 대한 급여액",   8,	    Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec,,,,"0"
			 ggoSpread.SSSetFloat    C_W15_2,   "금액", 12,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec ,,,,"0"
			 ggoSpread.SSSetFloat     C_W16_1,"(16)퇴직급여 지급대상이 " & vbCr & "되는 임원 또는 사용인에 대한 급여액",   8,     Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec,,,,"0"
			 ggoSpread.SSSetFloat    C_W16_2,   "금액", 12,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec  ,,,,"0","" 
			 ggoSpread.SSSetFloat     C_W17_1,"(17)기말현재전 임원 또는 " & vbCr & "사용인 전원의 퇴직시 "& vbCr & " 퇴직급여추계액",   8,     Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec,,,,"0" 
			 ggoSpread.SSSetFloat    C_W17_2,   "금액", 12,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec   ,,,,"0"
						 

			 .AddCellSpan  3, -1000, 2, 1
			 .AddCellSpan  5, -1000, 2, 1
			 .AddCellSpan  7, -1000, 2, 1
			 .AddCellSpan  9, -1000, 2, 1

				
			 .col = 3
			 .row =-999
			 .text =  "인원"
			 .col = 4
			 .row =-999
			 .text =  "금액"
			 .col = 5
			 .row =-999
			 .text =  "인원"
			 .col = 6
			 .row =-999
			 .text =  "금액"
			 .col = 7
			 .row =-999
			 .text =  "인원"
			 .col = 8
			 .row =-999
			 .text =  "금액"
			 .col = 9
			 .row =-999
			 .text =  "인원"
			 .col = 10
			 .row =-999
			 .text =  "금액"
	

				Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
	
				Call ggoSpread.SSSetColHidden(C_SEQ_NO,C_SEQ_NO,True)
										
					
				.ReDraw = true

				 
				Call SetSpreadLock 		
	
				
					 
	 End With

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
	 		ggoSpread.SSSetRequired C_ACCT,		    -1, C_ACCT
             ggoSpread.SpreadLock C_SEQ_NO,		    -1, C_SEQ_NO
	 		ggoSpread.SpreadLock C_w15_1, -1, C_w15_1
	 		ggoSpread.SpreadLock C_w15_2 , -1, C_w15_2
	 		ggoSpread.SpreadLock C_w17_1, -1, C_w17_1
	 		ggoSpread.SpreadLock C_w17_2 , -1, C_w17_2

	 	.vspdData.ReDraw = True

	 End With
   
End Sub

Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
dim sumRow
    ggoSpread.Source = frm1.vspdData
    With frm1

    .vspdData.ReDraw = False
    
     if .vspdData.row <>  .vspdData.maxrows then
         ggoSpread.SSSetRequired C_ACCT , pvStartRow, pvEndRow
         ggoSpread.SSSetProtected C_w17_2 , pvStartRow, pvEndRow
         ggoSpread.SSSetProtected  C_SEQ_NO  ,	pvStartRow, pvEndRow	
		 ggoSpread.SSSetProtected C_w15_1 , pvStartRow, pvEndRow
		 ggoSpread.SSSetProtected C_w15_2 , pvStartRow, pvEndRow
		 ggoSpread.SSSetProtected C_w17_1 , pvStartRow, pvEndRow
		 ggoSpread.SSSetProtected C_w17_2 , pvStartRow, pvEndRow
    end if	 
   	.vspdData.col = c_seq_no	 
    .vspdData.row = .vspdData.maxrows
	ggoSpread.SSSetProtected C_SEQ_NO ,  .vspdData.maxrows, .vspdData.maxrows	
    if .vspdData.text = "999999" and .vspdData.maxrows <> 0 then
	    ggoSpread.SSSetProtected C_ACCT ,  .vspdData.maxrows, .vspdData.maxrows	
	    ggoSpread.SSSetProtected C_w14_1 , .vspdData.maxrows, .vspdData.maxrows
		ggoSpread.SSSetProtected C_w14_2 , .vspdData.maxrows, .vspdData.maxrows 
	    ggoSpread.SSSetProtected C_w15_1 , .vspdData.maxrows, .vspdData.maxrows
		ggoSpread.SSSetProtected C_w15_2 , .vspdData.maxrows, .vspdData.maxrows
		ggoSpread.SSSetProtected C_w16_1 , .vspdData.maxrows, .vspdData.maxrows
		ggoSpread.SSSetProtected C_w16_2 , .vspdData.maxrows, .vspdData.maxrows
		ggoSpread.SpreadUnLock   C_w17_1 , .vspdData.maxrows,  C_w17_2, .vspdData.maxrows

    end if
        
    .vspdData.ReDraw = True
    
    End With
End Sub

Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos

    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
        
				C_SEQ_NO	= iCurColumnPos(1)
				C_ACCT		= iCurColumnPos(2)
				C_W14_1		= iCurColumnPos(3)	
				C_W14_2		= iCurColumnPos(4)
				C_W15_1		= iCurColumnPos(5)
				C_W15_2		= iCurColumnPos(6)
				C_W16_1		= iCurColumnPos(7)
				C_W16_2		= iCurColumnPos(8)
				C_W17_1		= iCurColumnPos(9)
				C_W17_2		= iCurColumnPos(10)
			 
	
    End Select    
End Sub

' ---------------------- 서식내 검증 -------------------------
Function  Verification()
Dim dblw14, dblw14Amt , dblw16, dblw16Amt,IntRetCD
	
	Verification = False
	
	  With frm1.vspdData

          .row = .maxrows
          .col = C_W14_1	 : dblw14 = unicdbl(.value)
          .row = .maxrows
          .col = C_W14_2	 : dblw14Amt = unicdbl(.value)
          .row = .maxrows
          .col = C_W16_1	 : dblw16 = unicdbl(.value)
          .row = .maxrows
          .col = C_W16_2	 : dblw16Amt = unicdbl(.value)
          if  unicdbl(dblw14) < unicdbl(dblw16) then
              IntRetCD = DisplayMsgBox("WC0010", parent.VB_INFORMATION, "1년간 계속근로한 임원 사용인의 인원", "총인원") 
              Exit Function
          end if
          if  unicdbl(dblw14Amt) < unicdbl(dblw16Amt) then
              IntRetCD = DisplayMsgBox("WC0010", parent.VB_INFORMATION, "1년간 계속근로한 임원 사용인의 급여액", "총급여액") 
              Exit Function
          end if
 
    
    End With

	Verification = True	
End Function

'========================================================================================

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
			arrParam(0) = "표준소득율"										' 팝업 명칭 
			arrParam(1) = "tb_std_income_rate" 								' TABLE 명칭 
			arrParam(2) = Trim(strCode)										' Code Condition
			arrParam(3) = ""												' Name Cindition
			arrParam(4) = ""												' Where Condition
			arrParam(5) = "표준소득율"										' 조건필드의 라벨 명칭 
            
			arrField(0) = "STD_INCM_RT_CD"									' Field명(0)
			arrField(1) = "BUSNSECT_NM"										' Field명(1)
			arrField(2) = "FULL_DETAIL_NM"									' Field명(1)
			arrField(3) = ""												' Field명(1)
			
			arrHeader(0) = "표준소득률 번호"								' Header명(0)
			arrHeader(1) = "업태"											' Header명(1)
			arrHeader(2) = "업종"											' Header명(1)
	
	
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
				.vspdData.Col = C_IND_CLASS
				.vspdData.Text = arrRet(1)
				.vspdData.Col = C_IND_TYPE
				.vspdData.Text = arrRet(2)
				.vspdData.Col = C_RATE_NO
				.vspdData.Text = arrRet(0)
				
				Call vspdData_Change(C_RATE_NO, frm1.vspdData.activerow )	 ' 변경이 읽어났다고 알려줌 
		
		End Select
	End With
	If iwhere  <> 0 Then
		lgBlnFlgChgValue = True
	End If
End Function

Function FncCalSum(byval col,byval row)
dim  sumCol
dim  w15_1,w15_2,w16_2

        w15_1=  unicdbl(FncSumSheet(frm1.vspdData,C_W14_1, row, row, false, -1, -1, "V"))   - unicdbl(FncSumSheet(frm1.vspdData,C_W16_1, row, row, false, -1, -1, "V"))  
        if w15_1  < 0 then
           w15_1 = 0
         End if  
        frm1.vspdData.Row = row
        frm1.vspdData.Col = C_W15_1
        frm1.vspdData.text = w15_1
        
        w15_2=  unicdbl(FncSumSheet(frm1.vspdData,C_W14_2, row, row, false, -1, -1, "V"))   - unicdbl(FncSumSheet(frm1.vspdData,C_W16_2, row, row, false, -1, -1, "V"))  
         if w15_2  < 0 then
           w15_2 = 0
         End if 
        frm1.vspdData.Row = row
        frm1.vspdData.Col = C_W15_2
        frm1.vspdData.text = w15_2
   
        sumCol=  unicdbl(FncSumSheet(frm1.vspdData,col, 1, frm1.vspdData.maxrows-1, false, -1, -1, "V")) 

        frm1.vspdData.Row = frm1.vspdData.maxrows   
        frm1.vspdData.Col = col
        frm1.vspdData.text = sumCol
        
        if col = C_W16_2 then
           frm1.txtW1.value = sumCol
		   Frm1.txtW3.value = unicdbl(Frm1.txtw1.value)*(5/100)
		   
        end if  
        
        sumCol=  unicdbl(FncSumSheet(frm1.vspdData,C_W15_1, 1, frm1.vspdData.maxrows-1, false, -1, -1, "V")) 
         
        frm1.vspdData.Row = frm1.vspdData.maxrows   
        frm1.vspdData.Col = C_W15_1
        frm1.vspdData.text = sumCol
        
        sumCol=  unicdbl(FncSumSheet(frm1.vspdData,C_W15_2, 1, frm1.vspdData.maxrows-1, false, -1, -1, "V")) 
      
        frm1.vspdData.Row = frm1.vspdData.maxrows   
        frm1.vspdData.Col = C_W15_2
        frm1.vspdData.text = sumCol
    	
End Function

'============================================  폼 함수  ====================================
Sub Form_Load()

    Call LoadInfTB19029     
                                                    <%'Load table , B_numeric_format%>
    Call ggoOper.LockField(Document, "N")                                   <%'Lock  Suitable  Field%>                         
   
    Call InitSpreadSheet()                                                    <%'Setup the Spread sheet%>
    Call InitVariables                                                      <%'Initializes local global variables%>
    
    Call SetToolbar("1100110100000111")										<%'버튼 툴바 제어 %>

	' 변경한곳 
    Call InitComboBox
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
    Call ggoOper.FormatDate(frm1.txtFISC_YEAR, parent.gDateFormat,3)

	Call SetDefaultVal

    Call FncQuery
End Sub


'============================================  이벤트 함수  ====================================
'============================== 레퍼런스 함수  ========================================
Function OpenRefMenu()

    Dim arrRet
    Dim arrParam(2)
    Dim arrField, arrHeader
    Dim iCalledAspName, IntRetCD
    
    If IsOpenPop = True Then Exit Function

    IsOpenPop = True

    iCalledAspName = AskPRAspName("W5105RA1")
    
    If Trim(iCalledAspName) = "" Then
        IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "W5105RA1", "x")
        IsOpenPop = False
        Exit Function
    End If
    
    With frm1
        If .vspdData.ActiveRow > 0 then 
            arrParam(0) = GetSpreadText(.vspdData, 3, .vspdData.ActiveRow, "X", "X")
            arrParam(1) = GetSpreadText(.vspdData, 4, .vspdData.ActiveRow, "X", "X")
        End If            
    End With    

    arrRet = window.showModalDialog(iCalledAspName, Array(Window.parent,arrParam, arrField, arrHeader), _
             "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
    
    IsOpenPop = False
    
    
End Function

Function GetRef()	' 그리드1의 금액가져오기 링크 클릭시 
	Dim sFiscYear, sRepType, sCoCd, IntRetCD, arrW1, arrW2, iMaxRows, sTmp, iRow, arrADDR
	sCoCd		= frm1.txtCO_CD.value
	sFiscYear	= frm1.txtFISC_YEAR.text
	sRepType	= frm1.cboREP_TYPE.value
	
	Dim sMesg

	' 온로드시 레퍼런스메시지 가져온다.
    wgRefDoc = GetDocRef(frm1.txtCO_CD.value, frm1.txtFISC_YEAR.text, frm1.cboREP_TYPE.value, BIZ_MNU_ID)

	sMesg = wgRefDoc & vbCrLf & vbCrLf
    call SelectColor(frm1.txtW4)  
    call SelectColor(frm1.txtW6)  
    call SelectColor(frm1.txtW7) 
    call SelectColor(frm1.txtW15HO) 
    call SelectColor(frm1.txtRemark) 

	IntRetCD = DisplayMsgBox("WC0003", parent.VB_YES_NO, sMesg, "X")           '⊙: "Will you destory previous data"
    Call ggoOper.LockField(Document, "N")
	If IntRetCD = vbNo Then
		 Exit Function
	End If
   
	frm1.vspdData.MaxRows = 0
    ggoSpread.ClearSpreadData
	
	Dim strVal
    
    With Frm1
    
		strVal = BIZ_REF_PGM_ID & "?txtMode="            & parent.UID_M0001						         
        strVal = strVal			& "&txtFISC_YEAR="       & Frm1.txtFISC_YEAR.Text      '☜: Query Key        
        strVal = strVal			& "&cboREP_TYPE="        & Frm1.cboREP_TYPE.Value      '☜: Query Key   
        
		Call RunMyBizASP(MyBizASP, strVal)   
    End With  

End Function




Sub txtFISC_YEAR_DblClick(Button)
    If Button = 1 Then
        frm1.txtFISC_YEAR.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtFISC_YEAR.Focus
    End If
End Sub


Sub txtw15ho_Change( )
    IF IsRunEvents = TRUE THEN EXIT Sub
    lgBlnFlgChgValue = True

End Sub

Sub txtRemark_Change( )
Dim dblw17
     IF IsRunEvents = TRUE THEN EXIT Sub
    lgBlnFlgChgValue = True
     With frm1.vspdData
        IF  .maxrows =< 0 THEN
            dblw17 = 0
             Frm1.txtW9.value  =  unicdbl(dblw17) * 0.3 + unicdbl(frm1.txtRemark.value)
        ELSE  
           .row = .maxrows
           .col = C_W17_2	 : dblw17 = unicdbl(.value)
           Frm1.txtW9.value  =  unicdbl(dblw17) * 0.3 + unicdbl(frm1.txtRemark.value)
        END IF
    
    End With
   
End Sub

Sub txtw1_Change( )
	 IF IsRunEvents = TRUE THEN EXIT Sub
    lgBlnFlgChgValue = True
    Frm1.txtW3.value = unicdbl(Frm1.txtw1.value)*(5/100)
End Sub

Sub txtw3_Change( )
     IF IsRunEvents = TRUE THEN EXIT Sub
    lgBlnFlgChgValue = True

    if unicdbl(Frm1.txtW3.value)  > unicdbl(Frm1.txtW10.value) then
       Frm1.txtW11.value  = unicdbl(Frm1.txtW10.value)
    else 
       Frm1.txtW11.value  = unicdbl(Frm1.txtW3.value)  
    end if
    
End Sub

Sub txtw4_Change( )
	IF IsRunEvents = TRUE THEN EXIT Sub
    lgBlnFlgChgValue = True
    Frm1.txtW8.value = unicdbl(Frm1.txtw4.value) - unicdbl(Frm1.txtw5.value) - unicdbl(Frm1.txtw6.value) - unicdbl(Frm1.txtw7.value)

End Sub

Sub txtw5_Change( )
	IF IsRunEvents = TRUE THEN EXIT Sub
    lgBlnFlgChgValue = True
    Frm1.txtW8.value = unicdbl(Frm1.txtw4.value) - unicdbl(Frm1.txtw5.value) - unicdbl(Frm1.txtw6.value) - unicdbl(Frm1.txtw7.value)

End Sub

Sub txtw6_Change( )
	IF IsRunEvents = TRUE THEN EXIT Sub
    lgBlnFlgChgValue = True
    Frm1.txtW8.value = unicdbl(Frm1.txtw4.value) - unicdbl(Frm1.txtw5.value) - unicdbl(Frm1.txtw6.value) - unicdbl(Frm1.txtw7.value)


End Sub
Sub txtw7_Change( )
    lgBlnFlgChgValue = True
    Frm1.txtW8.value =unicdbl(Frm1.txtw4.value) - unicdbl(Frm1.txtw5.value) - unicdbl(Frm1.txtw6.value) - unicdbl(Frm1.txtw7.value)


End Sub

Sub txtw8_Change( )
	IF IsRunEvents = TRUE THEN EXIT Sub
    lgBlnFlgChgValue = True
    if unicdbl(Frm1.txtw8.value) > 0 then
        Frm1.txtW10.value =unicdbl(Frm1.txtw9.value) - unicdbl(Frm1.txtw8.value)
    Else
        Frm1.txtW10.value =unicdbl(Frm1.txtw9.value)    
    end if
    

End Sub


Sub txtw9_Change( )
	IF IsRunEvents = TRUE THEN EXIT Sub
    lgBlnFlgChgValue = True
   Call txtw8_Change


End Sub

Sub txtw10_Change( )
	IF IsRunEvents = TRUE THEN EXIT Sub
    lgBlnFlgChgValue = True
    
    if unicdbl(Frm1.txtW3.value)  > unicdbl(Frm1.txtW10.value) then
       Frm1.txtW11.value  = Frm1.txtW10.value  
    else 
       Frm1.txtW11.value  = Frm1.txtW3.value  
    end if

End Sub

Sub txtw11_Change( )
	IF IsRunEvents = TRUE THEN EXIT Sub
    lgBlnFlgChgValue = True
    if unicdbl(Frm1.txtw12.value) - unicdbl(Frm1.txtw11.value) > 0 then 
       Frm1.txtW13.value= unicdbl(Frm1.txtw12.value) - unicdbl(Frm1.txtw11.value)

    Else
       Frm1.txtW13.value =0
    end if  
   

End Sub

Sub txtw12_Change( )
	IF IsRunEvents = TRUE THEN EXIT Sub
    lgBlnFlgChgValue = True
   Call txtw11_Change


End Sub
'========================================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : 
'========================================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	Dim strTemp
	Dim intPos1

	With frm1.vspdData
        ggoSpread.Source = frm1.vspdData
        
        If Row > 0 And Col = C_RATE_POPUP Then
            .Col = Col - 1
            .Row = Row
            Call OpenPopup(.Text, 5)

        End If
    End With
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
    Dim i
    Dim w13,w5,w6,w7
 
    Frm1.vspdData.Row = Row
    Frm1.vspdData.Col = Col
  '------ Developer Coding part (Start ) -------------------------------------------------------------- 
  
  '--------------------'그리드에 입력된 내역이 기존데이터에 있을때 체크----------------------------------
    Select Case Col
        Case C_W14_1
           
            Call FncCalSum(Col,Row)
        Case C_W14_2
           
            Call FncCalSum(Col,Row)     
       
        Case C_W16_1          
            Call FncCalSum(Col,Row)   
        Case C_W16_2          
            Call FncCalSum(Col,Row)
        Case C_W17_2          
			frm1.txtw9.value =  unicdbl(FncSumSheet(frm1.vspdData,C_W17_2, row, row, false, -1, -1, "V"))   * 0.3  + unicdbl(frm1.txtRemark.value)
       
    End Select
    
 '------ Developer Coding part (End   ) -------------------------------------------------------------- 
    If Frm1.vspdData.CellType = parent.SS_CELL_TYPE_FLOAT Then
      'If CDbl(Frm1.vspdData.text) < CDbl(Frm1.vspdData.TypeFloatMin) Then
      '   Frm1.vspdData.text = Frm1.vspdData.TypeFloatMin
      'End If
    End If
	
    ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row
    ggoSpread.UpdateRow Frm1.vspdData.maxrows

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
    Call GetSpreadColumnPos("B")
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
    
    Err.Clear                                                               'Protect system from crashing%>

  '-----------------------
  'Check previous data area
  '----------------------- %>
    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = True or lgBlnFlgChgValue = true Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "X", "X")			    '데이타가 변경되었습니다. 조회하시겠습니까?%>
    	If IntRetCD = vbNo Then
      	Exit Function
    	End If
    End If
    
  '-----------------------
  'Erase contents area
  '----------------------- %>
    Call ggoOper.ClearField(Document, "2")									'Clear Contents  Field%>
    ggoSpread.ClearSpreadData
    Call InitVariables                                                      'Initializes local global variables%>
	Call SetDefaultVal
    															
  '-----------------------
  'Check condition area
  '----------------------- %>
    If Not chkField(Document, "1") Then								'This function check indispensable field%>
       Exit Function
    End If    
	
    CALL DBQuery()
    
End Function

'========================================================================================================
' Name : FncDelete
' Desc : developer describe this line Called by MainDelete in Common.vbs
'========================================================================================================
Function FncDelete()
    Dim IntRetCd

    FncDelete = False                                                             '☜: Processing is NG
    
    
    '-----------------------
    'Check previous data area
    '----------------------- %>
    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = True or lgBlnFlgChgValue = true Then
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
    ggoSpread.Source = frm1.vspdData
	If Not  ggoSpread.SSDefaultCheck Then                                         '☜: Check contents area
	      Exit Function
	End If  
    
   if Verification = False then Exit Function
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
    if frm1.vspdData.maxrows <> frm1.vspdData.activerow then
       ggoSpread.EditUndo                                                  '☜: Protect system from crashing
     
    end if 
    
    
      Call vspdData_Change(C_W14_1,frm1.vspdData.activerow)
      Call vspdData_Change(C_W14_2,frm1.vspdData.activerow)
      Call vspdData_Change(C_W16_1,frm1.vspdData.activerow)
      Call vspdData_Change(C_W16_2,frm1.vspdData.activerow)
   if frm1.vspdData.maxrows = 1 then
      ggoSpread.EditUndo                                                  '☜: Protect system from crashing
    end if 
       
    
End Function

Function FncInsertRow(ByVal pvRowCnt) 
    Dim IntRetCD
    Dim imRow
    Dim iRow, iSeqNo

    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
    Dim uCountID,ii
    
    

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
		.vspdData.focus
		ggoSpread.Source = .vspdData
		IF  .vspdData.MaxRows = C_SHEETMAXROWS THEN
		    exit function
		End If  
		
		'.vspdData.ReDraw = False
		iSeqNo = .vspdData.MaxRows+1
	
      
        
		if 	.vspdData.MaxRows = 0 then
		
		     ggoSpread.InsertRow  imRow 
		     SetSpreadColor 1, 1
		     .vspdData.Col	= C_SEQ_NO
			.vspdData.Text	= 1
		     ggoSpread.InsertRow  imRow 
		     .row = .vspdData.MaxRows
		    .vspdData.Col	= C_SEQ_NO
			.vspdData.Text	= 999999
			.vspdData.Col	= C_ACCT
			.vspdData.Text	= "합계"
			 SetSpreadColor .vspdData.MaxRows, .vspdData.MaxRows
			
		else
				'.vspdData.ReDraw = False	' 이 행이 ActiveRow 값을 사라지게 함, 특별히 긴 로직이 아니라 ReDraw를 허용함. - 최영태 
				
		     
				iRow = .vspdData.ActiveRow
		
				If iRow = .vspdData.MaxRows Then
				    .vspdData.ActiveRow  = .vspdData.MaxRows -1
					ggoSpread.InsertRow iRow-1 , imRow 
					SetSpreadColor iRow, iRow
    
				
					For ii = .vspdData.ActiveRow To  .vspdData.ActiveRow + imRow - 1
					
						Call MaxSpreadVal(frm1.vspdData, C_SEQ_NO, ii)
						
					Next
					Call SetSpreadColor(iRow , (iRow-1) + imRow)   
				Else
				
				  
			
		            ggoSpread.InsertRow ,imRow
		            For ii = .vspdData.ActiveRow To  .vspdData.ActiveRow + imRow - 1
					
						Call MaxSpreadVal(frm1.vspdData, C_SEQ_NO, ii)
						
					Next
					Call SetSpreadColor(iRow + 1, (iRow+1) +  imRow - 1)   

					
				End If
        end if 	
    End With
    
   
    
    If Err.number = 0 Then
       FncInsertRow = True                                                          '☜: Processing is OK
    End If   
	
    Set gActiveElement = document.ActiveElement   
    
End Function

Function FncDeleteRow() 
   If frm1.vspdData.MaxRows < 1 Then Exit Function

    Dim lDelRows
    Dim iDelRowCnt, i
    
    With frm1  

    .vspdData.focus
    ggoSpread.Source = .vspdData 
    if frm1.vspdData.maxrows <> frm1.vspdData.activerow  then
       lDelRows = ggoSpread.DeleteRow                                              '☜: Protect system from crashing
      
       
    end if 
	
	  Call vspdData_Change(C_W14_1,frm1.vspdData.activerow)
      Call vspdData_Change(C_W14_2,frm1.vspdData.activerow)
      Call vspdData_Change(C_W16_1,frm1.vspdData.activerow)
      Call vspdData_Change(C_W16_2,frm1.vspdData.activerow)
 
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
    If ggoSpread.SSCheckChange = True or lgBlnFlgChgValue = true Then
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
        
        	With frm1
				If lgIntFlgMode = parent.OPMD_UMODE Then
		
					strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001							'☜: 
					strVal = strVal & "&txtCo_Cd=" & Trim(frm1.txtCo_Cd.value)				'☆: 조회 조건 데이타 
					strVal = strVal & "&txtFISC_YEAR=" & Trim(frm1.txtFISC_YEAR.text)				'☆: 조회 조건 데이타 
					strVal = strVal & "&cboREP_TYPE=" & Trim(frm1.cboREP_TYPE.value)				'☆: 조회 조건 데이타 
					strVal = strVal		& "&lgStrPrevKey=" & lgStrPrevKey
				Else
					strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001							'☜: 
  				    strVal = strVal & "&txtCo_Cd=" & Trim(frm1.txtCo_Cd.value)				'☆: 조회 조건 데이타 
					strVal = strVal & "&txtFISC_YEAR=" & Trim(frm1.txtFISC_YEAR.text)				'☆: 조회 조건 데이타 
					strVal = strVal & "&cboREP_TYPE=" & Trim(frm1.cboREP_TYPE.value)				'☆: 조회 조건 데이타 
					strVal = strVal		& "&lgStrPrevKey=" & lgStrPrevKey
				End If
				strVal = strVal		& "&lgPageNo=" & lgPageNo         
				strVal = strVal		& "&txtMaxRows=" & .vspdData.MaxRows		
				strVal = strVal		& "&lgMaxCount=" & CStr(C_SHEETMAXROWS)

		End With

		Call RunMyBizASP(MyBizASP, strVal)   
    End With                                           '☜:  Run biz logic

    DbQuery = True  

End Function

Function DbQuery2() 

    DbQuery2 = False
    
    Err.Clear                                                               
	
	Call LayerShowHide(1)
	
	Dim strVal


    With Frm1
    
		strVal = BIZ_PGM_ID & "?txtMode="            & parent.UID_M0005						         
        strVal = strVal     & "&txtCO_CD="			& Frm1.txtCO_CD.value      '☜: Query Key        
        strVal = strVal     & "&txtFISC_YEAR="		& Frm1.txtFISC_YEAR.Text      '☜: Query Key        
        strVal = strVal     & "&cboREP_TYPE="		& Frm1.cboREP_TYPE.Value      '☜: Query Key           
		
		Call RunMyBizASP(MyBizASP, strVal)   
    End With  

    DbQuery2 = True  
  
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

	strVal = BIZ_PGM_ID & "?txtMode=" &  parent.UID_M0003                                '☜: Delete
	strVal = strVal & "&txtCo_Cd=" & Trim(frm1.txtCo_Cd.value)				'☆: 조회 조건 데이타 
    strVal = strVal & "&txtFISC_YEAR=" & Trim(frm1.txtFISC_YEAR.text)				'☆: 조회 조건 데이타 
	strVal = strVal & "&cboREP_TYPE=" & Trim(frm1.cboREP_TYPE.value)				'☆: 조회 조건 데이타 
	strVal = strVal		& "&lgStrPrevKey=" & lgStrPrevKey

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
                    .vspdData.Col = C_ACCT          : strVal = strVal & Trim(.vspdData.Text) &  Parent.gColSep
                    .vspdData.Col = C_W14_1          : strVal = strVal & Trim(.vspdData.Text) &  Parent.gColSep
                    .vspdData.Col = C_W14_2          : strVal = strVal & Trim(.vspdData.Text) &  Parent.gColSep
                    .vspdData.Col = C_W15_1          : strVal = strVal & Trim(.vspdData.Text) &  Parent.gColSep
                    .vspdData.Col = C_W15_2          : strVal = strVal & Trim(.vspdData.Text) &  Parent.gColSep
                    .vspdData.Col = C_W16_1          : strVal = strVal & Trim(.vspdData.Text) &  Parent.gColSep
                    .vspdData.Col = C_W16_2          : strVal = strVal & Trim(.vspdData.Text) &  Parent.gColSep
                    .vspdData.Col = C_W17_1          : strVal = strVal & Trim(.vspdData.Text) &  Parent.gColSep
                    .vspdData.Col = C_W17_2          : strVal = strVal & Trim(.vspdData.Text) &  Parent.gRowSep   
 

                    lGrpCnt = lGrpCnt + 1
                    
               Case  ggoSpread.UpdateFlag                                      '☜: Update
                                                  strVal = strVal & "U"  &  Parent.gColSep
                                                  'strVal = strVal & lRow &  Parent.gColSep
                   .vspdData.Col = C_SEQ_NO          : strVal = strVal & Trim(.vspdData.Text) &  Parent.gColSep
                    .vspdData.Col = C_ACCT          : strVal = strVal & Trim(.vspdData.Text) &  Parent.gColSep
                    .vspdData.Col = C_W14_1          : strVal = strVal & Trim(.vspdData.Text) &  Parent.gColSep
                    .vspdData.Col = C_W14_2          : strVal = strVal & Trim(.vspdData.Text) &  Parent.gColSep
                    .vspdData.Col = C_W15_1          : strVal = strVal & Trim(.vspdData.Text) &  Parent.gColSep
                    .vspdData.Col = C_W15_2          : strVal = strVal & Trim(.vspdData.Text) &  Parent.gColSep
                    .vspdData.Col = C_W16_1          : strVal = strVal & Trim(.vspdData.Text) &  Parent.gColSep
                    .vspdData.Col = C_W16_2          : strVal = strVal & Trim(.vspdData.Text) &  Parent.gColSep
                    .vspdData.Col = C_W17_1          : strVal = strVal & Trim(.vspdData.Text) &  Parent.gColSep
                    .vspdData.Col = C_W17_2          : strVal = strVal & Trim(.vspdData.Text) &  Parent.gRowSep   
                    
                    lGrpCnt = lGrpCnt + 1
               Case  ggoSpread.DeleteFlag                                      '☜: Delete
                                                  strDel = strDel & "D"  &  Parent.gColSep
                                                  'strVal = strVal & lRow &  Parent.gColSep
                    .vspdData.Col = C_SEQ_NO          : strVal = strVal & Trim(.vspdData.Text) &  Parent.gRowSep
   
                    lGrpCnt = lGrpCnt + 1
           End Select
       Next
		 .txtMode.value        =  Parent.UID_M0002
		'.txtUpdtUserId.value  =  Parent.gUsrID
		'.txtInsrtUserId.value =  Parent.gUsrID
		.txtMaxRows.value     = lGrpCnt-1 
		.txtSpread.value      = strDel & strVal
		
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
		<TD <%=HEIGT_TYPE_00%>></TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD >
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" NOWRAP><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" width="200" ><font color=white><%=Request("strASPMnuMnuNm")%></font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* ALIGN=RIGHT>&nbsp;<a href="vbscript:GetRef">금액 불러오기</A> </TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
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
									<TD CLASS="TD5">사업연도</TD>
									<TD CLASS="TD6"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name=txtFISC_YEAR CLASS=FPDTYYYY title=FPDATETIME ALT="사업연도" tag="14X1" id=txtFISC_YEAR></OBJECT>');</SCRIPT>
									<TD CLASS="TD5">법인명</TD>
									<TD CLASS="TD6"><INPUT TYPE=TEXT NAME="txtCO_CD" Size=10 tag="14">
										<INPUT TYPE=TEXT NAME="txtCO_NM" Size=20 tag="14">
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
					<TD WIDTH=100% HEIGHT=* valign=top>
					 <DIV ID="TabDiv" STYLE="FLOAT: left; HEIGHT: 100%; WIDTH: 100%; overflow=auto">
						<TABLE id=tbl1  CLASS="TB3" CELLSPACING=0 BORDER=1>	
					      
								<TR>
									<TD id=td1 WIDTH=100% valign=top HEIGHT="100" >
									   
									      <FIELDSET CLASS="CLSFLD"><LEGEND ALIGN=LEFT>1.퇴직급여 충당금 조정 </LEGEND>
													<TABLE id=tbl2 width = 100% bgcolor = #696969  border = 0 cellpadding = 1 cellspacing = 1>
													   
														
														<TR>
															<TD CLASS="TD51" align =center  width =8%  ROWSPAN=2>
																법인세법 시행령 제60조 제1항에 따른 한도액 
															</TD>
															
														    <TD CLASS="TD51" align =center   COLSPAN=3>
																(1)1년간 계속 근로한 임원·사용인에게<BR> 지급한 총급여액{(16)의 계} 
															</TD>
															<TD CLASS="TD51" align =center   >
																(2)설정률 
															</TD>
															<TD CLASS="TD51" align =center COLSPAN=2 >
																(3)한도액<br>((1)×(2))
															</TD>
															<TD CLASS="TD51" align =center  >
																퇴직전환금 
															</TD>
															<TD CLASS="TD51" align =center >
																퇴직보험 수령액 
															</TD>
														</TR>
														
														<TR>
															
															<TD CLASS="TD61" align =center COLSPAN=3>
																<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW1" name=txtW1 CLASS=FPDS40 title=FPDOUBLESINGLE ALT="" tag="24X2Z" width = 100% ></OBJECT>');</SCRIPT>
															</TD>
															<TD CLASS="TD61" align =center >
																<INPUT TYPE=TEXT NAME="txtW2"  CLASS=FPDS40  tag="24" style="text-align: center">
															</TD>
															<TD CLASS="TD61" align =center  COLSPAN=2>
																<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW3" name=txtW3 CLASS=FPDS40 title=FPDOUBLESINGLE ALT="" tag="24X2Z" width = 100%></OBJECT>');</SCRIPT>
															</TD>
															<TD CLASS="TD61" align =left width = 12% >
																<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtRemark" name=txtRemark CLASS=FPDS40 title=FPDOUBLESINGLE ALT="" tag="21X2" width = 100%></OBJECT>');</SCRIPT>
															</TD>
															
															<TD CLASS="TD61" align =center >
																<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW15HO" name=txtW15HO CLASS=FPDS40 title=FPDOUBLESINGLE ALT="" tag="21X2" width = 100% ></OBJECT>');</SCRIPT>
															</TD>
															
														</TR>
														<TR>
															<TD CLASS="TD51" align =center  ROWSPAN=2>
																법인세법 시행령 제60조 제2항및제3항에 따른 한도액 
															</TD>
															
														    <TD CLASS="TD51" align =center  >
																(4)장부상<br>충당금기초잔액 
															</TD>
															<TD CLASS="TD51" align =center >
																(5)기중<br>충당금<br>환입액 
															</TD>
															<TD CLASS="TD51" align =center  >
																(6)충당금<br>부인<br>누계액 
															</TD>
															<TD CLASS="TD51" align =center >
																(7)기중<br>퇴직금<br>지급액 
															</TD>
															<TD CLASS="TD51" align =center   >
																(8)차감액<br>((4)-(5)-(6)-(7))
															</TD>
														    <TD CLASS="TD51" align =center >
																(9)누적한도액<br>{(17)×35(30)/100＋퇴직금전환금}
															</TD>
															<TD CLASS="TD51" align =center COLSPAN=2 >
																(10)한도액<br>((9)－(8))
															</TD>	
														</TR>
														
														<TR>
															
															<TD CLASS="TD61" align =center width = 10%>
																<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW4" name=txtW4 CLASS=FPDS40 title=FPDOUBLESINGLE ALT="" tag="21X2Z" width = 100% ></OBJECT>');</SCRIPT>
															</TD>
															<TD CLASS="TD61" align =center  width = 10%>
																<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW5" name=txtW5 CLASS=FPDS40 title=FPDOUBLESINGLE ALT="" tag="21X2Z" width = 100% ></OBJECT>');</SCRIPT>
															</TD>
															<TD CLASS="TD61" align =center  width = 10%>
																<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW6" name=txtW6 CLASS=FPDS40 title=FPDOUBLESINGLE ALT="" tag="21X2Z" width = 100%></OBJECT>');</SCRIPT>
															</TD>
															<TD CLASS="TD61" align =center width = 10%>
																<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW7" name=txtW7 CLASS=FPDS40 title=FPDOUBLESINGLE ALT="" tag="21X2Z" width = 100%></OBJECT>');</SCRIPT>
															</TD>
														    <TD CLASS="TD61" align =center width = 15%>
																<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW8" name=txtW8 CLASS=FPDS40 title=FPDOUBLESINGLE ALT="" tag="24X2" width = 100%></OBJECT>');</SCRIPT>
															</TD>
															<TD CLASS="TD61" align =center width = 15% >
																<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW9" name=txtW9 CLASS=FPDS40 title=FPDOUBLESINGLE ALT="" tag="21X2" width = 100%></OBJECT>');</SCRIPT>
															</TD>
															<TD CLASS="TD61" align =center  COLSPAN=2>
																<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW10" name=txtW10 CLASS=FPDS40 title=FPDOUBLESINGLE ALT="" tag="24X2" width = 100%></OBJECT>');</SCRIPT>
															</TD>	
															
														</TR>
														
														<TR>
															<TD CLASS="TD51" align =center   ROWSPAN=2>
																한도초과액 계산 
															</TD>
															
														    <TD CLASS="TD51" align =center  COLSPAN=2>
																(11)한  도  액<br>((3)과 (10)중 적은 금액)
															</TD>
															<TD CLASS="TD51" align =center   COLSPAN=3 >
																(12)회사계상액 
															</TD>
															<TD CLASS="TD51" align =center  COLSPAN=3>
																(13)한도초과액((12)－(11))
															</TD>
														
															
														</TR>
														
														<TR>
															
															<TD CLASS="TD61" align =center COLSPAN=2>
																<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW11" name=txtW11 CLASS=FPDS40 title=FPDOUBLESINGLE ALT="" tag="24X2" width = 100% ></OBJECT>');</SCRIPT>
															</TD>
															<TD CLASS="TD61" align =center COLSPAN=3>
																<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW12" name=txtW12 CLASS=FPDS40 title=FPDOUBLESINGLE ALT="" tag="21X2Z" width = 100% ></OBJECT>');</SCRIPT>
															</TD>
															<TD CLASS="TD61" align =center  COLSPAN=3>
																<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW13" name=txtW13 CLASS=FPDS40 title=FPDOUBLESINGLE ALT="" tag="24X2Z" width = 100%></OBJECT>');</SCRIPT>
															</TD>
																								
														</TR>									
													</TABLE>
										   </FIELDSET>				
										   			
									</TD>
								</TR>
								<TR>
										<TD WIDTH=100%  valign=top>
											<TABLE <%=LR_SPACE_TYPE_20%>>
											            <TR>
															<TD >
																 2.총급여액 및 퇴직급여추계액 명세 
															</TD>
														</TR>
														<TR>
															<TD HEIGHT="100%" valign=top>
																<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=200 tag="23" TITLE="SPREAD" id=vaSpread1> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
															</TD>
														</TR>
											</TABLE>
										</TD>
								</TR>		
							</TABLE>
					   </div>	
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
<TEXTAREA CLASS=hidden NAME=txtSpread tag="24" tabindex="-1" style="display: 'none'"></TEXTAREA>

<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode"     TAG="24">

</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
<FORM NAME=EBAction TARGET="MyBizASP" METHOD="POST">
</FORM>
</BODY>
</HTML>

