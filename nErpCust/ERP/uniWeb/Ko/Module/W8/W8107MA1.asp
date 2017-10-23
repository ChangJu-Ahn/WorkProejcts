
<%@ LANGUAGE="VBSCRIPT" %>
<%'======================================================================================================

'*  1. Module Name          : 법인세 
'*  2. Function Name        : 제1호 법인세과세표준 및 세액신고서 
'*  3. Program ID           : W8107MA1
'*  4. Program Name         : W8107MA1.asp
'*  5. Program Desc         :제1호 법인세과세표준 및 세액신고서 
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
<SCRIPT LANGUAGE="VBScript"	SRC="../WCM/inc_CliOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../WCM/incEB.vbs"></SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit

'============================================  상수/변수 선언  ====================================

<!-- #Include file="../../inc/lgvariables.inc" -->
Const BIZ_MNU_ID = "W8107MA1"											 '☆: 비지니스 로직 ASP명 
Const BIZ_PGM_ID = "W8107Mb1.asp"											 '☆: 비지니스 로직 ASP명 
Const EBR_RPT_ID = "W8113OA1"

Dim IsOpenPop          
Dim lgStrPrevKey2
Dim lgRefMode	' 레퍼런스 참조상태 : ERP 추출사용유무 

Dim dblOverRate , dblDownRate
Dim dblOverRate_View , dblDownRate_View

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

   lgKeyStream = UCASE(Frm1.txtCo_Cd.Value) &  parent.gColSep                                         '0
   lgKeyStream = lgKeyStream & (Frm1.txtFISC_YEAR.text) &  parent.gColSep                             '1 
   lgKeyStream = lgKeyStream & UCASE(Frm1.cboREP_TYPE.Value ) &  parent.gColSep						  '2   


    
   if pOpt ="S" then
			lgKeyStream = lgKeyStream &  Fn_Radio_Value(frm1.rdow1)            &  parent.gColSep      '3
			lgKeyStream = lgKeyStream &  unicdbl(frm1.txtw1_rate.value)		  &  parent.gColSep       '4
			lgKeyStream = lgKeyStream &  unicdbl(frm1.txtw1_rate.value) /100   &  parent.gColSep      '5
			 lgKeyStream = lgKeyStream &  Fn_Radio_Value(frm1.rdow2)           &  parent.gColSep      '6
	
       
			 lgKeyStream = lgKeyStream &  Fn_Radio_Value(frm1.rdow3)   &  parent.gColSep               '7
			 lgKeyStream = lgKeyStream &  Fn_Radio_Value(frm1.rdow4)   &  parent.gColSep			   '8
			 lgKeyStream = lgKeyStream &  frm1.txtw5.text			   &  parent.gColSep			   '9
			 lgKeyStream = lgKeyStream &  frm1.txtw6.text              &  parent.gColSep			   '10
             lgKeyStream = lgKeyStream &  Fn_Radio_Value(frm1.rdow8)	&  parent.gColSep   		   '11
		
			 lgKeyStream = lgKeyStream &  Fn_Radio_Value(frm1.rdow9)    &  parent.gColSep				'12
			 lgKeyStream = lgKeyStream &  Fn_Radio_Value(frm1.rdow10)   &  parent.gColSep			    '13
			 lgKeyStream = lgKeyStream &  Fn_Radio_Value(frm1.rdow11)   &  parent.gColSep				'14
			 lgKeyStream = lgKeyStream &  Trim(frm1.txtw12_1.text)			&  parent.gColSep		        '15
			 lgKeyStream = lgKeyStream &  Trim(frm1.txtw12_2.text)				&  parent.gColSep			'16
			 lgKeyStream = lgKeyStream &  Fn_Radio_Value(frm1.rdow13)   &  parent.gColSep				'17
			 lgKeyStream = lgKeyStream &  Fn_Radio_Value(frm1.rdow14)   &  parent.gColSep				'18
			 lgKeyStream = lgKeyStream &  Fn_Radio_Value(frm1.rdow15)   &  parent.gColSep				'19
			 
			 lgKeyStream = lgKeyStream &  unicdbl(frm1.txtw16.value)			&  parent.gColSep				'20

			 lgKeyStream = lgKeyStream &  unicdbl(frm1.txtw17_1.value)		&  parent.gColSep				'21
			 lgKeyStream = lgKeyStream &  unicdbl(frm1.txtw17_2.value)		&  parent.gColSep				'22
			 lgKeyStream = lgKeyStream &  unicdbl(frm1.txtw17_Sum.value)	&  parent.gColSep				'23
			 lgKeyStream = lgKeyStream &  unicdbl(frm1.txtw18_1.value)		&  parent.gColSep				'24
			 lgKeyStream = lgKeyStream &  unicdbl(frm1.txtw18_2.value)		&  parent.gColSep				'25
			 lgKeyStream = lgKeyStream &  unicdbl(frm1.txtw18_Sum.value)	&  parent.gColSep				'26
			 lgKeyStream = lgKeyStream &  unicdbl(frm1.txtw19_1.value)		&  parent.gColSep				'27
			 lgKeyStream = lgKeyStream &  unicdbl(frm1.txtw19_2.value)		&  parent.gColSep				'28
			 lgKeyStream = lgKeyStream &  unicdbl(frm1.txtw19_Sum.value)	&  parent.gColSep				'29
			 lgKeyStream = lgKeyStream &  unicdbl(frm1.txtw20_1.value)		&  parent.gColSep				'30
			 lgKeyStream = lgKeyStream &  unicdbl(frm1.txtw20_2.value)		&  parent.gColSep				'31
			 lgKeyStream = lgKeyStream &  unicdbl(frm1.txtw20_Sum.value)		&  parent.gColSep				'32
			 lgKeyStream = lgKeyStream &  unicdbl(frm1.txtw21_1.value)		&  parent.gColSep				'33
			 lgKeyStream = lgKeyStream &  unicdbl(frm1.txtw21_2.value)		&  parent.gColSep				'34
			 lgKeyStream = lgKeyStream &  unicdbl(frm1.txtw21_Sum.value)		&  parent.gColSep				'35
			 lgKeyStream = lgKeyStream &  unicdbl(frm1.txtw22_1.value)		&  parent.gColSep				'36
			 lgKeyStream = lgKeyStream &  unicdbl(frm1.txtw23_1.value)		&  parent.gColSep				'37
			 ' -- 200603 서식추가(법인유형별구분코드)
			 lgKeyStream = lgKeyStream &  Trim(frm1.txtW_TYPE.value)		&  parent.gColSep				'38
			 
			 
   end if
    
    

End Sub 

'============================================  신고구분 콤보 박스 채우기  ====================================

Sub InitComboBox()
	Dim IntRetCD1
	call CommonQueryRs("MINOR_CD,MINOR_NM"," B_MINOR "," MAJOR_CD = 'W1018' ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboREP_TYPE ,lgF0  ,lgF1  ,Chr(11))
End Sub

Sub InitSpreadSheet()
	Call AppendNumberPlace("8","15","0")	' -- 금액 15자리 고정 : 출하검사패치
    
End Sub








'============================================  그리드 함수  ====================================





'============================================  조회조건 함수  ====================================

'============================================  폼 함수  ====================================
Sub Form_Load()

    Call LoadInfTB19029     
                                             <%'Load table , B_numeric_format%>
    Call ggoOper.LockField(Document, "N")                                   <%'Lock  Suitable  Field%>                         
   
    Call InitSpreadSheet()                                                    <%'Setup the Spread sheet%>
                                                     <%'Initializes local global variables%>

    Call SetToolbar("1100100000000111")										<%'버튼 툴바 제어 %>

	' 변경한곳 
    Call InitComboBox
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
     Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
    Call ggoOper.FormatDate(frm1.txtFISC_YEAR, parent.gDateFormat,3)
    Call SetDefaultVal()
    Call InitVariables  

    ' 세무조정 체크호출 
	Call FncQuery
  
End Sub




Sub SetDefaultVal()
dim strWhere 
DIM strW1
Dim sFiscYear, sRepType, sCoCd, iGap

	
	frm1.txtFISC_YEAR.text 	= "<%=wgFISC_YEAR%>"
	frm1.txtCO_CD.value 	= "<%=wgCO_CD%>"
	frm1.txtCO_NM.value 	= "<%=wgCO_NM%>"
	frm1.cboREP_TYPE.value 	= "<%=wgREP_TYPE%>"

End Sub


Function Fn_Radio_Check(obj)
dim i ,chkObj ,IntRetCD

	Fn_Radio_Check = False
	    	for i =0 to  obj.length -1
			   if  obj(i).checked =false then
			       chkObj = false
			   else 
			       chkObj = True   
			      
			       Exit for
			 
			   end if 
			next
	          if chkObj = false then
	                   
					   IntRetCD = DisplayMsgBox("X","X",  obj(0).Alt & "(을)를 선택해 주세요", "X") 
					
					   obj(0).focus
	
					   Exit Function
	  		end if
	    
	     
	 Fn_Radio_Check = True  
End Function



Function Fn_Radio_value(obj)
dim i ,chkObj ,IntRetCD
   	for i =0 to  obj.length -1
			   if  obj(i).checked =false then
			       chkObj = false
			   else 
			       chkObj = True   
			       Fn_Radio_value  = obj(i).value
			       Exit for
			 
			   end if 
			next
	    

End Function

' ----------------------  검증 -------------------------
Function  Verification()
  dim IntRetCD
  dim i ,chkObj
	Verification = False
         
	
   
   
   ' if frm1.txtw6.text < 0 then
   '    IntRetCD = DisplayMsgBox("WC0006","X",  frm1.txtw6.Alt, "X") 
   '    Exit function 
   ' end if   
	
	'if unicdbl(frm1.txtw14.text) < unicdbl(frm1.txtw14.text) - unicdbl(frm1.txtw12.text)  then
    '   IntRetCD = DisplayMsgBox("WC0010","X",  "전사업연도 공제감면세액", "직전연도의 차감할 세액") 
    '   Exit function 
    'end if   
    
    'if unicdbl(frm1.txtw15.text) > unicdbl(frm1.txtw12.text) then
    '   IntRetCD = DisplayMsgBox("WC0010","X",  frm1.txtw12.Alt, frm1.txtw15.Alt)    '%1은 '%2보다 같거나 작아야됩니다 
    '   Exit function 
    'end if   
    
	Verification = True	
End Function
'============================================  사용자정의  ====================================


Function OpenPopup(Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	'B9003
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "법인유형별 구분"							' 팝업 명칭 
	arrParam(1) = "dbo.ufn_TB_MINOR('W1093', '" & C_REVISION_YM & "') "					' TABLE 명칭 
	arrParam(2) =  Trim(frm1.txtW_TYPE.value) 						' Code Condition
	arrParam(3) = ""									' Name Cindition
	arrParam(4) = " "									' Where Condition
	arrParam(5) = "구분 코드"

    arrField(0) = "MINOR_CD"							' Field명(0)
    arrField(1) = "MINOR_NM"							' Field명(1)

    arrHeader(0) = "코드"							' Header명(0)
    arrHeader(1) = "코드명"								' Header명(1)

	arrRet = window.showModalDialog("../../comasp/adoCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	If arrRet(0) = "" Then
		frm1.txtW18.focus
	    Exit Function
	Else
		frm1.txtW_TYPE.value = arrRet(0)
		frm1.txtW_TYPE_NM.value = arrRet(1)
		lgBlnFlgChgValue = True
	End If
End Function

Function GetRef()	' 금액가져오기 링크 클릭시 
	Dim sFiscYear, sRepType, sCoCd, IntRetCD
	Dim sMesg
	Dim w16 , w17_1, w17_2, w18_1, w18_2, w19_1, w19_2 , w20_1 ,w20_2, w21_1,  w21_2, w22_1,  w23_1
    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
	
	sCoCd		= frm1.txtCO_CD.value
	sFiscYear	= frm1.txtFISC_YEAR.text
	sRepType	= frm1.cboREP_TYPE.value
	
	
	' 온로드시 레퍼런스메시지 가져온다.
    wgRefDoc = GetDocRef(frm1.txtCO_CD.value, frm1.txtFISC_YEAR.text, frm1.cboREP_TYPE.value, BIZ_MNU_ID)
	
	sMesg = wgRefDoc & vbCrLf & vbCrLf
	Call selectColor(frm1.txtw16)
    Call selectColor(frm1.txtw17_1)
    Call selectColor(frm1.txtw17_2)
    Call selectColor(frm1.txtw18_1)
    Call selectColor(frm1.txtw18_2)
    Call selectColor(frm1.txtw19_1)
    Call selectColor(frm1.txtw19_2)
    Call selectColor(frm1.txtw20_1)
    Call selectColor(frm1.txtw20_2)
    Call selectColor(frm1.txtw21_1)
    Call selectColor(frm1.txtw21_2)
    Call selectColor(frm1.txtw22_1)
    Call selectColor(frm1.txtw23_1)
    
    IntRetCD = DisplayMsgBox("WC0003", parent.VB_YES_NO, sMesg, "X")           '⊙: "Will you destory previous data"
	Call ggoOper.LockField(Document, "N") 
	 If IntRetCD = vbNo Then
	 	Exit Function
	 End If


   '***참조 
   ' W16		 '17호 서식의 (112)합계의 4계 
   ' W17_1		 '3호 서식의 112	
   ' W17_2		  3호 서식의 138	
   ' W18_1        3호 서식의 115
   ' w18_2        3호 서식의 140
   ' w19_1        3호 서식의 125+133
   ' w19_2        3호 서식의 145
   ' w20_1        3호 서식의 132
   ' w20_2        3호 서식의 148
   ' w21_1        3호 서식의 134 
   ' w21_2        3호 서식의 149
   ' w22_1        3호 서식의 154
   ' w23_1        3호 서식의 157

   


	call CommonQueryRs("W16,W17_1,W17_2,W18_1,W18_2, W19_1","dbo.ufn_TB_1_GetRef('" & sCoCd & "', '" & sFiscYear & "', '" & sRepType & "')", "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

	     IF lgF0 = "" THEN EXIT Function 
	      W16	 = unicdbl(replace(lgF0, chr(11),""))		 
          W17_1  = unicdbl(replace(lgF1, chr(11),""))		
		  W17_2	 = unicdbl(replace(lgF2, chr(11),""))			
		  W18_1  = unicdbl(replace(lgF3, chr(11),""))		
		  w18_2  = unicdbl(replace(lgF4, chr(11),""))		
		  w19_1  = unicdbl(replace(lgF5, chr(11),""))		
		  w19_2  = unicdbl(replace(lgF6, chr(11),""))	
		  
		
		   frm1.txtW16.value	 = w16 
           frm1.txtW17_1.value   = w17_1	
		   frm1.txtW17_2.value	 = w17_2
		   frm1.txtW18_1.value  = w18_1
		   frm1.txtw18_2.value   = w18_2
		   frm1.txtw19_1.value   = w19_1
		   frm1.txtw19_2.value   = w19_2
		  
		  	
    
	call CommonQueryRs("W19_2, W20_1,W20_2, W21_1,W21_2, W22_1,W23_1","dbo.ufn_TB_1_GetRef('" & sCoCd & "', '" & sFiscYear & "', '" & sRepType & "')", "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	    IF lgF0 = "" THEN EXIT Function 
        w20_1  = unicdbl(replace(lgF1, chr(11),""))		
		w20_2  = unicdbl(replace(lgF2, chr(11),""))		
		w21_1  = unicdbl(replace(lgF3, chr(11),""))		
		w21_2  = unicdbl(replace(lgF4, chr(11),""))		
		w22_1  = unicdbl(replace(lgF5, chr(11),""))		
		w23_1 = unicdbl(replace(lgF6, chr(11),""))	
		
		
		 frm1.txtw20_1.value  =	w20_1		
		 frm1.txtw20_2.value  = w20_2	
		 frm1.txtw21_1.value  = w21_1
		 frm1.txtw21_2.value  = w21_2
		 frm1.txtw22_1.value  = w22_1
		 frm1.txtw23_1.value =  w23_1
		
	

  
    
      lgBlnFlgChgValue = TRUE

   
End Function


Function Fn_CalSum()
 
End Function

'============================================  이벤트 함수  ====================================
Sub txtFISC_YEAR_DblClick(Button)
    If Button = 1 Then
        frm1.txtFISC_YEAR.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtFISC_YEAR.Focus
    End If
End Sub


Sub txtW_TYPE_onChange()
	Dim arrVal
	
	If Len(frm1.txtW_TYPE.Value) > 0 Then
		If CommonQueryRs("MINOR_NM", "dbo.ufn_TB_MINOR('W1093', '" & C_REVISION_YM & "') " , " MINOR_CD = '" & Trim(frm1.txtW_TYPE.Value) &"' ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) Then
	    	arrVal				= Split(lgF0, Chr(11))
			frm1.txtW_TYPE_NM.Value	= arrVal(0)
		Else
			Call DisplayMsgBox("970000", "x",frm1.txtW_TYPE.alt & " '" & UCase(Me.Value) & "' " ,"x")
			'frm1.txtW_TYPE.Value	= ""
			frm1.txtW_TYPE_NM.Value	= ""
			frm1.txtW_TYPE.focus
		End If
	Else
		frm1.txtW_TYPE_NM.Value = ""
	End If
	lgBlnFlgChgValue = True
End Sub


Function Fn_CalSum()
  
    
   
      
     

end function


Function CheckMessage(ByVal Obj)
dim IntRetCD
    if  UNICDbl(Obj.value) < 0 then
        Obj.value = 0
        Obj.focus

    end if
    
end function




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

     Call SetToolbar("1100100000000111")          '⊙: 버튼 툴바 제어 
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
    
    If Not chkField(Document, "2") Then                                          '☜: Check contents area
       Exit Function
    End If
    
    
    if Fn_Radio_Check(frm1.rdow1) = false then exit Function
    if Fn_Radio_Check(frm1.rdow2) = false then exit Function
    
    
    if Fn_Radio_Check(frm1.rdow3) = false then exit Function
    if Fn_Radio_Check(frm1.rdow4) = false then exit Function
    if Fn_Radio_Check(frm1.rdow8) = false then exit Function
 
    if Fn_Radio_Check(frm1.rdow9) = false then exit Function
    if Fn_Radio_Check(frm1.rdow10) = false then exit Function
    if Fn_Radio_Check(frm1.rdow11) = false then exit Function
    if Fn_Radio_Check(frm1.rdoW13) = false then exit Function
    if Fn_Radio_Check(frm1.rdoW14) = false then exit Function
    if Fn_Radio_Check(frm1.rdoW15) = false then exit Function 
    
<%  '-----------------------
    'Save function call area
    '----------------------- %>
    Call ggoOper.LockField(Document, "N")
    
    
    
    

    If Verification = False Then Exit Function 
    Call MakeKeyStream("S")
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
	    Call SetToolbar("1100100000000111")										<%'버튼 툴바 제어 %>
	    
		
	Else
	   '2 디비환경값 , 로드시환경값 비교 
		 Call SetToolbar("1101100000000111")										<%'버튼 툴바 제어 %>
		
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


'==========================================================================================
Sub txtW5_KeyDown(KeyCode, Shift)
	 
End Sub

'======================================================================================================
'   Event Name : txtW5_KeyDown
'   Event Desc : 달력 Popup을 호출 
'=======================================================================================================
Sub txtW5_DblClick(Button)
	If Button = 1 Then
		Call SetFocusToDocument("M")	
		frm1.txtW5.Action = 7
		frm1.txtW5.focus
	End If
End Sub


'======================================================================================================
'   Event Name : txtW6_KeyDown
'   Event Desc : 달력 Popup을 호출 
'=======================================================================================================
Sub txtW6_DblClick(Button)
	If Button = 1 Then
		Call SetFocusToDocument("M")	
		frm1.txtW6.Action = 7
		frm1.txtW6.focus
	End If
End Sub



'======================================================================================================
'   Event Name : txtW2_S_KeyDown
'   Event Desc : 달력 Popup을 호출 
'=======================================================================================================
Sub txtW2_S_DblClick(Button)
	If Button = 1 Then
		Call SetFocusToDocument("M")	
		frm1.txtW2_S.Action = 7
		frm1.txtW2_S.focus
	End If
End Sub

'======================================================================================================
'   Event Name : txtW12_1_KeyDown
'   Event Desc : 달력 Popup을 호출 
'=======================================================================================================
Sub txtW12_1_DblClick(Button)
	If Button = 1 Then
		Call SetFocusToDocument("M")	
		frm1.txtW12_1.Action = 7
		frm1.txtW12_1.focus
	End If
End Sub
'======================================================================================================
'   Event Name : txtW12_1_KeyDown
'   Event Desc : 달력 Popup을 호출 
'=======================================================================================================
Sub txtW12_2_DblClick(Button)
	If Button = 1 Then
		Call SetFocusToDocument("M")	
		frm1.txtW12_2.Action = 7
		frm1.txtW12_2.focus
	End If
End Sub
Sub txtW5_Change()
	 lgBlnFlgChgValue = True
End Sub

Sub txtW6_Change()
	 lgBlnFlgChgValue = True
End Sub

Sub txtW12_1_Change()
	 lgBlnFlgChgValue = True
End Sub

Sub txtW12_2_Change()
	 lgBlnFlgChgValue = True
End Sub


sub DataChange()
  
     if frm1.rdoW1_3.checked = true then

        Call ggoOper.SetReqAttr(frm1.txtW1_rate, "N")
    else
         frm1.txtW1_rate.value = 0
         Call ggoOper.SetReqAttr(frm1.txtW1_rate, "Q")
    end if
    
    lgBlnFlgChgValue = True
End Sub 



Function CalSum(ByVal Obj, ByVal Obj1, byval Obj2)

    Obj.value =  UNICDbl(Obj1.value) + UNICDbl(Obj2.value)
    
end function



Sub txtW17_1_Change()
      Call CalSum( frm1.txtW17_Sum ,frm1.txtW17_1, frm1.txtW17_2) 
	lgBlnFlgChgValue = True
End Sub

Sub txtW17_2_Change()
       Call CalSum( frm1.txtW17_Sum ,frm1.txtW17_1, frm1.txtW17_2) 
	 lgBlnFlgChgValue = True
End Sub

Sub txtW18_1_Change()
     Call CalSum( frm1.txtW18_Sum ,frm1.txtW18_1, frm1.txtW18_2) 
	 lgBlnFlgChgValue = True
End Sub

Sub txtW18_2_Change()
      Call CalSum( frm1.txtW18_Sum ,frm1.txtW18_1, frm1.txtW18_2) 
	 lgBlnFlgChgValue = True
End Sub
Sub txtW19_1_Change()
      Call CalSum( frm1.txtW19_Sum ,frm1.txtW19_1, frm1.txtW19_2) 
	 lgBlnFlgChgValue = True
End Sub
Sub txtW19_2_Change()
    Call CalSum( frm1.txtW19_Sum ,frm1.txtW19_1, frm1.txtW19_2) 
	 lgBlnFlgChgValue = True
End Sub
Sub txtW20_1_Change()
	Call CalSum( frm1.txtW20_Sum ,frm1.txtW20_1, frm1.txtW20_2) 
	 lgBlnFlgChgValue = True
End Sub

Sub txtW20_2_Change()
   Call CalSum( frm1.txtW20_Sum ,frm1.txtW20_1, frm1.txtW20_2 )
	 lgBlnFlgChgValue = True
End Sub

Sub txtW21_1_Change()
    Call CalSum( frm1.txtW21_Sum ,frm1.txtW21_1, frm1.txtW21_2) 
	 lgBlnFlgChgValue = True
End Sub
Sub txtW21_2_Change()
     Call CalSum( frm1.txtW21_Sum ,frm1.txtW21_1, frm1.txtW21_2) 
	 lgBlnFlgChgValue = True
End Sub
Sub txtW22_1_Change()
	 lgBlnFlgChgValue = True
End Sub
Sub txtW23_1_Change()
	 lgBlnFlgChgValue = True
End Sub
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" width="250" ><font color=white><%=Request("strASPMnuMnuNm")%></font></td>
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
									<TD CLASS="TD6"><script language =javascript src='./js/w8107ma1_txtFISC_YEAR_txtFISC_YEAR.js'></script>
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
				
				<TR  HEIGHT= *>
				        <TD WIDTH=100% valign=top  >
						<DIV ID="TabDiv" STYLE="FLOAT: left; HEIGHT: 100%; WIDTH: 100%; overflow=auto">
						<TABLE WIDTH=100% HEIGHT=100% cellpadding = 0 cellspacing = 0>
						<TR>
							<TD WIDTH=100%>
							<TABLE width = 100%  bgcolor = #696969  border = 0 cellpadding = 1 cellspacing = 1 ID="Table1">
								<TR>
									<TD CLASS="TD51" align =center colspan =2 WIDTH="20%">법인구분</TD>
									<TD CLASS="TD61" align =left colspan =3  WIDTH="30%" valign=top><INPUT TYPE="RADIO" CLASS="RADIO" ID="rdoW1_1" NAME="rdoW1" TAG="21X" VALUE="1" CHECKED onClick = "Call DataChange() " alt ="법인구분"><LABEL FOR="rdoW1_1">1.내국</LABEL>
										<INPUT TYPE="RADIO" CLASS="RADIO" ID="rdoW1_2" NAME="rdoW1" TAG="21X" VALUE="2" onClick = "Call DataChange() "  alt ="종류별구분"><LABEL FOR="rdoW1_2">2.외국</LABEL>
										<INPUT TYPE="RADIO" CLASS="RADIO" ID="rdoW1_3" NAME="rdoW1" TAG="21X" VALUE="3" onClick = "Call DataChange() "  alt ="종류별구분"><LABEL FOR="rdoW1_3">3.외투</LABEL>
										(비율<script language =javascript src='./js/w8107ma1_txtW1_rate_txtW1_rate.js'></script>%)
									</TD>
									<TD CLASS="TD51" align =center WIDTH="20%">조정구분</TD>
									<TD CLASS="TD61" align =left WIDTH="30%"><INPUT TYPE="RADIO" CLASS="RADIO" ID="rdoW3_1" NAME="rdoW3" TAG="21X" VALUE="1" onClick = "Call DataChange()"  alt ="조정구분"><LABEL FOR="rdoW3_1">1.외부</LABEL>
										<INPUT TYPE="RADIO" CLASS="RADIO" ID="rdoW3_2" NAME="rdoW3" TAG="21X" VALUE="2" onClick = "Call DataChange()"   alt ="조정구분"><LABEL FOR="rdoW3_2">2.자기</LABEL>
									</TD>
								</TR>
								<TR>
									<TD CLASS="TD51" align =center colspan =2 WIDTH="20%">종류별구분</TD>
									<TD CLASS="TD51" align =center WIDTH="8%">중소</TD>
									<TD CLASS="TD51" align =center WIDTH="8%">일반</TD>
									<TD CLASS="TD51" align =center WIDTH="14%">당기순이익과세</TD>
									<TD CLASS="TD51" align =center WIDTH="20%">외부감사대상</TD>
									<TD CLASS="TD61" align =left WIDTH="30%"><INPUT TYPE="RADIO" CLASS="RADIO" ID="rdoW4_1" NAME="rdoW4" TAG="21X" VALUE="1" onClick = "Call DataChange()" alt ="외부감사대상" ><LABEL FOR="rdoW4_1">1.여</LABEL>
										<INPUT TYPE="RADIO" CLASS="RADIO" ID="rdoW4_2" NAME="rdoW4" TAG="21X" VALUE="2" onClick = "Call DataChange()" alt ="외부감사대상"><LABEL FOR="rdoW4_2">2.부</LABEL>
									</TD>
								</TR>
								<TR>
									<TD CLASS="TD51" align =center ROWspan =3 WIDTH="7%">영리<BR>법인</TD>
									<TD CLASS="TD51" align =center WIDTH="13%">주권상장법인</TD>
									<TD CLASS="TD61" align =center WIDTH="8%"><INPUT TYPE="RADIO"  CLASS="RADIO" ID="rdoW2_1_A" NAME="rdoW2" TAG="25X" VALUE="11"  onClick = "Call DataChange() "   alt ="주권상장법인의 중소"><LABEL FOR="rdoW2_1_A">11</LABEL></TD>
									<TD CLASS="TD61" align =center WIDTH="8%"><INPUT TYPE="RADIO" CLASS="RADIO" ID="rdoW2_1_B" NAME="rdoW2" TAG="25X" VALUE="12" onClick = "Call DataChange() "  alt ="주권상장법인의 일반"><LABEL FOR="rdoW2_1_B">12</LABEL></TD>
									<TD CLASS="TD61" align =center WIDTH="14%">&nbsp;</TD>
									<TD CLASS="TD51" align =center WIDTH="20%" rowspan=4>신고구분</TD>
									<TD CLASS="TD61" align =left WIDTH="30%"><INPUT TYPE="RADIO" CLASS="RADIO" ID="rdoW8_1" NAME="rdoW8" TAG="21X" VALUE="10" onClick = "Call DataChange() " alt ="신고구분"><LABEL FOR="rdoW8_1" >1.정기</LABEL>
									</TD>
								</TR>
								<TR>
									<TD CLASS="TD51" align =center WIDTH="13%">협회등록법인</TD>
									<TD CLASS="TD61" align =center WIDTH="8%"><INPUT TYPE="RADIO" CLASS="RADIO" ID="rdoW2_2_A" NAME="rdoW2" TAG="25X" VALUE="21" onClick = "Call DataChange() "   alt ="협회등록법인"><LABEL FOR="rdoW2_2_A">21</LABEL></TD>
									<TD CLASS="TD61" align =center WIDTH="8%"><INPUT TYPE="RADIO" CLASS="RADIO" ID="rdoW2_2_B" NAME="rdoW2" TAG="25X" VALUE="22" onClick = "Call DataChange() " alt ="협회등록법인" ><LABEL FOR="rdoW2_2_B">20</LABEL></TD>
									<TD CLASS="TD61" align =center WIDTH="14%">&nbsp;</TD>
									<TD CLASS="TD61" align =left WIDTH="30%" >&nbsp;&nbsp;&nbsp;2.수정신고(<INPUT TYPE="RADIO" CLASS="RADIO" ID="rdoW8_2_A" NAME="rdoW8" TAG="25X" VALUE="21" onClick = "Call DataChange() "><LABEL FOR="rdoW8_2_A" alt ="수정신고">가.서면분석</LABEL>,
										<INPUT TYPE="RADIO" CLASS="RADIO" ID="rdoW8_2_B" NAME="rdoW8" TAG="25X" VALUE="22" onClick = "Call DataChange()"  ><LABEL FOR="rdoW8_2_B"  alt ="수정신고">나.기타</LABEL>)
									</TD>
								</TR>
								<TR>
									<TD CLASS="TD51" align =center WIDTH="13%">기 타 법인</TD>
									<TD CLASS="TD61" align =center WIDTH="8%"><INPUT TYPE="RADIO" CLASS="RADIO" ID="rdoW2_3" NAME="rdoW2" TAG="21X" VALUE="30"  onClick = "Call DataChange() "   alt ="종류별구분" ><LABEL FOR="rdoW2_3">30</LABEL></TD>
									<TD CLASS="TD61" align =center WIDTH="8%"><INPUT TYPE="RADIO" CLASS="RADIO" ID="rdoW2_4" NAME="rdoW2" TAG="21X" VALUE="40"  onClick = "Call DataChange() "   alt ="종류별구분"><LABEL FOR="rdoW2_4">40</LABEL></TD>
									<TD CLASS="TD61" align =center WIDTH="14%">&nbsp;</TD>
									<TD CLASS="TD61" align =left WIDTH="30%"><INPUT TYPE="RADIO" CLASS="RADIO" ID="rdoW8_3" NAME="rdoW8" TAG="21X" VALUE="30" onClick = "Call DataChange() "  alt ="신고구분"><LABEL FOR="rdoW8_3" >3.기한후 신고</LABEL></TD>
								</TR>
								<TR>
									<TD CLASS="TD51" align =center colspan =2 WIDTH="20%">비 영 리 법 인</TD>
									<TD CLASS="TD61" align =center WIDTH="8%"><INPUT TYPE="RADIO" CLASS="RADIO" ID="rdoW2_6" NAME="rdoW2" TAG="21X" VALUE="30"  onClick = "Call DataChange() "   alt ="종류별구분" ><LABEL FOR="rdoW2_6">60</LABEL></TD>
									<TD CLASS="TD61" align =center WIDTH="8%"><INPUT TYPE="RADIO" CLASS="RADIO" ID="rdoW2_7" NAME="rdoW2" TAG="21X" VALUE="40"  onClick = "Call DataChange() "   alt ="종류별구분"><LABEL FOR="rdoW2_7">70</LABEL></TD>
									<TD CLASS="TD61" align =center WIDTH="14%"><INPUT TYPE="RADIO" CLASS="RADIO" ID="rdoW2_5" NAME="rdoW2" TAG="21X" VALUE="50"  onClick = "Call DataChange() "  alt ="종류별구분"><LABEL FOR="rdoW2_5">50</LABEL></TD>
									<TD CLASS="TD61" align =left WIDTH="30%"><INPUT TYPE="RADIO" CLASS="RADIO" ID="rdoW8_4" NAME="rdoW8" TAG="21X" VALUE="40" onClick = "Call DataChange() "  alt ="신고구분"><LABEL FOR="rdoW8_4" >4.중도폐업신고</LABEL> </TD>
								</TR>
								<TR>
									<TD CLASS="TD51" align =center colspan =2 WIDTH="20%">법인유형별구분</TD>
									<TD CLASS="TD61" align =left><INPUT NAME="txtW_TYPE" ALT="법인유형별구분" MAXLENGTH="3" SIZE="7" STYLE="TEXT-ALIGN:left" tag="22XXXU" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCalType" align=top TYPE="BUTTON" onclick="vbscript:Call OpenPopup(0)"></TD>
									<TD CLASS="TD61" align =center colspan=2><INPUT NAME="txtW_TYPE_NM" ALT="법인유형별구분" style="width:100%" tag = "24" ></TD>
									<TD CLASS="TD51" align =center WIDTH="20%">결산확정일</TD>
									<TD CLASS="TD61" align =left WIDTH="30%"><script language =javascript src='./js/w8107ma1_OBJECT1_txtW5.js'></script></TD>
								</TR>
								<TR>
									<TD CLASS="TD51" align =center colspan =2 WIDTH="20%">신 고 일</TD>
									<TD CLASS="TD61" align =left colspan=3><script language =javascript src='./js/w8107ma1_txtW6_txtW6.js'></script></TD>
									<TD CLASS="TD51" align =center WIDTH="20%">납 부 일</TD>
									<TD CLASS="TD61" align =left WIDTH="30%"></TD>
								</TR>
								<TR>
									<TD CLASS="TD51" align =center colspan =2 WIDTH="20%">신고기한 연장승인</TD>
									<TD CLASS="TD51" align =left colspan=2>1.신청일</TD>
									<TD CLASS="TD61" align =left><script language =javascript src='./js/w8107ma1_txtW12_1_txtW12_1.js'></script></TD>
									<TD CLASS="TD51" align =center WIDTH="20%">2.연장기한</TD>
									<TD CLASS="TD61" align =left WIDTH="30%"><script language =javascript src='./js/w8107ma1_txtW12_2_txtW12_2.js'></script></TD>
								</TR>
							</TABLE><br>
							</TD>
						<TR>
							<TD WIDTH=100%>
							<TABLE width = 100%  bgcolor = #696969  border = 0 cellpadding = 1 cellspacing = 1 ID="Table1">
								<TR>
									<TD CLASS="TD51" align =center width=30%>구 분</TD>
									<TD CLASS="TD51" align =center width=10%>여</TD>
									<TD CLASS="TD51" align =center width=10%>부</TD>
									<TD CLASS="TD51" align =center  width=30%>구 분</TD>
									<TD CLASS="TD51" align =center width=10%>여</TD>
									<TD CLASS="TD51" align =center width=10%>부</TD>
								</TR>
								<TR>
									<TD CLASS="TD51" align =left>주식변동</TD>
									<TD CLASS="TD61" align =center><INPUT TYPE="RADIO" CLASS="RADIO" ID="rdoW9_1" NAME="rdoW9" TAG="21X" VALUE="1" onClick = "Call DataChange() "  alt ="주식변동여부"><LABEL FOR="rdoW9_1" >1</LABEL> </TD>
									<TD CLASS="TD61" align =center><INPUT TYPE="RADIO" CLASS="RADIO" ID="rdoW9_2" NAME="rdoW9" TAG="21X" VALUE="2" onClick = "Call DataChange() "   alt ="주식변동여부"><LABEL FOR="rdoW9_2" >2</LABEL> </TD>
									<TD CLASS="TD51" align =left>장부전산화</TD>
									<TD CLASS="TD61" align =center><INPUT TYPE="RADIO" CLASS="RADIO" ID="rdoW10_1" NAME="rdoW10" TAG="21X" VALUE="1" onClick = "Call DataChange() "   alt ="장부전산화여부"><LABEL FOR="rdoW10_1" >1</LABEL> </TD>
									<TD CLASS="TD61" align =center><INPUT TYPE="RADIO" CLASS="RADIO" ID="rdoW10_2" NAME="rdoW10" TAG="21X" VALUE="2" onClick = "Call DataChange() " alt ="장부전산화여부"><LABEL FOR="rdoW10_2" >2</LABEL></TD>
								</TR>
								<TR>
									<TD CLASS="TD51" align =left>사업연도의제</TD>
									<TD CLASS="TD61" align =center><INPUT TYPE="RADIO" CLASS="RADIO" ID="rdoW11_1" NAME="rdoW11" TAG="21X" VALUE="1" onClick = "Call DataChange()" alt ="사업연도의제"><LABEL FOR="rdoW11_1" >1</LABEL> </TD>
									<TD CLASS="TD61" align =center><INPUT TYPE="RADIO" CLASS="RADIO" ID="rdoW11_2" NAME="rdoW11" TAG="21X" VALUE="2"onClick = "Call DataChange()" alt ="사업연도의제"><LABEL FOR="rdoW11_2" >2</LABEL> </TD>
									<TD CLASS="TD51" align =left>결손금소급공제 법인세환급신청</TD>
									<TD CLASS="TD61" align =center><INPUT TYPE="RADIO" CLASS="RADIO" ID="rdoW13_1" NAME="rdoW13" TAG="21X" VALUE="1"  onClick = "Call DataChange()" alt ="결손금소급공제 법인세환급신청여부"><LABEL FOR="rdoW13_1" >1</LABEL> </TD>
									<TD CLASS="TD61" align =center><INPUT TYPE="RADIO" CLASS="RADIO" ID="rdoW13_2" NAME="rdoW13" TAG="21X" VALUE="2" onClick = "Call DataChange()" alt ="결손금소급공제 법인세환급신청여부"><LABEL FOR="rdoW13_2" >2</LABEL> </TD>
								</TR>
								<TR>
									<TD CLASS="TD51" align =left>감가상각방법(내용연수)신고서제출</TD>
									<TD CLASS="TD61" align =center><INPUT TYPE="RADIO" CLASS="RADIO" ID="rdoW14_1" NAME="rdoW14" TAG="21X" VALUE="1" onClick = "Call DataChange()" alt ="감가상각방법신고서제출여부" ><LABEL FOR="rdoW14_1" >1</LABEL> </TD>
									<TD CLASS="TD61" align =center><INPUT TYPE="RADIO" CLASS="RADIO" ID="rdoW14_2" NAME="rdoW14" TAG="21X" VALUE="2" onClick = "Call DataChange()" alt ="감가상각방법신고서제출여부" ><LABEL FOR="rdoW14_2">2</LABEL></TD>
									<TD CLASS="TD51" align =left>재고자산등 평가방법신고서 제출</TD>
									<TD CLASS="TD61" align =center><INPUT TYPE="RADIO" CLASS="RADIO" ID="rdoW15_1" NAME="rdoW15" TAG="21X" VALUE="1" onClick = "Call DataChange()"  alt ="재고자산등 평가방법신고서 제출여부"><LABEL FOR="rdoW15_1">1</LABEL> </TD>
									<TD CLASS="TD61" align =center><INPUT TYPE="RADIO" CLASS="RADIO" ID="rdoW15_2" NAME="rdoW15" TAG="21X" VALUE="2" onClick = "Call DataChange()"    alt ="재고자산등 평가방법신고서 제출여부"><LABEL FOR="rdoW15_2">2</LABEL></TD>
								</TR>

							</TABLE><br>
							</TD>
						<TR>
							<TD WIDTH=100%>

							<TABLE width = 100%  bgcolor = #696969  border = 0 cellpadding = 1 cellspacing = 1 ID="Table2">
							  	        <TR>
							 				 <TD CLASS="TD51" align =center  rowspan =2    width=25%>
							 					구분 
							 				</TD>
							 				<TD CLASS="TD51" align =center colspan =3 >
							 					법인세 
							 				</TD>
																
							 			</TR>
																	
							 			<TR>
							 				 <TD CLASS="TD51" align =center  width=25%>
							 					법인세 
							 				</TD>
							 				<TD CLASS="TD51" align =center  width=25% >
							 					토지 등 양도소득에 대한 법인세 
							 				</TD>
							 				<TD CLASS="TD51" align =center   width=25%>
							 				  계 
							 				</TD>
																
							 			</TR>
							 			<TR>
							 				 <TD CLASS="TD51" align =LEFT  width=25%>
							 					&nbsp;&nbsp;(16) 수  입  금  액 
							 				</TD>
							 				<TD bgcolor =#eeeee align =center  width=25%  COLSPAN = 2>
							 					(<script language =javascript src='./js/w8107ma1_txtW16_txtW16.js'></script>)
							 				</TD>
							 				<TD bgcolor =#eeeee align =center   width=25%>
																						  
							 				</TD>
																
							 			</TR>
							 			<TR>
							 				 <TD CLASS="TD51" align =LEFT  width=25%>
							 					&nbsp;&nbsp;(17) 과  세  표  준 
							 				</TD>
							 				<TD bgcolor =#eeeee align =center  width=25%>
							 					<script language =javascript src='./js/w8107ma1_txtW17_1_txtW17_1.js'></script>
							 				</TD>
							 				<TD bgcolor =#eeeee align =center   width=25%>
							 				    <script language =javascript src='./js/w8107ma1_txtW17_2_txtW17_2.js'></script>
							 				</TD>
							 				<TD bgcolor =#eeeee align =center   width=25%>
																						
							 				    <script language =javascript src='./js/w8107ma1_txtW17_Sum_txtW17_Sum.js'></script>
																			
							 				</TD>
																
							 			</TR>
							 			<TR>
							 				 <TD CLASS="TD51" align =LEFT  width=25%>
							 					&nbsp;&nbsp;(18) 산  출  세  액 
							 				</TD>
							 				<TD bgcolor =#eeeee align =center  width=25%>
							 					<script language =javascript src='./js/w8107ma1_txtW18_1_txtW18_1.js'></script>
							 				</TD>
							 				<TD bgcolor =#eeeee align =center   width=25%>
							 				    <script language =javascript src='./js/w8107ma1_txtW18_2_txtW18_2.js'></script>
							 				</TD>
							 				<TD bgcolor =#eeeee align =center   width=25%>
																						
							 				    <script language =javascript src='./js/w8107ma1_txtW18_Sum_txtW18_Sum.js'></script>
																			
							 				</TD>
																
							 			</TR>
							 			<TR>
							 				 <TD CLASS="TD51" align =LEFT  width=25%>
							 					&nbsp;&nbsp;(19) 총  부  담  세  액 
							 				</TD>
							 				<TD bgcolor =#eeeee align =center  width=25%>
							 					<script language =javascript src='./js/w8107ma1_txtW19_1_txtW19_1.js'></script>
							 				</TD>
							 				<TD bgcolor =#eeeee align =center   width=25%>
							 				    <script language =javascript src='./js/w8107ma1_txtW19_2_txtW19_2.js'></script>
							 				</TD>
							 				<TD bgcolor =#eeeee align =center   width=25%>
																						
							 				    <script language =javascript src='./js/w8107ma1_txtW19_Sum_txtW19_Sum.js'></script>
																			
							 				</TD>
																
							 			</TR>
							 			<TR>
							 				 <TD CLASS="TD51" align =LEFT  width=25%>
							 					  &nbsp;&nbsp;(20) 기  납  부  세  액 
							 				</TD>
							 				<TD bgcolor =#eeeee align =center  width=25%>
							 					<script language =javascript src='./js/w8107ma1_txtW20_1_txtW20_1.js'></script>
							 				</TD>
							 				<TD bgcolor =#eeeee align =center   width=25%>
							 				    <script language =javascript src='./js/w8107ma1_txtW20_2_txtW20_2.js'></script>
							 				</TD>
							 				<TD bgcolor =#eeeee align =center   width=25%>
																						
							 				    <script language =javascript src='./js/w8107ma1_txtW20_Sum_txtW20_Sum.js'></script>
																			
							 				</TD>
																
							 			</TR>
							 			<TR>
							 				 <TD CLASS="TD51" align =LEFT  width=25%>
							 					&nbsp;&nbsp;(21) 차  감  납  부  할  세  액 
							 				</TD>
							 				<TD bgcolor =#eeeee align =center  width=25%>
							 					<script language =javascript src='./js/w8107ma1_txtW21_1_txtW21_1.js'></script>
							 				</TD>
							 				<TD bgcolor =#eeeee align =center   width=25%>
							 				    <script language =javascript src='./js/w8107ma1_txtW21_2_txtW21_2.js'></script>
							 				</TD>
							 				<TD bgcolor =#eeeee align =center   width=25%>
																						
							 				    <script language =javascript src='./js/w8107ma1_txtW21_Sum_txtW21_Sum.js'></script>
																			
							 				</TD>
																
							 			</TR>
							 			<TR>
							 				 <TD CLASS="TD51" align =LEFT  width=25%>
							 					&nbsp;&nbsp;(22) 분  납  할  세  액 
							 				</TD>
							 				<TD bgcolor =#eeeee align =center  width=25%>
							 					<script language =javascript src='./js/w8107ma1_txtW22_1_txtW22_1.js'></script>
							 				</TD>
																						
							 				<TD bgcolor =#eeeee align =center   width=25% rowspan = 2>
							 				  ※(24) 실 납 부 세 액 
																			
							 				</TD>
							 				<TD bgcolor =#eeeee align =center   width=25% rowspan = 2>
																			
																			
							 				</TD>
																
							 			</TR>
							 			<TR>
							 				 <TD CLASS="TD51" align =LEFT  width=25%>
							 					&nbsp;&nbsp;(23) 차  감  납  부  세  액 
							 				</TD>
							 				<TD bgcolor =#eeeee align =center  width=25%>
							 					<script language =javascript src='./js/w8107ma1_txtW23_1_txtW23_1.js'></script>
							 				</TD>
																						
																			
																
							 			</TR>			   
							 </TABLE>
							</TD>
						<TR>
						</TABLE>

						</DIV>

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
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 tabindex=-1></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN TABINDEX=-1 NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN TABINDEX=-1 NAME="txtw9_VALUE" tag="24">

<INPUT TYPE=HIDDEN TABINDEX=-1 NAME="txtKeyStream" tag="24" >
<INPUT TYPE=HIDDEN TABINDEX=-1 NAME="txtFlgMode" tag="24" >



</FORM>
<DIV ID="MousePT" NAME="MousePT">

<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm" tabindex=-1></iframe>
</DIV>
<FORM NAME=EBAction TARGET="MyBizASP" METHOD="POST">
	<input type="hidden" name="uname" TABINDEX="-1">
	<input type="hidden" name="dbname" TABINDEX="-1">
	<input type="hidden" name="filename" TABINDEX="-1">
	<input type="hidden" name="strUrl" TABINDEX="-1">
	<input type="hidden" name="date" TABINDEX="-1">
</FORM>
</BODY>
</HTML>

