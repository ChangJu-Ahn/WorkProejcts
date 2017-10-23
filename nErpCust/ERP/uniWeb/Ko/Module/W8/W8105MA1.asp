
<%@ LANGUAGE="VBSCRIPT" %>
<%'======================================================================================================

'*  1. Module Name          : 법인세 
'*  2. Function Name        : 제68호결손법인소급공제신청서 
'*  3. Program ID           : W8105MA1
'*  4. Program Name         : W8105MA1.asp
'*  5. Program Desc         : 제68호결손법인소급공제신청서 
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
Const BIZ_MNU_ID = "W8105MA1"											 '☆: 비지니스 로직 ASP명 
Const BIZ_PGM_ID = "W8105Mb1.asp"											 '☆: 비지니스 로직 ASP명 
Const EBR_RPT_ID = "W8111OA1"


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

   lgKeyStream = UCASE(Frm1.txtCo_Cd.Value) &  parent.gColSep  
   lgKeyStream = lgKeyStream & (Frm1.txtFISC_YEAR.text) &  parent.gColSep ' 
   lgKeyStream = lgKeyStream & UCASE(Frm1.cboREP_TYPE.Value ) &  parent.gColSep '   



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
                                                     <%'Initializes local global variables%>
    Call AppendNumberPlace("6","3","2")
    Call SetToolbar("1100100000000111")										<%'버튼 툴바 제어 %>

	' 변경한곳 
    Call InitComboBox
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.FormatField(Document, "2", ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
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
    
      Call fu_CompanyYYMMDD()
     
    
    
    call CommonQueryRs("REFERENCE_1,REFERENCE_2"," ufn_TB_Configuration('W2018','" & C_REVISION_YM & "')   "," Minor_cd = '1' ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)   '1억이하 
         dblDownRate = unicdbl(lgF0)
         dblDownRate_View = replace(lgF1,Chr(11),"")
    call CommonQueryRs("REFERENCE_1,REFERENCE_2"," ufn_TB_Configuration('W2018','" & C_REVISION_YM & "')   "," Minor_cd = '2' ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)   '1억초과 
         dblOverRate = unicdbl(lgF0)
         dblOverRate_View = replace(lgF1,Chr(11),"")
	   
	


End Sub

' ----------------------  검증 -------------------------
Function  Verification()
  dim IntRetCD

	Verification = False
    if frm1.txtw6.text < 0 then
       IntRetCD = DisplayMsgBox("WC0006","X",  frm1.txtw6.Alt, "X") 
       Exit function 
    end if   
	
	if unicdbl(frm1.txtw14.text) < unicdbl(frm1.txtw14.text) - unicdbl(frm1.txtw12.text)  then
       IntRetCD = DisplayMsgBox("WC0010","X",  "전사업연도 공제감면세액", "직전연도의 차감할 세액") 
       Exit function 
    end if   
    
    if unicdbl(frm1.txtw15.text) > unicdbl(frm1.txtw12.text) then
       IntRetCD = DisplayMsgBox("WC0010","X",  frm1.txtw12.Alt, frm1.txtw15.Alt)    '%1은 '%2보다 같거나 작아야됩니다 
       Exit function 
    end if   
    
	Verification = True	
End Function
'============================================  사용자정의  ====================================


function fu_CompanyYYMMDD 

  Dim sFiscYear, sRepType, sCoCd, iGap, IntRetCD

    sCoCd		= frm1.txtCO_CD.value
	sFiscYear	= frm1.txtFISC_YEAR.text
	sRepType	= frm1.cboREP_TYPE.value


   '사업일 
		IntRetCD =  CommonQueryRs("FISC_START_DT, FISC_END_DT"," TB_COMPANY_HISTORY "," CO_CD= '" & sCoCd & "' AND FISC_YEAR='" & sFiscYear & "' AND REP_TYPE='" & sRepType & "'",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
		if IntRetCD = false then
		   IntRetCD = DisplayMsgBox("X", "X", "조회된 연도의 사업이력이 없습니다.", "X") 
		
		else
		
		   	frm1.txtw1_s.text = replace(lgF0, Chr(11),"")
			frm1.txtw1_e.text = replace(lgF1, Chr(11),"")
	
        end if
			
		'직전연도사업일 
		call CommonQueryRs("FISC_START_DT, FISC_END_DT"," TB_COMPANY_HISTORY "," CO_CD= '" & sCoCd & "' AND FISC_YEAR='" & sFiscYear -1 & "' AND REP_TYPE='1'",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
 
		frm1.txtw2_s.text = replace(lgF0, Chr(11),"")
		frm1.txtw2_e.text = replace(lgF1, Chr(11),"")
End function


Function GetRef()	' 금액가져오기 링크 클릭시 
	Dim sFiscYear, sRepType, sCoCd, IntRetCD
	Dim arrW1 ,arrW2 ,  arrW3, arrW4, arrW5, arrW6, iRow, iCol
	Dim sMesg

	
	sCoCd		= frm1.txtCO_CD.value
	sFiscYear	= frm1.txtFISC_YEAR.text
	sRepType	= frm1.cboREP_TYPE.value
	
	' 세무정보 조사 : 메시지가져오기.
	
	
	if wgConfirmFlg = "Y" then    '확정시 
	   Exit function
	end if   
	
	' 온로드시 레퍼런스메시지 가져온다.
    wgRefDoc = GetDocRef(frm1.txtCO_CD.value, frm1.txtFISC_YEAR.text, frm1.cboREP_TYPE.value, BIZ_MNU_ID)
	
	sMesg = wgRefDoc & vbCrLf & vbCrLf
	Call selectColor(frm1.txtw6)
    Call selectColor(frm1.txtw8)
    Call selectColor(frm1.txtw11)

    IntRetCD = DisplayMsgBox("WC0003", parent.VB_YES_NO, sMesg, "X")           '⊙: "Will you destory previous data"
	Call ggoOper.LockField(Document, "N") 
	 If IntRetCD = vbNo Then
	 	Exit Function
	 End If


   '***참조 
   ' W6 제 3호 서식의 (107) 각 사업년도소득금액 
   ' W8 직전 사업연도 제 3호 서식의 (112) 과세표준 
   ' W9 직전 사업연도 제 3호 서식의 (117) 세율 
   ' W11 직전 사업연도 (121) 공제감면세액(ㄱ) + (123) 공제감면세액(ㄴ)


	call CommonQueryRs("w6,W8,  W11","dbo.ufn_TB_68_GetRef('" & sCoCd & "', '" & sFiscYear & "', '" & sRepType & "')", "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

	IF lgF0 = "" THEN EXIT Function 
    if unicdbl(lgF0) * (-1) > 0 then
       frm1.txtw6.value = unicdbl(replace(lgF0,chr(11),"")) * (-1) 
    else
       frm1.txtw6.value = 0
    end if
       frm1.txtw8.value = unicdbl(replace(lgF1,chr(11),""))
     
       frm1.txtw11.value = unicdbl(replace(lgF2,chr(11),""))    
    
    
    
      lgBlnFlgChgValue = TRUE
	  Call Fn_CalSum()
   
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


Function Fn_CalSum1()


    if  unicdbl(frm1.txtw11.text) <= 100000000 * unicdbl(dblDownRate) then														 '(11) * 1억 * 전기법인세율(1억미만)
				if ( unicdbl(frm1.txtw8.text)- unicdbl(frm1.txtw11.text)/unicdbl(dblDownRate)) <   unicdbl(frm1.txtw6.text) then      'Min[(8)-(11) / 전기법인세율(1억미만),6]
				     frm1.txtw7.text =    unicdbl(frm1.txtw8.text)- unicdbl(frm1.txtw11.text)/unicdbl(dblDownRate) 
				Else
				    frm1.txtw7.text =    unicdbl(frm1.txtw6.text)
				End if

    Else
				if  (unicdbl(frm1.txtw8.text)-(100000000 + (unicdbl(frm1.txtw11.text)-100000000*unicdbl(dblDownRate)) /unicdbl(dblDownRate))) < unicdbl(frm1.txtw6.text)  then
				     frm1.txtw7.text  = (unicdbl(frm1.txtw8.text)-(100000000 + (unicdbl(frm1.txtw11.text)-100000000*unicdbl(dblDownRate)) /unicdbl(dblOverRate) ))
				else
				     frm1.txtw7.text  =  unicdbl(frm1.txtw6.text)
				end if
    
    End if 
  	

 

end function


Function Fn_CalSum()
  
    
    if unicdbl(frm1.txtw8.text) > 100000000 then
       frm1.txtw9.value = dblOverRate_View
       frm1.txtw9_value.value = (dblOverRate)
    else
	   frm1.txtw9.value = dblDownRate_View
       frm1.txtw9_value.value = (dblDownRate)
    
    end if   
         
  	
    if unicdbl(frm1.txtw8.text) <=0 then

       frm1.txtw10.text =0
       
   
    elseif  unicdbl(frm1.txtw8.text) > 0 and unicdbl(frm1.txtw8.text)  <= 100000000 then
        frm1.txtw10.text = unicdbl(frm1.txtw8.text) * unicdbl(dblDownRate)
    
    elseif  unicdbl(frm1.txtw8.text) >  100000000 then
       
       frm1.txtw10.text =( 100000000  * unicdbl(dblDownRate))  + (unicdbl(frm1.txtw8.text) - 100000000) *  unicdbl(dblOverRate)
    end if
     
      frm1.txtw12.text = unicdbl(frm1.txtw10.text)-unicdbl(frm1.txtw11.text)
      frm1.txtw13.text = unicdbl(frm1.txtw10.text)
      
     if	unicdbl(frm1.txtw8.text)-unicdbl(frm1.txtw7.text)  <=0  then
			 frm1.txtw14.text = 0
     elseif (unicdbl(frm1.txtw8.text)-unicdbl(frm1.txtw7.text)) >0 and (unicdbl(frm1.txtw8.text)-unicdbl(frm1.txtw7.text))  <=100000000 then
			 frm1.txtw14.text  = (unicdbl(frm1.txtw8.text)-unicdbl(frm1.txtw7.text)) * dblDownRate			
	 elseif (unicdbl(frm1.txtw8.text)-unicdbl(frm1.txtw7.text)) > 100000000		 then
           frm1.txtw14.text  = (100000000 * dblDownRate) + ((unicdbl(frm1.txtw8.text)-unicdbl(frm1.txtw7.text))-100000000) *  unicdbl(dblOverRate)
           
     end if   
     
     
   
         frm1.txtw15.text = unicdbl(frm1.txtw13.text) - unicdbl(frm1.txtw14.text)
    

end function


Function CheckMessage(ByVal Obj)
dim IntRetCD
    if  UNICDbl(Obj.value) < 0 then
        Obj.value = 0
        Obj.focus

    end if
    
end function



Sub txtW6_CHANGE()
    lgBlnFlgChgValue = TRUE
    Call Fn_CalSum
End Sub


Sub txtW7_CHANGE()
    lgBlnFlgChgValue = TRUE
    Call Fn_CalSum
End Sub


Sub txtW8_CHANGE()
    lgBlnFlgChgValue = TRUE
    Call Fn_CalSum
End Sub


Sub txtW11_CHANGE()
    lgBlnFlgChgValue = TRUE
    Call Fn_CalSum
End Sub

Sub txtW11_CHANGE()
    lgBlnFlgChgValue = TRUE
    Call Fn_CalSum
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
     Call fu_CompanyYYMMDD()
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
    
    
    
    if  frm1.txtW1_S.text ="" or frm1.txtW1_E.text ="" then
       Call DisplayMsgBox("X","x","당기 시작일 또는 종료일이 없습니다.","x")  		
       Exit Function
    end if
    
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

    If Verification = False Then Exit Function 
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
Sub txtW1_S_KeyDown(KeyCode, Shift)
	 
End Sub

'======================================================================================================
'   Event Name : txtW1_S_KeyDown
'   Event Desc : 달력 Popup을 호출 
'=======================================================================================================
Sub txtW1_S_DblClick(Button)
	If Button = 1 Then
		Call SetFocusToDocument("M")	
		frm1.txtW1_S.Action = 7
		frm1.txtW1_S.focus
	End If
End Sub


'======================================================================================================
'   Event Name : txtW1_E_KeyDown
'   Event Desc : 달력 Popup을 호출 
'=======================================================================================================
Sub txtW1_E_DblClick(Button)
	If Button = 1 Then
		Call SetFocusToDocument("M")	
		frm1.txtW1_E.Action = 7
		frm1.txtW1_E.focus
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
'   Event Name : txtW2_E_KeyDown
'   Event Desc : 달력 Popup을 호출 
'=======================================================================================================
Sub txtW2_E_DblClick(Button)
	If Button = 1 Then
		Call SetFocusToDocument("M")	
		frm1.txtW2_E.Action = 7
		frm1.txtW2_E.focus
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
					<TD >
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" NOWRAP><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" idth="200" ><font color=white><%=Request("strASPMnuMnuNm")%></font></td>
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
									<TD CLASS="TD6"><script language =javascript src='./js/w8105ma1_txtFISC_YEAR_txtFISC_YEAR.js'></script>
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
					   
					    
									<TABLE bgcolor = #696969  border = 0 cellpadding = 1 cellspacing = 1 ID="Table1">
									  
									
										<TR>
											 <TD CLASS="TD51" align =LEFT width =15% >
												(1)결손사업연도 
											</TD>
											
										    <TD CLASS="TD61" align =LEFT >
													<script language =javascript src='./js/w8105ma1_txtW1_S_txtW1_S.js'></script>~
													<script language =javascript src='./js/w8105ma1_txtW1_E_txtW1_E.js'></script>
											</TD>
											
										    <TD CLASS="TD51" align =LEFT  >
												(2)직전사업연도 
											</TD>
											 <TD CLASS="TD61" align =LEFT     >
													<script language =javascript src='./js/w8105ma1_txtW2_S_txtW2_S.js'></script>~
													<script language =javascript src='./js/w8105ma1_txtW2_E_txtW2_E.js'></script>
											</TD>
											
											
										</TR>
										
										<TR>
											 <TD CLASS="TD51" align =LEFT  rowspan =2>
												(2)결손사업연도<br>&nbsp;&nbsp;결손금액 
											</TD>
											<TD CLASS="TD51" align =LEFT  colspan =2>
												(6)결손금액 
											</TD>
											 
										    <TD CLASS="TD61" align =LEFT   >
												<script language =javascript src='./js/w8105ma1_txtW6_txtW6.js'></script>
											</TD>
											
										</TR>
										<TR>
											 <TD CLASS="TD51" align =LEFT  >
												(7)소급공제받을 결손금액 
											</TD>
											 <TD  CLASS="TD51"  align =LEFT   ><BUTTON NAME="btnCb_autoisrt"  ONCLICK="VBScript: Fn_CalSum1()" >자동계산</BUTTON></TD>
										     <TD CLASS="TD61" align =LEFT  >
												<script language =javascript src='./js/w8105ma1_txtW7_txtW7.js'></script>
											</TD>
											
										</TR>
										<TR>
										   <TD CLASS="TD51" align =LEFT  rowspan =5>
												(4)직전사업연도<br>&nbsp;&nbsp;법인세액계산 
											</TD>
											 <TD CLASS="TD51" align =LEFT  colspan =2>
												(8)과세표준 
											</TD>
											
										    <TD CLASS="TD61" align =LEFT  >
												<script language =javascript src='./js/w8105ma1_txtW8_txtW8.js'></script>
											</TD>
											
										</TR>
										
										<TR>
										   
											 <TD CLASS="TD51" align =LEFT  colspan =2>
												(9)세율 
											</TD>
											
										    <TD CLASS="TD61" align=right nowrap  >
												<INPUT TYPE=TEXT id="txtw9" NAME="txtw9" Size=35 tag="24" style=""></OBJECT>
											</TD>
											
											
										</TR>
										<TR>
										   
											 <TD CLASS="TD51" align =LEFT  colspan =2>
												(10)산출세액 
											</TD>
											
										    <TD CLASS="TD61" align =LEFT  >
												<script language =javascript src='./js/w8105ma1_txtW10_txtW10.js'></script>
											</TD>
											
											
										</TR>
										<TR>
										   
											 <TD CLASS="TD51" align =LEFT  colspan =2>
												(11)공제감면세액 
											</TD>
											
										    <TD CLASS="TD61" align =LEFT  >
												<script language =javascript src='./js/w8105ma1_txtW11_txtW11.js'></script>
											</TD>
											
											
										</TR>
										<TR>
										   
											 <TD CLASS="TD51" align =LEFT  colspan =2>
												(12)차감세액((10)-(11))
											</TD>
											
										    <TD CLASS="TD61" align =LEFT  >
												<script language =javascript src='./js/w8105ma1_txtW12_txtW12.js'></script>
											</TD>
											
											
										</TR>
					
										<TR>
										   <TD CLASS="TD51" align =LEFT  rowspan =5>
												(5)환급신청<br>&nbsp;&nbsp;세액 계산 
											</TD>
											 <TD CLASS="TD51" align =LEFT  colspan =2>
												(13)직전사업연도법인세액((13)=(10))
											</TD>
											
										    <TD CLASS="TD61" align =LEFT  >
												<script language =javascript src='./js/w8105ma1_txtW13_txtW13.js'></script>
											</TD>
											
										</TR>
										
										<TR>
										   
											 <TD CLASS="TD51" align =LEFT  colspan =2>
											(14)차감할 세액[((8)-(7))*세율][(14) ≥(10)-(12)]
											</TD>
											
										    <TD CLASS="TD61" align =LEFT  >
												<script language =javascript src='./js/w8105ma1_txtW14_txtW14.js'></script>
											</TD>
											
											
										</TR>
										<TR>
										   
											 <TD CLASS="TD51" align =LEFT  colspan =2>
												(15)환급신청세액((13)-(14))(15≤(12))
											</TD>
											
										    <TD CLASS="TD61" align =LEFT  >
												<script language =javascript src='./js/w8105ma1_txtW15_txtW15.js'></script>
											</TD>
											
											
										</TR>
					
					
											
									</TABLE>
						
						   			
					</TD>
				</TR>
		    
				
			</TABLE>
		</TD>
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
		<TD <%=HEIGHT_TYPE_01%>></TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtw9_VALUE" tag="24">

<INPUT TYPE=HIDDEN NAME="txtKeyStream" tag="24" >
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24" >



</FORM>
<DIV ID="MousePT" NAME="MousePT">

<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

