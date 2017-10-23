
<%@ LANGUAGE="VBSCRIPT" %>
<%'======================================================================================================

'*  1. Module Name          : 법인세 
'*  2. Function Name        : 제12호 농어촌특별세 과세표준 및 세액조정계산서 
'*  3. Program ID           : W8111MA1
'*  4. Program Name         : W8111MA1.asp
'*  5. Program Desc         : 제12호 농어촌특별세 과세표준 및 세액조정계산서 
'*  6. Modified date(First) : 2005/03/18
'*  7. Modified date(Last)  : 2005/03/18
'*  8. Modifier (First)     : 홍지영 
'*  9. Modifier (Last)      : 홍지영 
'* 10. Comment              : 참조: dbo.ufn_TB_12_GetRef 
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
Const BIZ_MNU_ID = "W8111MA1"											 '☆: 비지니스 로직 ASP명 
Const BIZ_PGM_ID = "W8111Mb1.asp"											 '☆: 비지니스 로직 ASP명 
Const EBR_RPT_ID = "W8103OA1"

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
			 lgKeyStream = lgKeyStream &  unicdbl(frm1.txtw5_Amt.value)            &  parent.gColSep       '3
			 lgKeyStream = lgKeyStream &  unicdbl(frm1.txtw5_rate.value)		   &  parent.gColSep       '4
			 lgKeyStream = lgKeyStream &  unicdbl(frm1.txtw5_rate_val.value)       &  parent.gColSep       '5
			 lgKeyStream = lgKeyStream &  unicdbl(frm1.txtw5_Tax.value)            &  parent.gColSep       '6
			 
	         lgKeyStream = lgKeyStream &  Trim(frm1.txtw6.value)                   &  parent.gColSep      '7
             lgKeyStream = lgKeyStream &  unicdbl(frm1.txtw6_Amt.value)            &  parent.gColSep      '8
			 lgKeyStream = lgKeyStream &  unicdbl(frm1.txtw6_rate.value)		   &  parent.gColSep      '9
			 lgKeyStream = lgKeyStream &  unicdbl(frm1.txtw6_rate_val.value)       &  parent.gColSep      '10
			 lgKeyStream = lgKeyStream &  unicdbl(frm1.txtw6_Tax.value)            &  parent.gColSep      '11
			 
			 lgKeyStream = lgKeyStream &  Trim(frm1.txtw7.value)                   &  parent.gColSep      '12
             lgKeyStream = lgKeyStream &  unicdbl(frm1.txtw7_Amt.value)            &  parent.gColSep      '13
			 lgKeyStream = lgKeyStream &  unicdbl(frm1.txtw7_rate.value)		   &  parent.gColSep      '14
			 lgKeyStream = lgKeyStream &  unicdbl(frm1.txtw7_rate_val.value)       &  parent.gColSep      '15
			 lgKeyStream = lgKeyStream &  unicdbl(frm1.txtw7_Tax.value)            &  parent.gColSep      '16
			 
		
             lgKeyStream = lgKeyStream &  unicdbl(frm1.txtw8_Amt.value)            &  parent.gColSep      '17
			 lgKeyStream = lgKeyStream &  unicdbl(frm1.txtw8_Tax.value)            &  parent.gColSep      '18
			 
			 
             lgKeyStream = lgKeyStream &  unicdbl(frm1.txtw10_Amt.value)           &  parent.gColSep       '19
			 lgKeyStream = lgKeyStream &  unicdbl(frm1.txtw10_rate.value)		   &  parent.gColSep       '20
			 lgKeyStream = lgKeyStream &  unicdbl(frm1.txtw10_rate_val.value)      &  parent.gColSep       '21
			 lgKeyStream = lgKeyStream &  unicdbl(frm1.txtw10_Tax.value)           &  parent.gColSep       '22
			 
			 lgKeyStream = lgKeyStream &  Trim(frm1.txtw11.value)                  &  parent.gColSep       '23
             lgKeyStream = lgKeyStream &  unicdbl(frm1.txtw11_Amt.value)           &  parent.gColSep       '24
			 lgKeyStream = lgKeyStream &  unicdbl(frm1.txtw11_rate.value)		   &  parent.gColSep       '25
			 lgKeyStream = lgKeyStream &  unicdbl(frm1.txtw11_rate_val.value)      &  parent.gColSep       '26
			 lgKeyStream = lgKeyStream &  unicdbl(frm1.txtw11_Tax.value)           &  parent.gColSep       '27
			 
             lgKeyStream = lgKeyStream &  unicdbl(frm1.txtw12_Amt.value)           &  parent.gColSep       '28
			 lgKeyStream = lgKeyStream &  unicdbl(frm1.txtw12_Tax.value)           &  parent.gColSep       '29
			 
			 
			 
   end if
    
    

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

    Call SetToolbar("1100100000000111")										<%'버튼 툴바 제어 %>

	' 변경한곳 
    Call InitComboBox
    Call  AppendNumberPlace("7", "3", "0")
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
    Call ggoOper.FormatDate(frm1.txtFISC_YEAR, parent.gDateFormat,3)
    Call initData
    Call SetDefaultVal()
    Call InitVariables  
     
    ' 세무조정 체크호출 
	Call FncQuery
  
End Sub




Sub SetDefaultVal()
dim strWhere 
DIM strW1
Dim sFiscYear, sRepType, sCoCd, iGap
   '농특세율(w3001)
    call CommonQueryRs("REFERENCE_1,REFERENCE_2"," ufn_TB_Configuration('w3001','" & C_REVISION_YM & "')   "," Minor_cd = '1' ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)  
         frm1.txtW10_rate_val.value  = unicdbl(lgF0)
         frm1.txtW10_rate.value = unicdbl(replace(lgF1,"%",""))
         
         frm1.txtW5_rate_val.value  = unicdbl(lgF0)
         frm1.txtW5_rate.value = unicdbl(replace(lgF1,"%",""))
  

End Sub



' ----------------------  검증 -------------------------
Function  Verification()
  dim IntRetCD
  dim i ,chkObj
	Verification = False
      
	Verification = True	
End Function
'============================================  사용자정의  ====================================




Function GetRef()	' 금액가져오기 링크 클릭시 
	Dim sFiscYear, sRepType, sCoCd, IntRetCD
	Dim sMesg
	Dim w5 , w10
    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
	
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
	Call selectColor(frm1.txtW5_AMT)
    Call selectColor(frm1.txtW10_AMT)

    
    IntRetCD = DisplayMsgBox("WC0003", parent.VB_YES_NO, sMesg, "X")           '⊙: "Will you destory previous data"
	Call ggoOper.LockField(Document, "N") 
	 If IntRetCD = vbNo Then
	 	Exit Function
	 End If

 

   
   'W5 =  13호 서식의 1-⑩ 감면세액합계를 입력함.
   'W10 = 13호 서식의 2-⑦감면세액을 입력함.


	call CommonQueryRs("W5,W10","dbo.ufn_TB_12_GetRef('" & sCoCd & "', '" & sFiscYear & "', '" & sRepType & "')", "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

	     IF lgF0 = "" THEN EXIT Function 
	         W5	 = unicdbl(replace(lgF0, chr(11),""))		 
             W10    = unicdbl(replace(lgF1, chr(11),""))		
 	         frm1.txtW5_AMT.value	 = W5 
             frm1.txtW10_AMT.value   = W10	

    
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








'==========================================================================================
Sub InitData()
	Dim intRow
	Dim intIndex 
		
	frm1.txtFISC_YEAR.text 	= "<%=wgFISC_YEAR%>"
	frm1.txtCO_CD.value 	= "<%=wgCO_CD%>"
	frm1.txtCO_NM.value 	= "<%=wgCO_NM%>"
	frm1.cboREP_TYPE.value 	= "<%=wgREP_TYPE%>"

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
     Call initData
     Call SetDefaultVal()   
     
     lgBlnFlgChgValue = False        

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


Sub txtW6_Change()
	 lgBlnFlgChgValue = True
End Sub

Sub txtW7_Change()
	 lgBlnFlgChgValue = True
End Sub

Sub txtW11_Change()
	 lgBlnFlgChgValue = True
End Sub






Function CalSum(ByVal Obj, ByVal Obj1, byval Obj2)

    Obj.value =  UNICDbl(Obj1.value) * UNICDbl(Obj2.value)
    frm1.txtW8_AMT.Value		=  unicdbl(frm1.txtW5_AMT.value) + unicdbl(frm1.txtW6_AMT.value) + unicdbl(frm1.txtW7_AMT.value)
    frm1.txtW8_TAX.value		=  unicdbl(frm1.txtW5_TAX.value) + unicdbl(frm1.txtW6_TAX) + unicdbl(frm1.txtW7_TAX.value)
    frm1.txtW12_AMT.Value		=  unicdbl(frm1.txtW10_AMT.value) + unicdbl(frm1.txtW11_AMT.value)
    frm1.txtW12_TAX.Value		=  unicdbl(frm1.txtW10_TAX.value) + unicdbl(frm1.txtW11_TAX.value)
end function



Sub txtW5_AMT_Change()
      Call CalSum( frm1.txtW5_TAX ,frm1.txtW5_AMT, frm1.txtW5_RATE_VAL) 
	lgBlnFlgChgValue = True
End Sub

Sub txtW6_AMT_Change()
      Call CalSum( frm1.txtW6_TAX ,frm1.txtw6_AMT, frm1.txtW6_RATE_VAL) 
	lgBlnFlgChgValue = True
End Sub

Sub txtW7_AMT_Change()
      Call CalSum( frm1.txtW7_TAX ,frm1.txtW7_AMT, frm1.txtW7_RATE_VAL) 
	lgBlnFlgChgValue = True
End Sub

Sub txtW10_AMT_Change()
      Call CalSum( frm1.txtW10_TAX ,frm1.txtW10_AMT, frm1.txtW10_RATE_VAL) 
	lgBlnFlgChgValue = True
End Sub

Sub txtW11_AMT_Change()
      Call CalSum( frm1.txtW11_TAX ,frm1.txtW11_AMT, frm1.txtW11_RATE_VAL) 
	lgBlnFlgChgValue = True
End Sub


Sub txtW5_RATE_Change()
      Call CalSum( frm1.txtW5_TAX ,frm1.txtW5_AMT, frm1.txtW5_RATE_VAL) 
	lgBlnFlgChgValue = True
End Sub

Sub txtW6_rate_Change()
     frm1.txtW6_rate_val.value = unicdbl(frm1.txtW6_rate.value) * 0.01
     Call CalSum( frm1.txtW6_TAX ,frm1.txtw6_AMT, frm1.txtW6_RATE_VAL) 
	lgBlnFlgChgValue = True
End Sub

Sub txtW7_RATE_Change()

     frm1.txtW7_rate_val.value = unicdbl(frm1.txtW7_rate.value) * 0.01
      Call CalSum( frm1.txtW7_TAX ,frm1.txtW7_AMT, frm1.txtW7_RATE_VAL) 
	lgBlnFlgChgValue = True
End Sub

Sub txtW10_Rate_Change()
      frm1.txtW10_rate_val.value = unicdbl(frm1.txtW10_rate.value) * 0.01
      Call CalSum( frm1.txtW10_TAX ,frm1.txtW10_AMT, frm1.txtW10_RATE_VAL) 
	lgBlnFlgChgValue = True
End Sub

Sub txtW11_Rate_Change()
     frm1.txtW11_rate_val.value = unicdbl(frm1.txtW11_rate.value) * 0.01
     Call CalSum( frm1.txtW11_TAX ,frm1.txtW11_AMT, frm1.txtW11_RATE_VAL) 
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
									<TD CLASS="TD6"><script language =javascript src='./js/w8111ma1_txtFISC_YEAR_txtFISC_YEAR.js'></script>
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
				        <TD WIDTH=800 valign=top  >
					   
				
					    
							<TABLE width = 100%  bgcolor = #696969  border = 0 cellpadding = 1 cellspacing = 1 ID="Table2">
							  	        <TR>
							 				 <TD CLASS="TD51" align =center  rowspan =2   >
							 					 (1)법인유형 
							 				</TD>
							 				<TD CLASS="TD51" align =center colspan =2 >
							 					(2)과세표준 
							 				</TD>
							 				<TD CLASS="TD51" align =center rowspan =2 >
							 					세율 
							 				</TD>
							 		        <TD CLASS="TD51" align =center rowspan =2 >
							 					(3)세액 
							 				</TD>
							 			</TR>
							 			  <TR>
							 				 <TD CLASS="TD51" align =center  >
							 					 구분 
							 				</TD>
							 				<TD CLASS="TD51" align =center  >
							 					 금액 
							 				</TD>
																	
							 			</TR>
							 			<TR>
							 				 <TD CLASS="TD51" align =center   rowspan =4  >
							 					(4)일반법인 
							 				</TD>
							 			     <TD CLASS="TD51" align =left  >
							 					(5)법인세감면세액 
							 				</TD>
																		 
							 				 <TD CLASS="TD51" align =center   >
							 					<script language =javascript src='./js/w8111ma1_txtW5_AMT_txtW5_AMT.js'></script>
							 				</TD>
							 				<TD CLASS="TD51" align =center   >
							 					<script language =javascript src='./js/w8111ma1_txtW5_RATE_txtW5_RATE.js'></script>%
							 					<INPUT TYPE=HIDDEN  id=txtW5_rate_val name=txtW5_rate_val ALT="(5)법인세감면비율" >
							 				</TD>
							 				<TD CLASS="TD51" align =center   >
							 					<script language =javascript src='./js/w8111ma1_txtW5_tax_txtW5_tax.js'></script>
							 				</TD>
																		 
																
							 			</TR>
							 				<TR>
											
							 				 <TD CLASS="TD51" align =left  >
							 					(6)<INPUT type="text" id=txtW6 name=txtW6 tag="25X26" >
							 				</TD>
							 				  <TD CLASS="TD51" align =center   >
							 					<script language =javascript src='./js/w8111ma1_txtW6_AMT_txtW6_AMT.js'></script>
							 				</TD>
							 				<TD CLASS="TD51" align =center   >
							 					<script language =javascript src='./js/w8111ma1_txtW6_RATE_txtW6_RATE.js'></script>%
							 					<INPUT TYPE=HIDDEN  id=txtW6_rate_val name=txtW6_rate_val ALT="(6)" >
							 				</TD>
							 				<TD CLASS="TD51" align =center   >
							 					<script language =javascript src='./js/w8111ma1_txtW6_tax_txtW6_tax.js'></script>
							 				</TD>
																		 
																
							 			</TR>
										<TR>
																		
							 			  
							 				<TD CLASS="TD51" align =left  >
							 					(7)<INPUT type="text" id=txtW7 name=txtW7 tag="25X26" >
							 				</TD>
							 				  <TD CLASS="TD51" align =center   >
							 					<script language =javascript src='./js/w8111ma1_txtW7_AMT_txtW7_AMT.js'></script>
							 				</TD>
							 				<TD CLASS="TD51" align =center   >
							 					<script language =javascript src='./js/w8111ma1_txtW7_RATE_txtW7_RATE.js'></script>%
							 					<INPUT TYPE=HIDDEN  id=txtW7_rate_val name=txtW7_rate_val ALT="(6)" >
							 				</TD>
							 				<TD CLASS="TD51" align =center   >
							 					<script language =javascript src='./js/w8111ma1_txtW7_tax_txtW7_tax.js'></script>
							 				</TD>
																		 
																
							 			</TR>
												
															
							 			<TR>
																		
							 			
							 				<TD CLASS="TD51" align =left   >
							 					(8)소계 
							 				</TD>
							 				  <TD CLASS="TD51" align =center   >
							 					<script language =javascript src='./js/w8111ma1_txtW8_AMT_txtW8_AMT.js'></script>
							 				</TD>
							 				<TD CLASS="TD51" align =center   >
							 					
							 				</TD>
							 				<TD CLASS="TD51" align =center   >
							 					<script language =javascript src='./js/w8111ma1_txtW8_tax_txtW8_tax.js'></script>
							 				</TD>
																		 
																
							 			</TR>
															
							 		<TR>
							 				 <TD CLASS="TD51" align =center   rowspan =3  >
							 					(9)조합법인등 
							 				</TD>
							 			     <TD CLASS="TD51" align =left  >
							 					(10)법인세감면세액 
							 				</TD>
																		 
							 				 <TD CLASS="TD51" align =center   >
							 					<script language =javascript src='./js/w8111ma1_txtW10_AMT_txtW10_AMT.js'></script>
							 				</TD>
							 				<TD CLASS="TD51" align =center   >
							 					<script language =javascript src='./js/w8111ma1_txtW10_RATE_txtW10_RATE.js'></script>%
							 					<INPUT TYPE=HIDDEN  id=txtW10_rate_val name=txtW10_rate_val ALT="(10)" >
							 				</TD>
							 				<TD CLASS="TD51" align =center   >
							 					<script language =javascript src='./js/w8111ma1_txtW10_tax_txtW10_tax.js'></script>
							 				</TD>
																		 
																
							 			</TR>
							 			<TR>
																		
							 			    
							 				<TD CLASS="TD51" align =left  >
							 					(11)<INPUT type="text" id=txtW11 name=txtW11 tag="25X26" >
							 				</TD>
							 				  <TD CLASS="TD51" align =center   >
							 					<script language =javascript src='./js/w8111ma1_txtW11_AMT_txtW11_AMT.js'></script>
							 				</TD>
							 				<TD CLASS="TD51" align =center   >
							 					<script language =javascript src='./js/w8111ma1_txtW11_RATE_txtW11_RATE.js'></script>%
							 					<INPUT TYPE=HIDDEN  id=txtW11_rate_val name=txtW11_rate_val ALT="(11)" >
							 				</TD>
							 				<TD CLASS="TD51" align =center   >
							 					<script language =javascript src='./js/w8111ma1_txtW11_tax_txtW11_tax.js'></script>
							 				</TD>
																		 
																
							 			</TR>
												
																		
															
							 				<TR>
										
							 				<TD CLASS="TD51" align =left   >
							 					(12)소계 
							 				</TD>
							 				  <TD CLASS="TD51" align =center   >
							 					<script language =javascript src='./js/w8111ma1_txtW12_AMT_txtW12_AMT.js'></script>
							 				</TD>
							 				<TD CLASS="TD51" align =center   >
							 			
							 				</TD>
							 				<TD CLASS="TD51" align =center   >
							 					<script language =javascript src='./js/w8111ma1_txtW12_tax_txtW12_tax.js'></script>
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



</FORM>
<DIV ID="MousePT" NAME="MousePT">

<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

