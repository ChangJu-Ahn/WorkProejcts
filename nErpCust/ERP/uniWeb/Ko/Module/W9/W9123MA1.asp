<%@ LANGUAGE="VBScript" CODEPAGE=949%>
<% session.CodePage=949 %>
<%'======================================================================================================

'*  1. Module Name          : 법인세 
'*  2. Function Name        : 전산1호전산운용조직명세서 
'*  3. Program ID           : W9123MA1
'*  4. Program Name         : W9123MA1.asp
'*  5. Program Desc         : 전산1호전산운용조직명세서 
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
Const BIZ_MNU_ID = "W9123MA1"											 '☆: 비지니스 로직 ASP명 
Const BIZ_PGM_ID = "W9123MB1.asp"											 '☆: 비지니스 로직 ASP명 
Const EBR_RPT_ID = "W9123OA1"

Dim IsOpenPop          
Dim lgStrPrevKey2
Dim lgRefMode	' 레퍼런스 참조상태 : ERP 추출사용유무 

dim strW1_R
dim strW5_R

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
	Dim strW9_1, strW9_2, strW9_3, strW9_4, strW9_5, strW9_6
	Dim strW10_1, strW10_2, strW10_3, strW10_4, strW10_5, strW10_6, strW10_7, strW10_8, strW10_9
   
    lgKeyStream = UCASE(Frm1.txtCo_Cd.Value) &  parent.gColSep  
   lgKeyStream = lgKeyStream & (Frm1.txtFISC_YEAR.text) &  parent.gColSep ' 
   lgKeyStream = lgKeyStream & UCASE(Frm1.cboREP_TYPE.Value ) &  parent.gColSep '   
  

   if pOpt = "S" then
     if Frm1.chkW9_1.checked = True then 
         strW9_1 = "Y"
      else
         strW9_1 = "N"
      End if
      
      if Frm1.chkW9_2.checked = True then 
         strW9_2 = "Y"
      else
         strW9_2 = "N"
      End if                                    
      if Frm1.chkW9_3.checked = True then 
         strW9_3 = "Y"
      else
         strW9_3 = "N"
      End if                                    
      if Frm1.chkW9_4.checked = True then 
         strW9_4 = "Y"
      else
         strW9_4 = "N"
      End if                                    
      if Frm1.chkW9_5.checked = True then 
         strW9_5 = "Y"
      else
         strW9_5 = "N"
      End if                                    
      if Frm1.chkW9_6.checked = True then 
         strW9_6 = "Y"
      else
         strW9_6 = "N"
      End if                                    
      
      if Frm1.chkW10_1.checked = True then 
         strW10_1 = "Y"
      else
         strW10_1 = "N"
      End if                                    
      if Frm1.chkW10_2.checked = True then 
         strW10_2 = "Y"
      else
         strW10_2 = "N"
      End if                                    
      if Frm1.chkW10_3.checked = True then 
         strW10_3 = "Y"
      else
         strW10_3 = "N"
      End if                                    
      if Frm1.chkW10_4.checked = True then 
         strW10_4 = "Y"
      else
         strW10_4 = "N"
      End if
      
      if Frm1.chkW10_5.checked = True then 
         strW10_5 = "Y"
      else
         strW10_5 = "N"
      End if
      
      if Frm1.chkW10_6.checked = True then 
         strW10_6 = "Y"
      else
         strW10_6 = "N"
      End if     
      if Frm1.chkW10_7.checked = True then 
         strW10_7 = "Y"
      else
         strW10_7 = "N"
      End if     
      if Frm1.chkW10_8.checked = True then 
         strW10_8 = "Y"
      else
         strW10_8 = "N"
      End if    
      if Frm1.chkW10_9.checked = True then 
         strW10_9 = "Y"
      else
         strW10_9 = "N"
      End if  
      
    
    
    
      lgKeyStream = lgKeyStream &  Trim(Frm1.cboW1.Value ) &  parent.gColSep		'3  
      lgKeyStream = lgKeyStream &  Trim(Frm1.cboW2.Value ) &  parent.gColSep		'4
      lgKeyStream = lgKeyStream &  Trim(Frm1.txtw2_Etc.Value ) &  parent.gColSep	'5  
      lgKeyStream = lgKeyStream &  Trim(Frm1.cboW3.Value ) &  parent.gColSep		'6
      lgKeyStream = lgKeyStream &  Trim(Frm1.txtw3_Etc.Value ) &  parent.gColSep	'7   
      lgKeyStream = lgKeyStream &  Trim(Frm1.cboW4.Value ) &  parent.gColSep		'8     
      lgKeyStream = lgKeyStream &  Trim(Frm1.txtw4_Etc.Value ) &  parent.gColSep	'9     
      lgKeyStream = lgKeyStream &  Trim(Frm1.txtw4_1.Value ) &  parent.gColSep		'10    
      lgKeyStream = lgKeyStream &  Trim(Frm1.cboW5.Value ) &  parent.gColSep		'11
      lgKeyStream = lgKeyStream &  Trim(Frm1.txtw5_Etc.Value ) &  parent.gColSep	'12
      lgKeyStream = lgKeyStream &  Trim(Frm1.cboW6.Value ) &  parent.gColSep		'13
      lgKeyStream = lgKeyStream &  Trim(Frm1.txtw6_Etc.Value ) &  parent.gColSep	'14        
      lgKeyStream = lgKeyStream &  Trim(Frm1.txtw7_1.Value ) &  parent.gColSep		'15
      lgKeyStream = lgKeyStream &  Trim(Frm1.txtw7_2.Value ) &  parent.gColSep		'16
      lgKeyStream = lgKeyStream &  Trim(Frm1.cboW8.Value ) &  parent.gColSep		'17    
      lgKeyStream = lgKeyStream &  Trim(strW9_1) &  parent.gColSep					'18
      lgKeyStream = lgKeyStream &  Trim(strW9_2) &  parent.gColSep					'19        
      lgKeyStream = lgKeyStream &  Trim(strW9_3) &  parent.gColSep					'20    
      lgKeyStream = lgKeyStream &  Trim(strW9_4) &  parent.gColSep					'21    
      lgKeyStream = lgKeyStream &  Trim(strW9_5) &  parent.gColSep					'22    
      lgKeyStream = lgKeyStream &  Trim(strW9_6) &  parent.gColSep					'23    
      lgKeyStream = lgKeyStream &  Trim(frm1.txtW9_6_ETC.value ) &  parent.gColSep ' 24   
      lgKeyStream = lgKeyStream &  Trim(strW10_1) &  parent.gColSep					'25
      lgKeyStream = lgKeyStream &  Trim(strW10_2) &  parent.gColSep					'26        
      lgKeyStream = lgKeyStream &  Trim(strW10_3) &  parent.gColSep					'27   
      lgKeyStream = lgKeyStream &  Trim(strW10_4) &  parent.gColSep					'28    
      lgKeyStream = lgKeyStream &  Trim(strW10_5) &  parent.gColSep					'29    
      lgKeyStream = lgKeyStream &  Trim(strW10_6) &  parent.gColSep					'30    
      lgKeyStream = lgKeyStream &  Trim(strW10_7) &  parent.gColSep					'31    
      lgKeyStream = lgKeyStream &  Trim(strW10_8) &  parent.gColSep					'32    
      lgKeyStream = lgKeyStream &  Trim(strW10_9) &  parent.gColSep					'33    
      lgKeyStream = lgKeyStream &  Trim(frm1.txtw10_9_ETC.Value) &  parent.gColSep	'34    
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                      
      

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
    Call InitVariables                                                      <%'Initializes local global variables%>
    
    Call SetToolbar("1100100000000111")										<%'버튼 툴바 제어 %>

	' 변경한곳 
    Call InitComboBox
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.FormatField(Document, "2", ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
    Call ggoOper.FormatDate(frm1.txtFISC_YEAR, parent.gDateFormat,3)
    Call SetDefaultVal()
	Call FncQuery
     
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


'============================================  이벤트 함수  ====================================
Sub txtFISC_YEAR_DblClick(Button)
    If Button = 1 Then
        frm1.txtFISC_YEAR.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtFISC_YEAR.Focus
    End If
End Sub









Sub cboW1_Onchange()  

    lgBlnFlgChgValue  = True 
    if frm1.cboW1.value = "1" Or frm1.cboW1.value = "2"  then
       Call ggoOper.SetReqAttr(frm1.cboW3, "N")
    else   
       frm1.cboW3.value =""
       Call ggoOper.SetReqAttr(frm1.cboW3, "Q")
       
    end if
    if  frm1.cboW1.value = "1" Or frm1.cboW1.value = "2"  Or frm1.cboW1.value = "3" then
        Call ggoOper.SetReqAttr(frm1.cboW4, "N")
        Call ggoOper.SetReqAttr(frm1.txtw4_1, "N")
    else
        frm1.cboW4.value =""
        frm1.txtw4_1.value =""
        Call ggoOper.SetReqAttr(frm1.cboW4, "Q")
        Call ggoOper.SetReqAttr(frm1.txtw4_1, "Q")
    end if
     if frm1.cboW1.value = "3" then
        Call ggoOper.SetReqAttr(frm1.cboW5, "N")
     Else
          frm1.cboW5.value =""
        Call ggoOper.SetReqAttr(frm1.cboW5, "Q")   
     end if
     
     
     if frm1.cboW1.value = "4" then
        Call ggoOper.SetReqAttr(frm1.cboW6, "N")
     else
             frm1.cboW6.value =""
        Call ggoOper.SetReqAttr(frm1.cboW6, "Q")
     end if                    
End Sub


Sub cboW2_Onchange()  

    lgBlnFlgChgValue  = True 
    if frm1.cboW2.value = "6"   then
       Call ggoOper.SetReqAttr(frm1.txtw2_Etc, "N")
    else  
       frm1.txtw2_Etc.value ="" 
       Call ggoOper.SetReqAttr(frm1.txtw2_Etc, "Q")
    end if

End Sub

Sub txtw2_Etc_Onchange()  

    lgBlnFlgChgValue  = True 
    
End Sub

Sub cbow3_Onchange()  

    lgBlnFlgChgValue  = True 
    if frm1.cboW3.value = "8"   then
       Call ggoOper.SetReqAttr(frm1.txtw3_Etc, "N")
    else   
        frm1.txtw3_Etc.value ="" 
       Call ggoOper.SetReqAttr(frm1.txtw3_Etc, "Q")
    end if

End Sub

Sub txtw3_Etc_Onchange()  

    lgBlnFlgChgValue  = True 
   
End Sub


Sub cbow4_Onchange()  

    lgBlnFlgChgValue  = True 
    if frm1.cboW4.value = "8"  then
       Call ggoOper.SetReqAttr(frm1.txtw4_Etc, "N")

    else  
       frm1.txtw4_Etc.value =""  
       Call ggoOper.SetReqAttr(frm1.txtw4_Etc, "Q")
    
    end if

End Sub

Sub txtw4_Etc_Onchange()  

    lgBlnFlgChgValue  = True 
    
End Sub

Sub cbow5_Onchange()  

    lgBlnFlgChgValue  = True 
    if frm1.cboW5.value = "5"   then
       Call ggoOper.SetReqAttr(frm1.txtw5_Etc, "N")
    else   
       frm1.txtw5_Etc.value =""  
       Call ggoOper.SetReqAttr(frm1.txtw5_Etc, "Q")
    end if

End Sub

Sub txtw5_Etc_Onchange()  

    lgBlnFlgChgValue  = True 
    
End Sub



Sub cbow6_Onchange()  

    lgBlnFlgChgValue  = True 
    if frm1.cboW6.value = "5"   then
       Call ggoOper.SetReqAttr(frm1.txtw6_Etc, "N")
    else   
        frm1.txtw6_Etc.value =""  
       Call ggoOper.SetReqAttr(frm1.txtw6_Etc, "Q")
    end if
    

End Sub

Sub txtw6Etc_Onchange()  

    lgBlnFlgChgValue  = True 
    
End Sub


Sub cbow8_Onchange()  
dim i
    lgBlnFlgChgValue  = True 
    if frm1.cboW8.value = "1"   then

       Call ggoOper.SetReqAttr(frm1.chkW9_1, "D")
       Call ggoOper.SetReqAttr(frm1.chkW9_2, "D")
       Call ggoOper.SetReqAttr(frm1.chkW9_3, "D")
       Call ggoOper.SetReqAttr(frm1.chkW9_4, "D")
       Call ggoOper.SetReqAttr(frm1.chkW9_5, "D")
       Call ggoOper.SetReqAttr(frm1.chkW9_6, "D")

       
    
       
    else 
      
       
       frm1.chkW9_1.Checked = False
       frm1.chkW9_2.Checked = False
       frm1.chkW9_3.Checked = False
       frm1.chkW9_4.Checked = False
       frm1.chkW9_5.Checked = False
       frm1.chkW9_6.Checked = False
           
       Call ggoOper.SetReqAttr(frm1.chkW9_1, "Q")
       Call ggoOper.SetReqAttr(frm1.chkW9_2, "Q")
       Call ggoOper.SetReqAttr(frm1.chkW9_3, "Q")
       Call ggoOper.SetReqAttr(frm1.chkW9_4, "Q")
       Call ggoOper.SetReqAttr(frm1.chkW9_5, "Q")
       Call ggoOper.SetReqAttr(frm1.chkW9_6, "Q")
       Call ggoOper.SetReqAttr(frm1.txtW9_6_ETC, "Q")
       
    end if
    

End Sub


Sub chkw9_6_OnClick()  

    lgBlnFlgChgValue  = True 
    if frm1.chkW9_6.Checked =true    then
       Call ggoOper.SetReqAttr(frm1.txtW9_6_ETC, "N")
       
    else   
       Call ggoOper.SetReqAttr(frm1.txtW9_6_ETC, "Q")
        frm1.txtW9_6_ETC.value =""
    end if
    

End Sub


Sub chkw10_9_OnClick()  

    lgBlnFlgChgValue  = True 
    if frm1.chkw10_9.Checked =true    then
       Call ggoOper.SetReqAttr(frm1.txtw10_9_ETC, "N")
    else   
       Call ggoOper.SetReqAttr(frm1.txtw10_9_ETC, "Q")
        frm1.txtw10_9_ETC.value =""
    end if
    

End Sub



Sub txtW9_6_ETC_Onchange()  

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
    
        If Not chkField(Document, "2") Then
			Exit Function
	    End If
    
    
<%  '-----------------------
    'Save function call area
    '----------------------- %>
    Call ggoOper.LockField(Document, "N")
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
		IntRetCD = DisplayMsgBox("800442", parent.VB_YES_NO, "X", "X")			    <%'%>
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
	
	   Call cboW1_onchange
	   Call cboW2_onchange
	   Call cboW3_onchange
	   Call cboW4_onchange
	   Call cboW5_onchange
	   Call cboW6_onchange
	   Call cboW8_onchange
	   Call chkw9_6_OnClick
	   Call chkw10_9_OnClick
	   
	   '2 디비환경값 , 로드시환경값 비교 
		 Call SetToolbar("1101100000000111")										<%'버튼 툴바 제어 %>
		
	End If
   lgBlnFlgChgValue = false
		
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
									<TD CLASS="TD6"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name=txtFISC_YEAR CLASS=FPDTYYYY title=FPDATETIME ALT="당기시작일자" tag="14X1" id=txtFISC_YEAR></OBJECT>');</SCRIPT>
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
					<TD WIDTH=100% valign=top  >
					   
					    
							   <DIV ID="TabDiv" STYLE="FLOAT: left; HEIGHT: 100%; WIDTH: 100%; overflow=auto">
								
									<FIELDSET CLASS="CLSFLD"><LEGEND ALIGN=LEFT>전산조직운용명세서</LEGEND>
									    <table Height=100% WIDTH=620 CELLSPACING=0 CELLPADDING="7"  >
											<TR CLASS="TD51">
												<TD WIDTH=100% Height="30" ALIGN=CENTER  colspan = 2></TD>
											
											</TR>
											<TR CLASS="TD61">
												<TD WIDTH=60% Height="20" ALIGN=Left >1.회계프로그램(시스템) 사용현황	</TD>
												<TD WIDTH=40% Height="20" ALIGN=Left ><SELECT NAME="cboW1" ALT="1.회계프로그램(시스템) 사용현황" tag="25X1" onChange='Call cboW1_onchange'>
																							<option value = ""></option>
																							<option value = 1>①</option>
																							<option value = 2>②</option>
																							<option value = 3>③</option>
																							<option value = 4>④</option>
																										</SELECT>
												</TD>
												
												
											</TR>
											
											<TR CLASS="TD61">
												<TD CLASS="TD6"  colspan = 2>&nbsp;&nbsp;&nbsp;① 자체개발&nbsp;&nbsp;&nbsp;② 외주개발&nbsp;&nbsp;&nbsp;③ ERP&nbsp;&nbsp;&nbsp;④ 상업용 회계프로그램<Br>
				                			   
				                			    
				                			    </TD>
											
												
												
											</TR>
											
											   			
											<TR CLASS="TD61">
												<TD WIDTH=60% Height="20" ALIGN=Left >2.OS(운영체제)</TD>	
												<TD WIDTH=40% Height="20" ALIGN=Left ><SELECT NAME="cboW2" ALT="2.OS(운영체제)" tag="25X1" onChange='Call cboW2_onchange'>
																							<option value = ""></option>
																							<option value = 1>①</option>
																							<option value = 2>②</option>
																							<option value = 3>③</option>
																							<option value = 4>④</option>
																							<option value = 5>⑤</option>
																							<option value = 6>⑥</option>
																										</SELECT>
												</TD>
								
											</TR>
											<TR CLASS="TD61">
												<TD CLASS="TD6"  colspan = 2>&nbsp;&nbsp;&nbsp;① Unix계열(Linux, Zenix, HP-UX)&nbsp;&nbsp;&nbsp;② 서버급 Windows(NT, 2000, 2003)</TD>
				                				
											</TR>
											<TR CLASS="TD61">
												<TD CLASS="TD6"  colspan = 2>&nbsp;&nbsp;&nbsp;③ IBM OS계열(OS/2, OS/390, OS/400)&nbsp;&nbsp;&nbsp;④ PC용 Windows</TD>
				                				
											</TR>
											<TR CLASS="TD61">
												<TD CLASS="TD6"  colspan = 2>&nbsp;&nbsp;&nbsp;⑤ 세무대리인이 대리&nbsp;&nbsp;&nbsp⑥ 기타 (<INPUT NAME="txtw2_Etc" ALT="기타" TYPE="Text" MAXLENGTH=50 SiZE=20 tag=24>)</TD>
				                				
											</TR>
											
											
											<TR CLASS="TD61">
												<TD WIDTH=60%% Height="20" ALIGN=Left >3.프로그램 언어(1번 ①② 응답자만)</TD>	
												<TD WIDTH=40% Height="20" ALIGN=Left ><SELECT NAME="cboW3" ALT="3.프로그램 언어" tag="24XXXU">
																							<option value = ""></option>
																							<option value = 1>①</option>
																							<option value = 2>②</option>
																							<option value = 3>③</option>
																							<option value = 4>④</option>
																							<option value = 5>⑤</option>
																							<option value = 6>⑥</option>
																							<option value = 7>⑦</option>
																							<option value = 8>⑧</option>
																										</SELECT>
												</TD>
											
											</TR>
											<TR CLASS="TD61">
												<TD CLASS="TD6"  colspan = 2>&nbsp;&nbsp;&nbsp;① C++&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
												                                               ② Delphi&nbsp;&nbsp;&nbsp;&nbsp&nbsp;&nbsp;&nbsp;
												                                               ③ COBOL&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
												                                               ④ Power Builder</TD>
				                				
											</TR>
											<TR CLASS="TD61">
												<TD CLASS="TD6"  colspan = 2>&nbsp;&nbsp;&nbsp;⑤ Visual Basic&nbsp;&nbsp;
																							   ⑥ Visual C++&nbsp;&nbsp;&nbsp;
																							   ⑦ C#&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
																							   ⑧ 기타(<INPUT NAME="txtw3_Etc" ALT="기타" TYPE="Text" MAXLENGTH=50 SiZE=20 tag=24>)</TD>
				                				
											</TR>
											
											
											<TR CLASS="TD61">
												<TD WIDTH=40% Height="20" ALIGN=Left >4.DBMS(1번 ①②③  응답자만)</TD>	
												<TD WIDTH=40% Height="20" ALIGN=Left ><SELECT NAME="cboW4" ALT="4.DBM" tag="24X1">
																							<option value = ""></option>
																							<option value = 1>①</option>
																							<option value = 2>②</option>
																							<option value = 3>③</option>
																							<option value = 4>④</option>
																							<option value = 5>⑤</option>
																							<option value = 6>⑥</option>
																							<option value = 7>⑦</option>
																							<option value = 8>⑧</option>
																										</SELECT>
												</TD>
											
											</TR>
												<TR CLASS="TD61">
												<TD CLASS="TD6"  colspan = 2>&nbsp;&nbsp;&nbsp;① Oracle&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
												                                               ② AS/400&nbsp;&nbsp;&nbsp;&nbsp&nbsp;&nbsp;&nbsp;
												                                               ③ IBM DB2&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
												                                               ④ MS SQL Server</TD>
				                				
											</TR>
											<TR CLASS="TD61">
												<TD CLASS="TD6"  colspan = 2>&nbsp;&nbsp;&nbsp;⑤ My SQL&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
																							   ⑥ Sybase&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
																							   ⑦ Informix&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
																							   ⑧ 기타(<INPUT NAME="txtw4_Etc" ALT="기타" TYPE="Text" MAXLENGTH=50 SiZE=20 tag=24>)</TD>
				                				
											</TR>
											
											<TR CLASS="TD61">
												<TD WIDTH=40% Height="20" ALIGN=Left>4-1.DBMS Version</TD>	
												<TD WIDTH=40% Height="20" ALIGN=Left >(<INPUT NAME="txtw4_1" ALT="DBMS Version" TYPE="Text" MAXLENGTH=30 SiZE=20 tag=24>)</TD>
											
											</TR>
											
											<TR CLASS="TD61">
												<TD WIDTH=60% Height="20" ALIGN=Left >5.ERP(1번 ③응답자만)	</TD>
												<TD WIDTH=40% Height="20" ALIGN=Left ><SELECT NAME="cboW5" ALT="5.ERP" tag="24X1">
																							<option value = ""></option>
																							<option value = 1>①</option>
																							<option value = 2>②</option>
																							<option value = 3>③</option>
																							<option value = 4>④</option>
																							<option value = 5>⑤</option>
																							<option value = 5>⑥</option>
																										</SELECT>
												</TD>
												
												
											</TR>
											
											<TR CLASS="TD61">
												<TD CLASS="TD6"  colspan = 2>&nbsp;&nbsp;&nbsp;① SAP&nbsp;&nbsp;
																							   ② ORACLE&nbsp;
																							   ③ UNI-ERP&nbsp;
																							   ④ 자체개발&nbsp;
																							   ⑤ 기타(<INPUT NAME="txtw5_Etc" ALT="기타" TYPE="Text" MAXLENGTH=50 SiZE=20 tag=24>)&nbsp;&nbsp;
																							   ⑥ 더존ERP</TD>
				                			   
				                			    
				                			    </TD>
											
												
												
											</TR>
											
											
											
											<TR CLASS="TD61">
												<TD WIDTH=40% Height="20" ALIGN=Left >6.상업용 회계프로그램(1번 ④ 응답자만)</TD>	
												<TD WIDTH=40% Height="20" ALIGN=Left ><SELECT NAME="cboW6" ALT="6.상업용 회계프로그램" tag="24X1">
																							<option value = ""></option>
																							<option value = 1>①</option>
																							<option value = 2>②</option>
																							<option value = 3>③</option>
																							<option value = 4>④</option>
																							<option value = 5>⑤</option>
																										</SELECT>
												</TD>
											
											</TR>
											<TR CLASS="TD61">
												<TD CLASS="TD6"  colspan = 2>&nbsp;&nbsp;&nbsp;① 더존 NEOplus I,II&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
																							   ② 키컴 SA-Win&nbsp;&nbsp;&nbsp;
																							   ③ 아이퀘스트 얼마에요&nbsp;&nbsp;&nbsp;
																							  
				                			   
				                			    
				                			    </TD>
											
												
												
											</TR>
											<TR CLASS="TD61">
												<TD CLASS="TD6"  colspan = 2>&nbsp;&nbsp;&nbsp;④ 서울 마이크로 시스템 Micro Win&nbsp;&nbsp;&nbsp;&nbsp;
																							   ⑤ 기타(<INPUT NAME="txtw6_Etc" ALT="기타" TYPE="Text" MAXLENGTH=50 SiZE=20 tag=24>)</TD>
				                			   
				                			    
				                			    </TD>
											
												
												
											</TR>
											
											
											
											<TR CLASS="TD61">
												<TD WIDTH=60% Height="20" ALIGN=Left >7.수입금액관리 프로그램</TD>	
												<TD WIDTH=40% Height="20" ALIGN=Left ></TD>
											
											</TR>
											<TR CLASS="TD61">
												<TD WIDTH=60% Height="20" ALIGN=Left >&nbsp;&nbsp;&nbsp;제조회사(<INPUT NAME="txtw7_1" ALT="제조회사" TYPE="Text"  MAXLENGTH=50 SiZE=20 tag=21>)</TD>	
												<TD WIDTH=60% Height="20" ALIGN=Left >&nbsp;&nbsp;&nbsp;S/W명칭(<INPUT NAME="txtw7_2" ALT="S/W 명칭" TYPE="Text"  MAXLENGTH=50 SiZE=20 tag=21>)</TD>	
										
											</TR>
											
											<TR CLASS="TD61">
												<TD WIDTH=40% Height="20" ALIGN=Left >8.전자상거래 유무</TD>	
												<TD WIDTH=40% Height="20" ALIGN=Left ><SELECT NAME="cboW8" ALT="8.전자상거래 유무" tag="22X1">
																							<option value = ""></option>
																							<option value = 1>①</option>
																							<option value = 2>②</option>
																										</SELECT>
												</TD>
											
											</TR>
											<TR CLASS="TD61">
												<TD CLASS="TD6"  colspan = 2>&nbsp;&nbsp;&nbsp;① 있음&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
																							   ② 없음 
																							   
																							  
				                			   
				                			    
				                			    </TD>
											
												
												
											</TR>
											
											<TR CLASS="TD61">
												<TD WIDTH=40% Height="20" ALIGN=Left >9.전자상거래 유형(복수선택가능 8번 ①응답자만)</TD>	
												<TD WIDTH=40% Height="20" ALIGN=Left >
												
												</TD>
											
											</TR>
											<TR CLASS="TD61">
												<TD WIDTH=100% Height="20" ALIGN=Left colspan =2 >&nbsp;&nbsp;&nbsp;<INPUT TYPE=CHECKBOX CLASS="Check" id="chkW9_1" name=chkW9_1 tag="24" VALUE="N" >재화의 공급&nbsp;&nbsp;&nbsp;
																									    <INPUT TYPE=CHECKBOX CLASS="Check" id="chkW9_2" name=chkW9_2 tag="24" VALUE="N">서비스의 공급&nbsp;&nbsp;&nbsp;
																									     <INPUT TYPE=CHECKBOX CLASS="Check" id="chkW9_3" name=chkW9_3 tag="24" VALUE="N">컨텐츠의 제공&nbsp;&nbsp;&nbsp;
												</TD>	
												
											
											</TR>
											<TR CLASS="TD61">
												<TD WIDTH=100% Height="20" ALIGN=Left colspan =2 >&nbsp;&nbsp;&nbsp;<INPUT TYPE=CHECKBOX CLASS="Check"  id="chkW9_4" name=chkW9_4 tag="24" VALUE="N">광고의 게재&nbsp;&nbsp;&nbsp;
																									    <INPUT TYPE=CHECKBOX CLASS="Check"  id="chkW9_5" name=chkW9_5 tag="24" VALUE="N">거래중개&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
																									     <INPUT TYPE=CHECKBOX CLASS="Check" id="chkW9_6 " name=chkW9_6 tag="24" VALUE="N" >기타(<INPUT NAME="txtW9_6_ETC" ALT="9.전자상거래 유형" TYPE="Text"  MAXLENGTH=50 SiZE=20 tag=14>)
																									     
												</TD>	
												
											
											</TR>
											<TR CLASS="TD61">
												<TD WIDTH=40% Height="20" ALIGN=Left >10.단위업무시스템 종류</TD>	
												<TD WIDTH=40% Height="20" ALIGN=Left >
												</TD>
											
											</TR>
											<TR CLASS="TD61">
												<TD WIDTH=100% Height="20" ALIGN=Left colspan =2 >&nbsp;&nbsp;&nbsp;<INPUT TYPE=CHECKBOX CLASS="Check" id="chkw10_1" name=chkw10_1 tag="25" VALUE="N">재무회계(원장, 재무제표, 채권/채무)&nbsp;&nbsp;&nbsp;
																									    <INPUT TYPE=CHECKBOX CLASS="Check" id="chk10_2" name=chkw10_2 tag="25" VALUE="N">관리회계(원가,성과)&nbsp;&nbsp;&nbsp;
																									     
												</TD>	
												
											
											</TR>
											<TR CLASS="TD61">
												<TD WIDTH=100% Height="20" ALIGN=Left colspan =2 >&nbsp;&nbsp;&nbsp;<INPUT TYPE=CHECKBOX CLASS="Check" id="chkw10_3" name=chkw10_3 tag="25" VALUE="N">재무관리(자금, 예산)&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
																									    <INPUT TYPE=CHECKBOX CLASS="Check" id="chkw10_4" name=chkw10_4 tag="25" VALUE="N">판매관리(고객, 주문, 출하, 배송, 청구)&nbsp;&nbsp;&nbsp;
																									    
																									     
												</TD>	
												
											
											</TR>
											<TR CLASS="TD61">
												<TD WIDTH=100% Height="20" ALIGN=Left colspan =2 >&nbsp;&nbsp;&nbsp;<INPUT TYPE=CHECKBOX CLASS="Check" id="chkW10_5" name=chkW10_5 tag="25" VALUE="N">자재관리(구매, 검수, 입출고, 재고)&nbsp;&nbsp;&nbsp;&nbsp;
																									    <INPUT TYPE=CHECKBOX CLASS="Check" id="chkw10_6" name=chkw10_6 tag="25" VALUE="N">생산관리(생산계획,생산실적)&nbsp;&nbsp;&nbsp;
																									    
																									     
												</TD>	
												
											
											</TR>
											
											<TR CLASS="TD61">
												<TD WIDTH=100% Height="20" ALIGN=Left colspan =2 >&nbsp;&nbsp;&nbsp;<INPUT TYPE=CHECKBOX CLASS="Check" id="chkw10_7" name=chkw10_7 tag="25" VALUE="N">품질관리(픔질계획,품질검사)&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
																									    <INPUT TYPE=CHECKBOX CLASS="Check" id="chkw10_8" name=chkw10_8 tag="25" VALUE="N">인사관리(급여,복지후생,채용)&nbsp;&nbsp;&nbsp;
																									     
												</TD>	
												
											
											</TR>
											<TR CLASS="TD61">
												<TD WIDTH=100% Height="20" ALIGN=Left colspan =2 >&nbsp;&nbsp;&nbsp;<INPUT TYPE=CHECKBOX CLASS="Check" id="chkw10_9" name=chkw10_9 tag="25" VALUE="N">기타(<INPUT NAME="txtw10_9_ETC" ALT="10.단위업무시스템 종류" TYPE="Text"  MAXLENGTH=50 SiZE=20 tag=24>)
																									     
												</TD>	
												
											
											</TR>
											<TR CLASS="TD51">
												<TD WIDTH=100% Height="30" ALIGN=CENTER  colspan = 2></TD>
											
											</TR>
											  
										</TABLE>		
									</FIELDSET>
				
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

