
<%@ LANGUAGE="VBScript" CODEPAGE=949%>
<% session.CodePage=949 %>
<%'======================================================================================================

'*  1. Module Name          : 법인세 
'*  2. Function Name        : 조특제11호의5고용증대특별세액공제 
'*  3. Program ID           : W6113MA1
'*  4. Program Name         : W6113MA1.asp
'*  5. Program Desc         : 조특제11호의5고용증대특별세액공제 
'*  6. Modified date(First) : 2005/03/18
'*  7. Modified date(Last)  : 2006/02/10
'*  8. Modifier (First)     : 홍지영 
'*  9. Modifier (Last)      : HJO 
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
Const BIZ_MNU_ID = "W6113MA1"											 '☆: 비지니스 로직 ASP명 
Const BIZ_PGM_ID = "W6113Mb1.asp"											 '☆: 비지니스 로직 ASP명 
Const EBR_RPT_ID = "W6113OA1"


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

   lgKeyStream = UCASE(Frm1.txtCo_Cd.Value) &  parent.gColSep  
   lgKeyStream = lgKeyStream & (Frm1.txtFISC_YEAR.text) &  parent.gColSep ' 
   lgKeyStream = lgKeyStream & UCASE(Frm1.cboREP_TYPE.Value ) &  parent.gColSep '   
   lgKeyStream = lgKeyStream & strW1_R &  parent.gColSep ' 
   lgKeyStream = lgKeyStream &  (Frm1.txtW1_Rate.Value)   &  parent.gColSep ' 
   lgKeyStream = lgKeyStream & strW5_R &  parent.gColSep '  
   lgKeyStream = lgKeyStream &  (Frm1.txtW5_Rate.Value)   &  parent.gColSep '   


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
    Call  AppendNumberPlace("7", "7", "2")
    Call  AppendNumberPlace("6", "5", "2")
    
	' 변경한곳 
    Call InitComboBox
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
     Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
    Call ggoOper.FormatDate(frm1.txtFISC_YEAR, parent.gDateFormat,3)
  
    Call SetDefaultVal()
	frm1.txtFISC_YEAR.text 	= "<%=wgFISC_YEAR%>"
	frm1.txtCO_CD.value 	= "<%=wgCO_CD%>"
	frm1.txtCO_NM.value 	= "<%=wgCO_NM%>"
	frm1.cboREP_TYPE.value 	= "<%=wgREP_TYPE%>"
     
    ' 세무조정 체크호출 
	Call FncQuery
  
End Sub




Sub SetDefaultVal()
dim strWhere 
DIM strW1


dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
        strWhere = "A.MAJOR_CD = 'W4012' and A.MINOR_CD = B.MINOR_CD and   A.MAJOR_CD = B.MAJOR_CD"
           
    	call CommonQueryRs(" MINOR_NM , MAX(CASE when B.SEQ_NO = 1 then B.REFERENCE end) , MAX(CASE when B.SEQ_NO = 2 then B.REFERENCE end) "," B_MINOR A , B_CONFIGURATION B  ",strWhere & " and A.MINOR_CD = '1' GROUP BY MINOR_NM" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    	
    	     frm1.txtW1_Rate.value 	=  Replace(lgF2,chr(11),"") 
    	     strW1_R	=    Replace(lgF1,chr(11),"") 	         
   	         strWhere = "A.MAJOR_CD = 'W4012' and A.MINOR_CD = B.MINOR_CD and   A.MAJOR_CD = B.MAJOR_CD"
           
    	call CommonQueryRs(" MINOR_NM , MAX(CASE when B.SEQ_NO = 1 then B.REFERENCE end) , MAX(CASE when B.SEQ_NO = 2 then B.REFERENCE end) "," B_MINOR A , B_CONFIGURATION B  ",strWhere & " and A.MINOR_CD = '2' GROUP BY MINOR_NM" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    	
    	     frm1.txtW5_Rate.value 	=  Replace(lgF2,chr(11),"") 

    	     strW5_R	=    Replace(lgF1,chr(11),"")


End Sub


'============================================  이벤트 함수  ====================================
Sub txtFISC_YEAR_DblClick(Button)
    If Button = 1 Then
        frm1.txtFISC_YEAR.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtFISC_YEAR.Focus
    End If
End Sub


Function CalSum()
    frm1.txtw2.value =  (UNICDbl(frm1.txtw3.text ) - UNICDbl(frm1.txtw4.text))
     if  UNICDbl(frm1.txtw2.text) < 0 then
         frm1.txtw2.text = 0
     end if
    frm1.txtw1.value =  (UNICDbl(frm1.txtw2.text) *  UNICDbl(strW1_R))
    
     
     if UNICDbl(frm1.txtw7.text ) = 0 then
        frm1.txtw6.text = 0
     else   
		If (UNICDbl(frm1.txtw7.text ) > UNICDbl(frm1.txtw8.text) ) Then
		   frm1.txtw6.text =fix(( (UNICDbl(frm1.txtw7.text ) -  UNICDbl(frm1.txtw8.text)) / UNICDbl(frm1.txtw7.text )) *  UNICDbl(frm1.txtw9.text ))
		 Else
		        frm1.txtw6.text = 0
			 
		End IF
     end if 
     
     frm1.txtw5.text = UNICDbl(frm1.txtw6.text ) *  UNICDbl(strW5_R)
     
     frm1.txtw10.text = UNICDbl(frm1.txtw1.text ) + UNICDbl(frm1.txtw5.text ) 

end function

Function CheckMessage(ByVal Obj)
dim IntRetCD
    if  UNICDbl(Obj.text) < 0 then
        Obj.value = 0
        Obj.focus

    end if
    
end function






Sub txtw3_Change()  
     
    lgBlnFlgChgValue  = True 
    Call CheckMessage(frm1.txtw3)
    Call CalSum() 
    
End Sub


Sub txtw4_Change()  
     
    lgBlnFlgChgValue  = True 
    Call CheckMessage(frm1.txtw4)
    Call CalSum() 
    
End Sub

Sub txtw6_Change()  
     
    lgBlnFlgChgValue  = True 
    Call CheckMessage(frm1.txtw6)

    
End Sub

Sub txtw7_Change()  
     
    lgBlnFlgChgValue  = True 
    Call CheckMessage(frm1.txtw7)
    
      frm1.txtW7.Focus 
    
    if UNIcdbl(frm1.txtw7.value) > 24 then
       frm1.txtw7.text = 24
     
    end if   

    Call CalSum
    frm1.txtw7.focus
    
End Sub

Sub txtw8_Change()  
     
    lgBlnFlgChgValue  = True 
    Call CheckMessage(frm1.txtw8)
    if UNIcdbl(frm1.txtw8.value) > 24 then
       frm1.txtw8.text = 24
     
    end if   
     Call CalSum
     frm1.txtw8.focus
     
    
End Sub

Sub txtw9_Change()  
     
    lgBlnFlgChgValue  = True 
    Call CheckMessage(frm1.txtw9)
    Call CalSum() 
    frm1.txtw9.focus
    
    
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



'============================== 레퍼런스 함수  ========================================
Function GetRef()	' 금액가져오기 링크 클릭시 
	Dim sFiscYear, sRepType, sCoCd, IntRetCD
	sCoCd		= frm1.txtCO_CD.value
	sFiscYear	= frm1.txtFISC_YEAR.text
	sRepType	= frm1.cboREP_TYPE.value
	
	Dim sMesg

	' 세무정보 조사 : 메시지가져오기.
	
	
	if wgConfirmFlg = "Y" then    '확정시 
	   Exit function
	end if   
	
	' 온로드시 레퍼런스메시지 가져온다.
    wgRefDoc = GetDocRef(frm1.txtCO_CD.value, frm1.txtFISC_YEAR.text, frm1.cboREP_TYPE.value, BIZ_MNU_ID)
	
	sMesg = wgRefDoc & vbCrLf & vbCrLf
     call SelectColor(frm1.txtw3) 
     call SelectColor(frm1.txtw4) 
    IntRetCD = DisplayMsgBox("WC0003", parent.VB_YES_NO, sMesg, "X")           '⊙: "Will you destory previous data"
    Call ggoOper.LockField(Document, "N")
	 If IntRetCD = vbNo Then
	 	Exit Function
	 End If
    '접대비 프로그램 
    Dim arrW1 ,arrW2 ,  arrW3, arrW4, arrW5, arrW6, iRow, iCol
	call CommonQueryRs("W3,W4","dbo.ufn_TB_11_5_GetRef('" & sCoCd & "', '" & sFiscYear & "', '" & sRepType & "')", "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

	If lgF0 = "" Then	 
	 
	   Exit Function
	else   
		frm1.txtw3.value = unicdbl(lgF0)
	    frm1.txtw4.value = unicdbl(lgF1)
    end if
	 
 
     

end function
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
	    Call SetToolbar("1100100000000111")										<%'버튼 툴바 제어 %>
	    
		
	Else
	   '2 디비환경값 , 로드시환경값 비교 
		 Call SetToolbar("1101100000000111")										<%'버튼 툴바 제어 %>
		 if uniCDBL(strW1_R) <> uniCDBL(frm1.txtW1_RATE_NEW.value)  then
		    IntRetCD = DisplayMsgBox("WC0027", parent.VB_YES_NO, frm1.txtW1_RATE_NEW.alt,"X") 
				If IntRetCD = vbNo Then	
				     Exit Function
				else     
				    strW1_R = uniCDBL(frm1.txtW1_RATE_NEW.value)
				    Call calSum()
				End If
		 end if 
		  if uniCDBL(strW5_R) <> uniCDBL(frm1.txtW5_RATE_NEW.value)  then
		    IntRetCD = DisplayMsgBox("WC0027", parent.VB_YES_NO, frm1.txtW5_RATE_NEW.alt,"X") 
				If IntRetCD = vbNo Then	
				     Exit Function
				else     
				    strW5_R = uniCDBL(frm1.txtW5_RATE_NEW.value)
				    Call calSum()
				End If
		 end if 
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
					<TD WIDTH=* ALIGN=RIGHT>&nbsp;
						<a href="vbscript:GetRef">금액 불러오기</A>  
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
					<TD WIDTH=520 valign=top  >
					   
					    
									<TABLE bgcolor = #696969  border = 0 cellpadding = 1 cellspacing = 1 ID="Table1">
									   
									
										<TR>
											 <TD CLASS="TD51" align =LEFT >
												(1)고용증대세액공제액 : (2)×<INPUT TYPE=TEXT NAME="txtW1_Rate"  style="WIDTH:100 ; border-style:none;"  tag="34" >
											</TD>
											
										    <TD CLASS="TD61" align =LEFT width = 15% >
												<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW1" name=txtW1 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="24X2" SIZE = 10 ></OBJECT>');</SCRIPT>
											</TD>
											
										</TR>
										
										<TR>
											 <TD CLASS="TD51" align =LEFT>
												(2)고용증대인원수 : (3)-(4)
											</TD>
											
										    <TD CLASS="TD61" align =LEFT  >
												<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW2" name=txtW2 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="24X76" SIZE = 10 ></OBJECT>');</SCRIPT>
											</TD>
											
										</TR>
										<TR>
											 <TD CLASS="TD51" align =LEFT>
												(3)당해 과세연도 상시근로자수 
											</TD>
											
										    <TD CLASS="TD61" align =LEFT width = 15% >
												<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW3" name=txtW3 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="25X76" SIZE = 10 ></OBJECT>');</SCRIPT>
											</TD>
											
										</TR>
										<TR>
											 <TD CLASS="TD51" align =LEFT >
												(4)직전 과세연도 상시근로자수 
											</TD>
											
										    <TD CLASS="TD61" align =LEFT width = 15% >
												<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW4" name=txtW4 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="25X76" SIZE = 10 ></OBJECT>');</SCRIPT>
											</TD>
											
										</TR>
										<TR>
											 <TD CLASS="TD51" align =LEFT  >
												(5)고용유지세액공제액 : (6)×<INPUT TYPE=TEXT NAME="txtW5_Rate"  style="WIDTH:100 ; border-style:none;"  tag="34" >
											</TD>
											
										    <TD CLASS="TD61" align =LEFT width = 15% >
												<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW5" name=txtW5 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="24X26" SIZE = 10 ></OBJECT>');</SCRIPT>
											</TD>
											
										</TR>
										<TR>
											 <TD CLASS="TD51" align =LEFT  >
												(6)고용유지인원수 : [((7)-(8))÷(7)]×(9)<br>(소수점 미만 절사)
											</TD>
											
										    <TD CLASS="TD61" align =center width = 15% >
												<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW6" name=txtW6 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="24X2Z" SIZE = 10 ></OBJECT>');</SCRIPT>
											</TD>
											
										</TR>
										<TR>
											 <TD CLASS="TD51" align =LEFT >
												(7)고용유지제도 시행일 전 1월간 상시근로자 1인당 <BR>1일 평균근로시간 
											</TD>
											
										    <TD CLASS="TD61" align =LEFT width = 15% >
												<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW7" name=txtW7 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="21X76" SIZE = 10 ></OBJECT>');</SCRIPT>
											</TD>
											
										</TR>
										<TR>
											 <TD CLASS="TD51" align =LEFT  >
												(8)고용유지제도 시행일 이후 1월간 상시근로자 1인당 <BR>1일 평균근로시간 
											</TD>
											
										    <TD CLASS="TD61" align =LEFT >
												<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW8" name=txtW8 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="25X76" SIZE = 10 ></OBJECT>');</SCRIPT>
											</TD>
											
										</TR>
										<TR>
											 <TD CLASS="TD51" align =LEFT  >
												(9)직전 과세연도 상시근로자수 
											</TD>
											
										    <TD CLASS="TD61" align =LEFT width = 15% >
												<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW9" name=txtW9 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="25X6z" SIZE = 10 ></OBJECT>');</SCRIPT>
											</TD>
											
										</TR>
										<TR>
											 <TD CLASS="TD51" align =LEFT  >
												(10)세액공제액 계 : ((1)+(5)))
											</TD>
											
										    <TD CLASS="TD61" align =LEFT width = 15% >
												<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW10" name=txtW10 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="24X26" SIZE = 10 ></OBJECT>');</SCRIPT>
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

<INPUT TYPE=HIDDEN NAME="txtKeyStream" tag="24" >
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24" >
<INPUT TYPE=HIDDEN NAME="txtW1_RATE_NEW" tag="24" ALT="고용증대 인원에 대한 세액공제액">
<INPUT TYPE=HIDDEN NAME="txtW5_RATE_NEW" tag="24" ALT="고용유지 인원에 대한 세액공제액">


</FORM>
<DIV ID="MousePT" NAME="MousePT">

<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

