
<%@ LANGUAGE="VBScript" CODEPAGE=949%>
<% session.CodePage=949 %>
<%'======================================================================================================

'*  1. Module Name          : 법인세 
'*  2. Function Name        : 접대비조정명세서(갑)
'*  3. Program ID           : W3111MA1
'*  4. Program Name         : W3111MA1.asp
'*  5. Program Desc         : 접대비조정명세서(갑)
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
<SCRIPT LANGUAGE="VBScript"	SRC="../WCM/incGrid.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"	SRC="../WCM/inc_CliOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../WCM/incEB.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit

'============================================  상수/변수 선언  ====================================

<!-- #Include file="../../inc/lgvariables.inc" -->
Const BIZ_MNU_ID     =  "W3111MA1"	 
Const BIZ_PGM_ID = "W3111MB1.asp"											 '☆: 비지니스 로직 ASP명 
Const EBR_RPT_ID  = "W3111OA1"





Const C_SHEETMAXROWS = 100



Dim IsOpenPop          
Dim lgStrPrevKey2
Dim lgRefMode	' 레퍼런스 참조상태 : ERP 추출사용유무 




dim strW1_A
dim strW1_B
dim strW1_C
dim strW1_D
dim strW1_E
dim strW1_F
dim strW1_G
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

  
	Call AppendNumberPlace("8","15","0")	' -- 금액 15자리 고정 : 출하검사패치 
			
		
    
End Sub


'============================================  그리드 함수  ====================================

Sub InitSpreadComboBox()
    Dim iCodeArr , IntRetCD1
    Dim iNameArr
    Dim iDx



End Sub


Sub SetSpreadLock()



 
   
End Sub

Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)

End Sub






Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos

    Select Case UCase(pvSpdNo)
       Case "A"
         
          
       
	

    End Select    
End Sub
'============================== 레퍼런스 함수  ========================================
Function GetRef()	' 금액가져오기 링크 클릭시 
	Dim sFiscYear, sRepType, sCoCd, IntRetCD
	sCoCd		= frm1.txtCO_CD.value
	sFiscYear	= frm1.txtFISC_YEAR.text
	sRepType	= frm1.cboREP_TYPE.value
	
	Dim sMesg


	 wgRefDoc = GetDocRef(frm1.txtCO_CD.value, frm1.txtFISC_YEAR.text, frm1.cboREP_TYPE.value, BIZ_MNU_ID)
	
	sMesg = wgRefDoc & vbCrLf & vbCrLf
    call SelectColor(frm1.txtw6) 
    call SelectColor(frm1.txtw7) 
    IntRetCD = DisplayMsgBox("WC0003", parent.VB_YES_NO, sMesg, "X")           '⊙: "Will you destory previous data"\
    Call ggoOper.LockField(Document, "N")
	 If IntRetCD = vbNo Then
	 	Exit Function
	 End If
	 
    '접대비 프로그램 
    
   Call SetMajor()
    
    Dim arrW1 ,arrW2 ,  arrW3, arrW4, arrW5, arrW6, iRow, iCol
	
	call CommonQueryRs("WA,WB,W6,W7","dbo.ufn_TB_23A_GetRef_"&C_REVISION_YM&"('" & sCoCd & "', '" & sFiscYear & "', '" & sRepType & "')", "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

	If lgF0 = "" Then	 Exit Function

	    arrW1 = REPLACE(lgF0, chr(11),"")
	    
	    arrW2 = REPLACE(lgF1, chr(11),"")
		arrW3 = REPLACE(lgF2, chr(11),"")
		arrW4 = REPLACE(lgF3, chr(11),"")	


     
           frm1.txt23w6.value =  unicdbl(arrW1)
           Call txt23w6_Onchange
           frm1.txt23w2.value =  unicdbl(arrW2)
           Call txt23w2_Onchange
           frm1.txtw6.value =   unicdbl(arrW3)
           frm1.txtw7.value =   unicdbl(arrW4)
 
     

end function
'============================================  조회조건 함수  ====================================

'============================================  폼 함수  ====================================
Sub Form_Load()

    Call LoadInfTB19029    
                                             <%'Load table , B_numeric_format%>
    Call ggoOper.LockField(Document, "N")                                   <%'Lock  Suitable  Field%>                         
   
    Call InitSpreadSheet()                                                    <%'Setup the Spread sheet%>
    Call SetToolbar("1100100000000111")										<%'버튼 툴바 제어 %>

	' 변경한곳 
    Call InitComboBox
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.FormatField(Document, "2", ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
    Call ggoOper.FormatDate(frm1.txtFISC_YEAR, parent.gDateFormat,3)
    Call SetDefaultVal
    Call InitVariables
    Call FncQuery  
End Sub



Sub SetDefaultVal()

	    frm1.txtFISC_YEAR.text 	= "<%=wgFISC_YEAR%>"
		frm1.txtCO_CD.value 	= "<%=wgCO_CD%>"
		frm1.txtCO_NM.value 	= "<%=wgCO_NM%>"
		frm1.cboREP_TYPE.value 	= "<%=wgREP_TYPE%>"
   	 

		Call SetMajor()
    	
End Sub



Function SetMajor()


dim strWhere , strFrom , strSelect
DIM strW1


dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
 '중소기업 / 중소기업외 접대비 기본한도액 
        strSelect =  "B.REFERENCE_1 , Case when A.REP_TYPE= 1 then DATEDIFF(month, FISC_START_DT, FISC_END_DT) +1  else 6 end "
        strFrom  =   "TB_COMPANY_HISTORY A,  dbo.ufn_TB_Configuration('W2016', '" & C_REVISION_YM & "') B"
		strWhere =   "B.MINOR_CD =  A.COMP_TYPE1 AND CO_CD = "&  FilterVar(Trim(UCase(FRM1.txtCo_Cd.VALUE)),"","S") &"  AND "
		strWhere =  strWhere &       " FISC_YEAR =  "& FilterVar(Trim(UCase(FRM1.txtFISC_YEAR.TEXT)),"","S") &" AND "
		strWhere =  strWhere &       " REP_TYPE =   "& FilterVar(Trim(UCase(FRM1.cboREP_TYPE.VALUE)),"","S") &""

      
       call CommonQueryRs(strSelect,strFrom,strWhere  ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    	
    	
    	     strW1 =  Replace(lgF0,chr(11),"") 
    	
    	     frm1.txtW1_nm.value 	=    " (1) " & strW1 & " x "& Replace(lgF1,chr(11),"") & "/12"
             frm1.txtW1.value       = unicdbl(strW1) *( unicdbl(Replace(lgF1,chr(11),""))/12)
            
             strSelect  = " MINOR_NM , REFERENCE_1,  REFERENCE_2 "
             strFrom    = " dbo.ufn_TB_Configuration('w2005', '" & C_REVISION_YM & "')  "
		     strWhere   = " MINOR_CD = '1'  "

    	
    	  '수입금액 100억 이하 x  20/10000  lgF0 표시값, lgF1 Value 값 
    	
    	
    	call CommonQueryRs(strSelect, strFrom  ,strWhere  ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    	
    	     frm1.txtW1_Anm.value 	=  Replace(lgF0,chr(11),"")  &" x  " &  Replace(lgF2,chr(11),"") 
             frm1.txtW1_Dnm.value 	=  Replace(lgF0,chr(11),"")  &" x  " &  Replace(lgF2,chr(11),"") 
    	     strW1_A	=    Replace(lgF1,chr(11),"")
       	     strW1_D	=    Replace(lgF1,chr(11),"") 
   	     
    	     '수입금액 100억 초과 500억 이하    lgF0 표시값, lgF1 Value 값 
    	     strSelect  = " MINOR_NM , REFERENCE_1,  REFERENCE_2 "
             strFrom    = " dbo.ufn_TB_Configuration('w2005', '" & C_REVISION_YM & "')  "
		     strWhere   = " MINOR_CD = '2' "
		     
		    
		call CommonQueryRs(strSelect, strFrom  ,strWhere  ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)     
    	     frm1.txtW1_Bnm.value 	=  Replace(lgF0,chr(11),"") &" x  " &  Replace(lgF2,chr(11),"") 
             frm1.txtW1_Enm.value 	=  Replace(lgF0,chr(11),"")  &" x  " &  Replace(lgF2,chr(11),"") 
    	     strW1_B	=    Replace(lgF1,chr(11),"")
    	     strW1_E	=    Replace(lgF1,chr(11),"") 
            
            
            '수입금액 500억 초과    lgF0 표시값, lgF1 Value 값 
             strSelect  = " MINOR_NM , REFERENCE_1,  REFERENCE_2 "
             strFrom    = " dbo.ufn_TB_Configuration('w2005', '" & C_REVISION_YM & "')  "
		     strWhere   = " MINOR_CD = '3' "
       call CommonQueryRs(strSelect, strFrom  ,strWhere  ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    	     frm1.txtW1_Cnm.value 	=  Replace(lgF0,chr(11),"") &" x  " &  Replace(lgF2,chr(11),"")
             frm1.txtW1_Fnm.value 	=  Replace(lgF0,chr(11),"")  &" x  " &  Replace(lgF2,chr(11),"") 
    	     strW1_C	=    Replace(lgF1,chr(11),"")
    	     strW1_F	=    Replace(lgF1,chr(11),"") 
    	     
    	     '특수관계자 매출에 대한 조정율   lgF0 표시값, lgF1 Value 값 
    	     
    	     strSelect  = " MINOR_NM , REFERENCE_1,  REFERENCE_2 "
             strFrom    = " dbo.ufn_TB_Configuration('w2005', '" & C_REVISION_YM & "')  "
		     strWhere   = " MINOR_CD = '4' "
        call CommonQueryRs(strSelect, strFrom  ,strWhere  ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    	     strW1_G	=    Replace(lgF1,chr(11),"")
   	     
End Function


'============================================  이벤트 함수  ====================================


Function OpenRefMenu()

    Dim arrRet
    Dim arrParam(2)
    Dim arrField, arrHeader
    Dim iCalledAspName, IntRetCD
    
    If IsOpenPop = True Then Exit Function

    IsOpenPop = True
   
    arrRet = window.showModalDialog("../W5/W5105RA1.asp", Array(Window.parent,arrParam, arrField, arrHeader), _
             "dialogWidth=1080px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
    
    IsOpenPop = False
    
    
End Function



Sub txtFISC_YEAR_DblClick(Button)
    If Button = 1 Then
        frm1.txtFISC_YEAR.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtFISC_YEAR.Focus
    End If
End Sub


Function Find_Max(Byval a, byval b)
    if a > b then
       Find_Max = unicdbl(a)
    else
       Find_Max = unicdbl(b)
    end if 
    
    
End Function

Sub txt23w6_OnChange()    ' 23호 서식 6호 금액 
   dim sum23w6 , dblLimitAmt1 ,dblLimitAmt2


    lgBlnFlgChgValue  = True 
    sum23w6 = unicdbl(Frm1.txt23w6.value) 
    Frm1.txtw1_A.text = 0
    Frm1.txtw1_B.text = 0
    Frm1.txtw1_C.text = 0
    dblLimitAmt1 = 10000000000
    dblLimitAmt2 = 50000000000
 
      if unicdbl(sum23w6) >    unicdbl(dblLimitAmt1)  then
         Frm1.txtw1_A.text =   unicdbl(dblLimitAmt1) * unicdbl(strW1_A)
 
				if   unicdbl(sum23w6) > dblLimitAmt2  then    '500억보다 크믄 
				     Frm1.txtw1_B.text =   unicdbl(dblLimitAmt2 -dblLimitAmt1) * unicdbl(strW1_B)     ' (500억 - 100억) * 비율 
			
				     Frm1.txtw1_C.text =   (unicdbl(sum23w6) - dblLimitAmt2)   * unicdbl(strW1_C)          
				else
				     Frm1.txtw1_B.text =   unicdbl(unicdbl(sum23w6) -dblLimitAmt1) * unicdbl(strW1_B)    
				end if
      else

               Frm1.txtw1_A.text =  unicdbl(sum23w6) * unicdbl(strW1_A)

      end if
       Call Fnc_SumCal()
End Sub






Sub txt23w2_OnChange()    ' 23호 서식 6호 금액 
   dim sum23w2 , dblLimitAmt1 ,dblLimitAmt2

    lgBlnFlgChgValue  = True 
    sum23w2 = unicdbl(Frm1.txt23w2.value) 
    Frm1.txtw1_D.text = 0
    Frm1.txtw1_E.text = 0
    Frm1.txtw1_F.text = 0
    dblLimitAmt1 = 10000000000
    dblLimitAmt2 = 50000000000
    
    
      if unicdbl(sum23w2) >  unicdbl(dblLimitAmt1) then
                Frm1.txtw1_D.text =    unicdbl(dblLimitAmt1) * unicdbl(strW1_D)

				if   unicdbl(sum23w2) > dblLimitAmt2  then
				     Frm1.txtw1_E.text =    unicdbl(dblLimitAmt2 -dblLimitAmt1)  * unicdbl(strW1_E)
				     Frm1.txtw1_F.text =   (unicdbl(sum23w2) - dblLimitAmt2)  * unicdbl(strW1_F)          
				else
				     Frm1.txtw1_E.text =   unicdbl(unicdbl(sum23w2) -dblLimitAmt1) * unicdbl(strW1_E)    
				end if
				
				
         
      else

         Frm1.txtw1_D.text =  unicdbl(sum23w2) * unicdbl(strW1_D)
      end if
      
      
      Call Fnc_SumCal()

End Sub



Sub txtw1_Change()   
    lgBlnFlgChgValue  = True 
    Call Fnc_SumCal()

End Sub

Sub txtw6_Change()   
    lgBlnFlgChgValue  = True 
    Call Fnc_SumCal()

End Sub
Sub txtw7_Change()   
    lgBlnFlgChgValue  = True 
    Call Fnc_SumCal()

End Sub


function Fnc_SumCal()


             Frm1.txtw2.text  = unicdbl(Frm1.txtw1_A.text) + unicdbl(Frm1.txtw1_B.text)+ unicdbl(Frm1.txtw1_C.text)
	         Frm1.txtw3.text  = unicdbl(Frm1.txtw1_D.text) + unicdbl(Frm1.txtw1_E.text)+ unicdbl(Frm1.txtw1_F.text)
             Frm1.txtw4.text   = (unicdbl(Frm1.txtw2.text) - unicdbl(Frm1.txtw3.text))  * unicdbl(strW1_G)
             Frm1.txtw5.text   = unicdbl(Frm1.txtw1.text) + unicdbl(Frm1.txtw3.text) +unicdbl(Frm1.txtw4.text)
             Frm1.txtw8.text   = unicdbl(Frm1.txtw6.text) - unicdbl(Frm1.txtw7.text) 
             Frm1.txtw9.text   = Find_max(unicdbl(Frm1.txtw8.text) - unicdbl(Frm1.txtw5.text),0) 
              
             if unicdbl(Frm1.txtw8.text)  > unicdbl(Frm1.txtw5.text) then
                Frm1.txtw10.text  = unicdbl(Frm1.txtw5.text) 
             else   
                Frm1.txtw10.text  = unicdbl(Frm1.txtw8.text) 
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
     Call SetMajor()
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

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	 Call ggoOper.LockField(Document, "N")
    '-----------------------
    'Reset variables area
    '-----------------------
    lgIntFlgMode = parent.OPMD_UMODE
    Call SetSpreadColor(-1,-1)  

     Call CheckTaxDoc(frm1.txtCO_CD.value, frm1.txtFISC_YEAR.text, frm1.cboREP_TYPE.value, BIZ_MNU_ID)
    Call InitData
	'1 컨펌체크 
	If wgConfirmFlg = "Y" Then

	    Call SetToolbar("1100000000000111")	
		
	Else
	   '2 디비환경값 , 로드시환경값 비교 
	       Call SetToolbar("1101100000000111")										<%'버튼 툴바 제어 %>
	End if       

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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB_BAK"><font color=white><%=Request("strASPMnuMnuNm")%></font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					
					<TD WIDTH=* ALIGN=RIGHT>&nbsp;
								<a href="vbscript:GetRef()">금액불러오기</A>|<A href="vbscript:OpenRefMenu">소득금액합계표조회</A></TD>
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
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%> </TD>
					
				</TR>
				
				
				
				<TR>
					<TD WIDTH=100% HEIGHT=* valign=top>
					   <DIV ID="TabDiv" STYLE="FLOAT: left; HEIGHT: 100%; WIDTH: 100%; overflow=auto">
					   
					      <FIELDSET CLASS="CLSFLD"><LEGEND ALIGN=LEFT> </LEGEND>
									<TABLE width = 100% bgcolor = #696969  border = 0 cellpadding = 1 cellspacing = 1 ID="Table1">
									   
										
										
										<TR>
											<TD CLASS="TD51" align =center width = 10% COLSPAN=3>
												구                                     분 
											</TD>
											
										    <TD CLASS="TD51" align =center width = 15% >
											금       액 
											</TD>
											
											
											
										</TR>
										<TR>
											
											<TD CLASS="TD51" align =left COLSPAN=3>
											  <INPUT TYPE=TEXT NAME="txtW1_nm"  style="WIDTH: 100% ; border-style:none;"  tag="34" >
											</TD>
											
											<TD CLASS="TD61" align =center width = 15% >
												<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW1" name=txtW1 CLASS=FPDS40 title=FPDOUBLESINGLE ALT="" tag="21X8Z" width = 100% ></OBJECT>');</SCRIPT>
											</TD>
											
										</TR>
										
										<TR>
											 <TD CLASS="TD51" align =center width = 5% ROWSPAN=9 >
											      수  입<BR>금  액<BR>기  준 
											</TD>
											 <TD CLASS="TD51" align =center width = 15% ROWSPAN=4 >
											    총수입금액<BR>기준 
											</TD>
											 <TD CLASS="TD51" align =left width = 20%  >
								                <INPUT TYPE=TEXT NAME="txtW1_Anm"  style="WIDTH: 100% ; border-style:none;"   tag="34" >
											    
											</TD>
											
											<TD CLASS="TD61" align =center width = 15% >
												<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW1_A" name=txtW1_A CLASS=FPDS40 title=FPDOUBLESINGLE ALT="" tag="24X80" width = 100% ></OBJECT>');</SCRIPT>
											</TD>
											
										</TR>
										
										<TR>
											
											<TD CLASS="TD51" align =left width = 15%  >
											     <INPUT TYPE=TEXT NAME="txtW1_Bnm"  style="WIDTH: 100% ; border-style:none;"  tag="34" >
											</TD>
											<TD CLASS="TD61" align =center width = 15% >
												<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW1_B" name=txtW1_B CLASS=FPDS40 title=FPDOUBLESINGLE ALT="" tag="24X80" width = 100% ></OBJECT>');</SCRIPT>
											</TD>
										</TR>
										<TR>
											
											<TD CLASS="TD51" align =left width = 15%  >
											     <INPUT TYPE=TEXT NAME="txtW1_Cnm"  style="WIDTH: 100% ; border-style:none;"  tag="34" >
											</TD>
											<TD CLASS="TD61" align =center width = 15% >
												<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW1_C" name=txtW1_C CLASS=FPDS40 title=FPDOUBLESINGLE ALT="" tag="24X80" width = 100% ></OBJECT>');</SCRIPT>
											</TD>
										</TR>
										
										<TR>
											
											<TD CLASS="TD51" align =center width = 15%  >
											    (2)    소      계 
											</TD>
											<TD CLASS="TD61" align =center width = 15% >
												<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW2" name=txtW2 CLASS=FPDS40 title=FPDOUBLESINGLE ALT="" tag="24X80" width = 100% ></OBJECT>');</SCRIPT>
											</TD>
										</TR>
										
										<TR>
											
											 <TD CLASS="TD51" align =center width = 15% ROWSPAN=4 >
											   일반수입금액<BR>기준 
											</TD>
											 <TD CLASS="TD51" align =left width = 15%  >
											     <INPUT TYPE=TEXT NAME="txtW1_Dnm"  style="WIDTH: 100% ; border-style:none;"  tag="34" >
											</TD>
											
											<TD CLASS="TD61" align =center width = 15% >
												<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW1_D" name=txtW1_D CLASS=FPDS40 title=FPDOUBLESINGLE ALT="" tag="24X80" width = 100% ></OBJECT>');</SCRIPT>
											</TD>
											
										</TR>
										
										<TR>
											
											<TD CLASS="TD51" align =left width = 15%  >
											   <INPUT TYPE=TEXT NAME="txtW1_Enm"  style="WIDTH: 100% ; border-style:none;"  tag="34" >
											</TD>
											<TD CLASS="TD61" align =center width = 15% >
												<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW1_E" name=txtW1_E CLASS=FPDS40 title=FPDOUBLESINGLE ALT="" tag="24X80" width = 100% ></OBJECT>');</SCRIPT>
											</TD>
										</TR>
										<TR>
											
											<TD CLASS="TD51" align =left width = 15%  >
											   <INPUT TYPE=TEXT NAME="txtW1_Fnm"  style="WIDTH: 100% ; border-style:none;"  tag="34" >
											</TD>
											<TD CLASS="TD61" align =center width = 15% >
												<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW1_F" name=txtW1_F CLASS=FPDS40 title=FPDOUBLESINGLE ALT="" tag="24X80" width = 100% ></OBJECT>');</SCRIPT>
											</TD>
										</TR>
										
										<TR>
											
											<TD CLASS="TD51" align =center width = 15%  >
											   (3)    소      계 
											</TD>
											<TD CLASS="TD61" align =center width = 15% >
												<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW3" name=txtW3 CLASS=FPDS40 title=FPDOUBLESINGLE ALT="" tag="24X80" width = 100% ></OBJECT>');</SCRIPT>
											</TD>
										</TR>
										
										<TR>
											<TD CLASS="TD51" align =center width = 15% >
											  (4)기타수입금액 
											</TD>
											
										    <TD CLASS="TD51" align =left width = 15%  >
											   ((2)－(3))×20/100
											</TD>
										    
											<TD CLASS="TD61" align =center width = 15% >
												<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW4" name=txtW4 CLASS=FPDS40 title=FPDOUBLESINGLE ALT="" tag="24X80" width = 100% ></OBJECT>');</SCRIPT>
											</TD>
											
											
										</TR>
										
										<TR>
											<TD CLASS="TD51" align =left width = 15% COLSPAN=3 >
											  (5) 접대비한도액 ((1)＋(3)＋(4))
											</TD>
											
										  
											<TD CLASS="TD61" align =center width = 15% >
												<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW5" name=txtW5 CLASS=FPDS40 title=FPDOUBLESINGLE ALT="" tag="24X80" width = 100% ></OBJECT>');</SCRIPT>
											</TD>
											
											
										</TR>
										<TR>
											<TD CLASS="TD51" align =left width = 15% COLSPAN=3 >
											  (6) 접대비 해당금액 
											</TD>
											
										  
											<TD CLASS="TD61" align =center width = 15% >
												<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW6" name=txtW6 CLASS=FPDS40 title=FPDOUBLESINGLE ALT="" tag="24X80" width = 100% ></OBJECT>');</SCRIPT>
											</TD>
											
											
										</TR>
										<TR>
											<TD CLASS="TD51" align =left width = 15% COLSPAN=3 >
											  (7) 5만원  초과 접대비중 신용카드 등 미사용으로 인한 손금불산입액 
											</TD>
											
										  
											<TD CLASS="TD61" align =center width = 15% >
												<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW7" name=txtW7 CLASS=FPDS40 title=FPDOUBLESINGLE ALT="" tag="24X80" width = 100% ></OBJECT>');</SCRIPT>
											</TD>
											
											
										</TR>
										<TR>
											<TD CLASS="TD51" align =left width = 15% COLSPAN=3 >
											  (8) 차감 접대비 해당금액 ((6)－(7))
											</TD>
											
										  
											<TD CLASS="TD61" align =center width = 15% >
												<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW8" name=txtW8 CLASS=FPDS40 title=FPDOUBLESINGLE ALT="" tag="24X80" width = 100% ></OBJECT>');</SCRIPT>
											</TD>
											
											
										</TR>
										
						                <TR>
											<TD CLASS="TD51" align =left width = 15% COLSPAN=3 >
											  (9) 한도초과액 ((8)－(5))
											</TD>
											
										  
											<TD CLASS="TD61" align =center width = 15% >
												<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW18" name=txtW9 CLASS=FPDS40 title=FPDOUBLESINGLE ALT="" tag="24X80" width = 100% ></OBJECT>');</SCRIPT>
											</TD>
											
											
										</TR>
										  <TR>
											<TD CLASS="TD51" align =left width = 15% COLSPAN=3 >
											   (10) 손금산입한도내 접대비지출액 ((5)와 (8)중 적은 금액)
											</TD>
											
										  
											<TD CLASS="TD61" align =center width = 15% >
												<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW10" name=txtW10 CLASS=FPDS40 title=FPDOUBLESINGLE ALT="" tag="24X80" width = 100% ></OBJECT>');</SCRIPT>
											</TD>
											
											
										</TR>
										
										
						
											   
											
											
										</TR>
						
									</TABLE>
						   </FIELDSET>				
						   			
					</TD>
				</TR>
		 
				
			</TABLE>
			</DIV>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_01%>></TD>
		<INPUT TYPE=HIDDEN NAME="txt23w2" tag="21">
        <INPUT TYPE=HIDDEN NAME="txt23w6" tag="21">
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
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 tabindex="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>

<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" tabindex="-1">


<INPUT TYPE=HIDDEN NAME="txt23w10" tag="24" tabindex="-1">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24" tabindex="-1">
<INPUT TYPE=HIDDEN NAME="txtKeyStream" tag="24" tabindex="-1">

</FORM>
<DIV ID="MousePT" NAME="MousePT">

<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm" tabindex="-1"></iframe>
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

