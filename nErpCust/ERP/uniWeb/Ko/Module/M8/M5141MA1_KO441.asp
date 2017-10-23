<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% session.CodePage=949 %>

<%
Response.Expires = -1
%>
<!--
'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : M5141MA1
'*  4. Program Name         : 단가정산 
'*  5. Program Desc         :
'*  6. Modified date(First) : 2005/03/10
'*  7. Modified date(Last)  : 
'*  8. Modifier (First)     : Lee Sung Yong
'*  9. Modifier (Last)      : 
'* 10. Comment              :
'* 11. History              :
'*                            
'**********************************************************************************************
-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!--'#########################################################################################################
'												1. 선 언 부 
'##########################################################################################################!-->
<!-- '******************************************  1.1 Inc 선언   **********************************************
'	기능: Inc. Include
'********************************************************************************************************* !-->
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->
<!-- #Include file="../../inc/IncSvrVariables.inc" -->

<!-- #Include file="../../inc/incSvrDBAgent.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->


<!--'==========================================  1.1.1 Style Sheet  ======================================
'==========================================================================================================!-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'==========================================  1.1.2 공통 Include   ======================================
'==========================================================================================================!-->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs"></SCRIPT> 
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentA.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentVariables.vbs"></SCRIPT>

<SCRIPT LANGUAGE="JavaScript"   SRC="../../inc/TabScript.js"></SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit          

'******************************************  1.2 Global 변수/상수 선언  ***********************************
'	1. Constant는 반드시 대문자 표기.
'**********************************************************************************************************

Const TAB1 = 1									
Const TAB2 = 2

Const ivType = "ST"

Dim C_IvNo					'매입번호 
Dim C_IvSeq					'매입순번 
Dim C_PlantCd				'공장 
Dim C_PlantNm				'공장명 
Dim C_ItemCd				'품목코드 
Dim C_ItemNm				'품목명 
Dim C_Spec					'규격 
Dim C_IvQty					'매입수량 
Dim C_CtlQty				'조정매입수량				
Dim C_IvPrc					'매입단가 
Dim C_CtlPrc				'조정매입단가 
Dim C_Amt					'매입금액 
Dim C_CtlAmt				'조정매입금액 
Dim C_NetAmt				'매입순금액 
Dim C_VatYn					'VAT 포함여부 
Dim C_VatFlg				'VAT 포함여부 
Dim C_VatAmt				'VAT 금액 
Dim C_VatCtlAmt				'조정VAT 금액 
Dim C_LocAmt				'매입자국금액 
Dim C_CtlLocAmt				'조정매입자국금액 
Dim C_LocVatAmt				'VAT 자국금액 
Dim C_NetLocAmt				'자국순금액 
Dim C_CtlLocVatAmt			'조정자국순금액 
Dim C_PoNo					'발주번호 
Dim C_PoSeq					'발주순번 
Dim C_IvNohdn				'HIDDEN 매입번호 
Dim C_IvSeqhdn				'HIDDEN 매입번호순번 
Dim C_Stateflg				'상태 FLAG

Dim C_CtlQty_Old			'
Dim C_CtlPrc_Old
Dim C_CtlAmt_Old
Dim C_VatCtlAmt_Old
Dim C_CtlLocAmt_Old
Dim C_CtlLocVatAmt_Old
Dim C_NetAmt_Old
Dim C_NetLocAmt_Old

' ==== MJG ADD 20050412 START ===
Dim	C_ItemAcct
Dim	C_VatType
Dim	C_VatRt
Dim	C_TrackingNo
Dim	C_IvCostCd
Dim	C_IvBizArea
Dim	C_MvmtQty
Dim	C_MvmtFlg
Dim	C_RefIvNo
Dim	C_RefIvSeqNo

Dim C_ItemUnit
' === MJG ADD 20050412 END ===



Dim StrTime
Dim EndTime
Dim DifferTime


<!-- #Include file="../../inc/lgvariables.inc" -->
'==========================================  1.2.1 Global 상수 선언  ======================================
'==========================================================================================================
Const BIZ_PGM_ID 					= "M5141MB1_KO441.asp"											
'Const BIZ_OnLine_ID 				= "m3111ab1.asp"
'Const BIZ_PGM_JUMP_ID_PO_DTL 		= "M3112MA1"
'Const BIZ_PGM_JUMP_ID_PUR_CHARGE	= "M6111MA2"
'==========================================  1.2.2 Global 변수 선언  =====================================
'	1. 변수 표준에 따름. prefix로 g를 사용함.
'	2.Array인 경우는 ()를 반드시 사용하여 일반 변수와 구별해 됨 
'=========================================================================================================

'==========================================  1.2.3 Global Variable값 정의  ===============================
'========================================================================================================= 
'----------------  공통 Global 변수값 정의  ----------------------------------------------------------- 

'++++++++++++++++  Insert Your Code for Global Variables Assign  ++++++++++++++++++++++++++++++++++++++
Dim lgMpsFirmDate, lgLlcGivenDt								
Dim gSelframeFlg
Dim lgIntFlgMode_Dtl
Dim cboOldVal          
Dim IsOpenPop          
Dim lblnWinEvent
Dim lgCboKeyPress      
Dim lgOldIndex								
Dim lgOldIndex2 
Dim lgOpenFlag  
Dim lgTabClickFlag  
Dim arrCollectVatType
Dim StartDate, EndDate
Dim iDBSYSDate
Dim lgReqRefChk


Dim lgNextKey

iDBSYSDate = "<%=GetSvrDate%>"

	'------ ☆: 초기화면에 뿌려지는 마지막 날짜 ------
	EndDate = UniConvDateAToB(iDBSYSDate, parent.gServerDateFormat, parent.gDateFormat)
    'StartDate = UniDateAdd("m", -1, iDBSYSDate,gServerDateFormat)    '☆: 초기화면에 뿌려지는 시작 날짜 -----
    'StartDate = UniConvDateAToB(StartDate,gServerDateFormat,gDateFormat)    
'#########################################################################################################
'												2. Function부 
'
'	내용 : 개발자가 정의한 함수, 즉 Event관련 함수를 제외한 모든 사용자 정의 함수 기슬 
'	공통으로 적용 사항 : 1. Sub 또는 Function을 호출할 때 반드시 Call을 쓴다.
'		     	     	 2. Sub, Function 이름에 _를 쓰지 않도록 한다. (Event와 구별하기 위함) 
'#########################################################################################################

'========================================================================================
' Function Name : OnLineQueryOK
' Function Desc : fi
'========================================================================================
Function OnLineQueryOK() 

	'If Trim(frm1.txtSpplCd.value) <> "" Then Call SupplierLookUp()    
	'======================== 추후에 수정=======================
	'if Trim(frm1.txtIvTypeCd.Value) <> "" then Call ChangePotype()
	'======================== 추후에 수정=======================
End Function


'==========================================   ChangeSupplier()  ======================================
Sub ChangeSupplier(BpType)
	lgBlnFlgChgValue = true	
	if CheckRunningBizProcess = true then
		exit sub
	end if
	Call SpplRef(BpType)
End Sub



'==========================================   SpplRef()  ======================================
'	Name : SpplRef()
'	Description : It is Call at txtSupplier Change Event
'=========================================================================================================
Sub SpplRef(BpType)
	If gLookUpEnable = False Then
		Exit Sub
	End If

    Err.Clear                                                      '☜: Protect system from crashing
    
    Dim strVal, StrvalBpCd
	Select Case BpType
		Case "1"                                                   '공급처인경우 화폐 변동 
		    if Trim(frm1.txtSpplCd.Value) = "" then
    			Exit Sub
    		End if
    		
    		StrvalBpCd = FilterVar(Trim(frm1.txtSpplCd.value), "", "SNM")
    	    
    	    if Trim(frm1.txtIvDt.Text) = ""  then
	            Call DisplayMsgBox("970021","X","매입등록일","X")
	            Exit Sub
	        End if
    	   
    	    if Trim(frm1.txtSpplCd.value) = ""  then
	            Call DisplayMsgBox("970021","X","공급처","X")
	            Exit Sub
	        End if
    	    
    	    Call GetPayDt()                                        '지불예정일 setting
    	Case "2"                                                   '지급처인경우 결제기간,대금결제참조,지급유형변동 
    		if Trim(frm1.txtPayeeCd.Value) = "" then               '발주번호 no checked경우 결제방법변동 
    			Exit Sub
    		End if
			StrvalBpCd = FilterVar(Trim(frm1.txtPayeeCd.value), "", "SNM")
    	Case "3"                                                  '세금계산서발행처인 경우 VAT,VAT이름,사업자등록번호 
    		if Trim(frm1.txtBuildCd.Value) = "" then
    			Exit Sub
    		End if   	
    		StrvalBpCd = FilterVar(Trim(frm1.txtBuildCd.value), "", "SNM")
    		
	        Call GetTaxBizArea("BP")
	        
	End Select
 
    strVal = BIZ_PGM_ID & "?txtMode=" & "LookUpSupplier"			'☜: 비지니스 처리 ASP의 상태 
    strVal = strval & "&txtBpType=" & BpType
    strVal = strVal & "&txtBpCd=" & StrvalBpCd		'☆: 조회 조건 데이타 
    
    if LayerShowHide(1) = False then
		Exit sub
	end if

	Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동	
	
End Sub

<%'======================================   GetPayDt()  =====================================
'	Name : GetPayDt()
'	Description : 지불예정일을 가져온다.
'==================================================================================================== %>
Sub GetPayDt()
   	Dim strSelectList, strFromList, strWhereList
	Dim strSpplCd, strIvDt,temp
	Dim strRs
	Dim arrTaxBizArea(2), arrTemp

    	strSpplCd  = frm1.txtSpplCd.value                       '공급처	
    	temp    = UNIConvDate(frm1.txtIvDt.text)            '매입등록일 
		strIvDt = mid(temp,1,4)
		strIvDt = strIvDt & mid(temp,6,2)
		strIvDt = strIvDt & mid(temp,9,2) 
		<%'발행처와 영업 그룹이 모두 등록되어 있는 경우 종합코드에 설정된 rule을 따른다 %>
    
	
	strSelectList = " * "
	strFromList = " dbo.ufn_m_GetPayDt( " & FilterVar(strSpplCd, "''", "S") & " ,  " & FilterVar(strIvDt, "''", "S") & " ) "
	strWhereList = ""
	
	Err.Clear
    
	If CommonQueryRs2by2(strSelectList,strFromList,strWhereList,strRs) Then
		arrTemp = Split(strRs, Chr(11))
		
	Else
		If Err.number <> 0 Then
			MsgBox Err.description 
			Err.Clear 
			Exit Sub
		End If
		
	End if
End Sub
'==========================================   Cfm()  ======================================
'	Name : Cfm()
'	Description : 확정버튼,확정취소버튼의 Event 합수 
'=========================================================================================================
 Sub Cfm()
    Dim IntRetCD 
    
    Err.Clear                                                               
    
    if lgBlnFlgChgValue = True	then
		Call DisplayMsgBox("189217", "X", "X", "X")
		Exit sub
	End if
	
	if frm1.rdoRelease(0).checked = True then
		IntRetCD = DisplayMsgBox("900018", Parent.VB_YES_NO, "x", "x")
		If IntRetCD = vbNo Then
			Exit Sub
		End If
		Call DbSave("Posting")				                                    
					                                                 
	elseif frm1.rdoRelease(1).checked = True then
			
		IntRetCD = DisplayMsgBox("900018", Parent.VB_YES_NO, "x", "x")
		If IntRetCD = vbNo Then
			Exit Sub
		End If
		Call DbSave("UnPosting")
	End if
	
End Sub

'--------------------------------------------------------------------
'		Cookie 사용함수 
'--------------------------------------------------------------------

Function CookiePage(Byval Kubun)

	Dim strTemp, arrVal
	Dim IntRetCD

		
	If Kubun = 0 Then

		strTemp = ReadCookie("PoNo")
		
		If strTemp = "" then Exit Function
		
		frm1.txtIvNo.value = strTemp
	
		WriteCookie "PoNo" , ""
		
		Call dbQuery()
	
	elseIf Kubun = 1 Then
		
	    If lgIntFlgMode <> Parent.OPMD_UMODE Then                           
	        Call DisplayMsgBox("900002", "X", "X", "X")
	        Exit Function
	    End If
		
	    If lgBlnFlgChgValue = True Then
			IntRetCD = DisplayMsgBox("900017", Parent.VB_YES_NO, "X", "X")
			If IntRetCD = vbNo Then
				Exit Function
			End If
	    End If
    	
		WriteCookie "PoNo" , frm1.txtIvNo.value
		
		Call PgmJump(BIZ_PGM_JUMP_ID_PO_DTL)
	
	elseIf Kubun = 2 Then
	
	    If lgIntFlgMode <> Parent.OPMD_UMODE Then                           
	        Call DisplayMsgBox("900002", "X", "X", "X")
	        Exit Function
	    End If
	    	
	    If lgBlnFlgChgValue = True Then
			IntRetCD = DisplayMsgBox("900017", Parent.VB_YES_NO, "X", "X")
			If IntRetCD = vbNo Then
				Exit Function
			End If
	    End If
    	
	    WriteCookie "Process_Step" , "PO"
		WriteCookie "Po_No" , Trim(frm1.txtIvNo.value)
		WriteCookie "Pur_Grp", Trim(frm1.txtBillCd.Value)
		WriteCookie "Po_Cur", Trim(frm1.txtPayeeCd.Value)
		WriteCookie "Po_Xch", Trim(frm1.txtXchRt.Value)
		
		Call PgmJump(BIZ_PGM_JUMP_ID_PUR_CHARGE)
				
	End IF
	
End Function
'------------------------------------------------------------------------------------------
'Radio에서 Click을 할 경우 flag를 Setting
'------------------------------------------------------------------------------------------
Sub Setchangeflg(byval kubun)
	lgBlnFlgChgValue = True	
	If kubun = 1 Then
		if frm1.rdoRelease(0).checked = true then
			frm1.hdnRelease.value= "N"
		else
			frm1.hdnRelease.value= "Y"
		end if
	Elseif kubun = 2 Then 
		if frm1.rdoVatFlg1.checked = true then
			frm1.hdvatFlg.Value = "1"
		else
			frm1.hdvatFlg.Value = "2"
		End if
	End if
End Sub
'------------------------------------------------------------------------------------------
'사용자가 Radio Button을 Click할 때 마다 숨겨진 hdnRelease를 Setting
'------------------------------------------------------------------------------------------
Sub Changeflg()
	if frm1.rdoRelease(0).checked = true then
		frm1.hdnRelease.value= "N"
	else
		frm1.hdnRelease.value= "Y"
	end if
End Sub

'=============================================== 2.2.3 InitSpreadPosVariables() ========================================
' Function Name : InitSpreadPosVariables
' Function Desc : This method assign sequential number for Spreadsheet column position
'========================================================================================
Sub InitSpreadPosVariables()
  C_IvNo              = 1
  C_IvSeq             = 2
  C_PlantCd           = 3
  C_PlantNm           = 4
  C_ItemCd            = 5
  C_ItemNm            = 6
  C_Spec              = 7
  C_IvQty             = 8
  C_CtlQty            = 9
  C_IvPrc             = 10
  C_CtlPrc            = 11
  C_Amt               = 12
  C_CtlAmt            = 13
  C_NetAmt			  = 14
  C_VatYn             = 15
  C_VatFlg            = 16
  C_VatAmt            = 17
  C_VatCtlAmt         = 18
  C_LocAmt			  = 19
  C_CtlLocAmt		  = 20
  C_NetLocAmt		  = 21
  C_LocVatAmt		  = 22
  C_CtlLocVatAmt      = 23
  
  C_PoNo              = 24
  C_PoSeq             = 25
  C_IvNohdn			  = 26
  C_IvSeqhdn		  = 27
  
  C_CtlQty_Old			= 28
  C_CtlPrc_Old			= 29
  C_CtlAmt_Old			= 30
  C_VatCtlAmt_Old		= 31
  C_CtlLocAmt_Old		= 32
  C_CtlLocVatAmt_Old	= 33
  
  C_NetAmt_Old			= 34
  C_NetLocAMt_Old		= 35

' ==== MJG ADD 20050412 START ===
	C_ItemAcct		= 36
	C_VatType       = 37
	C_VatRt         = 38
	C_TrackingNo        = 39
	C_IvCostCd      = 40
	C_IvBizArea     = 41
	C_MvmtQty       = 42
	C_MvmtFlg       = 43
	C_RefIvNo       = 44
	C_RefIvSeqNo    = 45
	
	C_ItemUnit		= 46
' ==== MJG ADD 20050412 END ===
 
  C_Stateflg		  = 47
  
  
  

  
End Sub

'=============================================== 2.2.3 InitSpreadSheet() ========================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================
Sub InitSpreadSheet()

	Call InitSpreadPosVariables
	
	With frm1.vspdData
	
	ggoSpread.Source = frm1.vspdData
	ggoSpread.Spreadinit "V20050320",,Parent.gAllowDragDropSpread 
	
	.ReDraw = false

    .MaxCols = C_Stateflg+1
    .Col = .MaxCols:	.ColHidden = True
    .MaxRows = 0

    Call GetSpreadColumnPos("A")
	ggoSpread.SSSetEdit 	C_IvNo , "매입번호", 7,,,18,2
	ggoSpread.SSSetEdit 	C_IvSeq, "순번", 10
    ggoSpread.SSSetEdit 	C_PlantCd, "공장", 7,,,4,2
    ggoSpread.SSSetEdit 	C_PlantNm, "공장명", 20
    ggoSpread.SSSetEdit 	C_ItemCd, "품목", 18,,,18,2
    ggoSpread.SSSetEdit 	C_ItemNm, "품목명", 20    
    ggoSpread.SSSetEdit		C_Spec, "품목규격", 20
    SetSpreadFloatLocal		C_IvQty, "매입수량",15,1,3       
    SetSpreadFloatLocal		C_CtlQty, "조정수량",15,1,3       
    SetSpreadFloatLocal		C_IvPrc, "매입단가", 15, 1, 4    
    SetSpreadFloatLocal		C_CtlPrc, "조정단가",15,1,4    
    SetSpreadFloatLocal		C_Amt, "금액",15,1,2    
    SetSpreadFloatLocal		C_CtlAmt, "조정금액",15,1,2  
    SetSpreadFloatLocal		C_NetAmt, "조정순금액",15,1,2    
    ggoSpread.SSSetCombo	C_VatYn,"VAT포함여부", 10,2,False               '13 차 추가 
    ggoSpread.SetCombo		"포함" & vbtab & "별도",C_VatYn  
    'ggoSpread.SSSetCombo 	C_VatFlg, "VAT포함여부코드", 15,2,False
    'ggoSpread.SetCombo		"" & vbtab & "1",C_VatFlg    
    ggoSpread.SSSetEdit		C_VatFlg , "VAT포함여부코드", 7,,,15,2 
    SetSpreadFloatLocal		C_VatAmt, "VAT금액",15,1,2        
    SetSpreadFloatLocal		C_VatCtlAmt, "VAT조정금액",15,1,2        
    SetSpreadFloatLocal		C_LocAmt, "매입자국금액",15,1,2 
    SetSpreadFloatLocal		C_CtlLocAmt, "조정자국금액",15,1,2   
	SetSpreadFloatLocal		C_NetLocAmt, "조정자국순금액",15,1,2   
	         
    SetSpreadFloatLocal		C_LocVatAmt, "VAT자국금액",15,1,2   
    SetSpreadFloatLocal		C_CtlLocVatAmt, "VAT조정자국금액",15,1,2  
    ggoSpread.SSSetEdit 	C_PoNo , "발주번호", 11,,,15,2 
    ggoSpread.SSSetEdit 	C_PoSeq , "발주순번", 7,,,15,2
    ggoSpread.SSSetEdit 	C_IvNohdn , "매입참조번호", 11,,,18,2 
    ggoSpread.SSSetEdit 	C_IvSeqhdn , "매입참조순번", 7,,,5,2 
    
    SetSpreadFloatLocal		C_CtlQty_Old, "CtlQty",15,1,4   
    SetSpreadFloatLocal		C_CtlPrc_Old, "CtlPrc",15,1,4   
    SetSpreadFloatLocal		C_CtlAmt_Old, "CtlAmt",15,1,4   
    SetSpreadFloatLocal		C_VatCtlAmt_Old, "VatCtlAmt",15,1,4   
    SetSpreadFloatLocal		C_CtlLocAmt_Old, "CtlLocAmt",15,1,4   
    SetSpreadFloatLocal		C_CtlLocVatAmt_Old, "CtlLocVatAmt",15,1,4   
    SetSpreadFloatLocal		C_NetAmt_Old, "NetAmt",15,1,4   
    SetSpreadFloatLocal		C_NetLocAmt_Old, "NetLocAmt",15,1,4   

	ggoSpread.SSSetEdit	C_ItemAcct  	,	"품목계정",	15,,,15,2	
	ggoSpread.SSSetEdit	C_VatType       ,	"VAT 유형",	15,,,15,2	
	ggoSpread.SSSetEdit	C_VatRt         ,	"VAT RATE",	15,,,15,2	
	ggoSpread.SSSetEdit	C_TrackingNo        ,	"입출고번호",	15,,,15,2	
	ggoSpread.SSSetEdit	C_IvCostCd      ,	"코스트센터",	15,,,15,2	
	ggoSpread.SSSetEdit	C_IvBizArea     ,	"매입사업장",	15,,,15,2	
	SetSpreadFloatLocal	C_MvmtQty       ,	"입출고수량",	15,1,3   
	ggoSpread.SSSetEdit	C_MvmtFlg       ,	"입출고여부",	15,,,15,2	
	ggoSpread.SSSetEdit	C_RefIvNo       ,	"참조매입번호",	15,,,15,2	
	ggoSpread.SSSetEdit	C_RefIvSeqNo    ,	"참조매입순번",	15,,,15,2	  
	
	ggoSpread.SSSetEdit	C_ItemUnit   ,	"품목단위",	15,,,15,2	  
    
    ggoSpread.SSSetEdit		C_Stateflg , "C_Stateflg" , 10
    
    '완료후 Hidden 영역 설정을 위해 주석을 풀것 
	Call ggoSpread.SSSetColHidden(C_IvNo,C_IvNo,True)
	Call ggoSpread.SSSetColHidden(C_IvSeq,C_IvSeq,True)
	Call ggoSpread.SSSetColHidden(C_VatFlg,C_VatFlg,True)
	Call ggoSpread.SSSetColHidden(C_CtlQty_Old,C_CtlQty_Old,True)
	Call ggoSpread.SSSetColHidden(C_CtlPrc_Old,C_CtlPrc_Old,True)
	Call ggoSpread.SSSetColHidden(C_CtlAmt_Old,C_CtlAmt_Old,True)
	Call ggoSpread.SSSetColHidden(C_VatCtlAmt_Old,C_VatCtlAmt_Old,True)
	Call ggoSpread.SSSetColHidden(C_CtlLocAmt_Old,C_CtlLocAmt_Old,True)
	Call ggoSpread.SSSetColHidden(C_CtlLocVatAmt_Old,C_CtlLocVatAmt_Old,True)
	Call ggoSpread.SSSetColHidden(C_NetAmt_Old,C_NetAmt_Old,True)
	Call ggoSpread.SSSetColHidden(C_NetLocAMt_Old,C_NetLocAMt_Old,True)
	Call ggoSpread.SSSetColHidden(C_Stateflg,C_Stateflg,True)
	
	Call ggoSpread.SSSetColHidden(C_ItemAcct  	, C_ItemAcct  	,True)	
	Call ggoSpread.SSSetColHidden(C_VatType   	, C_VatType   	,True)    
	Call ggoSpread.SSSetColHidden(C_VatRt     	, C_VatRt     	,True)    
	Call ggoSpread.SSSetColHidden(C_TrackingNo    	, C_TrackingNo    	,True)    
	Call ggoSpread.SSSetColHidden(C_IvCostCd  	, C_IvCostCd  	,True)    
	Call ggoSpread.SSSetColHidden(C_IvBizArea 	, C_IvBizArea 	,True)    
	Call ggoSpread.SSSetColHidden(C_MvmtQty   	, C_MvmtQty   	,True)    
	Call ggoSpread.SSSetColHidden(C_MvmtFlg   	, C_MvmtFlg   	,True)    
	Call ggoSpread.SSSetColHidden(C_RefIvNo   	, C_RefIvNo   	,True)    
	Call ggoSpread.SSSetColHidden(C_RefIvSeqNo	, C_RefIvSeqNo	,True)
	Call ggoSpread.SSSetColHidden(C_ItemUnit	, C_ItemUnit	,True)    	    	

    
    Call SetSpreadLock
    
	.ReDraw = true
	
    End With
    
End Sub

'================================== 2.2.4 SetSpreadLock() ==================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadLock()
    
    With frm1
    ggoSpread.SpreadLock frm1.vspddata.maxcols,-1
    ggoSpread.SpreadLock C_IvNo , -1
    ggoSpread.SpreadLock C_IvSeq , -1
    ggoSpread.SpreadLock C_PlantCd, -1
    ggoSpread.SpreadLock C_PlantNm , -1
    ggoSpread.SpreadLock C_ItemCd, -1
    ggoSpread.SpreadLock C_ItemNm, -1
    ggoSpread.SpreadLock C_IvQty, -1
    
    ggoSpread.SpreadUnLock C_CtlQty, -1
    ggoSpread.SSSetrequired C_CtlQty, -1
    
    ggoSpread.SpreadLock C_IvPrc,-1
    
    ggoSpread.SpreadUnLock C_CtlPrc, -1
    ggoSpread.SSSetrequired C_CtlPrc, -1
    
    ggoSpread.SpreadLock C_Amt,-1
    
    ggoSpread.SpreadUnLock C_CtlAmt, -1
    ggoSpread.SSSetrequired C_CtlAmt, -1

    ggoSpread.SpreadLock C_NetAmt, -1
        
    ggoSpread.SpreadUnLock C_VatYn,-1    
    ggoSpread.SSSetrequired C_VatYn, -1
    
    ggoSpread.SpreadLock C_VAtFlg, -1
    
    ggoSpread.SpreadLock C_VatAmt, -1
    
    ggoSpread.SpreadUnLock C_VatCtlAmt, -1
    ggoSpread.SSSetrequired C_VatCtlAmt, -1
    
    ggoSpread.SpreadLock C_LocAmt, -1
    ggoSpread.SpreadUnLock C_CtlLocAmt, -1
    ggoSpread.SSSetrequired C_CtlLocAmt, -1
    
    ggoSpread.SpreadLock C_NetLocAmt, -1
    
    ggoSpread.SpreadLock C_LocVatAmt, -1
    ggoSpread.SpreadUnLock C_CtlLocVatAmt, -1
    ggoSpread.SSSetrequired C_CtlLocVatAmt, -1
    
    
    ggoSpread.SpreadLock C_PoNo, -1
    ggoSpread.SpreadLock C_PoSeq, -1
    ggoSpread.SpreadLock C_IvNohdn, -1
    ggoSpread.SpreadLock C_IvSeqhdn, -1
    
    End With
    
       
End Sub

Sub SetSpreadLockAfterQuery()

	Dim index,Count,index1 , strReqChk

    With frm1
    
   .vspdData.ReDraw = False
    
    if .vspdData.MaxRows < 1 then
		if .hdnRelease.Value <> "Y" then
			'Call SetToolbar("1110111111101")
		End if
		Exit sub
	end if
	
	if .hdnRelease.Value = "Y" then
		For index = C_SeqNo to C_Stateflg
			ggoSpread.SpreadLock index , -1
		Next
	Else
		
		For index1 = Cint(.hdnmaxrows.value) + 1 to .vspdData.MaxRows
		    ggoSpread.SpreadLock frm1.vspddata.maxcols, index1, frm1.vspddata.maxcols, index1
			ggoSpread.SpreadLock C_SeqNo , index1,C_SeqNo,index1
			ggoSpread.SpreadLock C_PlantCd ,index1,C_PlantCd,index1
			ggoSpread.SpreadLock C_Popup1 , index1,C_Popup1,index1
			ggoSpread.spreadlock C_PlantNm , index1,C_PlantNm,index1
			ggoSpread.SpreadLock C_ItemCd, index1,C_ItemCd,index1
			ggoSpread.SpreadLock C_Popup2 , index1,C_Popup2,index1
			ggoSpread.spreadlock C_ItemNm , index1,C_ItemNm,index1
			ggoSpread.spreadlock C_SpplSpec,index1,C_SpplSpec,index1         '품목규격 추가 
			ggoSpread.SpreadUnLock C_OrderQty,index1,C_OrderQty,index1
			ggoSpread.sssetrequired C_OrderQty, index1,index1
			
			if UCase(frm1.hdnRetflg.Value) = "N" then
				ggoSpread.SpreadUnLock C_OrderUnit , index1,C_OrderUnit,index1
				ggoSpread.sssetrequired C_OrderUnit, index1,index1
				ggoSpread.SpreadUnLock C_Popup3 , index1,C_Popup3,index1
				ggoSpread.SpreadUnLock C_Cost , index1,C_Cost,index1
				ggoSpread.sssetrequired C_Cost, index1,index1
			else
				ggoSpread.SpreadLock C_OrderUnit , index1,C_OrderUnit,index1
				ggoSpread.SpreadLock C_Popup3 , index1,C_Popup3,index1
				ggoSpread.SpreadLock C_Cost , index1,C_Cost,index1
			end if		

			ggoSpread.SpreadUnLock C_CostCon, index1,C_CostCon,index1
			ggoSpread.sssetrequired C_CostCon, index1,index1
			ggoSpread.spreadlock C_NetAmt, index1,C_NetAmt,index1		

			if .hdnImportflg.value = "Y" then
				ggoSpread.spreadUnlock C_HSCd , index1,C_HSCd,index1
				ggoSpread.sssetrequired C_HSCd, index1,index1
				ggoSpread.spreadUnlock C_Popup5 , index1,C_Popup5,index1
				ggoSpread.spreadlock C_HSNm , index1,C_HSNm,index1				
			else
				ggoSpread.spreadlock C_HSCd, index1,C_HSCd,index1
				ggoSpread.spreadlock C_Popup5, index1,C_Popup5,index1
				ggoSpread.spreadlock C_HSNm, index1,C_HSNm,index1
			End if	
			
'			If Trim(.hdnreference.value) = "N" then
'			     ggoSpread.SSSetProtected	C_OrderAmt, index1, index1
'			else 
			    ggoSpread.SSSetRequired  C_OrderAmt, index1, index1
'			end if
    
			ggoSpread.spreadlock C_TrackingNo , index1,C_TrackingNo,index1
			ggoSpread.SpreadUnLock C_IOFlg, index1,C_IOFlgCd,index1 
			ggoSpread.SSSetRequired	C_IOFlg, index1,index1 
			ggoSpread.SSSetRequired	C_IOFlgCd, index1,index1
		    
			ggoSpread.SpreadUnLock C_VatType , index1,C_VatType,index1
			ggoSpread.SpreadUnLock C_Popup7 , index1,C_Popup7,index1
			ggoSpread.spreadlock C_VatNm , index1,C_VatNm,index1
			ggoSpread.spreadlock C_VatRate , index1,C_VatRate,index1
			ggoSpread.spreadlock C_VatAmt , index1,C_VatAmt,index1
		'******************************************
		  '13차추가]
			if .hdnRetflg.Value = "Y" then
				ggoSpread.spreadUnLock C_RetCd , index1,C_RetCd,index1
				ggoSpread.SpreadUnLock C_Popup8 , index1,C_Popup8,index1
				ggoSpread.spreadlock   C_RetNm , index1,C_RetNm,index1
				ggoSpread.spreadUnLock C_Lot_No , index1,C_Lot_No,index1       
				ggoSpread.spreadUnLock C_Lot_Seq , index1,C_Lot_Seq,index1 
			else
				ggoSpread.spreadlock C_RetCd , index1,C_RetCd,index1		
				ggoSpread.spreadlock C_Popup8 , index1,C_Popup8,index1		
				ggoSpread.spreadlock C_RetNm , index1,C_RetNm,index1		
		        ggoSpread.spreadlock C_Lot_No , index1,C_Lot_No,index1        
		        ggoSpread.spreadlock C_Lot_Seq , index1,C_Lot_Seq,index1      
		    end if        
		'******************************************
		    ggoSpread.SpreadUnLock C_SLCd , index1,C_SLCd,index1
		    ggoSpread.sssetrequired C_SLCd, index1,index1
		    ggoSpread.SpreadUnLock C_Popup6 , index1,C_Popup6,index1
		    ggoSpread.spreadlock C_SLNm, index1,C_SLNm,index1
			
            .vspdData.Row = index1
			.vspdData.Col = C_TrackingNo
			if Trim(.vspdData.Text) = "*" then
				ggoSpread.spreadlock C_TrackingNo, index1, C_TrackingNoPop, index1
			else
				ggoSpread.spreadUnlock C_TrackingNo, index1, C_TrackingNoPop, index1
				ggoSpread.sssetrequired C_TrackingNo, index1, index1
			end if

			'************************************************ 13차	

			frm1.vspdData.Row = index1
		    frm1.vspdData.Col = C_PrNo

			if Trim(.vspdData.Text) <> "" then
				ggoSpread.spreadlock C_OrderUnit, index1, C_OrderUnit, index1
				ggoSpread.spreadlock C_Popup3 , index1, C_Popup3, index1
		        ggoSpread.spreadlock C_DlvyDT, index1,C_DlvyDT, index1
		        ggoSpread.spreadlock C_TrackingNo, index1, C_TrackingNoPop, index1
			
				ggoOper.SetReqAttr	frm1.txtBillCd, "Q"
				ggoOper.SetReqAttr	frm1.txtSpplCd, "Q"
			else
				ggoSpread.spreadUnlock C_OrderUnit, index1, C_OrderUnit, index1
				ggoSpread.sssetrequired C_OrderUnit, index1, index1
				ggoSpread.SpreadUnLock C_DlvyDT, index1,C_DlvyDT, index1
			    ggoSpread.sssetrequired C_DlvyDT, index1, index1
			end if
		    ggoSpread.spreadUnlock C_Under,index1,C_Under,index1
		    ggoSpread.spreadUnlock C_Over,index1,C_Over,index1
	    next
	End if

	.vspdData.ReDraw = True
	End With 
	
	if frm1.hdnImportflg.value = "Y" then
	    ggoOper.SetReqAttr	frm1.txtCnfmDt, "N"
	else     
		ggoOper.SetReqAttr	frm1.txtCnfmDt, "D"
		'ggoOper.SetReqAttr	frm1.txtOffDt, "Q"
		'ggoOper.SetReqAttr	frm1.txtApplicantCd, "Q"
		'ggoOper.SetReqAttr	frm1.txtApplicantNm, "Q"
		'ggoOper.SetReqAttr	frm1.txtIncotermsCd, "Q"
		'ggoOper.SetReqAttr	frm1.txtIncotermsNm, "Q"
		'ggoOper.SetReqAttr	frm1.txtTransCd, "Q"
		'ggoOper.SetReqAttr	frm1.txtTransNm, "Q"
	end if	

End Sub
'================================== 2.2.5 SetSpreadColor() ==================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
    With frm1

    ggoSpread.SSSetProtected	frm1.vspddata.maxcols, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected	C_SeqNo		, pvStartRow, pvEndRow
    ggoSpread.SSSetRequired		C_PlantCd	, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected	C_PlantNm	, pvStartRow, pvEndRow
    ggoSpread.SSSetRequired		C_ItemCd	, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected	C_ItemNm	, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected	C_SpplSpec	, pvStartRow, pvEndRow '품목규격 추가 
    ggoSpread.SSSetRequired		C_OrderQty	, pvStartRow, pvEndRow
    ggoSpread.SSSetRequired		C_OrderUnit	, pvStartRow, pvEndRow
    ggoSpread.SSSetRequired		C_Cost		, pvStartRow, pvEndRow
    ggoSpread.SSSetRequired		C_CostCon	, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected	C_CostConCd	, pvStartRow, pvEndRow
    
'    If Trim(.hdnreference.value) = "N" then
'        ggoSpread.SSSetProtected	C_OrderAmt, pvStartRow, pvEndRow
'    else 
        ggoSpread.SSSetRequired  C_OrderAmt, pvStartRow, pvEndRow
'    end if
    
    ggoSpread.SSSetProtected	C_NetAmt, pvStartRow, pvEndRow
    ggoSpread.SSSetRequired		C_DlvyDt, pvStartRow, pvEndRow
    
    if Trim(frm1.hdnImportflg.value) <> "Y" then
	    ggoSpread.SSSetProtected	C_HSCd	, pvStartRow, pvEndRow
	    ggoSpread.SSSetProtected	C_Popup5, pvStartRow, pvEndRow
	else
		ggoSpread.spreadUnlock	C_HSCd	, pvStartRow, C_HSCd, pvEndRow
		ggoSpread.sssetrequired	C_HSCd	, pvStartRow, pvEndRow
		ggoSpread.spreadUnlock	C_Popup5, pvStartRow, C_Popup5, pvEndRow
	end if
	
	ggoSpread.SSSetProtected		C_TrackingNo, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected		C_TrackingNoPop, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected		C_HSNm	, pvStartRow, pvEndRow
    ggoSpread.SSSetRequired		C_SLCd	, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected		C_SLNm	, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected		C_VatNm	, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected		C_VatRate, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected		C_VatAmt , pvStartRow, pvEndRow
    '******************************************
	ggoSpread.SSSetRequired		C_IOFlg	 , pvStartRow, pvEndRow
	ggoSpread.SSSetProtected		C_IOFlgCd, pvStartRow, pvEndRow  '13차추가 
	if .hdnRetflg.Value <> "Y" then
		ggoSpread.SSSetProtected C_RetCd	, pvStartRow, pvEndRow		
		ggoSpread.SSSetProtected C_Popup8, pvStartRow, pvEndRow		
		ggoSpread.SSSetProtected C_RetNm	, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_Lot_No, pvStartRow, pvEndRow        
		ggoSpread.SSSetProtected C_Lot_Seq, pvStartRow, pvEndRow      
	end if        
	'******************************************
    End With
End Sub

'================================== 2.2.5 SetSpreadColor() ==================================================
' Function Name : SetSpreadColorRef
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadColorRef(ByVal pvStartRow, ByVal pvEndRow)
    With frm1

	ggoSpread.SSSetRequired		C_CtlPrc	, pvStartRow, pvEndRow
	ggoSpread.SSSetRequired		C_CtlAmt	, pvStartRow, pvEndRow
	ggoSpread.SSSetRequired		C_VatYn		, pvStartRow, pvEndRow
	ggoSpread.SSSetRequired		C_VatFlg	, pvStartRow, pvEndRow
	ggoSpread.SSSetRequired		C_VatCtlAmt	, pvStartRow, pvEndRow
	ggoSpread.SSSetRequired		C_CtlLocAmt	, pvStartRow, pvEndRow
	ggoSpread.SSSetRequired		C_CtlLocVatAmt	, pvStartRow, pvEndRow
	
	
    End With
End Sub
'========================================================================================
' Function Name : GetSpreadColumnPos
' Description   : 
'========================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
	Dim iCurColumnPos
	
	Select Case UCase(pvSpdNo)
		Case "A"
			ggoSpread.Source = frm1.vspdData
			
			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

				C_IvNo              = iCurColumnPos(1)
				C_IvSeq             = iCurColumnPos(2)
  				C_PlantCd           = iCurColumnPos(3)
  				C_PlantNm           = iCurColumnPos(4)
  				C_ItemCd            = iCurColumnPos(5)
  				C_ItemNm            = iCurColumnPos(6)
  				C_Spec              = iCurColumnPos(7)
  				C_IvQty             = iCurColumnPos(8)
  				C_CtlQty            = iCurColumnPos(9)
  				C_IvPrc             = iCurColumnPos(10)
  				C_CtlPrc            = iCurColumnPos(11)
  				C_Amt               = iCurColumnPos(12)
  				C_CtlAmt            = iCurColumnPos(13)
  				C_NetAmt            = iCurColumnPos(14)
  				C_VatYn             = iCurColumnPos(15)
  				C_VatFlg            = iCurColumnPos(16)
				C_VatAmt            = iCurColumnPos(17)
  				C_VatCtlAmt         = iCurColumnPos(18)
  				C_LocAmt			= iCurColumnPos(19)
  				C_CtlLocAmt			= iCurColumnPos(20)
  				C_NetLocAmt			= iCurColumnPos(21)
  				C_LocVatAmt			= iCurColumnPos(22)
  				C_CtlLocVatAmt      = iCurColumnPos(23)
  				C_PoNo              = iCurColumnPos(24)
  				C_PoSeq             = iCurColumnPos(25)
  				C_IvNohdn			= iCurColumnPos(26)
  				C_IvSeqhdn			= iCurColumnPos(27)
  				
				C_CtlQty_Old		= iCurColumnPos(28)
				C_CtlPrc_Old		= iCurColumnPos(29)
				C_CtlAmt_Old		= iCurColumnPos(30)
				C_VatCtlAmt_Old		= iCurColumnPos(31)
				C_CtlLocAmt_Old		= iCurColumnPos(32)
				C_CtlLocVatAmt_Old	= iCurColumnPos(33)
				
				C_NetAmt_Old	= iCurColumnPos(34)
				C_NetLocAmt_Old	= iCurColumnPos(35)

				C_ItemAcct          =	iCurColumnPos(36)
				C_VatType           =	iCurColumnPos(37)
				C_VatRt             =	iCurColumnPos(38)
				C_TrackingNo            =	iCurColumnPos(39)
				C_IvCostCd          =	iCurColumnPos(40)
				C_IvBizArea         =	iCurColumnPos(41)
				C_MvmtQty           =	iCurColumnPos(42)
				C_MvmtFlg           =	iCurColumnPos(43)
				C_RefIvNo           =	iCurColumnPos(44)
				C_RefIvSeqNo		=	iCurColumnPos(45)
				C_ItemUnit			=	iCurColumnPos(46)

  				C_Stateflg			= iCurColumnPos(47)
	End Select

End Sub	
'==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
'=========================================================================================================
Sub InitVariables()

    lgIntFlgMode = Parent.OPMD_CMODE  
    lgIntFlgMode_Dtl = Parent.OPMD_CMODE                                        
    lgBlnFlgChgValue = False                                         
    lgIntGrpCount = 0                                                
    IsOpenPop = False												
    lgCboKeyPress = False
    lgOldIndex = -1
    lgOldIndex2 = -1
    lgMpsFirmDate=""
    lgLlcGivenDt=""
    lgStrPrevKey = ""                  
    frm1.vspdData.MaxRows = 0
    

End Sub

'=============================== 2.1.2 LoadInfTB19029() ========================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== 
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029A("I", "*","NOCOOKIE","MA") %>
<% Call LoadBNumericFormatA("I", "*","NOCOOKIE","MA") %>
End Sub

'******************************************  2.2 화면 초기화 함수  ***************************************
'	기능: 화면초기화 
'	설명: 화면초기화, Combo Display, 화면 Clear 등 화면 초기화 작업을 한다. 
'*********************************************************************************************************

'******************************************  2.2 화면 초기화 함수  ***************************************
'	기능: 화면초기화 
'	설명: 화면초기화, Combo Display, 화면 Clear 등 화면 초기화 작업을 한다. 
'*********************************************************************************************************

'==========================================  2.2.1 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
'=========================================================================================================
Sub SetDefaultVal()
	
	lgOpenFlag	= False
	lgTabClickFlag	= False
	gSelframeFlg = TAB1
	lgReqRefChk = False
	
    Call SetToolbar("1110100100001111")
    frm1.rdoRelease(0).checked = true
    frm1.hdnRelease.value = "N"
    'frm1.txtOffDt.text = EndDate
    frm1.txtIvDt.text = EndDate
    frm1.txtCnfmDt.text = EndDate
    frm1.hdnCurr.value = Parent.gCurrency   
    frm1.btnCfm.disabled = true
    frm1.btnSelect.disabled = true
    
    frm1.btnSend.disabled = true
    frm1.txtXchRt.Text = ""
	frm1.btnCfm.value = "확정"
	frm1.txtIvNo.focus	

	frm1.cboXchop.value = "*"
	frm1.hdnxchrateop.value ="*"
	frm1.hdnMergPurFlg.value = "N"


	Set gActiveElement = document.activeElement	

End Sub
'==========================================================================================
'   Event Name : InitComboBox
'   Event Desc : 콤보 박스 초기화 
'==========================================================================================

Sub InitComboBox()
	Call SetCombo(frm1.cboXchop,"*","*")
	Call SetCombo(frm1.cboXchop,"/","/")  
End Sub

'******************************************  2.3 Operation 처리함수  *************************************
'	기능: Operation 처리부분 
'	설명: Tab처리, Reference등을 행한다. 
'*********************************************************************************************************
Function ClickTab1()

	If gSelframeFlg = TAB1 Then Exit Function
	
	Call changeTabs(TAB1)	<% '~~~ 첫번째 Tab %>
	gSelframeFlg = TAB1
	
   	'Call setFocus(CLICK_HEADER)
   	frm1.txtIvNo.focus
	'Call SetToolbar("11111000001111")
	Call BtnToolCtrl(TAB1)
	
	Set gActiveElement = document.activeElement
	
End Function

Function ClickTab2()

	If gSelframeFlg = TAB2 Then Exit Function
	
	if frm1.txtIvTypeCd.value = "" then
		Call DisplayMsgBox("171800", "X", "X", "X")  
		frm1.txtIvTypeCd.focus
		Exit Function
	End if
   	
   	Call changeTabs(TAB2)	
	gSelframeFlg = TAB2
	
	frm1.txtIvNo.focus
	'Call BtnToolCtrl(TAB2)
	
	Set gActiveElement = document.activeElement
	
	
End Function



'------------------------------------------  SetClickflag, ResetClickflag()  -----------------------------
'	Name : SetClickflag, ResetClickflag()
'	Description :  
'---------------------------------------------------------------------------------------------------------

Function SetClickflag()
	lgTabClickFlag = True	
End Function

Function ResetClickflag()
	lgTabClickFlag = False
End Function

'Sub InitCollectType()
'	Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6, i
'	Dim iCodeArr, iNameArr, iRateArr
'
'    Err.Clear

'	Call CommonQueryRs(" Minor.MINOR_CD,  Minor.MINOR_NM, Config.REFERENCE ", " B_MINOR Minor,B_CONFIGURATION Config ", " Minor.MAJOR_CD='B9001' And Config.MAJOR_CD = Minor.MAJOR_CD And Config.MINOR_CD = Minor.MINOR_CD And Config.SEQ_NO = 1 ", lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

'    iCodeArr = Split(lgF0, Chr(11))
'    iNameArr = Split(lgF1, Chr(11))
'    iRateArr = Split(lgF2, Chr(11))

'	If Err.number <> 0 Then
'		MsgBox Err.description 
'		Err.Clear 
'		Exit Sub
'	End If

'	Redim arrCollectVatType(UBound(iCodeArr) - 1, 2)

'	For i = 0 to UBound(iCodeArr) - 1
'		arrCollectVatType(i, 0) = iCodeArr(i)
'		arrCollectVatType(i, 1) = iNameArr(i)
'		arrCollectVatType(i, 2) = iRateArr(i)
'	Next
'End Sub

'========================================================================================
' Function Name : GetCollectTypeRef
' Function Desc : 
'========================================================================================
'Sub GetCollectTypeRef(ByVal VatType, ByRef VatTypeNm, ByRef VatRate)
'
'	Dim iCnt

'	For iCnt = 0 To Ubound(arrCollectVatType)  
'		If arrCollectVatType(iCnt, 0) = UCASE(VatType) Then
'			VatTypeNm = arrCollectVatType(iCnt, 1)
'			VatRate   = arrCollectVatType(iCnt, 2)
'			Exit Sub
'		End If
'	Next
'	VatTypeNm = ""
'	VatRate = ""
'End Sub

'=====================================  SetVatType()  =========================================
Sub SetVatType()

	Dim VatType, VatTypeNm, VatRate

	VatType = Trim(frm1.txtVatCd.value)
	Call InitCollectType
	Call GetCollectTypeRef(VatType, VatTypeNm, VatRate)
    
	frm1.txtVatNm.value = VatTypeNm
	frm1.txtVatRt.text = UNIFormatNumber(UNICDbl(VatRate), ggExchRate.DecPoint, -2, 0, ggExchRate.RndPolicy, ggExchRate.RndUnit)

End Sub

'******************************************  2.4 POP-UP 처리함수  ****************************************
'	기능: 기준 POP-UP
'   설명: 기준 POP-UP에 관한 Open은 Include한다. 
'	      하나의 ASP에서 Popup이 중복되면 하나는 링크시켜 사용하고 나머지는 재정의하여 사용한다.
'*********************************************************************************************************

'------------------------------------------  OpenReqRef()  -------------------------------------------------
'	Name : OpenReqRef()
'	Description :구매요청참조 
'---------------------------------------------------------------------------------------------------------

Function OpenReqRef()

	Dim strRet
	Dim arrParam(11)
	Dim iCalledAspName
	Dim IntRetCD
	
	'if lgIntFlgMode = Parent.OPMD_CMODE then
	'	Call DisplayMsgBox("900002", "X", "X", "X")
	'	Exit Function
	'End if 
	
    If CheckRunningBizProcess = True Then
		Exit Function
	End If

	'이성룡 주석 
	'if frm1.txtIvTypeCd.value = "" then
	'	Call DisplayMsgBox("179010", "X", "X", "X")  
	'	frm1.txtIvTypeCd.focus
	'	Exit Function
	'End if
	
	if frm1.hdnRelease.Value = "Y" then
		
		Call DisplayMsgBox("17a008", "X", "X", "X")
		Exit Function
	End if
	
	If IsOpenPop = True Then Exit Function
			
	IsOpenPop = True
	
	'이성룡 추가 
'====== 2005.06.22 재고반영여부 관련 수정 ========
'	if Trim(frm1.txtIvTypeCd.value) = "" then
'        arrParam(0) = ""
'	Else
'        arrParam(0) = Trim(frm1.txtIvTypeCd.value)
'        arrParam(1) = Trim(frm1.txtIvTypeNm.value)
'	End if
'====== 2005.06.22 재고반영여부 관련 수정 ========
	
	if Trim(frm1.txtGrpCd.value) = "" then
        arrParam(2) = ""
	Else
        arrParam(2) = Trim(frm1.txtGrpCd.value)
        arrParam(3) = Trim(frm1.txtGrpNm.value)
	End if
	
	if Trim(frm1.txtVatCd.value) = "" then
        arrParam(4) = ""
	Else
        arrParam(4) = Trim(frm1.txtVatCd.value)
        arrParam(5) = Trim(frm1.txtVatNm.value)
	End if
	
	if Trim(frm1.txtSpplCd.value) = "" then
        arrParam(6) = ""
	Else
        arrParam(6) = Trim(frm1.txtSpplCd.value)
        arrParam(7) = Trim(frm1.txtSpplNm.value)
	End if
		
	if Trim(frm1.txtBuildCd.value) = "" then
        arrParam(8) = ""
	Else
        arrParam(8) = Trim(frm1.txtBuildCd.value)
        arrParam(9) = Trim(frm1.txtBuildNm.value)
	End if
	
	if Trim(frm1.txtIvNo2.value) = "" then
        arrParam(10) = ""
	Else
        arrParam(10) = Trim(frm1.txtIvNo2.value)
    End if
    
    if Trim(frm1.txtCur.value) = "" then
		arrParam(11) = ""
	Else 
		arrParam(11) = Trim(frm1.txtCur.value)
	End if
	
	'이성룡 추가 끝 
	

	
	iCalledAspName = AskPRAspName("M5141RA1_KO441")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "M5141RA1_KO441", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	strRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
		"dialogWidth=760px; dialogHeight=560px; center: Yes; help: No; resizable: No; status: No;")


	IsOpenPop = False
	
	If strRet(0,0) = "" Then
		Exit Function
	Else
		Call SetReqRef(strRet)
	End If
		
End Function
'------------------------------------------  OpenSppl()  -------------------------------------------------
'	Name : OpenSppl()
'	Description :공급처,세금계산서발행처,지급처 
'---------------------------------------------------------------------------------------------------------
Function OpenSppl(Byval BpType)
	Dim arrRet
	Dim arrHeader(6), arrField(6), arrParam(5) 
	
	arrHeader(2) = "사업자등록번호"									' Header명(2)
    arrField(0) = "B_BIZ_PARTNER.BP_Cd"									' Field명(0)
    arrField(1) = "B_BIZ_PARTNER.BP_Nm"								    ' Field명(1)
    arrField(2) = "B_BIZ_PARTNER.BP_RGST_NO"							' Field명(2)
    
	Select Case BpType
		Case "1"  '공급처 
			If lblnWinEvent = True Or UCase(frm1.txtSpplCd.className) = UCase(parent.UCN_PROTECTED) Then Exit Function    
			lblnWinEvent = True
			arrHeader(0) = "공급처"											' Header명(0)
			arrHeader(1) = "공급처명"										' Header명(1)

		    arrParam(0) = "공급처"											' 팝업 명칭 
			arrParam(1) = "B_BIZ_PARTNER "					                    ' TABLE 명칭 
			arrParam(2) = Trim(frm1.txtSpplCd.Value)		
			'arrParam(2) = Trim(frm1.txtSpplCd.Value)							' Code Condition
			arrParam(4) = "B_BIZ_PARTNER.BP_TYPE In (" & FilterVar("S", "''", "S") & " ," & FilterVar("CS", "''", "S") & ") And B_BIZ_PARTNER.usage_flag=" & FilterVar("Y", "''", "S") & " "		' Where Condition
			arrParam(5) = "공급처"											' TextBox 명칭 
		Case "2"   '지급처 
			If lblnWinEvent = True Or UCase(frm1.txtPayeeCd.className) = UCase(parent.UCN_PROTECTED) Then Exit Function    
			lblnWinEvent = True

			arrHeader(0) = "지급처"											' Header명(0)
			arrHeader(1) = "지급처명"										' Header명(1)

			arrParam(0) = "지급처"											' 팝업 명칭 
			arrParam(1) = "B_BIZ_PARTNER, B_BIZ_PARTNER_FTN	"					' TABLE 명칭 
			arrParam(2) = Trim(frm1.txtPayeeCd.Value)							' Code Condition%>
			'arrParam(2) = Trim(frm1.txtPayeeCd.Value)							' Code Condition%>
			arrParam(4) = "B_BIZ_PARTNER.BP_TYPE In (" & FilterVar("S", "''", "S") & " ," & FilterVar("CS", "''", "S") & ") And B_BIZ_PARTNER.usage_flag=" & FilterVar("Y", "''", "S") & " "
			arrParam(4) = arrParam(4) & " AND B_BIZ_PARTNER.BP_CD = B_BIZ_PARTNER_FTN.PARTNER_BP_CD  AND B_BIZ_PARTNER_FTN.BP_CD = " 				<%' Where Condition%>
            arrParam(4) = arrParam(4) & FilterVar(Trim(frm1.txtSpplCd.Value), "''", "S") & " AND  B_BIZ_PARTNER_FTN.PARTNER_FTN = " & FilterVar("MPA", "''", "S") & " "
			arrParam(5) = "지급처"											' TextBox 명칭 
		Case "3"   '세금계산서발행처 
			If lblnWinEvent = True Or UCase(frm1.txtBuildCd.className) = UCase(parent.UCN_PROTECTED) Then Exit Function    
			lblnWinEvent = True

			arrHeader(0) = "세금계산서발행처"											' Header명(0)
			arrHeader(1) = "세금계산서발행처명" 										' Header명(1)

			arrParam(0) = "세금계산서발행처"											' 팝업 명칭 
			arrParam(1) = "B_BIZ_PARTNER, B_BIZ_PARTNER_FTN	"           					' TABLE 명칭 
			arrParam(2) = Trim(frm1.txtBuildCd.Value)					            		' Code Condition
			'arrParam(2) = Trim(frm1.txtBuildCd.Value)					            		' Code Condition
			arrParam(4) = "B_BIZ_PARTNER.BP_TYPE In (" & FilterVar("S", "''", "S") & " ," & FilterVar("CS", "''", "S") & ") And B_BIZ_PARTNER.usage_flag=" & FilterVar("Y", "''", "S") & " "
			arrParam(4) = arrParam(4) & " AND B_BIZ_PARTNER.BP_CD = B_BIZ_PARTNER_FTN.PARTNER_BP_CD  AND B_BIZ_PARTNER_FTN.BP_CD = " 				<%' Where Condition%>
            arrParam(4) = arrParam(4) & FilterVar(Trim(frm1.txtSpplCd.Value), "''", "S") & " AND  B_BIZ_PARTNER_FTN.PARTNER_FTN = " & FilterVar("MBI", "''", "S") & " "
			arrParam(5) = "세금계산서발행처"											' TextBox 명칭 
	End Select
	
    If OpenCommPopup(arrHeader, arrField, arrParam, arrRet) Then
		Select Case BpType
			Case "1"   '공급처 
				frm1.txtSpplCd.Value = arrRet(0) : frm1.txtSpplNm.Value = arrRet(1)
				frm1.txtSpplCd.focus
			Case "2"   '지급처 
				frm1.txtPayeeCd.Value = arrRet(0) : frm1.txtPayeeNm.Value = arrRet(1)
				frm1.txtPayeeCd.focus
			Case "3"   '세금계산서발행처 
				frm1.txtBuildCd.Value = arrRet(0) : frm1.txtBuildNm.Value = arrRet(1) ': frm1.txtSpplRegNo.Value = arrRet(2)				
		        Call GetTaxBizArea("BP")
		        frm1.txtBuildCd.focus
		End Select 
		Call ChangeSupplier(BpType)
    End If
    lblnWinEvent = False
    Set gActiveElement = document.activeElement
End Function
'------------------------------------------  OpenIvNo()  -------------------------------------------------
Function OpenIvNo()
	Dim strRet
	Dim arrParam(0)
	Dim iCalledAspName
		
	If lblnWinEvent = True Or UCase(frm1.txtIvNo.className) = UCase(parent.UCN_PROTECTED) Then Exit Function
		
	lblnWinEvent = True
	arrParam(0) = ivType
		
	iCalledAspName = AskPRAspName("m5111pa1_KO441")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "m5111pa1_KO441", "X")
		lgIsOpenPop = False
		Exit Function
	End If
	
	strRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	lblnWinEvent = False
	
	If strRet(0) = "" Then
		frm1.txtIvNo.focus
		Exit Function
	Else
		frm1.txtIvNo.value = strRet(0)
		frm1.txtIvNo.focus	
		Set gActiveElement = document.activeElement
	End If	
		
End Function

'------------------------------------------  OpenIvType()  -------------------------------------------------
Function OpenIvType()
	Dim arrRet
	Dim arrHeader(6), arrField(6), arrParam(5) 
	If lblnWinEvent = True Or UCase(frm1.txtIvTypeCd.className) = UCase(parent.UCN_PROTECTED) Then Exit Function

	lblnWinEvent = True
	
	arrHeader(0) = "매입형태"						' Header명(0)
    arrHeader(1) = "매입형태명"						' Header명(1)
    
    arrField(0) = "IV_TYPE_CD"							' Field명(0)
    arrField(1) = "IV_TYPE_NM"							' Field명(1)
    
	arrParam(0) = "매입형태"						' 팝업 명칭 
	arrParam(1) = "M_IV_TYPE"							' TABLE 명칭 
	arrParam(2) = Trim(frm1.txtIvTypeCd.Value)			' Code Condition
	'arrParam(2) = Trim(frm1.txtIvTypeCd.Value)			' Code Condition
	'arrParam(3) = Trim(frm1.txtIvTypeNm.Value)			' Name Cindition
	'====== 2005.06.22 재고반영여부 관련 수정 ========
	arrParam(4) = "USAGE_FLG=" & FilterVar("Y", "''", "S") & "  and import_flg=" & FilterVar("N", "''", "S") & " and except_flg= " & FilterVar("Y", "''", "S") & " and stock_coverge_flg= " & FilterVar("Y", "''", "S") & " "						' Where Condition
	'====== 2005.06.22 재고반영여부 관련 수정 ========
	arrParam(5) = "매입형태"						' TextBox 명칭 
	
    If OpenCommPopup(arrHeader, arrField, arrParam, arrRet) then
		frm1.txtIvTypeCd.Value = arrRet(0)
		frm1.txtIvTypeNm.Value = arrRet(1)
    end if
    lblnWinEvent = False
    frm1.txtIvTypeCd.focus
	Set gActiveElement = document.activeElement
End Function

'------------------------------------------  OpenCommPopup()  -------------------------------------------------
Function OpenCommPopup(arrHeader, arrField, arrParam, arrRet)


	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	If arrRet(0) = "" Then
		OpenCommPopup = False
	Else
		OpenCommPopup = True
		lgBlnFlgChgValue = True
	End If
	
End Function





'------------------------------------------  OpenPayMeth()  -------------------------------------------------
Function OpenPayMeth()
	Dim arrRet
	Dim arrHeader(6), arrField(6), arrParam(5) 
	
	If lblnWinEvent = True Or UCase(frm1.txtPayTermCd.className) = UCase(parent.UCN_PROTECTED) Then Exit Function
	lblnWinEvent = True
	arrHeader(0) = "결제방법"						        ' Header명(0)
    arrHeader(1) = "결제방법명"						        ' Header명(1)
    arrHeader(2) = "Reference"
    
    arrField(0) = "B_Minor.MINOR_CD"							' Field명(0)
    arrField(1) = "B_Minor.MINOR_NM"							' Field명(1)
    arrField(2) = "b_configuration.REFERENCE"
    
	arrParam(0) = "결제방법"						        ' 팝업 명칭 
	arrParam(1) = "B_Minor,b_configuration"				        ' TABLE 명칭 
	arrParam(2) = Trim(frm1.txtPayTermCd.Value)			        ' Code Condition
	'arrParam(2) = Trim(frm1.txtPayTermCd.Value)			        ' Code Condition
	'arrParam(3) = Trim(frm1.txtPayTermNm.Value)			    ' Name Cindition
	arrParam(4) = "B_Minor.Major_Cd=" & FilterVar("B9004", "''", "S") & " and B_Minor.minor_cd =b_configuration.minor_cd and " & _
	              " b_configuration.SEQ_NO=1 AND b_configuration.major_cd= B_Minor.Major_Cd"	 
	arrParam(5) = "결제방법"						        ' TextBox 명칭 
	
    If OpenCommPopup(arrHeader, arrField, arrParam, arrRet) then
		frm1.txtPayTermCd.Value = arrRet(0) : frm1.txtPayTermNm.Value = arrRet(1)
		Call changePayMeth()
    end if
    lblnWinEvent = False
    frm1.txtPayTermCd.focus
    Set gActiveElement = document.activeElement
End Function

'========================================================================================
' Function Name : changePayMeth
'========================================================================================
Sub changePayMeth()
	
	frm1.txtPayTypeCd.Value = ""
	frm1.txtPayTypeNm.Value = ""
	frm1.txtPayDur.Text = 0	

End Sub


'------------------------------------------  SetReqRef()  -------------------------------------------------
'	Name : SetReqRef()
'	Description :구매요청참조 
'---------------------------------------------------------------------------------------------------------
Function SetReqRef(strRet)

	Dim Index1,Index3,Count1
	Dim IntIflg
	Dim strMessage
	Dim intstartRow,intEndRow, TempRow
	Dim iInsRow,intInsertRowsCount
	
	Const C_IvNo_Req			  = 0
    Const C_IvSeq_Req             = 1
    Const C_PlantCd_Req           = 2
    Const C_PlantNm_Req           = 3
    Const C_ItemCd_Req            = 4
    Const C_ItemNm_Req            = 5
    Const C_Spec_Req              = 6
    Const C_IvQty_Req             = 7
    
    Const C_IvPrc_Req             = 9
    
    Const C_Amt_Req               = 10
    
    Const C_VatYn_Req             = 12
    Const C_VatFlg_Req            = 11
    Const C_VatAmt_Req            = 13
    
    Const C_LocAmt_Req			  = 14
    Const C_LocVatAmt_Req		  = 15
    
    Const C_PoNo_Req              = 16
    Const C_PoSeq_Req             = 17
    
	Const	C_ItemAcct_Req          =	18
	Const	C_VatType_Req           =	19
	Const	C_VatRt_Req             =	20
	Const	C_TrackingNo_Req            =	21
	Const	C_IvCostCd_Req          =	22
	Const	C_IvBizArea_Req         =	23
	Const	C_MvmtQty_Req           =	24
	Const	C_MvmtFlg_Req           =	25
	Const	C_RefIvNo_Req           =	26
	Const	C_RefIvSeqNo_Req		=	27   
	
	Const	C_ItemUnit_Req				=	8 
    
    
    Count1 = Ubound(strRet,1)
	
	strMessage = ""
	
	IntIflg=true
    
    '	frm1.txtIvTypeCd.value		= strRet(Count1,0)
'	frm1.txtIvTypeNm.value		= strRet(Count1,8)
	frm1.txtSpplCd.value		= strRet(Count1,1)
	frm1.txtSpplNm.value		= strRet(Count1,9)
	frm1.txtPayeeCd.value		= strRet(Count1,3)
	frm1.txtPayeeNm.value		= strRet(Count1,11)
	frm1.txtVatCd.value			= strRet(Count1,5)
	frm1.txtVatNm.value			= strRet(Count1,12)
	frm1.txtVatrt.value			= strRet(Count1,15)
	frm1.txtGrpCd.value			= strRet(Count1,6)
	frm1.txtGrpNm.value			= strRet(Count1,13)
	frm1.txtBuildCd.value		= strRet(Count1,2)
	frm1.txtBuildNm.value		= strRet(Count1,10)
	frm1.txtCur.value			= strRet(Count1,4)
	frm1.txtBizAreaCd.value		= strRet(Count1,7)
	frm1.txtBizAreaNm.value		= strRet(Count1,14)
	frm1.txtPayTermCd.value		= strRet(Count1,16)
	frm1.txtPayTermNm.value		= strRet(Count1,17)
	
	frm1.txtSpplRegNo.value		= strRet(Count1,18)
	frm1.txtSpplIvNo.value		= strRet(Count1,19)
	frm1.txtPayDur.value		= strRet(Count1,20)
	frm1.txtPayTypeCd.value		= strRet(Count1,21)
	frm1.txtPayTypeNm.value		= strRet(Count1,22)
	frm1.txtPayTermstxt.value	= strRet(Count1,23)
	frm1.txtRemark.value		= strRet(Count1,24)
	
	with frm1
	
		
	Call changeTabs(TAB1)	
	gSelframeFlg = TAB1
		
	.vspdData.focus
	ggoSpread.Source = .vspdData
	intStartRow = .vspdData.MaxRows + 1
	.vspdData.Redraw = False
	
	TempRow = .vspdData.MaxRows					'리스트 max값 
			
	intInsertRowsCount = 0 '중복 안될때만 MAXROW에 1을 추가하기 위한변수 
	
	'중복된 요청건참조시 MAXROW값계산 로직 수정 200308
	for index1 = 0 to Count1 - 1
	
		.vspdData.Row=Index1+1
		
		If TempRow <> 0 Then
			For Index3 = 1 to TempRow
				if GetSpreadText(.vspdData,C_IvNohdn,index3,"X","X") = strRet(index1,C_IvNo_Req) And _
					GetSpreadText(.vspdData,C_IvSeqhdn,index3,"X","X") = strRet(index1,C_IvSeq_Req) then
					strMessage = strMessage & strRet(Index1,C_IvNo_Req&","&C_IvSeq_Req) & ";"
					intIflg=False					
					intInsertRowsCount = 0		'중복될땐 MAXROW를 증가시키지 않음.					
					Exit for
				Else 
					intInsertRowsCount =  1
				End if
			Next
		Else 		
			intInsertRowsCount =  1				
		End If
		
		if IntIflg <> False then
			lgReqRefChk = true

			.vspdData.MaxRows = CLng(TempRow) + CLng(intInsertRowsCount) 
			iInsRow = CLng(TempRow) + CLng(intInsertRowsCount) 
			
			TempRow = CLng(TempRow) + CLng(intInsertRowsCount) '다음 MAXROW계산시 베이스가 될 TempRow 를 증가시킴.
			lgBlnFlgChgValue = True
			
			Call .vspdData.SetText(0		,	iInsRow, ggoSpread.InsertFlag)

			Call SetState("R",iInsRow)
			
			Call .vspdData.SetText(C_PlantCd	,	iInsRow, strRet(index1,C_PlantCd_Req))
			Call .vspdData.SetText(C_PlantNm	,	iInsRow, strRet(index1,C_PlantNm_Req))
			Call .vspdData.SetText(C_ItemCd		,	iInsRow, strRet(index1,C_ItemCd_Req))
			Call .vspdData.SetText(C_ItemNm		,	iInsRow, strRet(index1,C_ItemNm_Req))
			Call .vspdData.SetText(C_Spec		,	iInsRow, strRet(index1,C_Spec_Req))
			Call .vspdData.SetText(C_IvQty		,	iInsRow, strRet(index1,C_IvQty_Req))
			Call .vspdData.SetText(C_ItemUnit		,	iInsRow, strRet(index1,C_ItemUnit_Req))
			
			Call .vspdData.SetText(C_IvPrc		,	iInsRow, strRet(index1,C_IvPrc_Req))
			
			Call .vspdData.SetText(C_Amt		,	iInsRow, strRet(index1,C_Amt_Req))  '금액 

			Call .vspdData.SetText(C_VatYn		,	iInsRow, strRet(index1,C_VatYn_Req))
			Call .vspdData.SetText(C_VatFlg		,	iInsRow, strRet(index1,C_VatFlg_Req))
			Call .vspdData.SetText(C_VatAmt		,	iInsRow, strRet(index1,C_VatAmt_Req))
			Call .vspdData.SetText(C_LocAmt		,	iInsRow, strRet(index1,C_LocAmt_Req))
			Call .vspdData.SetText(C_LocVatAmt	,	iInsRow, strRet(index1,C_LocVatAmt_Req))
			
			Call .vspdData.SetText(C_PoNo		,	iInsRow, strRet(index1,C_PoNo_Req))
			Call .vspdData.SetText(C_PoSeq		,	iInsRow, strRet(index1,C_PoSeq_Req))
			Call .vspdData.SetText(C_IvNohdn	,	iInsRow, strRet(index1,C_IvNo_Req))
			Call .vspdData.SetText(C_IvSeqhdn	,	iInsRow, strRet(index1,C_IvSeq_Req))
			
			Call .vspdData.SetText(C_ItemAcct  	,	iInsRow, strRet(index1,C_ItemAcct_Req))  
			Call .vspdData.SetText(C_VatType   	,	iInsRow, strRet(index1,C_VatType_Req))   
			Call .vspdData.SetText(C_VatRt     	,	iInsRow, strRet(index1,C_VatRt_Req))     
			Call .vspdData.SetText(C_TrackingNo    	,	iInsRow, strRet(index1,C_TrackingNo_Req))    
			Call .vspdData.SetText(C_IvCostCd  	,	iInsRow, strRet(index1,C_IvCostCd_Req))  
			Call .vspdData.SetText(C_IvBizArea 	,	iInsRow, strRet(index1,C_IvBizArea_Req)) 
			Call .vspdData.SetText(C_MvmtQty   	,	iInsRow, strRet(index1,C_MvmtQty_Req))   
			Call .vspdData.SetText(C_MvmtFlg   	,	iInsRow, strRet(index1,C_MvmtFlg_Req))   
			Call .vspdData.SetText(C_RefIvNo   	,	iInsRow, strRet(index1,C_RefIvNo_Req))   
			Call .vspdData.SetText(C_RefIvSeqNo	,	iInsRow, strRet(index1,C_RefIvSeqNo_Req))
			
			
			
			
			'조정금액 Setting
			Call .vspdData.SetText(C_CtlQty	,	iInsRow, 0)
			Call .vspdData.SetText(C_CtlPrc	,	iInsRow, 0)
			Call .vspdData.SetText(C_CtlAmt	,	iInsRow, 0)
			Call .vspdData.SetText(C_VatCtlAmt,	iInsRow, 0)
			Call .vspdData.SetText(C_CtlLocAmt,	iInsRow, 0)
			Call .vspdData.SetText(C_CtlLocVatAmt,	iInsRow, 0)
			Call .vspdData.SetText(C_NetAmt,	iInsRow, 0)
			Call .vspdData.SetText(C_NetLocAmt,	iInsRow, 0)
			
			'OLD 조정금액 Setting
			Call .vspdData.SetText(C_CtlQty_Old	,	iInsRow, 0)
			Call .vspdData.SetText(C_CtlPrc_Old	,	iInsRow, 0)
			Call .vspdData.SetText(C_CtlAmt_Old	,	iInsRow, 0)
			Call .vspdData.SetText(C_VatCtlAmt_Old,	iInsRow, 0)
			Call .vspdData.SetText(C_CtlLocAmt_Old,	iInsRow, 0)
			Call .vspdData.SetText(C_CtlLocVatAmt_Old,	iInsRow, 0)
			Call .vspdData.SetText(C_NetAmt_Old,	iInsRow, 0)
			Call .vspdData.SetText(C_NetLocAmt_Old,	iInsRow, 0)
			
			.vspddata.row = index1+1

		Else
			IntIFlg=True
		End if 
	next
	
	
	intEndRow = iInsRow

	Call ReFormatSpreadCellByCellByCurrency2(.vspdData,1,.vspdData.MaxRows,.txtCur.value, C_IvPrc,   "C" ,"I","X","X")  ' 매입단가 
	Call ReFormatSpreadCellByCellByCurrency2(.vspdData,1,.vspdData.MaxRows,.txtCur.value, C_CtlPrc,   "C" ,"I","X","X") ' 조정단가 
	Call ReFormatSpreadCellByCellByCurrency2(.vspdData,1,.vspdData.MaxRows,.txtCur.value, C_Amt,   "A" ,"I","X","X")  ' 금액 
	Call ReFormatSpreadCellByCellByCurrency2(.vspdData,1,.vspdData.MaxRows,.txtCur.value, C_CtlAmt,   "A" ,"I","X","X")  ' 조정금액 
	Call ReFormatSpreadCellByCellByCurrency2(.vspdData,1,.vspdData.MaxRows,.txtCur.value, C_NetAmt,   "A" ,"I","X","X")  ' 조정순금액 
	Call ReFormatSpreadCellByCellByCurrency2(.vspdData,1,.vspdData.MaxRows,.txtCur.value, C_VatAmt,   "A" ,"I","X","X") ' vat금액 
	Call ReFormatSpreadCellByCellByCurrency2(.vspdData,1,.vspdData.MaxRows,.txtCur.value, C_VatCtlAmt,   "A" ,"I","X","X") ' vat조정금액 
	Call ReFormatSpreadCellByCellByCurrency2(.vspdData,1,.vspdData.MaxRows,parent.gCurrency, C_LocAmt,   "A" ,"I","X","X") ' 매입자국금액 
	Call ReFormatSpreadCellByCellByCurrency2(.vspdData,1,.vspdData.MaxRows,parent.gCurrency, C_CtlLocAmt,   "A" ,"I","X","X") ' 조정자국금액 
	Call ReFormatSpreadCellByCellByCurrency2(.vspdData,1,.vspdData.MaxRows,parent.gCurrency, C_NetLocAmt,   "A" ,"I","X","X") ' 조정자국순금액 
	Call ReFormatSpreadCellByCellByCurrency2(.vspdData,1,.vspdData.MaxRows,parent.gCurrency, C_LocVatAmt,   "A" ,"I","X","X") ' VAT자국금액 
	Call ReFormatSpreadCellByCellByCurrency2(.vspdData,1,.vspdData.MaxRows,parent.gCurrency, C_CtlLocVatAmt,   "A" ,"I","X","X") ' VAT조정자국금액 
	
	
	
	Call ChangeCurr()
	
	Call ggoOper.LockField(Document, "Q")
	
	if strMessage <> "" then
		Call DisplayMsgBox("17a005", "X",strmessage,"구매요청번호")
		.vspdData.ReDraw = True
		Exit Function
	End if
		
	Call SetSpreadLock

	Call BtnToolCtrl(TAB1)
	
	.vspdData.ReDraw = True
	 
     End with
End Function

'==========================================  2.4.3 Set???()  =============================================
'	Name : Set???()
'	Description : PopUp에서 전송된 값을 특정 Tag Object에 지정 
'=========================================================================================================
Function SetVatName()	
	Dim index1
	with frm1

		For Index1 = 1 to .vspdData.MaxRows step 1
				'Insert Row 시 헤더의 부가세관련 정보 초기값으로 2002.2.19
				.vspdData.Row = index1
		
				.vspdData.Col = C_VatType
				.vspdData.Text = .hdntxtVatCd.value
		
				.vspdData.Col  = C_VatNm
				.vspdData.Text = .hdntxtVatNm.value
		
				.vspdData.Col  = C_VatRate
				.vspdData.Text = .hdntxtVatrt.value
		Next
	End With 
	'lgReqRefChk = False
		
End Function


'==========================================  2.4.3 Set???()  =============================================
'	Name : Set???()
'	Description : PopUp에서 전송된 값을 특정 Tag Object에 지정 
'=========================================================================================================
Function SetVat(byval arrRet)	
	
    Dim price, chk_vat_flg

    With frm1
		.vspdData.Col = C_VatType
		.vspdData.Text = arrRet(0)
		.vspdData.Col = C_VatNm
		.vspdData.Text = arrRet(1)
		.vspdData.Col = C_VatRate
		.vspdData.Text = arrRet(2)
		
		.vspdData.Col = C_OrderAmt
		price = UNICDbl(.vspdData.Text)
'	vat 금액계산 
' 부가세 포함/불포함 부가세 계산 변경 2002.3.9 L.I.P
		.vspdData.Col		= C_IOFlgCd
		chk_vat_flg	= .vspdData.text
		
		.vspdData.Col = C_VatAmt 
		if chk_vat_flg = "2"		Then
			.vspdData.Text = UNIConvNumPCToCompanyByCurrency(price * UNICDbl(arrRet(2))/(100 + UNICDbl(arrRet(2))),frm1.txtPayeeCd.value,Parent.ggAmtOfMoneyNo,Parent.gTaxRndPolicyNo,"X")
		Else
			.vspdData.Text = UNIConvNumPCToCompanyByCurrency(price * UNICDbl(arrRet(2))/(100 + UNICDbl(arrRet(2))),frm1.txtPayeeCd.value,Parent.ggAmtOfMoneyNo,Parent.gTaxRndPolicyNo,"X")
		End If

	End With
    Call vspdData_Change(C_VatType, frm1.vspdData.ActiveRow)   
	
End Function


'======================================   GetTaxBizArea()  =====================================
Sub GetTaxBizArea(Byval strFlag)
   	Dim strSelectList, strFromList, strWhereList
	Dim strBilltoParty, strSalesGrp, strTaxBizArea
	Dim strRs
	Dim arrTaxBizArea(2), arrTemp
	
    
	If strFlag = "NM" Then                              '세금신고사업장 변경시 이름값만 가져온다 
		strTaxBizArea = frm1.txtBizAreaCd.value
	Else
		strBilltoParty = frm1.txtBuildCd.value          '계산서 발행처 
		'strSalesGrp    = frm1.txtGrpCd.value            '구매그룹 
		
		<%'발행처와 영업 그룹이 모두 등록되어 있는 경우 종합코드에 설정된 rule을 따른다 %>
		If Len(strBillToParty) > 0 And Len(strSalesGrp) > 0	Then strFlag = "*"
	End if
		
	strSelectList = " * "
	strFromList = " dbo.ufn_m_GetTaxBizArea ( " & FilterVar(strBilltoParty, "''", "S") & " ,  " & FilterVar(strSalesGrp, "''", "S") & " ,  " & FilterVar(strTaxBizArea, "''", "S") & " ,  " & FilterVar(strFlag, "''", "S") & " ) "
	strWhereList = ""
	
	Err.Clear
    
	If CommonQueryRs2by2(strSelectList,strFromList,strWhereList,strRs) Then
		arrTemp = Split(strRs, Chr(11))
		frm1.txtBizAreaCd.value = arrTemp(1)
		frm1.txtBizAreaNm.value = arrTemp(2)
	Else
		If Err.number <> 0 Then
			MsgBox Err.description 
			Err.Clear 
			Exit Sub
		End If

		' 세금 신고 사업장을 Editing한 경우 
		'If strFlag = "NM" Then
		'	If Not OpenBillHdr(3) Then
				frm1.txtBizAreaCd.value = ""
				frm1.txtBizAreaNm.value = ""
		'	End if
		'End if
	End if
End Sub
'========================================================================================
' Function Name : SetRetCd
' Function Desc : 반납유형 직접 입력시 처리 
'========================================================================================
Sub SetRetCd()
	Dim iRetCd, iRetNm, strQUERY, tmpData
	Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6, i

	with frm1.vspdData

		Err.Clear
    
	   .Col = C_RetCd

		strQUERY = " Minor.MAJOR_CD='B9017' and  Minor.MINOR_CD = " & "'" & FilterVar(Trim( .text), " " , "SNM") & "' "
    
		Call CommonQueryRs(" Minor.MINOR_CD,  Minor.MINOR_NM ", " B_MINOR Minor ", strQUERY, lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
		If Err.number = 0 Then
			
			if lgF0 <> "" then
				iRetNm = Split(lgF1, Chr(11))
			   .Col = C_RetNm  
			   .text = iRetNm(0)
			  else
			   .Col = C_RetNm  
			   .text = ""
			end if
		else
			MsgBox Err.description, VbInformation, parent.gLogoName
			Err.Clear 
			Exit Sub
		End If
     
	End With
	   
End Sub

Function OpenMpOrderRef()

	Dim strRet
	Dim strParam
	
	if frm1.rdoRelease(1).checked = true then
		Call DisplayMsgBox("17a008", "X", "X", "X")
		Exit Function
	End if
	
	If IsOpenPop = True Then Exit Function
		
	IsOpenPop = True
	
	strParam = Parent.gColSep'strParam & Trim(frm1.txtSold_to_party.value) & Parent.gColSep
	strParam = strParam & Trim(frm1.txtIvDt.Text)

	strRet = window.showModalDialog("m3011ra1.asp", strParam, _
			"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	
	If strRet(0) = "" Then
		Exit Function
	Else
		Call SetMpOrder(strRet)
	End If	
		
End Function

Function SetMpOrder(strRet)
	
	frm1.txtMaintNo.value = strRet(0)
	'frm1.RefOnLine.value = "Y"

	Call OnLineQuery()

	lgBlnFlgChgValue = true

End Function
'========================================== 2.4.2 Open???()  =============================================
'	Name : Open???()
'	Description : 중복되어 있는 PopUp을 재정의, 재정의가 필요한 경우는 반드시 CommonPopUp.vbs 와 
'				  ManufactPopUp.vbs 에서 Copy하여 재정의한다.
'=========================================================================================================
'++++++++++++++++  Insert Your Code for PopUp(Open)  +++++++++++++++++++++++++++++++++++++++++++++++++++


Function SetPoNo(strRet)
	frm1.txtIvNo.value = strRet
End Function


'------------------------------------------  OpenCur()  -------------------------------------------------
'	Name : OpenCur()
'	Description : OpenCurr PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenCur()
	Dim arrRet
	Dim arrHeader(6), arrField(6), arrParam(5) 

	If lblnWinEvent = True Or UCase(frm1.txtCur.className) = UCase(parent.UCN_PROTECTED) Then Exit Function
	lblnWinEvent = True
	arrHeader(0) = "화폐"						' Header명(0)
    arrHeader(1) = "화폐명"						' Header명(1)
    
    arrField(0) = "Currency"						' Field명(0)
    arrField(1) = "Currency_Desc"					' Field명(1)
    
	arrParam(0) = "화폐"						' 팝업 명칭 
	arrParam(1) = "B_Currency"						' TABLE 명칭 
	arrParam(2) = Trim(frm1.txtCur.Value)			' Code Condition
	'arrParam(2) = Trim(frm1.txtCur.Value)			' Code Condition
	arrParam(3) = ""								' Name Cindition
	arrParam(4) = ""								' Where Condition
	arrParam(5) = "화폐"						' TextBox 명칭 
	
    If OpenCommPopup(arrHeader, arrField, arrParam, arrRet) Then
		frm1.txtCur.Value 	= arrRet(0)
		frm1.txtCurNm.Value = arrRet(1)
		Call ChangeCurr()
    End If
	lblnWinEvent = False
	frm1.txtCur.focus
	Set gActiveElement = document.activeElement
End Function

Function SetCurr(byval arrRet)
	frm1.txtPayeeCd.Value    = arrRet(0)		
	frm1.txtPayerNm.Value  = arrRet(1)		
	Call ChangeCurr()
	lgBlnFlgChgValue = True
End Function




Function OpenVat(byVal chk)

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtVatCd.className) = Ucase(Parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "VAT형태"				
	arrParam(1) = "B_MINOR,b_configuration"	
	
	arrParam(2) = Trim(frm1.txtVatCd.Value)
		
	arrParam(4) = "b_minor.MAJOR_CD='b9001' and b_minor.minor_cd=b_configuration.minor_cd "	
	arrParam(4) = arrParam(4) & "and b_minor.major_cd=b_configuration.major_cd and b_configuration.SEQ_NO=1"
	arrParam(5) = "VAT형태"					
	
    arrField(0) = "b_minor.MINOR_CD"			
    arrField(1) = "b_minor.MINOR_NM"
    arrField(2) = "b_configuration.REFERENCE"	
    
    arrHeader(0) = "VAT형태"					
    arrHeader(1) = "VAT형태명"				
    arrHeader(2) = "VAT율"
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		if chk = 1 then
			Call SetVat_H(arrRet)
		Else
			Call SetVat(arrRet)
		End if
	End If	
	
	Call SetVatAmt1()
		
End Function

' Vat 조정금액 , Vat 조정자국금액 Setting
Function SetVatAmt1()

	Dim vat , vat1 , vat2 , xChRt , VatRt , DocAmt , locvat 
	Dim VatFlg
	Dim index 

	If frm1.vspddata.maxrows < 1 then Exit Function
	
	xChRt = UNICDbl(Trim(frm1.txtXchRt.value))
	VatRt = UNICDbl(Trim(frm1.txtvatRt.value))
	

	With frm1.vspddata
	
		For index =	1	to .MaxRows
		
			.Row = index
			.Col = C_Stateflg
			
			ChangeVatAmt(index)
			
			ChangeNetAmt(index)
			
			ChangeVatLocAmt(index)
			
			ChangeNetLocAmt(index)
			
			
			If Trim(.Text) = "Q" Then
			
				ggoSpread.UpdateRow index
				
			Elseif  Trim(.Text) = "R" Then
			
				ggoSpread.InsertRow index
				
			End if 
		
		Next 

	HSumAmtNewCalc()		'Header Setting
	
	End With

End Function



Function SetVat_H(byval arrRet)
	frm1.txtVatCd.Value		 = arrRet(0)		
	frm1.txtVatNm.Value      = arrRet(1)		
	frm1.txtVatRt.Value = UNIFormatNumber(UNICDbl(arrRet(2)), ggExchRate.DecPoint, -2, 0, ggExchRate.RndPolicy, ggExchRate.RndUnit)
	
	lgBlnFlgChgValue = True
End Function


'------------------------------------------  OpenBizArea()  -------------------------------------------------
Function OpenBizArea()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lblnWinEvent = True Or UCase(frm1.txtBizAreaCd.className) = UCase(parent.UCN_PROTECTED) Then Exit Function

	lblnWinEvent = True

	arrParam(0) = "세금신고사업장"	
	arrParam(1) = "B_TAX_BIZ_AREA"
	
	arrParam(2) = Trim(frm1.txtBizAreaCd.Value)
	'arrParam(2) = Trim(frm1.txtBizAreaCd.Value)
	
	'arrParam(4) = "Tax_Flag = 'Y'"
	arrParam(4) = ""
	arrParam(5) = "세금신고사업장"			
	
    arrField(0) = "TAX_BIZ_AREA_CD"
    arrField(1) = "TAX_BIZ_AREA_NM"
    
    arrHeader(0) = "세금신고사업장"
    arrHeader(1) = "세금신고사업장명"
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	lblnWinEvent = False
	
	If arrRet(0) = "" Then
		frm1.txtBizAreaCd.focus
		Exit Function
	Else
		frm1.txtBizAreaCd.Value = arrRet(0)
		frm1.txtBizAreaNm.Value = arrRet(1)
		lgBlnFlgChgValue = True
	End If	
	frm1.txtBizAreaCd.focus
	Set gActiveElement = document.activeElement
End Function

'------------------------------------------  OpenGrp()  -------------------------------------------------
Function OpenGrp()
	Dim arrRet
	Dim arrHeader(6), arrField(6), arrParam(5) 
	If lblnWinEvent = True Or UCase(frm1.txtGrpCd.className) = UCase(parent.UCN_PROTECTED) Then Exit Function    
	lblnWinEvent = True	
	arrHeader(0) = "구매그룹"									' Header명(0)
    arrHeader(1) = "구매그룹명"									' Header명(1)
    
    arrField(0) = "PUR_GRP"											' Field명(0)
    arrField(1) = "PUR_GRP_NM"										' Field명(1)
    
	arrParam(0) = "구매그룹"									' 팝업 명칭 
	arrParam(1) = "B_PUR_GRP"										' TABLE 명칭 
	arrParam(2) = Trim(frm1.txtGrpCd.Value)							' Code Condition
	
																	' Where Condition
	arrParam(4) = "USAGE_FLG = " & FilterVar("Y", "''", "S") & " "
	arrParam(5) = "구매그룹"									' TextBox 명칭 
	
    If OpenCommPopup(arrHeader, arrField, arrParam, arrRet) then
		frm1.txtGrpCd.Value = arrRet(0)
		frm1.txtGrpNm.Value = arrRet(1)  
    end if
    Call GetTaxBizArea("*")
    lblnWinEvent = False
    frm1.txtGrpCd.focus
    Set gActiveElement = document.activeElement
End Function

Function OpenPayType()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	if Trim(frm1.txtPayTermCd.Value) = "" then
		Call DisplayMsgBox("17a002", Parent.VB_YES_NO,"결제방법", "X")
		Exit Function
	End if

	If IsOpenPop = True Or UCase(frm1.txtPayTypeCd.className) = Ucase(Parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "지급유형"				
	arrParam(1) = "B_MINOR,B_CONFIGURATION," _
	& "(Select REFERENCE From B_CONFIGURATION Where MAJOR_CD ='B9004'"_
		& "And MINOR_CD='" & FilterVar(Trim(frm1.txtPayTermCd.value), "", "SNM") & "' And SEQ_NO>=2)C"
	
	arrParam(2) = Trim(frm1.txtPayTypeCd.Value)
	
	arrParam(4) = "B_MINOR.MINOR_CD = C.REFERENCE And B_CONFIGURATION.MINOR_CD = B_MINOR.MINOR_CD And B_MINOR.MAJOR_CD = 'A1006' " _
				& "AND B_CONFIGURATION.REFERENCE IN('RP','P')"	
	arrParam(5) ="지급유형"					
	
	arrField(0) = "B_MINOR.MINOR_CD"						
	arrField(1) = "B_MINOR.MINOR_NM"				
        
    arrHeader(0) = "지급유형"				
    arrHeader(1) = "지급유형명"				
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		frm1.txtPayTypeCd.Value = arrRet(0)
		frm1.txtPayTypeNm.Value = arrRet(1)
		lgBlnFlgChgValue 		= True
	End If	
End Function

'전표조회 클릭시 호출 
'------------------------------------------  OpenGLRef()  ----------------------------------------------
Function OpenGLRef()
	Dim strRet
	Dim arrParam(1)
	Dim iCalledAspName
	
	If lblnWinEvent = True Then Exit Function
		
	lblnWinEvent = True
	
	arrParam(0) = Trim(frm1.hdnGlNo.value)

   If frm1.hdnGlType.Value = "A" Then               '회계전표팝업 
   		iCalledAspName = AskPRAspName("a5120ra1")
		If Trim(iCalledAspName) = "" Then
			IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a5120ra1" ,"x")
			IsOpenPop = False
			Exit Function
		End If
	   strRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam), _
			"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
    Elseif frm1.hdnGlType.Value = "T" Then          '결의전표팝업 
		iCalledAspName = AskPRAspName("a5130ra1")

		If Trim(iCalledAspName) = "" Then
			IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a5130ra1" ,"x")
			IsOpenPop = False
			Exit Function
		End If
	   strRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam), _
			"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
    Elseif frm1.hdnGlType.Value = "B" Then
     	Call DisplayMsgBox("205154","X" , "X","X")   '아직 전표가 생성되지 않았습니다. 
    End if

	lblnWinEvent = False
	
End Function

'============================================================================================================
' Name : SubGetGlNo
' Desc : Get Gl_no : 2003.03 KJH 전표번호 가져오는 로직 수정 
'============================================================================================================
Sub SubGetGlNo()
	Dim lgStrFrom
	Dim strTempGlNo, strGlNo
	Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
	
	On Error Resume Next
	Err.Clear 
	
	lgStrFrom =  " ufn_a_GetGlNo( " & FilterVar(frm1.hdnIvNo.Value, "''", "S") & " )"
	
	Call CommonQueryRs(" TEMP_GL_NO, GL_NO ", lgStrFrom, "", lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
	If lgF0 <> "" then 
		strTempGlNo = Split(lgF0, Chr(11))
		strGlNo		= Split(lgF1, Chr(11))
					
		If strGlNo(0) = "" and strTempGlNo(0) = "" then 
			frm1.hdnGlNo.Value		=	""
			frm1.hdnGlType.value	=	"B"
		Elseif strGlNo(0) = "" and strTempGlNo(0) <> "" then
			frm1.hdnGlNo.Value		=	strTempGlNo(0) 
			frm1.hdnGlType.value	=	"T"
		Elseif strGlNo(0) <> "" then 
			frm1.hdnGlNo.Value		=	strGlNo(0) 
			frm1.hdnGlType.value	=	"A"
		End If
	Else
		frm1.hdnGlNo.Value		=	""
		frm1.hdnGlType.value	=	"B"
	End if
	
End Sub
'발주내역에서 참조한 값을 찍어준다 

'==========================================  2.4.3 Set???()  =============================================
'	Name : Set???()
'	Description : PopUp에서 전송된 값을 특정 Tag Object에 지정 
'=========================================================================================================
'++++++++++++++++  Insert Your Code for PopUp(Open)  ++++++++++++++++++++++++++++++++++++++++++++++++++

'++++++++++++++++++++++++++++++++++++++++++  2.5 개발자 정의 함수  +++++++++++++++++++++++++++++++++++++++
'    개별적 프로그램 마다 필요한 개발자 정의 Procedure (Sub, Function, Validation & Calulation 관련 함수)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Sub SetSpreadFloatLocal(ByVal iCol , ByVal Header , _
                    ByVal dColWidth , ByVal HAlign , _
                    ByVal iFlag )
	     
   Select Case iFlag
        Case 2                                                              '금액 
            ggoSpread.SSSetFloat iCol, Header, dColWidth, Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec, HAlign
        Case 3                                                              '수량 
            ggoSpread.SSSetFloat iCol, Header, dColWidth, Parent.ggQtyNo       ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec, HAlign
        Case 4                                                              '단가 
            ggoSpread.SSSetFloat iCol, Header, dColWidth, Parent.ggUnitCostNo  ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec, HAlign
        Case 5                                                              '환율 
            ggoSpread.SSSetFloat iCol, Header, dColWidth, Parent.ggExchRateNo  ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec, HAlign,,"Z"
        Case 6                                                              '과부족허용율 
            ggoSpread.SSSetFloat iCol, Header, dColWidth, Parent.ggExchRateNo ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,,"0","999.9999"
    End Select
         
End Sub



'===================================== CurFormatNumericOCX()  =======================================
'	Name : CurFormatNumericOCX()
'	Description : 화면에서 일괄적으로 Rounding되는 Numeric OCX
'====================================================================================================
Sub CurFormatNumericOCX()

	With frm1
		'ggoOper.FormatFieldByObjectOfCur .txtIvAmt, .txtPayeeCd.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec
		'ggoOper.FormatFieldByObjectOfCur .txtGrossVatAmt,.txtPayeeCd.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec
		'ggoOper.FormatFieldByObjectOfCur .txtVatAmt, .txtPayeeCd.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec

		'ggoOper.FormatFieldByObjectOfCur .txtDetailNetAmt, .txtPayeeCd.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec
		'ggoOper.FormatFieldByObjectOfCur .txtDetailVatAmt, .txtPayeeCd.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec
		'ggoOper.FormatFieldByObjectOfCur .txtDetailGrossAmt, .txtPayeeCd.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec
		ggoOper.FormatFieldByObjectOfCur .txtXchRt, .txtCur.value, parent.ggExchRateNo, gBCurrency, gBDataType, gBDecimals, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec		
	End With

End Sub


'===================================== CurFormatNumSprSheet()  ======================================
'	Name : CurFormatNumSprSheet()
'	Description : 화면에서 일괄적으로 Rounding되는 Numeric Spread Sheet
'====================================================================================================
Sub CurFormatNumSprSheet()
	
	With frm1

		ggoSpread.Source = frm1.vspdData
		'단가 
		ggoSpread.SSSetFloatByCellOfCur C_Cost,-1, .txtPayeeCd.value, Parent.ggUnitCostNo, gBCurrency, gBDataType, gBDecimals, Parent.gComNum1000, Parent.gComNumDec ,,,"Z"
		'금액 
		ggoSpread.SSSetFloatByCellOfCur C_OrderAmt,-1, .txtPayeeCd.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gComNum1000, Parent.gComNumDec,,,"Z"
		'VAT금액 
		ggoSpread.SSSetFloatByCellOfCur C_VatAmt,-1, .txtPayeeCd.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gComNum1000, Parent.gComNumDec,,,"Z"
        ggoSpread.SSSetFloatByCellOfCur C_NetAmt,-1, .txtPayeeCd.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gComNum1000, Parent.gComNumDec,,,"Z"
	    ggoSpread.SSSetFloatByCellOfCur C_OrgNetAmt,-1, .txtPayeeCd.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gComNum1000, Parent.gComNumDec ,,,"Z"
        ggoSpread.SSSetFloatByCellOfCur C_OrgNetAmt1,-1, .txtPayeeCd.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gComNum1000, Parent.gComNumDec ,,,"Z"
	End With

End Sub	

'========================================================================================
' Function Name : ChangeCurr()
' Function Desc : 
'========================================================================================

Sub ChangeCurr()

	if UCase(Trim(frm1.txtCur.value)) = UCase(parent.gCurrency) then
		frm1.txtXchRt.Text = 1
		Call ggoOper.SetReqAttr(frm1.txtXchRt,"Q")
	else
		frm1.txtXchRt.Text = ""
		Call ggoOper.SetReqAttr(frm1.txtXchRt,"D")
	end if 
	Call CurFormatNumericOCX()
	Call SetAmtByCurAmt()	'환율 변화에 의한 값 변경 
End Sub
'========================================================================================
' Function Name : changePayterm
' Function Desc : 
'========================================================================================
Sub changePayterm()
	
	frm1.txtPayTypeCd.Value = ""
	frm1.txtPayTypeNm.Value = ""
	frm1.txtPayDur.Text = 0	

End Sub


<%
'================================== =====================================================
' Function Name : InitCollectType
' Function Desc : 소비세유형코드/명/율 저장하기 
' 여기부터 키보드에서 소비세유형코드를 변경시 소비세유형명,소비세율,매입금액,NetAmount를 변경시키는 함수 
'========================================================================================
%>
Sub InitCollectType()
	Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6, i
	Dim iCodeArr, iNameArr, iRateArr

    Err.Clear

	Call CommonQueryRs(" Minor.MINOR_CD,  Minor.MINOR_NM, Config.REFERENCE ", " B_MINOR Minor,B_CONFIGURATION Config ", " Minor.MAJOR_CD='B9001' And Config.MAJOR_CD = Minor.MAJOR_CD And Config.MINOR_CD = Minor.MINOR_CD And Config.SEQ_NO = 1 ", lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

    iCodeArr = Split(lgF0, Chr(11))
    iNameArr = Split(lgF1, Chr(11))
    iRateArr = Split(lgF2, Chr(11))

	If Err.number <> 0 Then
		MsgBox Err.description 
		Err.Clear 
		Exit Sub
	End If

	Redim arrCollectVatType(UBound(iCodeArr) - 1, 2)

	For i = 0 to UBound(iCodeArr) - 1
		arrCollectVatType(i, 0) = iCodeArr(i)
		arrCollectVatType(i, 1) = iNameArr(i)
		arrCollectVatType(i, 2) = iRateArr(i)
	Next
End Sub

'========================================================================================
' Function Name : GetCollectTypeRef
' Function Desc : 
'========================================================================================
Sub GetCollectTypeRef(ByVal VatType, ByRef VatTypeNm, ByRef VatRate)

	Dim iCnt

	For iCnt = 0 To Ubound(arrCollectVatType)  
		If arrCollectVatType(iCnt, 0) = UCASE(VatType) Then
			VatTypeNm = arrCollectVatType(iCnt, 1)
			VatRate   = arrCollectVatType(iCnt, 2)
			Exit Sub
		End If
	Next
	VatTypeNm = ""
	VatRate = ""
End Sub


'#########################################################################################################
'												3. Event부 
'	기능: Event 함수에 관한 처리 
'	설명: Window처리, Single처리, Grid처리 작업.
'         여기서 Validation Check, Calcuration 작업이 가능한 Event가 발생.
'         각 Object단위로 Grouping한다.
'##########################################################################################################


'******************************************  3.1 Window 처리  *********************************************
'	Window에 발생 하는 모든 Even 처리	
'*********************************************************************************************************
'========================================================================================
' Function Name : vspdData_Click
' Function Desc : 
'========================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
	gMouseClickStatus = "SPC"   
	Set gActiveSpdSheet = frm1.vspdData
	
	If lgIntFlgMode = Parent.OPMD_UMODE Then
		if Trim(frm1.hdnRelease.Value) = "N" then
			Call SetPopupMenuItemInf("1101111111")
		else
			Call SetPopupMenuItemInf("0000111111")
		end if
	
	Else
		Call SetPopupMenuItemInf("0000111111")
	End If

	If frm1.vspdData.MaxRows = 0 Then Exit Sub
	   
	If Row <= 0 Then
		ggoSpread.Source = frm1.vspdData
		If lgSortKey = 1 Then
			ggoSpread.SSSort Col				'Sort in Ascending
			lgSortkey = 2
		Else
			ggoSpread.SSSort Col, lgSortKey	'Sort in Descending
			lgSortkey = 1
		End If
		
		Exit Sub
	End If    		
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

'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

End Sub

'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
End Sub

'========================================================================================================
'   Event Name : vspdData_DblClick
'   Event Desc : 
'========================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)				
    If Row <= 0 Then
		Exit Sub
	End If
	If frm1.vspddata.MaxRows=0 Then
		Exit Sub
	End If

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
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

'========================================================================================
' Function Name : PopSaveSpreadColumnInf
' Description   : 
'========================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub

'========================================================================================
' Function Name : PopRestoreSpreadColumnInf
' Description   : 
'========================================================================================
Sub PopRestoreSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()      
    Call CurFormatNumSprSheet() 
    Call ggoSpread.ReOrderingSpreadData()
    Call SetSpreadLockAfterQuery()
End Sub

'==========================================================================================
'   Event Name : vspdData_LeaveCell
'   Event Desc : Cell을 벗어날시 무조건발생 이벤트 
'==========================================================================================

'==========================================================================================
'   Event Name : vspdData_Change(변경)
'   Event Desc :
'==========================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )


	Dim chkState , tmpAmt
	Dim CtlQty , IvQty , CtlAmt , Amt , VatAmt , VatCtlAmt , LocAmt , CtlLocAmt , LocVatAmt , CtlLocVatAmt  

		
    ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row 
    
    with frm1.vspdData 		
		.Row = Row
		.Col = 0
		
		if Trim(.Text) = ggoSpread.DeleteFlag  then
		    Exit Sub
		end if    
		
		.Col = C_Stateflg:	.Row = Row
		chkState = .Text
	
		if Trim(.Text) = "" then
			.Text = "U"
		End if


	.Col = Col
	tmpAmt	=	.Text
				
	Select Case Col

	
		Case C_CtlQty

			.Col	= C_IvQty
			IvQty	= UNICDbl(frm1.vspdData.Text)
						
			.Col = C_CtlQty
			If Trim(frm1.vspdData.Text) = "" Or IsNull(frm1.vspdData.Text) Then
				CtlQty = UNICDbl(0)
			Else
				CtlQty = UNICDbl(frm1.vspdData.Text)
			End If
			
			If abs(CtlQty) > IvQty Then
				Call DisplayMsgBox("970025","X","조정수량의절대값","매입수량")	
				.Col	=	C_CtlQty_Old
				tmpAmt	=	.Text
				.Col	=	C_CtlQty
				.Text	=	tmpAmt
				Exit Sub
			End If
			
		Case C_CtlAmt
			
			.Col = C_CtlAmt
			CtlAmt	= UNICDbl(.Text)
			.Col = C_Amt
			Amt		= UNICDbl(.Text)
			
			If CtlAmt < 0 and (CtlAmt*(-1)) > Amt Then
				Call DisplayMsgBox("970025","X","조정금액","금액")	
				.Col	=	C_CtlAmt_Old
				tmpAmt	=	.Text
				.Col	=	C_CtlAmt
				.Text	=	tmpAmt
				Exit Sub
			End If
			
		Case C_CtlLocAmt
			.Col = C_CtlLocAmt
			CtlLocAmt = UNICDbl(.Text)
			.Col = C_LocAmt
			LocAmt =	UNICDbl(.Text)
			
			If CtlLocAmt < 0 and (CtlLocAmt*(-1)) > LocAmt Then
				Call DisplayMsgBox("970025","X","조정자국금액","자국금액")	
				.Col	=	C_CtlLocAmt_Old
				tmpAmt	=	.Text
				.Col	=	C_CtlLocAmt
				.Text	=	tmpAmt
				Exit Sub
			End If
						
			
		Case C_VatCtlAmt
			.Col = C_VatCtlAmt
			VatCtlAmt = UNICDbl(.Text)
			.Col = C_VatAmt
			VatAmt =	UNICDbl(.Text)
			
			If VatCtlAmt < 0 and (VatCtlAmt*(-1)) > VatAmt Then
				Call DisplayMsgBox("970025","X","VAT조정금액","VAT금액")	
				.Col	=	C_VatCtlAmt_Old
				tmpAmt	=	.Text
				.Col	=	C_VatCtlAmt
				.Text	=	tmpAmt
				Exit Sub
			End If
			
			
		Case C_CtlLocVatAmt
			.Col = C_CtlLocVatAmt
			CtlLocVatAmt = UNICDbl(.Text)
			.Col = C_LocVatAmt
			LocVatAmt	 = UNICDbl(.Text)
			
			If CtlLocVatAmt < 0 and (CtlLocVatAmt*(-1)) > LocVatAmt Then
				Call DisplayMsgBox("970025","X","VAT조정자국금액","VAT자국금액")
				.Col	=	C_CtlLocVatAmt_Old
				tmpAmt	=	.Text
				.Col	=	C_CtlLocVatAmt
				.Text	=	tmpAmt	
				Exit Sub
			End If
			
	End select
	
	End With
      
	Frm1.vspdData.Row = Row
	Frm1.vspdData.Col = Col
	
    Call CheckMinNumSpread(frm1.vspdData, Col, Row) 
    
    '이성룡 추가(매입부분 관련하여 값 Setting)
 
	Dim Qty , Price , DocAmt , changeVatflg
    
    changeVatflg = "N"
    
    Select Case col
    
    Case C_CtlQty,C_CtlPrc , C_CtlAmt       '매입수량,매입단가,발주참조,입고참조인경우(C_Cost)= 매입금액 
    
		If col <> C_CtlAmt then
		
			frm1.vspdData.Col = C_CtlQty
		
			If Trim(frm1.vspdData.Text) = "" Or IsNull(frm1.vspdData.Text) Then
			Qty = 0
			Else
				Qty = UNICDbl(frm1.vspdData.Text)
			End If
		
			frm1.vspdData.Col = C_CtlPrc
			If Trim(frm1.vspdData.Text) = "" Or IsNull(frm1.vspdData.Text) Then
				Price = 0
			Else
				Price = UNICDbl(frm1.vspdData.Text)
			End If		
		
			DocAmt = Qty * Price           '(매입수량) * (단가)
			frm1.vspdData.Col = C_CtlAmt   '매입금액 
			frm1.vspdData.Text = UNIConvNumPCToCompanyByCurrency(DocAmt,frm1.txtCur.value,parent.ggAmtOfMoneyNo,"X","X") 
			
		End If
		
				
		Call ChangeVatAmt(Row)	'Vat 금액 
		
		Call ChangeNetAmt(Row)	'순금액 
		
		Call ChangeLocAmt(Row)	'자국금액 
		
		Call ChangeVatLocAmt(Row)	'Vat자국금액 
		
		Call ChangeNetLocAmt(Row)	'자국순금액 
		
		Call HSumAmtNewCalc			'Header 의 Sum 금액 
	
	Case C_VatCtlAmt	
	
		Call ChangeNetAmt(Row)	'순금액 
		
		Call ChangeVatLocAmt(Row)	'Vat자국금액 
		
		Call ChangeNetLocAmt(Row)	'자국순금액 
		
		Call HSumAmtNewCalc			'Header 의 Sum 금액 
		
	Case C_CtlLocVatAmt
	
		Call ChangeNetLocAmt(Row)		'자국순금액 
		
		Call HSumAmtNewCalc			'Header 의 Sum 금액 
		
	Case C_VatYn
		
		Call ChangeVatAmt(Row)	'Vat 금액 
		
		Call ChangeNetAmt(Row)	'순금액 
		
		Call ChangeLocAmt(Row)	'자국금액 
		
		Call ChangeVatLocAmt(Row)	'Vat자국금액 
		
		Call ChangeNetLocAmt(Row)	'자국순금액 
			
		Call HSumAmtNewCalc			'Header 의 Sum 금액 
		
		
	End Select
	
End Sub

'==========================================   ChangeVatLocAmt()  ===============================
'	Name : ChangeNetLocAmt()
'	Description : detail 금액이 변할때 자국순금액 변경 함수 
'==============================================================================================
Function ChangeNetLocAmt(Row)
	
	Dim VatFlg 
	Dim DocAmt	'자국금액 
	Dim NetAmt	'자국순금액 
	
	With frm1.vspdData
		
		.Row = Row
		.Col = C_VatFlg
		VatFlg = .Text
		
		.Col = C_CtlLocAmt
		DocAmt = UNICDbl(.Text)
		
		If VatFlg = "1" Then		'별도 
			NetAmt = UNIConvNumPCToCompanyByCurrency(DocAmt,frm1.txtCur.value,parent.ggAmtOfMoneyNo,parent.gTaxRndPolicyNo , "X")'DocAmt * VatRt / 100                  'vat 별도 vat 금액 
		Else
			.Col = C_CtlLocVatAmt
			NetAmt = UNIConvNumPCToCompanyByCurrency(DocAmt - UNICDbl(.Text),frm1.txtCur.value,parent.ggAmtOfMoneyNo,parent.gTaxRndPolicyNo , "X")'DocAmt * VatRt / 100                  'vat 별도 vat 금액 
		End If
		
		.Col = C_NetLocAmt
		
		.Text = NetAmt
		
	End With
	

End Function

'==========================================   ChangeVatLocAmt()  ===============================
'	Name : ChangeVatLocAmt()
'	Description : detail 금액이 변할때 VAT 자국금액 변경 함수 
'==============================================================================================
Function ChangeVatLocAmt(Row)
	
	Dim VatAmt , XchRt , Xchop
	
	XchRt = UNICDbl(Trim(frm1.txtXchRt.Text)) 
	Xchop = Trim(frm1.cboXchop.Value)
	
	With frm1.vspdData
		
		.Row = Row	
		.Col = C_VatCtlAmt
		VatAmt = UNICDbl(.Text)
		
		.Col = C_CtlLocVatAmt
		If Xchop = "*" Then
			.Text = UNIConvNumPCToCompanyByCurrency(VatAmt * XchRt, frm1.txtCur.value,parent.ggAmtOfMoneyNo,"X", "X")
		Else
			.Text = UNIConvNumPCToCompanyByCurrency(VatAmt / XChRt, frm1.txtCur.value,parent.ggAmtOfMoneyNo,"X", "X")
		End If
		
	End With
	

End Function

'==========================================   ChangeLocAmt()  ===============================
'	Name : ChangeLocAmt()
'	Description : detail 금액이 변할때 자국금액 변경 함수 
'==============================================================================================
Function ChangeLocAmt(Row)
	
	Dim DocAmt , XchRt , Xchop , LocAmt
	
	XchRt = UNICDbl(Trim(frm1.txtXchRt.Text))
	Xchop = Trim(frm1.cboXchop.Value)
	
	With frm1.vspdData
		.Row = Row
		.Col = C_CtlAmt
		DocAmt = UNICDbl(Trim(.Text))
		
		If Xchop = "*" Then
			LocAmt = DocAmt * XchRt  
		Else
			LocAmt = DocAmt / XchRt  
		End If
		
		.Col = C_CtlLocAmt
		frm1.vspdData.text = UNIConvNumPCToCompanyByCurrency(UNICDbl(LocAmt) ,parent.gCurrency,parent.ggAmtOfMoneyNo,parent.gTaxRndPolicyNo , "X")'vatloc 라운딩 
		
	End With
	

End Function

'==========================================   ChangeVatAmt()  ===============================
'	Name : ChangeVatAmt()
'	Description : detail 금액이 변할때 Vat금액 변경 함수 
'==============================================================================================
Function ChangeVatAmt(Row)
	
	Dim VatFlg , VatAmt , DocAmt , VatRt
	
	With frm1.vspdData
		
		.Row = Row
		.Col = C_VatFlg
		VatFlg = .Text
		
		.Col = C_CtlAmt
		DocAmt = UNICDbl(.Text)
		
		VatRt = UNICDbl(Trim(frm1.txtVatRt.Text))
		
		If VatFlg = "1" Then		'별도 
			VatAmt = UNIConvNumPCToCompanyByCurrency((DocAmt * VatRt) / 100,frm1.txtCur.value,parent.ggAmtOfMoneyNo,parent.gTaxRndPolicyNo , "X")'DocAmt * VatRt / 100                  'vat 별도 vat 금액 
		Else
			VatAmt = UNIConvNumPCToCompanyByCurrency((DocAmt * VatRt) / (VatRt + 100),frm1.txtCur.value,parent.ggAmtOfMoneyNo,parent.gTaxRndPolicyNo , "X")'CInt(DocAmt * VatRt / (VatRt + 100))  'vat 포함 vat 금액 
		End If
		
		.Col = C_VatCtlAmt
		
		.Text = VatAmt
		
	End With
	

End Function

'==========================================   ChangeNetAmt()  ===============================
'	Name : ChangeNetAmt()
'	Description : detail 금액이 변할때 조회부 총액변경 Event 합수 
'==============================================================================================
Function ChangeNetAmt(Row)
	
	Dim VatFlg , DocAmt , NetAmt
	
	With frm1.vspdData
		
		.Row = Row
		.Col = C_VatFlg
		VatFlg = .Text
		
		.Col = C_CtlAmt
		DocAmt = UNICDbl(.Text)
		
		If VatFlg = "1" Then		'별도 
			NetAmt = UNIConvNumPCToCompanyByCurrency(DocAmt,frm1.txtCur.value,parent.ggAmtOfMoneyNo,parent.gTaxRndPolicyNo , "X")'DocAmt * VatRt / 100                  'vat 별도 vat 금액 
		Else
			.Col = C_VatCtlAmt
			NetAmt = UNIConvNumPCToCompanyByCurrency(DocAmt - UNICDbl(.Text),frm1.txtCur.value,parent.ggAmtOfMoneyNo,parent.gTaxRndPolicyNo , "X")'DocAmt * VatRt / 100                  'vat 별도 vat 금액 
		End If
		
		.Col = C_NetAmt
		
		.Text = NetAmt
		
	End With
	

End Function

'==========================================   HSumAmtNewCalc()  ===============================
'	Name : HSumAmtNewCalc()
'	Description : detail 금액이 변할때 조회부 총액변경 Event 합수 
'==============================================================================================
Function HSumAmtNewCalc()

	Dim iIndex
	Dim SumIvAmt, SumLocAmt , SumVatAmt , SumVatLocAmt
	Dim IvAmt , LocAmt , VatAmt , VatLocAmt
	
	With frm1
	
		If .hdnIvAmt.value = ""				then .hdnIvAmt.value = "0"
		If .hdnIvLocAmt.value = ""			then .hdnIvLocAmt.value = "0"
		If .hdnGrossVatAmt.value = ""		then .hdnGrossVatAmt.value = "0"
		If .hdnGrossVatLocAmt.value = ""	then .hdnGrossVatLocAmt.value = "0"
		
	End With 
	
	SumIvAmt		= UNICDbl(frm1.hdnIvAmt.value)
	SumLocAmt		= UNICDbl(frm1.hdnIvLocAmt.value)
	SumVatAmt		= UNICDbl(frm1.hdnGrossVatAmt.value)
	SumVatLocAmt	= UNICDbl(frm1.hdnGrossVatLocAmt.value)

	With frm1.vspdData
	
		If .Maxrows >= 1 then 
		
			For iIndex = 1to .MaxRows
	
				.Row = iIndex
				.Col = C_NetAmt_Old
				SumIvAmt = SumIvAmt - UNICDbl(.Text)
		
				.Col = C_NetLocAmt_Old
				SumLocAmt = SumLocAmt - UNICDbl(.Text)
		
				.Col = C_VatCtlAmt_Old
				SumVatAmt = SumVatAmt - UNICDbl(.Text)
		
				.Col = C_CtlLocVatAmt_Old
				SumVatLocAmt = SumVatLocAmt - UNICDbl(.Text)
		
			Next
		
			For iIndex = 1 to .Maxrows
				.Row = iIndex
				.Col = 0
				If Trim(.text) <> ggoSpread.DeleteFlag then 			
				
					'매입단가 조정금액 
					.Col = C_NetAmt
					IvAmt = UNICDbl(.text)
					SumIvAmt = SumIvAmt + IvAmt
					
					'매입단가 조정자국금액 
					.Col = C_NetLocAmt
					LocAmt = UNICDbl(.text)
					SumLocAmt = SumLocAmt + LocAmt
										
					'VAT조정금액 
					.Col = C_VatCtlAmt
					VatAmt = UNICDbl(.text)
					SumVatAmt = SumVatAmt + VatAmt
										
					'VAT조정자국금액 
					.Col = C_CtlLocVatAmt
					VatLocAmt = UNICDbl(.text)
					SumVatLocAmt = SumVatLocAmt + VatLocAmt
				End if
				
			Next
		Else
			SumIvAmt		= 0
			SumLocAmt		= 0
			SumVatAmt		= 0
			SumVatLocAmt	= 0
		End if
			
	End with				
			
	frm1.txtIvAmt.Text			= UNIConvNumPCToCompanyByCurrency(SumIvAmt,frm1.txtCur.value,parent.ggAmtOfMoneyNo,"X" , "X")
	frm1.txtIvLocAmt.Text		= UNIConvNumPCToCompanyByCurrency(SumLocAmt,frm1.txtCur.value,parent.ggAmtOfMoneyNo,"X" , "X")
	frm1.txtGrossVatAmt.Text	= UNIConvNumPCToCompanyByCurrency(SumVatAmt,frm1.txtCur.value,parent.ggAmtOfMoneyNo,parent.gTaxRndPolicyNo , "X")
	frm1.txtGrossVatLocAmt.Text = UNIConvNumPCToCompanyByCurrency(SumVatLocAmt,frm1.txtCur.value,parent.ggAmtOfMoneyNo,parent.gTaxRndPolicyNo , "X")

End Function


Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
	Dim intIndex 

	With frm1.vspdData
	
		.Row = Row
		.Col = Col

		if Col = C_VatYn then 
			intIndex = .Value
			.Col	= C_VatFlg
			if intIndex = 0 then
				.text	= "2"
			else
				.text	= "1"
			end if 
				
        end if
         		
  End With
 
End Sub

Sub vspdData_LeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
    
    With frm1.vspdData 

    If Row >= NewRow Then
        Exit Sub
    End If

    If NewRow = .MaxRows Then
        'DbQuery
    End if    

    End With

End Sub


'================ vspdData_TopLeftChange() ==========================================================
Sub vspdData_TopLeftChange(ByVal OldLeft, ByVal OldTop, ByVal NewLeft, ByVal NewTop)
	
	If OldLeft <> NewLeft Then
	    Exit Sub
	End If
    

	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	    '☜: 재쿼리 체크	
		If lgStrPrevKey <> "" Then		                                                    '다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
			If DbQuery = False Then
				Exit Sub
			End if
		End If
	End If
End Sub

'==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분 
'=========================================================================================================
 Sub Form_Load()
	
    Call LoadInfTB19029  	
    Call AppendNumberRange("0","0","999")

    '이성룡 
    Call initFormatField()
    Call InitComboBox			
    Call SetDefaultVal
    Call InitVariables
    Call InitSpreadSheet           
    
    '----------  Coding part  -------------------------------------------------------------
    Call Changeflg
    
    Call CookiePage(0)

	Call changeTabs(TAB1)

    gIsTab     = "Y"
	gTabMaxCnt = 2

End Sub

'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )
   
End Sub

'==========================================================================================
'   Event Name : OCX Event
'   Event Desc :
'==========================================================================================
 Sub txtIvDt_DblClick(Button)
	if Button = 1 then
		frm1.txtIvDt.Action = 7
	End if
End Sub

 Sub txtIvDt_Change()
	lgBlnFlgChgValue = true	
	frm1.hdnPoDt.value = frm1.txtIvDt.text End Sub

Sub txtXchRt_OnBlur()
	lgBlnFlgChgValue = true	
	Call SetAmtByCurAmt
End Sub

 Sub txtPayDur_Change()
	lgBlnFlgChgValue = true	
End Sub
 Sub txtCnfmDt_DblClick(Button)
	if Button = 1 then
		frm1.txtCnfmDt.Action = 7
	End if
End Sub
 Sub txtCnfmDt_Change()
	lgBlnFlgChgValue = true	
End Sub
 Sub txtExpiryDt_DblClick(Button)
	if Button = 1 then
		frm1.txtExpiryDt.Action = 7
	End if
End Sub
 Sub txtExpiryDt_Change()
	lgBlnFlgChgValue = true	
End Sub
Sub rdoVatFlg1_OnClick()
	lgBlnFlgChgValue = true	
End Sub

Sub rdoVatFlg2_OnClick()
	lgBlnFlgChgValue = true	
End Sub

'==========================================================================================
'   Event Name : txtVat_Type_OnChange
'   Event Desc : 수주형태별로 무역정보 필수입력 처리 
'==========================================================================================
Sub cboXchop_OnChange()
	lgBlnFlgChgValue = True	

	if frm1.cboXchop.value ="*" then
		frm1.hdnxchrateop.value = "*"
	Else 
		frm1.hdnxchrateop.value = "/"
	End if
	
	Call SetAmtByCurAmt()
	
End Sub

Sub SetAmtByCurAmt()

	Dim index
	
	With frm1.vspddata
	
		If .MaxRows < 1 Then Exit Sub
		
		For index = 1 to .Maxrows

			ChangeLocAmt(index)
			
			ChangeVatLocAmt(index)
			
			ChangeNetLocAmt(index)
			
		Next
		
		Call HSumAmtNewCalc()		'Header Setting	
	
	End with
	


End Sub



'--------------------------------------------------------------------
'		Name        : SetState()
'		Description : Spread의 Row상태를 "R","C"로 Setting
'					  R-reference 참조      C-InsertRow
'--------------------------------------------------------------------

Sub SetState(byval strState,byval IRow)	
	frm1.vspdData.Row=IRow
	frm1.vspdData.Col=C_Stateflg
	frm1.vspdData.Text=strState
End Sub

Sub setVatAmt()
 dim sum
  
 with frm1
     sum = UNICDbl(.txtVatrt.text) * UNICDbl(.txtIvAmt.text)/100
     '.txtVatAmt.text = UNIFormatNumber(UNICDbl(sum), ggExchRate.DecPoint, -2, 0, ggExchRate.RndPolicy, ggExchRate.RndUnit)
     '.txtVatAmt.text = uniFormatNumberByTax(UNICDbl(sum),.txtPayeeCd.value,Parent.ggAmtOfMoneyNo)'vatloc 라운딩 

 end with
end sub 
'#########################################################################################################
'												4. Common Function부 
'	기능: Common Function
'	설명: 환율처리함수, VAT 처리 함수 
'#########################################################################################################


'#########################################################################################################
'												5. Interface부 
'	기능: Interface
'	설명: 각각의 Toolbar에 대한 처리를 행한다. 
'	      Toolbar의 위치순서대로 기술하는 것으로 한다. 
'	<< 공통변수 정의 부분 >>
' 	공통변수 : Global Variables는 아니지만 각각의 Sub나 Function에서 자주 사용하는 변수로 변수명은 
'				통일하도록 한다.
' 	1. 공통컨트롤을 Call하는 변수 
'    	   ADF (ADS, ADC, ADF는 그대로 사용)
'    	   - ADF는 Set하고 사용한 뒤 바로 Nothing 하도록 한다.
' 	2. 공통컨트롤에서 Return된 값을 받는 변수 
'    		strRetMsg
'#########################################################################################################
'*******************************  5.1 Toolbar(Main)에서 호출되는 Function *******************************
'	설명 : Fnc함수명 으로 시작하는 모든 Function
'*********************************************************************************************************
'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================
Function FncQuery() 
    Dim IntRetCD 
       
    FncQuery = False                                                
    Err.Clear                                                       

    If lgBlnFlgChgValue = True Then
    	
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    Call InitVariables												
	
	If Not chkFieldByCell(frm1.txtIvNo, "A",1)	then
        If gPageNo > 0 Then
            gSelframeFlg = gPageNo
        End If
            
        Exit Function
    End If 

    If DbQuery = False Then Exit Function
    Call Changeflg       
    
 '   lgBlnFlgChgValue = False
    FncQuery = True		
    
End Function

'========================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'========================================================================================
Function FncNew() 
    Dim IntRetCD 
    
    FncNew = False                                                  
    
    If lgBlnFlgChgValue = True Then
       IntRetCD = DisplayMsgBox("900015", Parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    Call ClickTab1()
    Call ggoOper.ClearField(Document, "1")                          
    Call ggoOper.ClearField(Document, "2")                          
    Call ggoOper.ClearField(Document, "3")                          
    Call ggoOper.LockField(Document, "N")    
    
    Call SetDefaultVal
    Call InitVariables
    Call InitSpreadSheet													
   

    frm1.txtIvNo.focus
	Set gActiveElement = document.activeElement
    
    FncNew = True														

End Function

'========================================================================================
' Function Name : FncDelete
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncDelete() 

	Dim IntRetCD,lRow

    FncDelete = False
    
    IntRetCD = DisplayMsgBox("900003", Parent.VB_YES_NO, "X", "X")
    
    If IntRetCD = vbNo Then Exit Function
    
    if frm1.vspdData.Maxrows < 1	then exit function
    
    With frm1.vspdData 

	'	.focus
		 ggoSpread.Source = frm1.vspdData 
		 
		 For lRow = 1 To .MaxRows step 1
		    .Row  = lRow
	       	.Col  = 0
			.Text = ggoSpread.DeleteFlag
		 Next
		'lDelRows = ggoSpread.DeleteRow
    End With
    If DbDelete = False Then Exit Function
    
    FncDelete = True                                        
        
End Function

'========================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Save Button of Main ToolBar
'========================================================================================
Function FncSave() 
    Dim IntRetCD 
    
    FncSave = False                                                         '⊙: Processing is NG
    
    Err.Clear   

	if CheckRunningBizProcess = true then
		exit function
	end if                                                            '☜: Protect system from crashing
    
    ggoSpread.Source = frm1.vspdData
    
    If lgBlnFlgChgValue = False and ggoSpread.SSCheckChange = False Then
        IntRetCD = DisplayMsgBox("900001","X","X","X")
        Exit Function
    End If

	If Not ggoSpread.SSDefaultCheck         Then            
	   Exit Function
	End If
    
    If Len(Trim(frm1.txtxchRt.text)) <= 0 Then
 	    IntRetCD =  DisplayMsgBox("200095", "X", "X", "X")
	    Call ClickTab1
	    frm1.txtxchRt.focus
	    Exit Function
    End If 
    
    If frm1.txtxchRt.text = 0 Then
 	    IntRetCD =  DisplayMsgBox("200095", "X", "X", "X")
	    Call ClickTab1
	    frm1.txtxchRt.focus
	    Exit Function
    End If 
    
    If Not chkField(Document, "2") Then                                  '⊙: Check contents area
       Exit Function
    End If
 
    'vat 포함여부 
    if frm1.rdoVatFlg1.checked = true then
    	frm1.hdvatFlg.Value = "1"
    else
    	frm1.hdvatFlg.Value = "2"
    End if
   
    If DbSave("toolbar") = False Then Exit Function                         '☜: Save db data
    
    If frm1.hdnIvNo.value <> frm1.txtIvNo.value then
		frm1.txtIvNo.value =	frm1.hdnIvNo.value		
	End If   
    
    FncSave = True                                                          '⊙: Processing is OK
    
End Function

'========================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================
Function FncCopy() 


End Function

'========================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================
Function FncCancel() 

    Dim maxrow,maxrow1,SumTotal,tmpGrossAmt,index,index1,orgtmpGrossAmt
    Dim SumVatTotal, tmpVatAmt, orgtmpVatAmt
	Dim starindex ,endindex,delflag
	
	if frm1.vspdData.Maxrows < 1	then exit function
	
	maxrow = frm1.vspdData.Maxrows
	index1 = 0
	
	starindex = frm1.vspdData.SelBlockRow
	endindex  = frm1.vspdData.SelBlockRow2
    
    Redim orgtmpGrossAmt(endindex - starindex)
    Redim tmpGrossAmt(endindex - starindex)
    Redim tmpVatAmt(endindex - starindex)
    Redim orgtmpVatAmt(endindex - starindex)
    Redim delflag(endindex - starindex)


	for index = starindex to endindex
		frm1.vspdData.Row = index
	    

	    frm1.vspdData.Col = 0
	    delflag(index1) = frm1.vspdData.Text
	    index1 = index1 + 1
	    
	next

	ggoSpread.Source = frm1.vspdData
	index = frm1.vspdData.ActiveRow - starindex
		


     ggoSpread.EditUndo                                     
     maxrow1 = frm1.vspdData.Maxrows

	HSumAmtNewCalc	'Header 의 4개 Sum값 원복 

End Function

'========================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================
Function FncInsertRow(ByVal pvRowCnt) 
                                          
End Function

'========================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================
Function FncDeleteRow() 
    Dim lDelRows
    Dim iDelRowCnt, i
    Dim index,SumTotal,SumVatTotal,idel
    if frm1.vspdData.Maxrows < 1	then exit function
    
    With frm1.vspdData 
    
		.focus
		ggoSpread.Source = frm1.vspdData 
    
		lDelRows = ggoSpread.DeleteRow

		
		for index = .SelBlockRow to .SelBlockRow2
			.Row = index
			.Col = C_Stateflg
			idel = .text
			.Col = 0

			if Trim(.text) <> ggoSpread.InsertFlag and Trim(idel) <> "D" then

		         .Col = C_Stateflg
			     frm1.vspdData.text = "D"

		    end if
		Next
   End With
   
   HSumAmtNewCalc		'4번 Sum 값 재설정.
   
End Function

'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================
Function FncPrint() 
	Call parent.FncPrint()
End Function

'========================================================================================
' Function Name : FncPrev
' Function Desc : This function is related to Previous Button
'========================================================================================
Function FncPrev() 
End Function

'========================================================================================
' Function Name : FncNext
' Function Desc : This function is related to Next Button
'========================================================================================
Function FncNext() 
End Function

'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================
Function FncExcel() 
    Call parent.FncExport(Parent.C_SINGLE)										
End Function

'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================
Function FncFind() 
    Call parent.FncFind(Parent.C_SINGLE , False)                               
End Function

'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================
Function FncExit()
	Dim IntRetCD
	FncExit = False
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    FncExit = True
End Function

'========================================================================================
' Function Name : DbDelete
' Function Desc : This function delete data
'========================================================================================
 Function DbDelete() 
    Err.Clear                                                           
    
    DbDelete = False													
    
    Dim strVal
	
	strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0003						
	strVal = strVal & "&txtIvNo=" & Trim(frm1.txtIvNo2.value)
	strVal = strVal & "&hdnRelease=" & frm1.hdnRelease.value
	strVal = strVal & "&txtMaxRows=" & frm1.txtMaxRows.value
    
    If LayerShowHide(1) = False Then Exit Function    

	Call RunMyBizASP(MyBizASP, strVal)								

    DbDelete = True                                                 

End Function

'========================================================================================
' Function Name : DbDeleteOk
' Function Desc : DbDelete가 성공적일때 수행 
'========================================================================================
Function DbDeleteOk()												
	lgBlnFlgChgValue = False
	Call MainNew()
End Function

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
 Function DbQuery() 
 
    Dim LngLastRow      
    Dim LngMaxRow       
    Dim LngRow          
    Dim strTemp         
    Dim StrNextKey      
    Dim strVal
        
    Err.Clear                                                       
    DbQuery = False       
    
    If LayerShowHide(1) = False Then Exit Function                                          
    
    With frm1
    
    If lgIntFlgMode = Parent.OPMD_UMODE Then
        strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001
	    strVal = strVal & "&txtIvNo=" & .hdnIvNo.value
    Else
	    strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001
	    strVal = strVal & "&txtIvNo=" & UCASE(Trim(.txtIvNo.value))
   End If
	strVal = strVal & "&lgPageNo=" & lgPageNo 
	strVal = strVal & "&lgNextKey=" & lgStrPrevKey
	strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
	
    frm1.hdnMaxRows.value = frm1.vspdData.MaxRows
    
    End with
    
    Call RunMyBizASP(MyBizASP, strVal)								
	
    DbQuery = True                                                  

End Function
'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================
Function DbQueryOk()												
	Dim lRow, lgTab, chkSts
    '-----------------------
    'Reset variables area
    '-----------------------
    'set vat
    '*************************
    call setVatAmt   
    '************************** 
    
    Call ggoOper.LockField(Document, "Q")							

	Call SetSpreadLock

	lgIntFlgMode = Parent.OPMD_UMODE	
	lgIntFlgMode_Dtl = Parent.OPMD_UMODE
	
	 chkSts = "DB"
	Call BtnToolCtrl(chkSts)
    lgBlnFlgChgValue = False

	Set gActiveElement = document.activeElement
	
'	Call HSumAmtNewCalc()

	Call SubGetGlNo()

	frm1.vspdData.ReDraw = True
	
	Call changeTabs(TAB1)
	
	If frm1.hdnRelease.value = "Y" Then
		ggoSpread.SSSetProtected        C_CtlQty, 1, frm1.vspddata.maxrows            
		ggoSpread.SSSetProtected        C_CtlPrc, 1, frm1.vspddata.maxrows            
		ggoSpread.SSSetProtected        C_CtlAmt, 1, frm1.vspddata.maxrows            
		ggoSpread.SSSetProtected        C_VatCtlAmt, 1, frm1.vspddata.maxrows            
		ggoSpread.SSSetProtected        C_CtlLocAmt, 1, frm1.vspddata.maxrows            
		ggoSpread.SSSetProtected        C_CtlLocVatAmt, 1, frm1.vspddata.maxrows            
		ggoSpread.SSSetProtected        C_VatYn, 1, frm1.vspddata.maxrows            
	End If
	
    frm1.vspdData.Focus 
	
End Function

Sub RemovedivTextArea()
	Dim ii

	For ii = 1 To divTextArea.children.length
	    divTextArea.removeChild(divTextArea.children(0))
	Next
End Sub

Sub BtnToolCtrl(byval chkSts)

	Dim lgTab
	lgTab = gSelframeFlg
	
	If frm1.hdnRelease.value = "Y" Then
	
		Call SetToolbar("11100000000111")
		
		frm1.btnCfm.value = "확정취소"
		
		frm1.btnCfm.disabled = False
		
		frm1.btnSelect.disabled = False
	
	Else
	
		Call SetToolbar("1111101100001111")
		
		frm1.btnCfm.value = "확정"
		
		frm1.btnCfm.disabled = False
		
		frm1.btnSelect.disabled = True
	
	End If
	
	

End Sub


'========================================================================================
' Function Name : 
' Function Desc : 실제 저장 로직을 수행 , 성공적이면 DBSaveOk 호출됨 
'========================================================================================
 Function DbSave(byval btnflg) 
    Dim lRow        
    Dim lGrpCnt     
	Dim strVal, strDel
	Dim ColSep, RowSep
	Dim msgCreditlimit
	Dim GmQty
	Dim MvmtIvQty
	Dim IvQty1,OldIvQty1
	Dim chkVatAmt

	Dim iVatDocAmt
	Dim iChkVatDocAmt
	Dim iRefVatRateFlg
	
	Dim strCUTotalvalLen '버퍼에 채워지는 양이 102399byte를 넘어 가는가를 체크하기위한 누적 데이타 크기 저장[수정,신규] 
	Dim strDTotalvalLen  '버퍼에 채워지는 양이 102399byte를 넘어 가는가를 체크하기위한 누적 데이타 크기 저장[삭제]

	Dim objTEXTAREA '동적인 HTML객체(TEXTAREA)를 만들기위한 임시 버퍼 

	Dim iTmpCUBuffer         '현재의 버퍼 [수정,신규] 
	Dim iTmpCUBufferCount    '현재의 버퍼 Position
	Dim iTmpCUBufferMaxCount '현재의 버퍼 Chunk Size

	Dim iTmpDBuffer          '현재의 버퍼 [삭제] 
	Dim iTmpDBufferCount     '현재의 버퍼 Position
	Dim iTmpDBufferMaxCount  '현재의 버퍼 Chunk Size

    DbSave = False                                                          '⊙: Processing is NG
    
    ColSep = parent.gColSep														'⊙: Column 분리 파라메타 
	RowSep = parent.gRowSep														'⊙: Row 분리 파라메타 
	
	iTmpCUBufferMaxCount = parent.C_CHUNK_ARRAY_COUNT '한번에 설정한 버퍼의 크기 설정[수정,신규]
	iTmpDBufferMaxCount  = parent.C_CHUNK_ARRAY_COUNT '한번에 설정한 버퍼의 크기 설정[삭제]

	ReDim iTmpCUBuffer(iTmpCUBufferMaxCount) '초기 버퍼의 설정[수정,신규]
	ReDim iTmpDBuffer (iTmpDBufferMaxCount)  '초기 버퍼의 설정[삭제]

	iTmpCUBufferCount = -1
	iTmpDBufferCount = -1

	strCUTotalvalLen = 0
	strDTotalvalLen  = 0

	With frm1
'	.hdnUsrId.value = parent.gUsrID
	.txtMode.value = parent.UID_M0002
	.txtFlgMode.value = lgIntFlgMode
		
    lGrpCnt = 1
    
    strVal = ""
    strDel = ""       		    
	
	If btnflg = "Posting" Then
		.txtMode.value = "Release" 				'☜: 확정 버튼 
	ElseIf btnflg = "UnPosting" then
		.txtMode.value = "UnRelease"
	End If

    For lRow = 1 To .vspdData.MaxRows
    
        .vspdData.Row = lRow
        .vspdData.Col = 0

        Select Case .vspdData.Text
        
        Case ggoSpread.InsertFlag, ggoSpread.UpdateFlag
        
			If .vspdData.Text = ggoSpread.InsertFlag Then
				strVal = strVal & "C" & ColSep				'☜: C=Create
			ElseIf .vspdData.Text = ggoSpread.UpdateFlag Then
				strVal = strVal & "U" & ColSep				'☜: U=Update
			End If
			
			.vspdData.Col = C_CtlQty
			If Trim(UNICDbl(.vspdData.Text)) = "" Or Trim(UNICDbl(.vspdData.Text)) = "0" Then
				Call DisplayMsgBox("970021","X","조정수량","X")
				Call LayerShowHide(0)
				Exit Function
			End If
				
        	.vspdData.Col = C_CtlPrc
			If Trim(UNICDbl(.vspdData.Text)) = "" Or Trim(UNICDbl(.vspdData.Text)) = "0" Then
				Call DisplayMsgBox("970021","X","조정단가","X")
				Call LayerShowHide(0)
				Exit Function
			End If

        	.vspdData.Col = C_CtlAmt
			If Trim(UNICDbl(.vspdData.Text)) = "" Or Trim(UNICDbl(.vspdData.Text)) = "0" Then
				Call DisplayMsgBox("970021","X","조정금액","X")
				Call LayerShowHide(0)
				Exit Function
			End If
			
			'.vspdData.Col = C_VatCtlAmt
			If Trim(UNICDbl(.vspdData.Text)) = "" Then
				Call DisplayMsgBox("970021","X","VAT조정금액","X")
				Call LayerShowHide(0)
				Exit Function
			End If
			
			.vspdData.Col = C_CtlLocAmt
			If Trim(UNICDbl(.vspdData.Text)) = "" Or Trim(UNICDbl(.vspdData.Text)) = "0" Then
				Call DisplayMsgBox("970021","X","조정자국금액","X")
				Call LayerShowHide(0)
				Exit Function
			End If			
						
			.vspdData.Col = C_CtlLocVatAmt
			If Trim(UNICDbl(.vspdData.Text)) = "" Then
				Call DisplayMsgBox("970021","X","VAT조정자국금액","X")
				Call LayerShowHide(0)
				Exit Function
			End If		
			
			
			.vspdData.Col = C_IvNo	 :		strVal = strVal & Trim(.vspdData.Text) & ColSep
			.vspdData.Col = C_IvSeq	 :		strVal = strVal & Trim(.vspdData.Text) & ColSep
			.vspdData.Col = C_PlantCd:		strVal = strVal & Trim(.vspdData.Text) & ColSep
			.vspdData.Col = C_ItemCd :		strVal = strVal & Trim(.vspdData.Text) & ColSep
			
			.vspdData.Col = C_CtlQty :		strVal = strVal & UNICDbl(Trim(.vspdData.Text)) & ColSep
			.vspdData.Col = C_CtlPrc :		strVal = strVal & UNICDbl(Trim(.vspdData.Text)) & ColSep
			.vspdData.Col = C_CtlAmt :		strVal = strVal & UNICDbl(Trim(.vspdData.Text)) & ColSep
			.vspdData.Col = C_VatFlg :		strVal = strVal & Trim(.vspdData.Text) & ColSep
			.vspdData.Col = C_VatCtlAmt:	strVal = strVal & UNICDbl(Trim(.vspdData.Text)) & ColSep
			.vspdData.Col = C_CtlLocAmt:	strVal = strVal & UNICDbl(Trim(.vspdData.Text)) & ColSep
			.vspdData.Col = C_CtlLocVatAmt:	strVal = strVal & UNICDbl(Trim(.vspdData.Text)) & ColSep
			.vspdData.Col = C_PoNo:			strVal = strVal & Trim(.vspdData.Text) & ColSep
			.vspdData.Col = C_PoSeq:		strVal = strVal & Trim(.vspdData.Text) & ColSep
			.vspdData.Col = C_IvNohdn:		strVal = strVal & Trim(.vspdData.Text) & ColSep
			.vspdData.Col = C_IvSeqhdn:		strVal = strVal & Trim(.vspdData.Text) & ColSep
			
			.vspdData.Col = C_ItemUnit:		strVal = strVal & Trim(.vspdData.Text) & ColSep		
			.vspdData.Col = C_ItemAcct:		strVal = strVal & Trim(.vspdData.Text) & ColSep	
			
			.vspdData.Col = C_VatType:		strVal = strVal & Trim(.vspdData.Text) & ColSep				
			.vspdData.Col = C_VatRt:		strVal = strVal & Trim(.vspdData.Text) & ColSep		
			
			.vspdData.Col = C_IvBizArea:	strVal = strVal & Trim(.vspdData.Text) & ColSep	
			.vspdData.Col = C_MvmtQty:		strVal = strVal & Trim(.vspdData.Text) & ColSep	
			.vspdData.Col = C_MvmtFlg:		strVal = strVal & Trim(.vspdData.Text) & ColSep	
			.vspdData.Col = C_TrackingNo:		strVal = strVal & Trim(.vspdData.Text) & ColSep	
			.vspdData.Col = C_IvCostCd:		strVal = strVal & Trim(.vspdData.Text) & ColSep	
			

			strVal = strVal & lRow & RowSep
		
		Case ggoSpread.DeleteFlag
			
			strDel = strDel & "D" & ColSep				'☜: D=Delete
			.vspdData.Col = C_IvNo :		strDel = strDel & Trim(.vspdData.Text) & ColSep
			.vspdData.Col = C_IvSeq:		strDel = strDel & Trim(.vspdData.Text) & ColSep
			strDel = strDel & lRow & RowSep
			
		End Select  
		
		
		lGrpCnt = lGrpCnt + 1		         
        '=====================
        .vspdData.Col = 0
		Select Case .vspdData.Text
		    Case ggoSpread.InsertFlag,ggoSpread.UpdateFlag
		    
		         If strCUTotalvalLen + Len(strVal) >  parent.C_FORM_LIMIT_BYTE Then  '한개의 form element에 넣을 Data 한개치가 넘으면 
		                            
		            Set objTEXTAREA = document.createElement("TEXTAREA")                 '동적으로 한개의 form element를 동저으로 생성후 그곳에 데이타 넣음 
		            objTEXTAREA.name = "txtCUSpread"
		            objTEXTAREA.value = Join(iTmpCUBuffer,"")
		            divTextArea.appendChild(objTEXTAREA)     
		 
		            iTmpCUBufferMaxCount = parent.C_CHUNK_ARRAY_COUNT                  ' 임시 영역 새로 초기화 
		            ReDim iTmpCUBuffer(iTmpCUBufferMaxCount)
		            iTmpCUBufferCount = -1
		            strCUTotalvalLen  = 0
		         End If
		       
		         iTmpCUBufferCount = iTmpCUBufferCount + 1
		      
		         If iTmpCUBufferCount > iTmpCUBufferMaxCount Then                              '버퍼의 조정 증가치를 넘으면 
		            iTmpCUBufferMaxCount = iTmpCUBufferMaxCount + parent.C_CHUNK_ARRAY_COUNT '버퍼 크기 증성 
		            ReDim Preserve iTmpCUBuffer(iTmpCUBufferMaxCount)
		         End If   
		         iTmpCUBuffer(iTmpCUBufferCount) =  strVal         
		         strCUTotalvalLen = strCUTotalvalLen + Len(strVal)
		   Case ggoSpread.DeleteFlag
		         If strDTotalvalLen + Len(strDel) >  parent.C_FORM_LIMIT_BYTE Then   '한개의 form element에 넣을 한개치가 넘으면 
		            Set objTEXTAREA   = document.createElement("TEXTAREA")
		            objTEXTAREA.name  = "txtDSpread"
		            objTEXTAREA.value = Join(iTmpDBuffer,"")
		            divTextArea.appendChild(objTEXTAREA)     
		          
		            iTmpDBufferMaxCount = parent.C_CHUNK_ARRAY_COUNT              
		            ReDim iTmpDBuffer(iTmpDBufferMaxCount)
		            iTmpDBufferCount = -1
		            strDTotalvalLen = 0 
		         End If
		       
		         iTmpDBufferCount = iTmpDBufferCount + 1

		         If iTmpDBufferCount > iTmpDBufferMaxCount Then                         '버퍼의 조정 증가치를 넘으면 
		            iTmpDBufferMaxCount = iTmpDBufferMaxCount + parent.C_CHUNK_ARRAY_COUNT
		            ReDim Preserve iTmpDBuffer(iTmpDBufferMaxCount)
		         End If   
		         
		         iTmpDBuffer(iTmpDBufferCount) =  strDel         
		         strDTotalvalLen = strDTotalvalLen + Len(strDel)
		End Select  
        strVal = ""
        strDel = ""
        '=====================
       
    Next
	
	If iTmpCUBufferCount > -1 Then   ' 나머지 데이터 처리 
	   Set objTEXTAREA = document.createElement("TEXTAREA")
	   objTEXTAREA.name   = "txtCUSpread"
	   objTEXTAREA.value = Join(iTmpCUBuffer,"")
	   divTextArea.appendChild(objTEXTAREA)     
	End If   

	If iTmpDBufferCount > -1 Then    ' 나머지 데이터 처리 
	   Set objTEXTAREA = document.createElement("TEXTAREA")
	   objTEXTAREA.name = "txtDSpread"
	   objTEXTAREA.value = Join(iTmpDBuffer,"")
	   divTextArea.appendChild(objTEXTAREA)     
	End If 

	'msgbox objTEXTAREA.value
	
	If lGrpCnt > 1 Or btnflg = "Posting" Then
		If LayerShowHide(1) = False Then
			Exit function
		End If
		Call ExecMyBizASP(frm1, BIZ_PGM_ID)										'☜: 비지니스 ASP 를 가동 
	End If
	
	End With
	
    DbSave = True                                                           '⊙: Processing is NG
    
End Function
'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncSave에 있는것을 옮김 
'========================================================================================
Function DbSaveOk()									
	lgBlnFlgChgValue = False
	Call fncQuery()		
End Function

'========================================================================================
' Function Name : chkEachFieldDomestic, chkEachFieldImport
' Function Desc : Manual check whether a value is entered at required field 
'========================================================================================
Function chkEachFieldDomestic()
	chkEachFieldDomestic = True
	
	If Not chkFieldByCell (frm1.txtIvTypeCd, "A",1) then 
		chkEachFieldDomestic = False
		Exit Function
	End If
	
	If Not chkFieldByCell (frm1.txtSpplCd, "A",1) then 
		chkEachFieldDomestic = False
		Exit Function
	End If
	
	If Not chkFieldByCell (frm1.txtIvDt, "A",1) then 
		chkEachFieldDomestic = False
		Exit Function
	End If
	
	If Not chkFieldByCell (frm1.txtBillCd, "A",1) then 
		chkEachFieldDomestic = False
		Exit Function
	End If
	
	If Not chkFieldByCell (frm1.txtPayeeCd, "A",1) then 
		chkEachFieldDomestic = False
		Exit Function
	End If
	
	If Not chkFieldByCell (frm1.txtPayTermCd, "A",1) then 
		chkEachFieldDomestic = False
		Exit Function
	End If
	
End Function

Function chkEachFieldImport()
	chkEachFieldImport	= True
	
	If Not chkFieldByCell (frm1.txtCnfmDt, "A",1) then 
		chkEachFieldImport = False
		Exit Function
	End If
	
	'If Not chkFieldByCell (frm1.txtOffDt, "A",1) then 
	'	chkEachFieldImport = False
	'	Exit Function
	'End If
	

	

	
	'If Not chkFieldByCell (frm1.txtApplicantCd, "A",1) then 
	'	chkEachFieldImport = False
	'	Exit Function
	'End If
	
End Function

'========================================================================================
' Function Name : initFormatField()
' Function Desc : Manual Formatting fields as amount or date 
'========================================================================================
Function  initFormatField()
	
	'Header
	call FormatDateField(frm1.txtIvDt)
	call FormatDateField(frm1.txtCnfmDt)

	call FormatDoubleSingleField(frm1.txtXchRt)
	call FormatDoubleSingleField(frm1.txtIvAmt)
	call FormatDoubleSingleField(frm1.txtIvLocAmt)
	call FormatDoubleSingleField(frm1.txtGrossVatAmt)
	call FormatDoubleSingleField(frm1.txtGrossVatLocAmt)

	call FormatDoubleSingleField(frm1.txtVatrt)


	
	call LockobjectField(frm1.txtIvDt,"R")
	call LockobjectField(frm1.txtCnfmDt,"O")

	
	call LockobjectField(frm1.txtXchRt,"O")
	call LockobjectField(frm1.txtIvAmt,"P")
	call LockobjectField(frm1.txtIvLocAmt,"P")
	call LockobjectField(frm1.txtGrossVatAmt,"P")
	call LockobjectField(frm1.txtGrossVatLocAmt,"P")

	call LockobjectField(frm1.txtVatrt,"P")


     
	call ggoOper.SetReqAttr(frm1.txtCnfmDt, "D")


End Function 

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>
<BODY TABINDEX="-1" SCROLL="NO">
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
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="ClickTab1()">
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>단가정산매입</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="ClickTab2()" onMouseOver="vbscript:SetClickflag" onMouseOut="vbscript:ResetClickflag" onFocus="vbscript:SetClickflag" onBlur="vbscript:ResetClickflag">
							<TR>
								<td background="../../../CShared/image/table/tab_up_bg.gif"><img src="../../../CShared/image/table/tab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/tab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>단가정산기타</font></td>
								<td background="../../../CShared/image/table/tab_up_bg.gif" align="right"><img src="../../../CShared/image/table/tab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					
					<TD WIDTH=*>&nbsp;</TD>
					<TD WIDTH=* align=right><A href="vbscript:OpenReqRef">매입참조</A></TD>
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
									<TD CLASS="TD5" NOWRAP>매입번호</TD>
									<TD CLASS="TD6" NOWRAP><INPUT Class = required  TYPE=TEXT NAME="txtIvNo" SIZE=32  MAXLENGTH=18 ALT="매입번호" tag="12NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnORGCd1" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenIvNo()"></TD>
									<TD CLASS=TD6></TD>
									<TD CLASS=TD6></TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR height="*">
					<TD WIDTH=100% valign=top>
						<DIV ID="TabDiv" STYLE="FLOAT: left; HEIGHT: 100%; WIDTH: 100%;" SCROLL=no>
							<TABLE <%=LR_SPACE_TYPE_60%>>
								<TR>
									<TD CLASS="TD5" NOWRAP>매입번호</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT AlT="매입번호" NAME="txtIvNo2"  MAXLENGTH=18 SIZE=34 tag="23NXXU" ></TD>
									<TD CLASS="TD5" NOWRAP>확정여부</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=radio AlT="발주확정" NAME="rdoRelease" CLASS="RADIO" checked tag="24" ONCLICK="vbscript:SetChangeflg(1)"><label for="rdoRelease">&nbsp;미확정&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</label>
														   <INPUT TYPE=radio AlT="발주확정" NAME="rdoRelease" CLASS="RADIO" ONCLICK="vbscript:setChangeflg(1)" tag="24"><label for="rdoRelease">&nbsp;확정&nbsp;</label></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>매입형태</TD>
									<TD CLASS="TD6" nowrap><INPUT CLASS = required TYPE=TEXT NAME="txtIvTypeCd" ALT="매입형태" MAXLENGTH=5 style="HEIGHT: 20px; WIDTH: 70px" tag="22NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnGrp" style="HEIGHT: 21px; WIDTH: 16px"  align=top TYPE="BUTTON"ONCLICK="vbscript:OpenIvType()">
														   <INPUT CLASS = protected readonly TYPE=TEXT NAME="txtIvTypeNm" ALT="매입형태" style="HEIGHT: 20px; WIDTH: 150px" tag="24X"></TD>
									<TD CLASS="TD5" NOWRAP>매입일</TD>
									<TD CLASS="TD6" NOWRAP>
										<Table cellpadding=0 cellspacing=0>
										   <TR>
									          <TD><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASS = required ALT=매입일 NAME="txtIvDt" CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 style="HEIGHT: 20px; WIDTH: 100px" tag="22X1" Title="FPDATETIME"></OBJECT>');</SCRIPT>
											  </TD>
											  <TD NOWRAP>
												&nbsp;확정일													
											  </TD>
											  <TD NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT=확정일 NAME="txtCnfmDt" CLASSID=<%=gCLSIDFPDT%> id=fpDateTime2 style="HEIGHT: 20px; WIDTH: 100px" tag="21X1" Title="FPDATETIME"></OBJECT>');</SCRIPT>
											  </TD>
											</TR>
										</Table>				
									</TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>공급처</TD>
									<TD CLASS="TD6" NOWRAP><INPUT CLASS = required TYPE=TEXT AlT="공급처" NAME="txtSpplCd" MAXLENGTH=10 SIZE=10 tag="23NXXU" ONChange="vbscript:ChangeSupplier(1)" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnGroup1Cd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenSppl(1)" OnMouseOver="vbscript:PopUpMouseOver()" OnMouseOut="vbscript:PopUpMouseOut()">
														   <INPUT TYPE=TEXT AlT="공급처" ID="txtSpplNm" NAME="arrCond" tag="24X" CLASS = protected readonly = True TabIndex = -1 ></TD>									
									<TD CLASS="TD5" NOWRAP>세금계산서발행처</TD>
									<TD CLASS="TD6" NOWRAP><INPUT CLASS = required TYPE=TEXT AlT="세금계산서발행처" NAME="txtBuildCd" MAXLENGTH=10 SIZE=10 tag="23NXXU" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnGroupCd2" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenSppl(3)">
														   <INPUT TYPE=TEXT AlT="세금계산서발행처" NAME="txtBuildNm" SIZE=20 tag="24X" CLASS = protected readonly = True TabIndex = -1 ></TD>
								</TR>							
								<TR>
									<TD CLASS="TD5" NOWRAP>지급처</TD>
									<TD CLASS="TD6" NOWRAP><INPUT CLASS = required TYPE=TEXT NAME="txtPayeeCd" ALT="지급처" style="HEIGHT: 20px; WIDTH: 70px" MAXLENGTH=10 SIZE=10 tag="22NXXU" ONChange="vbscript:ChangeSupplier(2)"><IMG SRC="../../../CShared/image/btnPopup.gif"  style="HEIGHT: 21px; WIDTH: 16px" NAME="btnSupplier" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenSppl(2)" OnMouseOver="vbscript:PopUpMouseOver()" OnMouseOut="vbscript:PopUpMouseOut()">
														   <INPUT CLASS = protected readonly = True TYPE=TEXT NAME="txtPayeeNm" ALT="지급처" tag="24X"></TD>
									<TD CLASS="TD5" NOWRAP>화폐</TD>
									<TD CLASS="TD6" NOWRAP>
									<Table cellpadding=0 cellspacing=0>
										   <TR>
											<TD>
									          <INPUT CLASS = required TYPE=TEXT AlT="화폐" NAME="txtCur" MAXLENGTH=3 SIZE=10 tag="23NXXU" onChange="ChangeCurr()" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnGroupCd2" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenCur()">
														   <INPUT TYPE=HIDDEN AlT="화폐" NAME="txtCurNm" tag="24X" CLASS = protected readonly = True TabIndex = -1 >
											</TD>			   
											  <TD NOWRAP>
												&nbsp;환율													
											  </TD>
											  <TD><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT=환율 NAME="txtXchRt" CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle2 style="HEIGHT: 20px; WIDTH: 70px" tag="21X5Z" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT></td>			
													<TD NOWRAP>
														&nbsp;<SELECT NAME="cboXchop" tag="22" STYLE="WIDTH:82px:" Alt="환율"></SELECT>
											  </TD>
											</TR>
										</Table>	
									</TD>

								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>매입단가조정순금액</TD>
									<TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT=발주금액 NAME="txtIvAmt" CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle1 style="HEIGHT: 20px; WIDTH: 250px" tag="24X2" Title="FPDOUBLESINGLE" CLASS = protected readonly = True TabIndex = -1 ></OBJECT>');</SCRIPT></TD>
									<TD CLASS="TD5" NOWRAP>매입단가조정자국순금액</TD>
									<TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT=발주자국금액 NAME="txtIvLocAmt" CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle1 style="HEIGHT: 20px; WIDTH: 250px" tag="24X2" Title="FPDOUBLESINGLE" CLASS = protected readonly = True TabIndex = -1 ></OBJECT>');</SCRIPT></td>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>VAT</TD>
									<TD CLASS="TD6" NOWRAP>
										<Table cellpadding=0 cellspacing=0>	
											<TR>
												<TD><INPUT TYPE=TEXT CLASS = required NAME="txtVatCd" ALT="VAT"  MAXLENGTH=5 SIZE=10 ONChange="vbscript:SetVatType()" tag="21NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnGroupCd2" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenVat(1)">&nbsp;</TD>
												<TD><INPUT TYPE=TEXT AlT="VAT" NAME="txtVatNm" SIZE=15 tag="24X" CLASS = protected readonly = True TabIndex = -1 >&nbsp;</TD>
												<TD><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT=VAT율 NAME="txtVatrt" CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle3 style="HEIGHT: 20px; WIDTH: 50px" tag="24X5" Title="FPDOUBLESINGLE" CLASS = protected readonly = True TabIndex = -1 ></OBJECT>');</SCRIPT></TD>
												<TD>&nbsp;%</TD>
												
											</TR>
										</TABLE>															
									<TD CLASS="TD5" nowrap>VAT포함구분</TD>
								    <TD CLASS="TD6" nowrap>
								    <INPUT TYPE=radio NAME="rdoVatFlg" ALT="VAT포함구분" CLASS="RADIO" checked id="rdoVatFlg1" ONCLICK="vbscript:SetChangeflg(2)" tag="21X"><label for="rdoVatFlg">별도 </label>
									<INPUT TYPE=radio NAME="rdoVatFlg" ALT="VAT포함구분" CLASS="RADIO" id="rdoVatFlg2" ONCLICK="vbscript:SetChangeflg(2)" tag="21X"><label for="rdoVatFlg">포함&nbsp;</label></TD>
								</TR>					

								<TR>
								    <TD CLASS="TD5" NOWRAP>VAT조정금액</TD>
									<TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT=VAT조정금액 NAME="txtGrossVatAmt" CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle1 style="HEIGHT: 20px; WIDTH: 250px" tag="24X2" Title="FPDOUBLESINGLE" CLASS = protected readonly = True TabIndex = -1 ></OBJECT>');</SCRIPT></TD>
									<TD CLASS="TD5" NOWRAP>VAT조정자국금액</TD>
									<TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT=VAT조정자국금액 NAME="txtGrossVatLocAmt" CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle1 style="HEIGHT: 20px; WIDTH: 250px" tag="24X2" Title="FPDOUBLESINGLE" CLASS = protected readonly = True TabIndex = -1 ></OBJECT>');</SCRIPT></td>
								</TR>
								
								<TR>
									<TD CLASS="TD5" nowrap>구매그룹</TD>
									<TD CLASS="TD6" nowrap><INPUT TYPE=TEXT CLASS = required NAME="txtGrpCd" ALT="구매그룹" style="HEIGHT: 20px; WIDTH: 70px" MAXLENGTH=4 tag="22NXXU" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnGrp" style="HEIGHT: 21px; WIDTH: 16px" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenGrp()" >
														   <INPUT TYPE=TEXT CLASS = protected readonly = True NAME="txtGrpNm" ALT="구매그룹" style="HEIGHT: 20px; WIDTH: 150px" tag="24X"></TD>
									<TD CLASS="TD5" nowrap>세금신고사업장</TD>
									<TD CLASS="TD6" nowrap><INPUT TYPE=TEXT NAME="txtBizAreaCd" ALT="세금신고사업장" style="HEIGHT: 20px; WIDTH: 70px" MAXLENGTH=10 tag="21NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" style="HEIGHT: 21px; WIDTH: 16px" NAME="btnPayMethod" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenBizArea()">
														   <INPUT TYPE=TEXT CLASS = protected readonly = True NAME="txtBizAreaNm" ALT="세금신고사업장" style="HEIGHT: 20px; WIDTH: 150px" tag="24X" ></TD>
								</TR>								
								<TR>
									<TD HEIGHT=100% WIDTH=100% COLSPAN=4>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=OBJECT1> <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
									</TD>
								</TR>	
															
							</TABLE>
							
						</div>
						<!--두번째 탭 -->
						<DIV ID="TabDiv" SCROLL=no>
							<TABLE <%=LR_SPACE_TYPE_60%>>
								<TR> 
									<TD CLASS="TD5" NOWRAP>사업자등록번호</TD>
									<TD CLASS="TD6" nowrap><INPUT TYPE=TEXT AlT="사업자등록번호" NAME="txtSpplRegNo" SIZE=20 tag="24X" CLASS = protected readonly = True TabIndex = -1 ></TD>
									<TD CLASS="TD5" NOWRAP>공급처 INVOICE NO.</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtSpplIvNo" ALT="공급처 INVOICE NO."  style="HEIGHT: 20px; WIDTH:250px" MAXLENGTH=50 tag="21"></TD>
								</TR>									
								<TR>
									<TD CLASS="TD5" nowrap>결제방법</TD>
									<TD CLASS="TD6" nowrap><INPUT TYPE=TEXT NAME="txtPayTermCd" CLASS = required ALT="결제방법" style="HEIGHT: 20px; WIDTH: 80px" MAXLENGTH=5 tag="22NXXU" OnChange="VBScript:changePayMeth()"><IMG SRC="../../../CShared/image/btnPopup.gif" style="HEIGHT: 21px; WIDTH: 16px" NAME="btnPayMethod" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPayMeth()">
														   <INPUT TYPE=TEXT NAME="txtPayTermNm" CLASS = protected readonly = True ALT="결제방법" style="HEIGHT: 20px; WIDTH: 150px" tag="24X" ></TD>
									<TD CLASS="TD5" nowrap>결제기간</TD>
									<TD CLASS="TD6" NOWRAP>
										<Table cellpadding=0 cellspacing=0>
											<TR>
												<TD NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT=결제기간 NAME="txtPayDur" CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle5 style="HEIGHT: 20px; WIDTH: 80px" tag="21X70" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT>
												</TD>
												<TD NOWRAP>
													&nbsp;일 
												</TD>
											</TR>
										</Table>
									</TD>				   
								</TR>
								<TR>
									<TD CLASS="TD5" nowrap>지급유형</TD>
									<TD CLASS="TD6" nowrap><INPUT TYPE=TEXT NAME="txtPayTypeCd" ALT="지급유형" style="HEIGHT: 20px; WIDTH: 80px" MAXLENGTH=5 tag="21NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" style="HEIGHT: 21px; WIDTH: 16px" NAME="btnPayMethod" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPayType()">
														   <INPUT TYPE=TEXT AlT="지급유형" NAME="txtPayTypeNm" SIZE=20 tag="24X" CLASS = protected readonly = True TabIndex = -1 ></TD>
									<TD CLASS="TD5" nowrap>
									<TD CLASS="TD6" nowrap>				   
								</TR>	
								<TR>
									<TD CLASS="TD5">대금결제참조</TD>
									<TD CLASS="TD6" colspan=3 width=100% NOWRAP><INPUT TYPE=TEXT AlT="대금결제참조" Size="90" NAME="txtPayTermstxt" MAXLENGTH=120 tag="21"></TD>
								</TR>
								<TR>
									<TD CLASS="TD5">비고</TD>
									<TD CLASS=TD6 Colspan=3 WIDTH=100% NOWRAP><INPUT TYPE=TEXT  NAME="txtRemark" ALT="비고" tag = "21" SIZE=90 MAXLENGTH=70></TD>
								</TR>
							</TABLE>
						</DIV>
						
					</TD>	
				</TR>
			</table>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_01%>></TD>
	</TR>
    <tr HEIGHT="20">
		<td WIDTH="100%">
			<table <%=LR_SPACE_TYPE_30%>>
				<tr>
					<TD WIDTH=10>&nbsp;</TD>
					<td >
					<button name="btnCfmSel" id="btnCfm" class="clsmbtn" ONCLICK="vbscript:Cfm()">확정</button>&nbsp;&nbsp;
					 <Div  STYLE="DISPLAY: none"><button name="btnSend" id="btnSend" class="clsmbtn" ONCLICK="Sending()">주문서발송</button></Div>
					<button name="btnSelect" class="clsmbtn" ONCLICK="OpenGlRef()">전표조회</button>&nbsp;
					</td>   
					<TD WIDTH=10>&nbsp;</TD>
				</tr>
			</table>
		</td>
    </tr>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX=-1></IFRAME>
		</TD>
	</TR>
</TABLE>
<!--	추가부분 시작	-->
<P ID="divTextArea"></P>
<!--	추가부분 끝	    -->
<TEXTAREA CLASS="hidden"  NAME="txtSpread" tag="24" TABINDEX="-1"></TEXTAREA>

<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24"><INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">

<INPUT TYPE=HIDDEN NAME="hdnCurr" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">

<INPUT TYPE=HIDDEN NAME="hdnreference"  tag="14">
<INPUT TYPE=HIDDEN NAME="hdnImportflg" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnBLflg" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnCCflg" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnRcptflg" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnSubcontraflg" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnRetflg" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnIvflg" tag="24">
<INPUT TYPE=HIDDEN NAME="hdvatFlg" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnRcptType" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnIssueType" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnIvType" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnMergPurFlg" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaintNo" tag="24">
<INPUT TYPE=HIDDEN NAME="hdclsflg" tag="24">
<INPUT TYPE=HIDDEN NAME="hdntotPoAmt" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnVATINCFLG" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnxchrateop" tag="2">
<!-- 20031117-->
<INPUT TYPE=HIDDEN NAME="hdnMaxRows" tag="14">
<INPUT TYPE=HIDDEN NAME="hdntxtVatCd" tag="14">
<INPUT TYPE=HIDDEN NAME="hdntxtVatNm" tag="14">
<INPUT TYPE=HIDDEN NAME="hdntxtVatrt" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnSupplierCd" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnPoDt" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnchgValue"  tag="14">
<INPUT TYPE=HIDDEN NAME="hdnSSCheckValue" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnIvAmt" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnIvLocAmt" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnGrossVatAmt" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnGrossVatLocAmt" tag="24">



<INPUT TYPE=HIDDEN NAME="hdnIvNo" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnRelease" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnVatFlg" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnGlType" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnGlNo" tag="24">


<!-- 화면디자인을 위해 추가 -->

<INPUT TYPE=HIDDEN NAME="txtReference" tag="14">








</FORM>
<DIV ID="MousePT" NAME="MousePT" STYLE="visible:true">
	<IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm"></IFRAME>
</DIV>
</BODY>
</HTML>