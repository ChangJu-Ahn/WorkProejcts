<%Option Explicit%>
<%
'**********************************************************************************************
'*  1. Module Name          : MES 
'*  2. Function Name        : 제품재고현황 
'*  3. Program ID           : xi315ma1_KO119
'*  4. Program Name         : 제품재고현황(한솔)
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2000/04/17
'*  8. Modified date(Last)  : 2002/04/11
'*  9. Modifier (First)     : 
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*                            -2000/04/17 : 화면 Layout & ASP Coding
'*                            -2001/12/19 : Date 표준적용 
'*                            -2002/04/11 : ADO 변환 
'**********************************************************************************************

'								'☜ : ASP가 캐쉬되지 않도록 한다.
'								'☜ : ASP가 버퍼에 저장되지 않고 바로 Client에 내려간다.

'☜ : 항상 서버 사이드 구문의 시작점인 좌꺽쇠(<)% 와 %우꺽쇠(>)는 New Line에 위치하여 
'	  서버 사이드 구문과 클라이언트 사이드 구문의 위치를 가늠할 수 있도록 한다.
'☜ : 아래 HTML 구문은 변경되어서는 안된다. 
%>
<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<!-- #Include file="../../inc/IncSvrNumber.inc"  -->
<!-- #Include file="../../inc/IncSvrDate.inc"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<!-- #Include file="../../inc/incSvrDBAgent.inc"  -->
<!-- #Include file="../../inc/incSvrDBAgentVariables.inc"  -->

<%													'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 

'On Error Resume Next
'err.clear

Call LoadBasisGlobalInf()
Call LoadInfTB19029B("Q","*", "NOCOOKIE", "QB") 

'============================================  2002-04-10 시작  =============================================
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3, rs4, rs5, rs6, rs7, rs8, rs9, rs10, rs11
															'☜ : DBAgent Parameter 선언 
Dim lgStrData                                               '☜ : Spread sheet에 보여줄 데이타를 위한 변수 
Dim lgMaxCount                                              '☜ : Spread sheet 의 visible row 수 
Dim lgTailList
Dim lgSelectList
Dim lgSelectListDT
Dim lgDataExist
Dim lgPageNo
Dim strQryMode

'--------------- 개발자 coding part(변수선언,Start)--------------------------------------------------------
DIM strPlantNm		'공장이름
DIM strSecItemNm	'SEC품목명
Dim MsgDisplayFlag
Dim lgADF
Dim lgstrRetMsg

Dim lgtotMQty          '양산입고
Dim lgtotSampleQty     'Sample입고 
Dim lgtotMOutQty       '양산출고 
Dim lgtotSampleOutQty  'Sample출고 
Dim lgtotVOutQty       '가상출고  

Dim lgtotMHoldQty      '양산Hold재고 
Dim lgtotSHoldQty      'Sample Hold재고 
Dim lgtotMUseQty       '양산가용재고 
Dim lgtotSUseQty       'Sample가용재고 

															'☜ : 받을 레코드셋의 갯수만큼 배열 크기 선언			
	MsgDisplayFlag = False															
'--------------- 개발자 coding part(변수선언,End)----------------------------------------------------------
   	
	Call LoadBasisGlobalInf()
	Call HideStatusWnd
	
    lgPageNo         = UNICInt(Trim(Request("lgPageNo")),0)              '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
    lgMaxCount       = 100						                       '☜ : 한번에 가져올수 있는 데이타 건수 
    lgSelectList     = Request("lgSelectList")
    lgTailList       = Request("lgTailList")
    lgSelectListDT   = Split(Request("lgSelectListDT"), gColSep)         '☜ : 각 필드의 데이타 타입 
    lgDataExist      = "No"
	strQryMode		 = Request("lgIntFlgMode")

    Call  FixUNISQLData1()                                                '☜ : DB-Agent로 보낼 parameter 데이타 set
    Call  FixUNISQLData()
    call  QueryData()                                                               '☜ : DB-Agent를 통한 ADO query

'/////////////////////////////////////////////////////////////////////////////////////////////
'Dim lgADF                                                   '☜ : ActiveX Data Factory 지정 변수선언 
'Dim lgstrRetMsg                                             '☜ : Record Set Return Message 변수선언 
'Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0				'☜ : DBAgent Parameter 선언 
'Dim rs1, rs2, rs3, rs4, rs5, rs6							'☜ : DBAgent Parameter 선언 
'Dim lgStrData                                               '☜ : Spread sheet에 보여줄 데이타를 위한 변수 
'Dim lgStrPrevKey                                            '☜ : 이전 값 
'Dim lgMaxCount                                              '☜ : Spread sheet 의 visible row 수 
'Dim lgTailList
'Dim lgSelectList
'Dim lgSelectListDT
'============================================  2002-04-10 끝  ===============================================
'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData1()

	Dim strVal
	Dim strVal2
	Dim strVal3
	Redim UNISqlId(1)                                                     '☜: SQL ID 저장을 위한 영역확보 
'--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------	
	Redim UNIValue(1, 1)

	UNISqlId(0) = "180000saa"    ' 공장코드
	UNISqlId(1) = "commonqry"    'SEC품목코드		
				
	'--------------- 개발자 coding part(실행로직,End)------------------------------------------------------

	UNIValue(0,0) = Trim(lgSelectList)		                              '☜: Select 절에서 Summary    필드 

	'--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
		          		   
	'공장코드
    UNIValue(0,0) = FilterVar(UCase(Request("txtPlantCd")), "''", "S")					'공장코드
    'SEC품목명
    UNIValue(1,0)  = "SELECT FR2_CD PAR_ITEM_NM from J_CODE_MAPPING (nolock) where major_Cd = 'J0010' and minor_cd = '0000' and FR1_CD = " & FilterVar(Request("txtSecItemCd"), "''", "S") 	

	'--------------- 개발자 coding part(실행로직,End)----------------------------------------------------
'	UNIValue(0,UBound(UNIValue,2)    ) = Trim(lgTailList)	'---Order By 조건 
	UNILock = DISCONNREAD :	UNIFlag = "1"                                 '☜: set ADO read mode
		   
    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    
	lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs1, rs2)
    
   	If SetConditionData = False Then Exit Sub

End Sub



' Make srpread sheet data
'----------------------------------------------------------------------------------------------------------
Sub MakeSpreadSheetData()

    Dim iLoopCount
    Dim iRowStr
    Dim ColCnt
    
    lgDataExist    = "Yes"
    lgstrData      = ""
  
    If CLng(lgPageNo) > 0 Then
       rs0.Move     = CLng(lgMaxCount) * CLng(lgPageNo)   'lgMaxCount:Max Fetched Count at once , lgStrPrevKeyIndex : Previous PageNo
    End If
    
    iLoopCount = -1
    
   Do while Not (rs0.EOF Or rs0.BOF)
   
        iLoopCount =  iLoopCount + 1
        iRowStr = ""
        
		For ColCnt = 0 To UBound(lgSelectListDT) - 1 
            iRowStr = iRowStr & Chr(11) & FormatRsString(lgSelectListDT(ColCnt),rs0(ColCnt))
		Next
 
        If iLoopCount < lgMaxCount Then
           lgstrData = lgstrData & iRowStr & Chr(11) & Chr(12)
        Else
           lgPageNo = lgPageNo + 1
           Exit Do
        End If
        
        rs0.MoveNext
	Loop

    If iLoopCount < lgMaxCount Then                                      '☜: Check if next data exists
       lgPageNo = ""
    End If        
    rs0.Close                                                       '☜: Close recordset object
    Set rs0 = Nothing	                                            '☜: Release ADF

End Sub

'----------------------------------------------------------------------------------------------------------
' Name : SetConditionData
' Desc : set value in condition area
'----------------------------------------------------------------------------------------------------------
Function SetConditionData()
	
'	On Error Resume Next
 
	SetConditionData = False
    
    If Not(rs1.EOF Or rs1.BOF) Then
        strPlantNm =  rs1(1)
        rs1.Close	:	Set rs1 = Nothing 
    Else
		rs1.Close	:	Set rs1 = Nothing
		If Len(Request("txtPlantCd")) And MsgDisplayFlag = False Then
			Call DisplayMsgBox("125000", vbOKOnly, "", "", I_MKSCRIPT)
		    MsgDisplayFlag = True
            ' Modify Focus Events    
            %>
                <Script language=vbs>
                parent.frm1.txtPlantCd.Focus()   
                </Script>
            <% 
            Set lgADF = Nothing       		    	
		End If
	End If 
	
	
	If Not(rs2.EOF Or rs2.BOF) Then
        strSecItemNm =  rs2(0)
        rs2.Close	:	Set rs2 = Nothing 
    Else
		rs2.Close	:	Set rs2 = Nothing
		If Len(Request("txtSecItemCd")) And MsgDisplayFlag = False Then
			Call DisplayMsgBox("XX1137", vbOKOnly, "", "", I_MKSCRIPT)
		    MsgDisplayFlag = True
            ' Modify Focus Events    
            %>
                <Script language=vbs>
                parent.frm1.txtSecItemCd.Focus()   
                </Script>
            <% 
            Set lgADF = Nothing       		    	
		End If
	End If 
	
	If MsgDisplayFlag = True Then Exit Function

	SetConditionData = True

End Function

'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()

	Dim strVal
	Dim strVal2
	Dim strVal3
	Dim strVal4
	Dim strVal5
	Redim UNISqlId(9)                                                     '☜: SQL ID 저장을 위한 영역확보 
'--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------	
  if Trim(Request("txtRadio")) = "MINV" or Trim(Request("txtRadio")) = "SINV" or Trim(Request("txtRadio")) = "MOUT" or Trim(Request("txtRadio")) = "SOUT" then
	Redim UNIValue(9, 3)
  else
    Redim UNIValue(9, 2)
  end if	

  if Trim(Request("txtRadio")) = "VOUT" then   '가상출고
	UNISqlId(0) = "xi315ma2_KO119"
  elseif Trim(Request("txtRadio")) = "MU" or Trim(Request("txtRadio")) = "SU" then  '양산가용/Sample가용
    UNISqlId(0) = "xi315ma3_KO119"
  elseif Trim(Request("txtRadio")) = "MH" or Trim(Request("txtRadio")) = "SH" then  '양산Hold재고/Sample Hold재고
    UNISqlId(0) = "xi315ma1_KO119"
  else
    UNISqlId(0) = "xi315ma4_KO119"      
  end if	
	UNISqlId(1) = "commonqry"    '양산입고		
	UNISqlId(2) = "commonqry"    'Sample입고	
	UNISqlId(3) = "commonqry"    '양산출고	
	UNISqlId(4) = "commonqry"    'Sample출고
	UNISqlId(5) = "commonqry"    '가상출고
	UNISqlId(6) = "commonqry"    '양산Hold재고	
	UNISqlId(7) = "commonqry"    'Sample Hold재고		 
	UNISqlId(8) = "commonqry"    '양산가용재고
	UNISqlId(9) = "commonqry"    'Sample가용재고
				
	'--------------- 개발자 coding part(실행로직,End)------------------------------------------------------

	UNIValue(0,0) = Trim(lgSelectList)		                              '☜: Select 절에서 Summary    필드 

	'--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
	
	strVal = " "
	strVal2 = " "
	strVal3 = " "
	strVal4 = " "
	strVal5 = " "
		
    '---공장코드(필수)
    If Len(Trim(Request("txtPlantCd"))) Then
    	strVal = strval & "  a.plant_cd =  " & FilterVar(Request("txtPlantCd"), "''", "S") & " " 
    	strVal2 = strval2 & "  a.plant_cd =  " & FilterVar(Request("txtPlantCd"), "''", "S") & " "    			
    	strVal3 = strval3 & "  a.plant_cd =  " & FilterVar(Request("txtPlantCd"), "''", "S") & " "    			
    End If
	
	'---생산일자(필수) 
    If Len(Trim(Request("txtPlanStartDt"))) Then
    	strVal = strval & " AND a.production_dt >=  " & FilterVar(uniConvDate(Trim(Request("txtPlanStartDt"))), "''", "S") & ""
    	strVal3 = strval3 & " AND a.production_dt >=  " & FilterVar(uniConvDate(Trim(Request("txtPlanStartDt"))), "''", "S") & ""
    	strVal4 = strval4 & " product_date >=  " & FilterVar(uniConvDate(Trim(Request("txtPlanStartDt"))), "''", "S") & ""
    End If
    
    If Len(Trim(Request("txtPlanEndDt"))) Then
    	strVal = strval & " AND a.production_dt <=  " & FilterVar(uniConvDate(Trim(Request("txtPlanEndDt"))), "''", "S") & ""
    	strVal3 = strval3 & " AND a.production_dt <=  " & FilterVar(uniConvDate(Trim(Request("txtPlanEndDt"))), "''", "S") & ""
    	strVal4 = strval4 & " AND product_date <=  " & FilterVar(uniConvDate(Trim(Request("txtPlanEndDt"))), "''", "S") & ""
    End If
    
    '---MES송신기간 
    If Len(Trim(Request("txtSendStartDt"))) Then
    	strVal = strval & " AND a.send_dt >=  " & FilterVar(uniConvDate(Trim(Request("txtSendStartDt"))), "''", "S") & ""
    	strVal3 = strval3 & " AND a.send_dt >=  " & FilterVar(uniConvDate(Trim(Request("txtSendStartDt"))), "''", "S") & ""
    End If
    
    If Len(Trim(Request("txtSendEndDt"))) Then
    	strVal = strval & " AND a.send_dt <=  " & FilterVar(uniConvDate(Trim(Request("txtSendEndDt"))), "''", "S") & ""
    	strVal3 = strval3 & " AND a.send_dt <=  " & FilterVar(uniConvDate(Trim(Request("txtSendEndDt"))), "''", "S") & ""
    End If
	
	'---SEC품목코드 
	If Len(Trim(Request("txtSecItemCd"))) Then
    	strVal5 = strval5 & " AND b.sec_item_cd =  " & FilterVar(Request("txtSecItemCd"), "''", "S") & " "
    End If    
	
	
	'---구분 
	If Trim(Request("txtRadio")) = "MU"  Then  '양산가용재고
'		strVal = strval & " AND ISNULL(A.SEC_INVOICE_NO,'') = '' AND UPPER(LEFT(A.PRODT_ORDER_NO,2)) <> 'PS' AND A.DELIVERY_HOLD_FG <> 'Y' and a.pallet_no not in (select pallet_no from S_Prev_Integrate_Lbl_Ko119)"
		strVal = strval & " AND ISNULL(A.SEC_INVOICE_NO,'') = '' AND UPPER(LEFT(A.PRODT_ORDER_NO,2)) <> 'PS' AND A.DELIVERY_HOLD_FG <> 'Y'"   	
	elseif Trim(Request("txtRadio")) = "SU" then  'Sample가용재고
'		strVal = strval & " AND ISNULL(A.SEC_INVOICE_NO,'') = '' AND UPPER(LEFT(A.PRODT_ORDER_NO,2)) = 'PS' AND A.DELIVERY_HOLD_FG <> 'Y' and a.pallet_no not in (select pallet_no from S_Prev_Integrate_Lbl_Ko119)"
		strVal = strval & " AND ISNULL(A.SEC_INVOICE_NO,'') = '' AND UPPER(LEFT(A.PRODT_ORDER_NO,2)) = 'PS' AND A.DELIVERY_HOLD_FG <> 'Y'"   	
	elseif Trim(Request("txtRadio")) = "MH" then  '양산Hold재고
	    strVal = strval & " AND UPPER(LEFT(A.PRODT_ORDER_NO,2)) <> 'PS' AND A.DELIVERY_HOLD_FG = 'Y' "
	elseif Trim(Request("txtRadio")) = "SH" then  'Sample Hold재고
		strVal = strval & " AND UPPER(LEFT(A.PRODT_ORDER_NO,2)) = 'PS' AND A.DELIVERY_HOLD_FG = 'Y' "  			
	elseif Trim(Request("txtRadio")) = "MINV" then  '양산입고
		strVal = strval & " AND UPPER(LEFT(A.PRODT_ORDER_NO,2)) <> 'PS'"
	elseif Trim(Request("txtRadio")) = "SINV" then  'Sample 입고
		strVal = strval & " AND UPPER(LEFT(A.PRODT_ORDER_NO,2)) = 'PS'"
	elseif Trim(Request("txtRadio")) = "MOUT" then  '양산출고
		strVal = strval & " AND UPPER(LEFT(A.PRODT_ORDER_NO,2)) <> 'PS' AND ISNULL(A.SEC_INVOICE_NO,'') <> '' "
	elseif Trim(Request("txtRadio")) = "SOUT" then  'Sample출고
		strVal = strval & " AND UPPER(LEFT(A.PRODT_ORDER_NO,2)) = 'PS' AND ISNULL(A.SEC_INVOICE_NO,'') <> '' "
	elseif Trim(Request("txtRadio")) = "VOUT" then  '가상출고
	    	
    End If   	    
     
   if Trim(Request("txtRadio")) = "MINV" or Trim(Request("txtRadio")) = "SINV" or Trim(Request("txtRadio")) = "MOUT" or Trim(Request("txtRadio")) = "SOUT" then
      UNIValue(0,1) = strVal
      UNIValue(0,2) = strVal4
	  UNIValue(0,3) = strVal5 
   else 		   
		UNIValue(0,1) = strVal
		UNIValue(0,2) = strVal5
   end if 
    
	'양산입고
	UNIValue(1,0)  = "SELECT sum(isnull(A.tray_item_qty,0)) from T_IF_RCV_TRAY_INFO_KO119 A (nolock) inner join b_item_mapping_ko119 b (nolock) on a.item_cd = b.item_cd where " & strVal3 & "AND UPPER(LEFT(A.PRODT_ORDER_NO,2)) <> 'PS'" & strVal5 & ""
	'Sample입고
	UNIValue(2,0)  = "SELECT sum(isnull(A.tray_item_qty,0)) from T_IF_RCV_TRAY_INFO_KO119 A (nolock) inner join b_item_mapping_ko119 b (nolock) on a.item_cd = b.item_cd where " & strVal3 & "AND UPPER(LEFT(A.PRODT_ORDER_NO,2)) = 'PS'"	& strVal5 & ""
	'양산출고
	UNIValue(3,0)  = "SELECT sum(isnull(A.tray_item_qty,0)) from T_IF_RCV_TRAY_INFO_KO119 A (nolock) inner join b_item_mapping_ko119 b (nolock) on a.item_cd = b.item_cd where " & strVal3 & "AND UPPER(LEFT(A.PRODT_ORDER_NO,2)) <> 'PS' AND ISNULL(A.SEC_INVOICE_NO,'') <> '' " & strVal5 & ""
	'Sample출고
	UNIValue(4,0)  = "SELECT sum(isnull(A.tray_item_qty,0)) from T_IF_RCV_TRAY_INFO_KO119 A (nolock) inner join  b_item_mapping_ko119 b (nolock) on a.item_cd = b.item_cd where " & strVal3 & "AND UPPER(LEFT(A.PRODT_ORDER_NO,2)) = 'PS' AND ISNULL(A.SEC_INVOICE_NO,'') <> '' " & strVal5 & ""
    '가상출고
    UNIValue(5,0)  = "SELECT sum(isnull(A.tray_item_qty,0)) from T_IF_RCV_TRAY_INFO_KO119 A (nolock) inner join b_item_mapping_ko119 b (nolock) on a.item_cd = b.item_cd inner join  S_Prev_Integrate_Lbl_Ko119  C on  A.PALLET_NO  =  C.PALLET_NO Where " & strVal3 & strVal5 & ""	
    '양산Hold재고
    UNIValue(6,0)  = "SELECT sum(isnull(A.tray_item_qty,0)) from T_IF_RCV_TRAY_INFO_KO119 A (nolock) inner join  b_item_mapping_ko119 b (nolock) on a.item_cd = b.item_cd where " & strVal3 & "AND UPPER(LEFT(A.PRODT_ORDER_NO,2)) <> 'PS' AND A.DELIVERY_HOLD_FG = 'Y' " & strVal5 & ""
	'Sample Hold 재고
	UNIValue(7,0)  = "SELECT sum(isnull(A.tray_item_qty,0)) from T_IF_RCV_TRAY_INFO_KO119 A (nolock) inner join  b_item_mapping_ko119 b (nolock) on a.item_cd = b.item_cd where " & strVal3 & "AND UPPER(LEFT(A.PRODT_ORDER_NO,2)) = 'PS' AND A.DELIVERY_HOLD_FG = 'Y' " & strVal5 & ""
    '양산가용재고
    UNIValue(8,0)  = "SELECT sum(isnull(A.tray_item_qty,0)) from T_IF_RCV_TRAY_INFO_KO119 A (nolock) inner join  b_item_mapping_ko119 b (nolock) on a.item_cd = b.item_cd "
	UNIValue(8,0)  = UNIValue(8,0) & "where " & strVal3 & "AND ISNULL(A.SEC_INVOICE_NO,'') = '' AND UPPER(LEFT(A.PRODT_ORDER_NO,2)) <> 'PS' AND A.DELIVERY_HOLD_FG <> 'Y' and a.pallet_no not in (select pallet_no from S_Prev_Integrate_Lbl_Ko119 (nolock)) " & strVal5 & ""
	'샘플가용재고
	UNIValue(9,0)  = "SELECT sum(isnull(A.tray_item_qty,0)) from T_IF_RCV_TRAY_INFO_KO119 A (nolock) inner join  b_item_mapping_ko119 b (nolock) on a.item_cd = b.item_cd "
	UNIValue(9,0)  = UNIValue(9,0) & "where " & strVal3 & "AND ISNULL(A.SEC_INVOICE_NO,'') = '' AND UPPER(LEFT(A.PRODT_ORDER_NO,2)) = 'PS' AND A.DELIVERY_HOLD_FG <> 'Y' and a.pallet_no not in (select pallet_no from S_Prev_Integrate_Lbl_Ko119 (nolock)) " & strVal5 & ""
			

'	UNIValue(3,0)  = "SELECT sum(isnull(A.tray_item_qty,0)) from T_IF_RCV_TRAY_INFO_KO119 A (nolock) where " & strVal3 & "AND UPPER(LEFT(A.PRODT_ORDER_NO,2)) <> 'PS' "
'	UNIValue(4,0)  = "SELECT sum(isnull(A.tray_item_qty,0)) from T_IF_RCV_TRAY_INFO_KO119 A (nolock) where " & strVal3 & "AND UPPER(LEFT(A.PRODT_ORDER_NO,2)) = 'PS' "
'	UNIValue(5,0)  = "SELECT sum(isnull(A.tray_item_qty,0)) from T_IF_RCV_TRAY_INFO_KO119 A (nolock) where " & strVal3 & ""
'	UNIValue(6,0)  = "SELECT sum(isnull(A.tray_item_qty,0)) from T_IF_RCV_TRAY_INFO_KO119 A (nolock) where " & strVal3 & "AND ISNULL(A.SEC_INVOICE_NO,'') <> '' "
'	UNIValue(7,0)  = "SELECT sum(isnull(A.tray_item_qty,0)) from T_IF_RCV_TRAY_INFO_KO119 A (nolock) where " & strVal2 & "AND ISNULL(A.SEC_INVOICE_NO,'') = '' AND UPPER(LEFT(A.PRODT_ORDER_NO,2)) <> 'PS' AND A.DELIVERY_HOLD_FG <> 'Y' "
'	UNIValue(8,0)  = "SELECT sum(isnull(A.tray_item_qty,0)) from T_IF_RCV_TRAY_INFO_KO119 A (nolock) where " & strVal3 & "AND A.DELIVERY_HOLD_FG = 'Y' "
'	UNIValue(9,0)  = "SELECT sum(isnull(A.tray_item_qty,0)) from T_IF_RCV_TRAY_INFO_KO119 A (nolock) inner join  S_Prev_Integrate_Lbl_Ko119  C on  A.PALLET_NO  =  C.PALLET_NO Where " & strVal3 & ""

	'--------------- 개발자 coding part(실행로직,End)----------------------------------------------------
'	UNIValue(0,UBound(UNIValue,2)    ) = Trim(lgTailList)	'---Order By 조건 
	UNILock = DISCONNREAD :	UNIFlag = "1"                                 '☜: set ADO read mode

End Sub
'----------------------------------------------------------------------------------------------------------
' Query Data
' ADO의 Record Set이용하여 Query를 하고 Record Set을 넘겨서 MakeSpreadSheetData()으로 Spreadsheet에 데이터를 
' 뿌림 
' ADO 객체를 생성할때 prjPublic.dll파일을 이용한다.(상세내용은 vb로 작성된 prjPublic.dll 소스 참조)
'----------------------------------------------------------------------------------------------------------
'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub QueryData()

	Dim lgstrRetMsg                                             '☜ : Record Set Return Message 변수선언 
	Dim lgADF                                                   '☜ : ActiveX Data Factory 지정 변수선언 
	Dim iStr
	    
	Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    
	lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs3, rs4, rs5, rs6, rs7, rs8, rs9, rs10, rs11)
	    
	Set lgADF   = Nothing

	iStr = Split(lgstrRetMsg,gColSep)
   
   
	If iStr(0) <> "0" Then
		Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
	End If    
		

   	IF NOT (rs3.EOF or rs3.BOF) then          '양산입고
		lgtotMQty     = rs3(0)				
	ELSE
		lgtotMQty     = ""
	End if
    rs3.Close
    Set rs3 = Nothing 

    IF NOT (rs4.EOF or rs4.BOF) then			'Sample입고
		lgtotSampleQty = rs4(0)			
	ELSE
		lgtotSampleQty = ""
	End if
    rs4.Close
    Set rs4 = Nothing 
    
    IF NOT (rs5.EOF or rs5.BOF) then		'양산출고
		lgtotMOutQty = rs5(0)			
	ELSE
		lgtotMOutQty = ""
	End if
    rs5.Close
    Set rs5 = Nothing
        
    IF NOT (rs6.EOF or rs6.BOF) then		'Sample출고
		lgtotSampleOutQty = rs6(0)			
	ELSE
		lgtotSampleOutQty = ""
	End if
    rs6.Close
    Set rs6 = Nothing  
    
        
    IF NOT (rs7.EOF or rs7.BOF) then		'가상출고
		lgtotVOutQty = rs7(0)				
	ELSE
		lgtotVOutQty = ""
	End if
    rs7.Close
    Set rs7 = Nothing      


    IF NOT (rs8.EOF or rs8.BOF) then		'양산Hold재고
		lgtotMHoldQty = rs8(0)				
	ELSE
		lgtotMHoldQty = ""
	End if
    rs8.Close
    Set rs8 = Nothing   
    
    
    IF NOT (rs9.EOF or rs9.BOF) then		'Sample Hold재고
		lgtotSHoldQty = rs9(0)				
	ELSE
		lgtotSHoldQty = ""
	End if
    rs9.Close
    Set rs9 = Nothing   
    
    IF NOT (rs10.EOF or rs10.BOF) then		'양산가용재고
		lgtotMUseQty = rs10(0)				
	ELSE
		lgtotMUseQty = ""
	End if
   rs10.Close
    Set rs10 = Nothing   
    
    IF NOT (rs11.EOF or rs11.BOF) then		'Sample가용재고
		lgtotSUseQty = rs11(0)				
	ELSE
		lgtotSUseQty = ""
	End if
	   rs11.Close
    Set rs11 = Nothing     

	If rs0.EOF And rs0.BOF Then
		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)		'No Data Found!!
		rs0.Close
		Set rs0 = Nothing
		MsgDisplayFlag = True
        ' Modify Focus Events    
        %>
            <Script language=vbs>
            Parent.frm1.txtPlantCd.focus    
            </Script>
        <%
	Else
		Call  MakeSpreadSheetData()
	End If

'	Call  SetConditionData()
    
End Sub

%>
<Script Language=vbscript>

With Parent 	

	.frm1.txtPlantNm.value			= "<%=ConvSPChars(strPlantNm)%>"
	.frm1.txtSecItemNm.value			= "<%=ConvSPChars(strSecItemNm)%>"
	
	.frm1.txtMassSumQty.text     = "<%=UNINumClientFormat(lgtotMQty,ggQty.DecPoint, 0)%>"      '양산입고 
    .frm1.txtSampleSumQty.text		= "<%=UNINumClientFormat(lgtotSampleQty,ggQty.DecPoint, 0)%>"      'Sample입고 
    .frm1.txtMOutSumQty.text    = "<%=UNINumClientFormat(lgtotMOutQty,ggQty.DecPoint, 0)%>"      '양산출고 
    .frm1.txtSampleOutSumQty.text	= "<%=UNINumClientFormat(lgtotSampleOutQty,ggQty.DecPoint, 0)%>"          'Sample출고 
    .frm1.txtVOutSumQty.text		= "<%=UNINumClientFormat(lgtotVOutQty,ggQty.DecPoint, 0)%>"			  '가상출고 
    .frm1.txtMHoldSumQty.text = "<%=UNINumClientFormat(lgtotMHoldQty,ggQty.DecPoint, 0)%>"      '양산Hold재고 
    .frm1.txtSampleHoldSumQty.text = "<%=UNINumClientFormat(lgtotSHoldQty,ggQty.DecPoint, 0)%>"      'Sample Hold재고 
	.frm1.txtMUseSumQty.text = "<%=UNINumClientFormat(lgtotMUseQty,ggQty.DecPoint, 0)%>"      '양산가용재고 
	.frm1.txtSampleUseSumQty.text = "<%=UNINumClientFormat(lgtotSUseQty,ggQty.DecPoint, 0)%>"      'Sample가용재고 

	If "<%=lgDataExist%>" = "Yes" Then
       'Set condition data to hidden area				
		If "<%=lgPageNo%>" = "1" Then   ' "1" means that this query is first and next data exists
			.frm1.txthPlantCd.value			= "<%=ConvSPChars(Request("txtPlantCd"))%>"		
			.frm1.txthSecItemCd.value		= "<%=ConvSPChars(Request("txtSecItemCd"))%>"									
			.frm1.txthPlanStartDt.value		= "<%=ConvSPChars(Request("txtPlanStartDt"))%>"
			.frm1.txthPlanEndDt.value		= "<%=ConvSPChars(Request("txtPlanEndDt"))%>"
			.frm1.txthSendStartDt.value		= "<%=ConvSPChars(Request("txtSendStartDt"))%>"
			.frm1.txthSendEndDt.value		= "<%=ConvSPChars(Request("txthSendEndDt"))%>"			
		End If      
		
		With parent
			.ggoSpread.Source = .frm1.vspdData 
			.frm1.vspdData.Redraw = False
			.ggoSpread.SSShowData "<%=lgstrData%>"									'☜: Display data 
			.lgPageNo =  "<%=ConvSPChars(lgPageNo)%>"								'☜: set next data tag
			.frm1.vspdData.Redraw = True
			.DbQueryOk
		End with 
	else
	    parent.dbquerynotok	    	      
    End If  
    
End With     
</Script>
