<%'======================================================
'*  1. Module Name          : 영업 
'*  2. Function Name        : 매출채권관리 
'*  3. Program ID           : S5112QA1
'*  4. Program Name         : 매출채권상세조회 
'*  5. Program Desc         : ADO Query
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2000/12/09
'*  8. Modified date(Last)  : 2001/12/18
'*  9. Modifier (First)     : Byun Jee Hyun
'* 10. Modifier (Last)      : Kim Hyungsuk
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*                            2000/12/09
'*                            2001/12/18	Date표준적용 
'=======================================================
%>
<!-- #Include file="../../inc/incSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDBAgent.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%                                                                         '☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 

On Error Resume Next

Dim lgADF                                                                  '☜ : ActiveX Data Factory 지정 변수선언 
Dim lgstrRetMsg                                                            '☜ : Record Set Return Message 변수선언 
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0                              '☜ : DBAgent Parameter 선언 
Dim lgstrData                                                              '☜ : data for spreadsheet data
Dim lgPageNo                                                           '☜ : 이전 값 
Dim lgMaxCount                                                             '☜ : 한번에 가져올수 있는 데이타 건수 
Dim lgTailList                                                             '☜ : Orderby절에 사용될 field 리스트 
Dim lgSelectList
Dim lgSelectListDT
'--------------- 개발자 coding part(변수선언,Start)--------------------------------------------------------
Dim strPoType	                                                           '⊙ : 발주형태 
Dim strPoFrDt	                                                           '⊙ : 발주일 
Dim strPoToDt	                                                           '⊙ :
Dim strSpplCd	                                                           '⊙ : 공급처 
Dim strPurGrpCd	                                                           '⊙ : 구매그룹 
Dim strItemCd	                                                           '⊙ : 품목 
Dim strTrackNo	                                                           '⊙ : Tracking No
Dim arrRsVal(9)
Dim lgDataExist
Dim FalsechkFlg

Dim iFrPoint
iFrPoint=0
'--------------- 개발자 coding part(변수선언,End)----------------------------------------------------------
  
    Call HideStatusWnd 
	Call LoadBasisGlobalInf()
	Call LoadInfTB19029B("Q", "S", "NOCOOKIE", "QB")
	Call LoadBNumericFormatB("Q", "S", "NOCOOKIE", "QB")

	lgPageNo       = UNICInt(Trim(Request("lgPageNo")),0)				   '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
    lgMaxCount     = 100							                       '☜ : 한번에 가져올수 있는 데이타 건수 
    lgSelectList   = Request("lgSelectList")                               '☜ : select 대상목록 
    lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)             '☜ : 각 필드의 데이타 타입 
    lgTailList     = Request("lgTailList")                                 '☜ : Orderby value


    Call FixUNISQLData()
    Call QueryData()
    
'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------

Sub MakeSpreadSheetData()

    Dim iLoopCount                                                                     
    Dim iRowStr
    Dim ColCnt
    
    lgDataExist    = "Yes"
    lgstrData      = ""
  
    If CLng(lgPageNo) > 0 Then
       rs0.Move     = CLng(lgMaxCount) * CLng(lgPageNo)                  'lgMaxCount:Max Fetched Count at once , lgPageNoIndex : Previous PageNo
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

    If iLoopCount < lgMaxCount Then                                 '☜: Check if next data exists
       lgPageNo = ""
    End If
    
    rs0.Close                                                       '☜: Close recordset object
    Set rs0 = Nothing	                                            '☜: Release ADF
    Set lgADF = Nothing                                                    '☜: ActiveX Data Factory Object Nothing
End Sub
'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()

    Dim strVal
    Dim arrVal(4)
    Redim UNISqlId(4)                                                     '☜: SQL ID 저장을 위한 영역확보 
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------

    Redim UNIValue(4,2)
     
    UNISqlId(0) = "S5112QA102"
    UNISqlId(1) = "s0000qa002"
    UNISqlId(2) = "s0000qa011"
    UNISqlId(3) = "S0000QA005"
    UNISqlId(4) = "S0000QA001"
     
    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    UNIValue(0,0) = lgSelectList                                          '☜: Select list
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------

	strVal = " "
	
	If Len(Request("txtSoldToParty")) Then
		strVal = "AND A.SOLD_TO_PARTY = " & FilterVar(Request("txtSoldToParty"), "''", "S") & " "		
	Else
		strVal = ""
	End If
	arrVal(0) = FilterVar(Request("txtSoldToParty"),"","S")

	If Len(Request("txtBillType")) Then
		strVal = strVal & " AND A.BILL_TYPE = " & FilterVar(Request("txtBillType"), "''", "S") & " "				
	End If		
	arrVal(1) = FilterVar(Request("txtBillType"),"","S")
		   
	If Len(Request("txtSalesGrp")) Then
		strVal = strVal & " AND A.SALES_GRP = " & FilterVar(Request("txtSalesGrp"), "''", "S") & " "			
	End If		
	arrVal(2) = FilterVar(Request("txtSalesGrp"),"","S")
    
	If Len(Request("txtItemCd")) Then
		strVal = strVal & " AND B.Item_Cd = " & FilterVar(Request("txtItemCd"), "''", "S") & " "				
	End If		
	arrVal(4) = FilterVar(Request("txtItemCd"),"","S")
    
        If Len(Request("txtBillFrDt")) Then
		strVal = strVal & " AND A.BILL_DT >= " & FilterVar(UNIConvDate(Request("txtBillFrDt")), "''", "S") & ""		
	End If		
	
	If Len(Request("txtBillToDt")) Then
		strVal = strVal & " AND A.BILL_DT <= " & FilterVar(UNIConvDate(Request("txtBillToDt")), "''", "S") & ""		
	End If

        If Request("txtRadio") = "Y" Then
		strVal = strVal & "AND A.POST_FLAG = " & FilterVar("Y", "''", "S") & " "
	ElseIf Request("txtRadio") = "N" Then
		strVal = strVal & "AND A.POST_FLAG = " & FilterVar("N", "''", "S") & " "			
	End If
	If Len(Request("txtProjectCd")) Then
		strVal = strVal & " AND B.TRACKING_NO = " & FilterVar(Request("txtProjectCd"), "''", "S") & " "				
	End If		
    
    UNIValue(0,1) = strVal   '---발주일 
    UNIValue(1,0) = arrVal(0)
    UNIValue(2,0) = arrVal(1)
    UNIValue(3,0) = arrVal(2)
	UNIValue(4,0) = arrVal(4)
    
    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    UNIValue(0,UBound(UNIValue,2)) = UCase(Trim(lgTailList))
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '☜: set ADO read mode
 
End Sub
'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub QueryData()
    Dim iStr
    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3, rs4)
	
	FalsechkFlg = False 
   
    If  rs1.EOF And rs1.BOF Then
        rs1.Close
        Set rs1 = Nothing

		If Len(Request("txtSoldToParty")) And FalsechkFlg = False Then
		   Call DisplayMsgBox("970000", vbInformation, "주문처", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code
	       Response.Write "<Script language=vbs>  " & vbCr   
           Response.Write "   Parent.frm1.txtSoldToParty.focus " & vbCr    
           Response.Write "</Script>      " & vbCr
	       FalsechkFlg = True
		End If	

    Else    
		arrRsVal(0) = rs1(0)
		arrRsVal(1) = rs1(1)
        rs1.Close
        Set rs1 = Nothing
    End If
    
    If  rs2.EOF And rs2.BOF Then
        rs2.Close
        Set rs2 = Nothing

		If Len(Request("txtBillType")) And FalsechkFlg = False Then
		   Call DisplayMsgBox("970000", vbInformation, "매출채권형태", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code
		   Response.Write "<Script language=vbs>  " & vbCr   
           Response.Write "   Parent.frm1.txtBillType.focus " & vbCr    
           Response.Write "</Script>      " & vbCr    
	       FalsechkFlg = True
		End If	

    Else    
		arrRsVal(2) = rs2(0)
		arrRsVal(3) = rs2(1)
        rs2.Close
        Set rs2 = Nothing
    End If

    If  rs4.EOF And rs4.BOF Then
        rs4.Close
        Set rs4 = Nothing

		If Len(Request("txtItemCd")) And FalsechkFlg = False Then
		   Call DisplayMsgBox("970000", vbInformation, "품목", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code
		   Response.Write "<Script language=vbs>  " & vbCr
           Response.Write "   Parent.frm1.txtItemCd.focus " & vbCr    
           Response.Write "</Script>      " & vbCr    
	       FalsechkFlg = True
		End If	

    Else    
		arrRsVal(8) = rs4(0)
		arrRsVal(9) = rs4(1)
        rs4.Close
        Set rs4 = Nothing
    End If

    If  rs3.EOF And rs3.BOF Then
        rs3.Close
        Set rs3 = Nothing

		If Len(Request("txtSalesGrp")) And FalsechkFlg = False Then
		   Call DisplayMsgBox("970000", vbInformation, "영업그룹", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code
		   Response.Write "<Script language=vbs>  " & vbCr
           Response.Write "   Parent.frm1.txtSalesGrp.focus " & vbCr    
           Response.Write "</Script>      " & vbCr    
	       FalsechkFlg = True
		End If	

    Else    
		arrRsVal(4) = rs3(0)
		arrRsVal(5) = rs3(1)
        rs3.Close
        Set rs3 = Nothing
    End If

    iStr = Split(lgstrRetMsg,gColSep)
    
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If    
        
    If  rs0.EOF And rs0.BOF And FalsechkFlg = False Then
        Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
        rs0.Close
        Set rs0 = Nothing
        Response.Write "<Script language=vbs>  " & vbCr   
        Response.Write "   Parent.frm1.txtSoldToParty.focus " & vbCr    
        Response.Write "</Script>      " & vbCr
    Else    
        Call  MakeSpreadSheetData()
    End If
   
End Sub

%>

<Script Language=vbscript>

    With parent
        .ggoSpread.Source    = .frm1.vspdData 
        .frm1.vspdData.Redraw = False
		.ggoSpread.SSShowDataByClip "<%=lgstrData%>","F"          '☜ : Display data
        .frm1.txtSoldToPartyNm.value		= "<%=ConvSPChars(arrRsVal(1))%>" 
        .frm1.txtBillTypeNm.value	= "<%=ConvSPChars(arrRsVal(3))%>" 
        .frm1.txtSalesGrpNm.value	= "<%=ConvSPChars(arrRsVal(5))%>" 
        .frm1.txtItemNm.value		= "<%=ConvSPChars(arrRsVal(9))%>"
        .frm1.txtHSoldToParty.value	= "<%=ConvSPChars(Request("txtSoldToParty"))%>"
		.frm1.txtHBillType.value	= "<%=ConvSPChars(Request("txtBillType"))%>"
		.frm1.txtHSalesGrp.value	= "<%=ConvSPChars(Request("txtSalesGrp"))%>"
		.frm1.txtHBillFrDt.value	= "<%=Request("txtBillFrDt")%>"
		.frm1.txtHItemCd.value		= "<%=ConvSPChars(Request("txtItemCd"))%>"
		.frm1.txtHBillToDt.value	= "<%=Request("txtBillToDt")%>"
        .frm1.txtHRadio.value	= "<%=Request("txtRadio")%>"
        .lgPageNo        =  "<%=lgPageNo%>"                       '☜: set next data tag
		If "<%=FalsechkFlg%>" = False Then
        .DbQueryOk
		End If
        .frm1.vspdData.Redraw = True                
	End with

</Script>	

<%
    Response.End													'☜: 비지니스 로직 처리를 종료함 
%>
