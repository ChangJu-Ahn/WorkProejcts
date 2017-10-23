<%
'**********************************************************************************************
'*  1. Module Name          : Sales
'*  2. Function Name        : 
'*  3. Program ID           : s1111qb1
'*  4. Program Name         : 품목단가조회 
'*  5. Program Desc         : 품목단가조회 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2002/07/04
'*  8. Modified date(Last)  : 2002/07/04
'*  9. Modifier (First)     : SonBumYeol		
'* 10. Modifier (Last)      : SonBumYeol
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDBAgent.inc" -->
<%                                                                         '☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 

On Error Resume Next

Dim lgADF                                                                  '☜ : ActiveX Data Factory 지정 변수선언 
Dim lgstrRetMsg                                                            '☜ : Record Set Return Message 변수선언 
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3               '☜ : DBAgent Parameter 선언 
Dim lgstrData                                                              '☜ : data for spreadsheet data
Dim lgStrPrevKey                                                           '☜ : 이전 값 
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
Dim BlankchkFlg

Dim iFrPoint
iFrPoint=0
'--------------- 개발자 coding part(변수선언,End)----------------------------------------------------------
  
Call LoadBasisGlobalInf()
Call LoadInfTB19029B("Q", "S", "NOCOOKIE", "QB")
Call LoadBNumericFormatB("Q", "S", "NOCOOKIE", "QB")

Call HideStatusWnd 


lgStrPrevKey   = Request("lgStrPrevKey")                               '☜ : Next key flag
lgMaxCount     = CInt(100)
lgSelectList   = Request("lgSelectList")                               '☜ : select 대상목록 
lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)             '☜ : 각 필드의 데이타 타입 
lgTailList     = Request("lgTailList")                                 '☜ : Orderby value

Call TrimData()
Call FixUNISQLData()
Call QueryData()
    
'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------

Sub MakeSpreadSheetData()
    Dim  RecordCnt
    Dim  ColCnt
    Dim  iCnt
    Dim  iRCnt
    Dim  iStr
    Dim strPriceFlagVal

    iCnt = 0
    lgstrData = ""

    If Len(Trim(lgStrPrevKey)) Then                                        '☜ : Chnage Nextkey str into int value
       If Isnumeric(lgStrPrevKey) Then
          iCnt = CInt(lgStrPrevKey)
       End If   
    End If   

	For iRCnt = 1 to iCnt  *  lgMaxCount                                   '☜ : Discard previous data
		rs0.MoveNext
		iFrPoint	= iCnt  *  lgMaxCount
    Next

    iRCnt = -1
    
    Do while Not (rs0.EOF Or rs0.BOF)
        iRCnt =  iRCnt + 1
        iStr = ""

		For ColCnt = 0 To UBound(lgSelectListDT) - 1
			'단가구분의 값을 설정한다. 20050217 by HJO
			If ColCnt = 11 Then
				If Trim((rs0(ColCnt)))="T" Then 
				strPriceFlagVal=	"진단가"
				Else
				strPriceFlagVal=	"가단가"
				End IF
				iStr = iStr & Chr(11) & FormatRsString(lgSelectListDT(ColCnt),strPriceFlagVal)
			Else
				iStr = iStr & Chr(11) & FormatRsString(lgSelectListDT(ColCnt),rs0(ColCnt))
			End If
		Next
        
        If  iRCnt < lgMaxCount Then			
            lgstrData      = lgstrData      & iStr & Chr(11) & Chr(12)
        Else
            iCnt = iCnt + 1
            lgStrPrevKey = CStr(iCnt)
            Exit Do
        End If
        rs0.MoveNext
	Loop

    If  iRCnt < lgMaxCount Then                                            '☜: Check if next data exists
        lgStrPrevKey = ""                                                  '☜: 다음 데이타 없다.
    End If
  	
	rs0.Close
    Set rs0 = Nothing 
    Set lgADF = Nothing                                                    '☜: ActiveX Data Factory Object Nothing
End Sub
'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()

    Dim strVal
    Dim arrVal(4)
    Redim UNISqlId(5)                                                     '☜: SQL ID 저장을 위한 영역확보 
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------

    Redim UNIValue(5,2)

    UNISqlId(0) = "S1111QA101"
    UNISqlId(1) = "s0000qa016"   
    UNISqlId(2) = "s0000qa000"
    UNISqlId(3) = "s0000qa000"
    UNISqlId(4) = "s0000qa003"
    UNISqlId(5) = "S0000QA014"
    

    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    UNIValue(0,0) = lgSelectList                                          '☜: Select list
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------

	strVal = ""


	If Len(Request("txtconValid_from_dt")) Then
		strVal = strVal & "A.VALID_FROM_DT <= " & FilterVar(UNIConvDate(Request("txtconValid_from_dt")), "''", "S") & ""		
	End If		

	
	If Len(Request("txtconItem_cd")) Then
		strVal = strVal & " AND A.ITEM_CD = " & FilterVar(UCase(Request("txtconItem_cd")), "''", "S") & " "
		arrVal(0) = Trim(Request("txtconItem_cd"))
	End If

 	If Len(Request("txtconPay_terms")) Then
		strVal = strVal & " AND A.PAY_METH = " & FilterVar(UCase(Request("txtconPay_terms")), "''", "S") & " "		
		arrVal(1) = Trim(Request("txtconPay_terms")) 
	End If		
    
    If Len(Request("txtconDeal_type")) Then
		strVal = strVal & " AND A.DEAL_TYPE = " & FilterVar(UCase(Request("txtconDeal_type")), "''", "S") & " "		
		arrVal(2) = Trim(Request("txtconDeal_type")) 
	End If	

	If Len(Request("txtconSales_unit")) Then
		strVal = strVal & " AND A.SALES_UNIT = " & FilterVar(UCase(Request("txtconSales_unit")), "''", "S") & " "		
		arrVal(3) = Trim(UCase(Request("txtconSales_unit")))
	End If

	If Len(Request("txtconCurrency")) Then
		strVal = strVal & " AND A.CURRENCY = " & FilterVar(UCase(Request("txtconCurrency")), "''", "S") & " "		
		arrVal(4) = Trim(UCase(Request("txtconCurrency")))
	End If
	

    UNIValue(0,1) = strVal   
    UNIValue(1,0) = FilterVar(arrVal(0), " " , "S")        
    UNIValue(2,0) = FilterVar("B9004", "", "S")
    UNIValue(2,1) = FilterVar(arrVal(1), " " , "S") 
    UNIValue(3,0) = FilterVar("S0001", "", "S")
    UNIValue(3,1) = FilterVar(arrVal(2), " " , "S")
    UNIValue(4,0) = FilterVar(arrVal(3), " " , "S")
    UNIValue(5,0) = FilterVar(arrVal(4), " " , "S")
    
    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    UNIValue(0,UBound(UNIValue,2)) = UCase(Trim(lgTailList))
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '☜: set ADO read mode
 
End Sub
'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub QueryData()
    Dim lgstrRetMsg                                             '☜ : Record Set Return Message 변수선언 
    Dim lgADF                                                   '☜ : ActiveX Data Factory 지정 변수선언 
    Dim iStr
    BlankchkFlg = False
    
        
    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3, rs4, rs5)

    Set lgADF   = Nothing

    iStr = Split(lgstrRetMsg,gColSep)
    
    If  rs1.EOF And rs1.BOF Then
        rs1.Close
        Set rs1 = Nothing

		If Len(Request("txtconItem_cd")) And BlankchkFlg = False Then
		   Call DisplayMsgBox("970000", vbInformation, "품목", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code
	       BlankchkFlg = True
            ' Modify Focus Events    
            %>
                <Script language=vbs>
                Parent.frm1.txtconItem_cd.focus    
                </Script>
            <%	       	
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

		If Len(Request("txtconPay_terms")) And BlankchkFlg = False Then
		   Call DisplayMsgBox("970000", vbInformation, "결제방법", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code
	       BlankchkFlg = True	
            ' Modify Focus Events    
            %>
                <Script language=vbs>
                Parent.frm1.txtconPay_terms.focus    
                </Script>
            <%	       
		End If	
    Else    
		arrRsVal(2) = rs2(0)
		arrRsVal(3) = rs2(1)
        rs2.Close
        Set rs2 = Nothing
    End If

    If  rs3.EOF And rs3.BOF Then
        rs3.Close
        Set rs3 = Nothing

		If Len(Request("txtconDeal_type")) And BlankchkFlg = False Then
		   Call DisplayMsgBox("970000", vbInformation, "판매유형", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code
	       BlankchkFlg = True	
            ' Modify Focus Events    
            %>
                <Script language=vbs>
                Parent.frm1.txtconDeal_type.focus    
                </Script>
            <%	       
		End If	
    Else    
		arrRsVal(4) = rs3(0)
		arrRsVal(5) = rs3(1)
        rs3.Close
        Set rs3 = Nothing
    End If
     
    If  rs4.EOF And rs4.BOF Then
        rs4.Close
        Set rs4 = Nothing

		If Len(Request("txtconSales_unit")) And BlankchkFlg = False Then
		   Call DisplayMsgBox("970000", vbInformation, "단위", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code
	       BlankchkFlg = True	
            ' Modify Focus Events    
            %>
                <Script language=vbs>
                Parent.frm1.txtconSales_unit.focus    
                </Script>
            <%	       
		End If	
    Else    
		arrRsVal(6) = rs4(0)
		arrRsVal(7) = rs4(1)
        rs4.Close
        Set rs4 = Nothing
    End If
   
   If  rs5.EOF And rs5.BOF Then
        rs5.Close
        Set rs5 = Nothing

		If Len(Request("txtconCurrency")) And BlankchkFlg = False Then
		   Call DisplayMsgBox("970000", vbInformation, "화폐", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code
	       BlankchkFlg = True	
            ' Modify Focus Events    
            %>
                <Script language=vbs>
                Parent.frm1.txtconCurrency.focus    
                </Script>
            <%	       
		End If	
    Else    
		arrRsVal(8) = rs5(0)
		arrRsVal(9) = rs5(1)
        rs5.Close
        Set rs5 = Nothing
    End If

    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If    
        
    If BlankchkFlg = False Then	
		If  rs0.EOF And rs0.BOF And BlankchkFlg =  False Then
		    Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
		    rs0.Close
		    Set rs0 = Nothing
            ' Modify Focus Events    
            %>
                <Script language=vbs>
                Parent.frm1.txtconItem_cd.focus    
                </Script>
            <%		    
    
			' 이 위치에 있는 Response.End 를 삭제하여야 함. Client 단에서 Name을 모두 뿌려준 후에 Response.End 를 기술함.
		Else    
		    Call  MakeSpreadSheetData()
		End If
	End If	

   
End Sub

'----------------------------------------------------------------------------------------------------------
' Set default value or preset value
'----------------------------------------------------------------------------------------------------------
Sub TrimData()
End Sub

%>
<Script Language=vbscript>
    With parent
		.ggoSpread.Source    = .frm1.vspdData 
        .frm1.vspdData.Redraw = False
		.ggoSpread.SSShowDataByClip "<%=lgstrData%>","F"						'☜ : Display data
		Call .ReFormatSpreadCellByCellByCurrency(.frm1.vspdData,"<%=iFrPoint+1%>",.frm1.vspdData.MaxRows,.GetKeyPos("A",10),.GetKeyPos("A",11),"C","Q","X","X")
        .lgStrPrevKey        =  "<%=lgStrPrevKey%>"                     '☜ : set next data tag
        
		.frm1.txtHconItem_cd.value		= "<%=ConvSPChars(Request("txtconItem_cd"))%>"
		.frm1.txtHconDeal_type.value	= "<%=ConvSPChars(Request("txtconDeal_type"))%>"
		.frm1.txtHconPay_terms.value	= "<%=ConvSPChars(Request("txtconPay_terms"))%>"
		.frm1.txtHconValid_from_dt.value = "<%=Request("txtconValid_from_dt")%>"
		.frm1.txtHconSales_unit.value	= "<%=ConvSPChars(Request("txtconSales_unit"))%>"
		.frm1.txtHconCurrency.value			= "<%=ConvSPChars(Request("txtconCurrency"))%>"
    

        
        .frm1.txtconItem_nm.value = "<%=ConvSPChars(arrRsVal(1))%>" 
        .frm1.txtconPay_terms_nm.value = "<%=ConvSPChars(arrRsVal(3))%>" 
        .frm1.txtconDeal_type_nm.value = "<%=ConvSPChars(arrRsVal(5))%>" 
		
        .DbQueryOk
        .frm1.vspdData.Redraw = True
	End with
</Script>	
<%
	Response.End													'☜: 비지니스 로직 처리를 종료함 
%>
