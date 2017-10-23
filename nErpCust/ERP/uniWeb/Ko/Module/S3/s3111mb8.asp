<%
'**********************************************************************************************
'*  1. Module Name          : 영업 
'*  2. Function Name        : 
'*  3. Program ID           : s3111mb8.asp	
'*  4. Program Name         : 수주현황조회 
'*  5. Program Desc         : 수주현황조회 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2003/05/28
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Cho in kuk
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************
%>
<!-- #Include file="../../inc/incSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDBAgent.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%                                                                         '☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 
Call LoadBasisGlobalInf()
Call LoadInfTB19029B("Q", "S", "NOCOOKIE", "MB")
Call LoadBNumericFormatB("Q", "S", "NOCOOKIE", "MB")

On Error Resume Next

Dim lgADF                                                                  '☜ : ActiveX Data Factory 지정 변수선언 
Dim lgstrRetMsg                                                            '☜ : Record Set Return Message 변수선언 
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0                              '☜ : DBAgent Parameter 선언 
Dim lgstrData                                                              '☜ : data for spreadsheet data
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
Const C_SHEETMAXROWS_D  = 100                                   '☆: Server에서 한번에 fetch할 최대 데이타 건수 
Dim lgPageNo
'--------------- 개발자 coding part(변수선언,End)----------------------------------------------------------
  
    Call HideStatusWnd 

	lgPageNo       = UNICInt(Trim(Request("lgPageNo")),0)                  '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
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
    Dim  ColCnt
    Dim  iLoopCount
    Dim  iRowStr
    
    lgstrData      = ""
    
    If CInt(lgPageNo) > 0 Then
       rs0.Move  =  C_SHEETMAXROWS_D * CInt(lgPageNo)
    End If

    iLoopCount = -1
    
    Do while Not (rs0.EOF Or rs0.BOF)
        iLoopCount =  iLoopCount + 1
        iRowStr = ""
		For ColCnt = 0 To UBound(lgSelectListDT) - 1
            iRowStr = iRowStr & Chr(11) & FormatRsString(lgSelectListDT(ColCnt),rs0(ColCnt))
		Next
 
        If  iLoopCount < C_SHEETMAXROWS_D Then
            lgstrData      = lgstrData      & iRowStr & Chr(11) & Chr(12)
        Else
            lgPageNo = lgPageNo + 1
            Exit Do
        End If
        rs0.MoveNext
	Loop

    If  iLoopCount < C_SHEETMAXROWS_D Then                             '☜: Check if next data exists
        lgPageNo = ""                                                  '☜: 다음 데이타 없다.
    End If
  	
	rs0.Close
    Set rs0 = Nothing                          '☜: ActiveX Data Factory Object Nothing
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

    UNISqlId(0) = "S3111QA101"
    UNISqlId(1) = "S0000QA002"
    UNISqlId(2) = "S0000QA007"
    UNISqlId(3) = "S0000QA005"
    UNISqlId(4) = "s0000qa000"
    UNISqlId(5) = "s0000qa000"

    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    UNIValue(0,0) = lgSelectList                                          '☜: Select list
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------

	strVal = " "
	
	If Len(Request("txtconBp_cd")) Then
		strVal = "AND A.SOLD_TO_PARTY = " & FilterVar(Trim(Request("txtconBp_cd")), "" , "S") & " "
		arrVal(0) = Trim(Request("txtconBp_cd"))
	Else
		strVal = ""
	End If

	If Len(Request("txtSOType_cd")) Then
		strVal = strVal & " AND A.SO_TYPE = " & FilterVar(Trim(Request("txtSOType_cd")), "" , "S") & " "		
		arrVal(1) = Trim(Request("txtSOType_cd"))
	End If		
		   
	If Len(Request("txtSalesGroup_cd")) Then
		strVal = strVal & " AND A.SALES_GRP = " & FilterVar(Trim(Request("txtSalesGroup_cd")), "" , "S") & " "		
		arrVal(2) = Trim(Request("txtSalesGroup_cd"))
	End If		
    
 	If Len(Request("txtPayTerms_cd")) Then
		strVal = strVal & " AND A.PAY_METH = " & FilterVar(Trim(Request("txtPayTerms_cd")), "" , "S") & " "		
		arrVal(3) = Trim(Request("txtPayTerms_cd")) 
	End If		
    
    If Len(Request("txtdeal_type_cd")) Then
		strVal = strVal & " AND A.DEAL_TYPE = " & FilterVar(Trim(Request("txtdeal_type_cd")), "" , "S") & " "		
		arrVal(4) = Trim(Request("txtdeal_type_cd")) 
	End If	
	
    If Len(Request("txtSOFrDt")) Then
		strVal = strVal & " AND A.SO_DT >= " & FilterVar(UNIConvDate(Request("txtSOFrDt")), "''", "S") & ""		
	End If		
	
	If Len(Request("txtSOToDt")) Then
		strVal = strVal & " AND A.SO_DT <= " & FilterVar(UNIConvDate(Request("txtSOToDt")), "''", "S") & ""		
	End If

	If Trim(Request("txtRadio")) = "Y" Or Trim(Request("txtRadio")) = "N" Then
		strVal = strVal & " AND A.CFM_FLAG = " & FilterVar(Request("txtRadio"), "''", "S") & ""		
	End If

    UNIValue(0,1) = strVal   '---발주일 
    UNIValue(1,0) = FilterVar(arrVal(0), "" , "S")
    UNIValue(2,0) = FilterVar(arrVal(1), "" , "S")
    UNIValue(3,0) = FilterVar(arrVal(2), "" , "S")        
    UNIValue(4,0) = FilterVar("B9004", "" , "S")
    UNIValue(4,1) = FilterVar(arrVal(3), "" , "S")
    UNIValue(5,0) = FilterVar("S0001", "" , "S")
    UNIValue(5,1) = FilterVar(arrVal(4), "" , "S")
    
    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    UNIValue(0,UBound(UNIValue,2)) = UCase(Trim(lgTailList))
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '☜: set ADO read mode
 
End Sub
'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub QueryData()
    Dim iStr
    Dim FalsechkFlg
    
    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3, rs4, rs5)

    FalsechkFlg = False

    If  rs2.EOF And rs2.BOF Then
        rs2.Close
        Set rs2 = Nothing

		If Len(Request("txtSOType_cd")) And FalsechkFlg = False Then
		   Call DisplayMsgBox("970000", vbInformation, "수주형태", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code
	       FalsechkFlg = True	
            ' Modify Focus Events    
            %>
                <Script language=vbs>
                Parent.frm1.txtSOType_cd.focus    
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

		If Len(Request("txtSalesGroup_cd")) And FalsechkFlg = False Then
		   Call DisplayMsgBox("970000", vbInformation, "영업그룹", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code
	       FalsechkFlg = True	
            ' Modify Focus Events    
            %>
                <Script language=vbs>
                Parent.frm1.txtSalesGroup_cd.focus    
                </Script>
            <%		
		End If
    Else    
		arrRsVal(4) = rs3(0)
		arrRsVal(5) = rs3(1)
        rs3.Close
        Set rs3 = Nothing
    End If
	
	If  rs1.EOF And rs1.BOF Then
        rs1.Close
        Set rs1 = Nothing

		If Len(Request("txtconBp_cd")) And FalsechkFlg = False Then
		   Call DisplayMsgBox("970000", vbInformation, "주문처", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code
	       FalsechkFlg = True	
            ' Modify Focus Events    
            %>
                <Script language=vbs>
                Parent.frm1.txtconBp_cd.focus    
                </Script>
            <%	       
		End If	
    Else    
		arrRsVal(0) = rs1(0)
		arrRsVal(1) = rs1(1)
        rs1.Close
        Set rs1 = Nothing
    End If
    
    If  rs4.EOF And rs4.BOF Then
        rs4.Close
        Set rs4 = Nothing

		If Len(Request("txtPayTerms_cd")) And FalsechkFlg = False Then
		   Call DisplayMsgBox("970000", vbInformation, "결제방법", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code
	       FalsechkFlg = True	
            ' Modify Focus Events    
            %>
                <Script language=vbs>
                Parent.frm1.txtPayTerms_cd.focus    
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

		If Len(Request("txtdeal_type_cd")) And FalsechkFlg = False Then
		   Call DisplayMsgBox("970000", vbInformation, "판매유형", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code
	       FalsechkFlg = True	
            ' Modify Focus Events    
            %>
                <Script language=vbs>
                Parent.frm1.txtdeal_type_cd.focus    
                </Script>
            <%			       
		End If	
    Else    
		arrRsVal(8) = rs5(0)
		arrRsVal(9) = rs5(1)
        rs5.Close
        Set rs5 = Nothing
    End If
    
    iStr = Split(lgstrRetMsg,gColSep)
    
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If    
        
    If  rs0.EOF And rs0.BOF And FalsechkFlg =  False Then
        Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
        rs0.Close
        Set rs0 = Nothing
        ' Modify Focus Events    
        %>
            <Script language=vbs>
            Parent.frm1.txtSOType_cd.focus    
            </Script>
        <%        
    Else    
        Call  MakeSpreadSheetData()
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
        .ggoSpread.SSShowDataByClip "<%=lgstrData%>", "F"                            '☜: Display data 
        
        Call Parent.ReFormatSpreadCellByCellByCurrency(Parent.Frm1.vspdData,-1,-1,Parent.GetKeyPos("A",12),Parent.GetKeyPos("A",5),"A", "Q" ,"X","X")
        Call Parent.ReFormatSpreadCellByCellByCurrency(Parent.Frm1.vspdData,-1,-1,Parent.GetKeyPos("A",12),Parent.GetKeyPos("A",6),"A", "Q" ,"X","X")
                
        .lgPageNo = "<%=lgPageNo%>"
		.frm1.HconBp_cd.value		= "<%=ConvSPChars(Request("txtconBp_cd"))%>"
		.frm1.HSOType_cd.value		= "<%=ConvSPChars(Request("txtSOType_cd"))%>"
		.frm1.HSalesGroup_cd.value	= "<%=ConvSPChars(Request("txtSalesGroup_cd"))%>"
		.frm1.HPayTerms_cd.value	= "<%=ConvSPChars(Request("txtPayTerms_cd"))%>"
		.frm1.Hdeal_type_cd.value	= "<%=ConvSPChars(Request("txtdeal_type_cd"))%>"
		.frm1.HSOFrDt.value			= "<%=Request("txtSOFrDt")%>"
		.frm1.HSOToDt.value			= "<%=Request("txtSOToDt")%>"
		.frm1.HRadio.value			= "<%=Request("txtRadio")%>"            
        .frm1.txtconBp_nm.value = "<%=ConvSPChars(arrRsVal(1))%>" 
        .frm1.txtSOType_nm.value = "<%=ConvSPChars(arrRsVal(3))%>" 
        .frm1.txtSalesGroup_nm.value = "<%=ConvSPChars(arrRsVal(5))%>" 
		.frm1.txtPayTerms_nm.value = "<%=ConvSPChars(arrRsVal(7))%>"
		.frm1.txtdeal_type_nm.value = "<%=ConvSPChars(arrRsVal(9))%>"
		
        .DbQueryOk
        .frm1.vspdData.Redraw = True
	End with
</Script>	
<%
	Response.End													'☜: 비지니스 로직 처리를 종료함 
%>
