<%
'**********************************************************************************************
'*  1. Module Name          : Sales
'*  2. Function Name        : 
'*  3. Program ID           : S2111QB1
'*  4. Program Name         : 조직별 품목판매계획현황조회 
'*  5. Program Desc         : 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 
'*  8. Modified date(Last)  : 2001/12/18
'*  9. Modifier (First)     : Byun Jee Hyun
'* 10. Modifier (Last)      : sonbumyeol
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
<%                                                                         '☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 

On Error Resume Next

Call LoadBasisGlobalInf()
	
Dim lgADF                                                                  '☜ : ActiveX Data Factory 지정 변수선언 
Dim lgstrRetMsg                                                            '☜ : Record Set Return Message 변수선언 
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3, rs4, rs5     '☜ : DBAgent Parameter 선언 
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
Dim FalsechkFlg
Dim arrRsVal(7)

'--------------- 개발자 coding part(변수선언,End)----------------------------------------------------------
  
    Call HideStatusWnd 


    lgStrPrevKey   = Request("lgStrPrevKey")                               '☜ : Next key flag
    lgMaxCount     = CInt(100)											   '☜ : 한번에 가져올수 있는 데이타 건수 
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

    iCnt = 0
    lgstrData = ""

    If Len(Trim(lgStrPrevKey)) Then                                        '☜ : Chnage Nextkey str into int value
       If Isnumeric(lgStrPrevKey) Then
          iCnt = CInt(lgStrPrevKey)
       End If   
    End If   

    For iRCnt = 1 to iCnt  *  lgMaxCount                                   '☜ : Discard previous data
        rs0.MoveNext
    Next

    iRCnt = -1
    
    Do while Not (rs0.EOF Or rs0.BOF)
        iRCnt =  iRCnt + 1
        iStr = ""
		For ColCnt = 0 To UBound(lgSelectListDT) - 1 
            Select Case  lgSelectListDT(ColCnt)
               Case "DD"   '날짜 
                           iStr = iStr & Chr(11) & UNIDateClientFormat(rs0(ColCnt))
               Case "F2"  ' 금액 
                           iStr = iStr & Chr(11) & UNINumClientFormat(rs0(ColCnt), ggAmtOfMoney.DecPoint, 0)
               Case "F3"  '수량 
                           iStr = iStr & Chr(11) & UNINumClientFormat(rs0(ColCnt), ggQty.DecPoint       , 0)
               Case "F4"  '단가 
                           iStr = iStr & Chr(11) & UNINumClientFormat(rs0(ColCnt), ggUnitCost.DecPoint  , 0)
               Case "F5"   '환율 
                           iStr = iStr & Chr(11) & UNINumClientFormat(rs0(ColCnt), ggExchRate.DecPoint  , 0)
               Case Else
                    iStr = iStr & Chr(11) & ConvSPChars(rs0(ColCnt))
            End Select
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
    Dim arrVal(3)
    Dim Majorcd(1)
    Dim strConPlanNum

	strConPlanNum = Trim(Request("txtConPlanNum"))
    
    Redim UNISqlId(5)                                                     '☜: SQL ID 저장을 위한 영역확보 
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
    Redim UNIValue(5,4)

    UNISqlId(0) = "S2111QA101"	
    UNISqlId(1) = "B1254MA803"	' 기존 S0000QA006    
    UNISqlId(2) = "S0000QA000"
    UNISqlId(3) = "S0000QA000"
    UNISqlId(4) = "s0000qa016"
    UNISqlId(5) = "ConPlanNum"
    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    UNIValue(0,0) = lgSelectList                                          '☜: Select list
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
	strVal = ""

	If Len(Request("txtConSalesOrg")) Then
		strVal = "WHERE SP.SALES_ORG = " & FilterVar(UCase(Request("txtConSalesOrg")), "''", "S") & " "
		arrVal(0) = Trim(Request("txtConSalesOrg"))
	End If
	
	If Len(Request("txtConSpYear")) Then
		strVal = strVal + " AND SP.SP_YEAR = " & FilterVar(Request("txtConSpYear"), "''", "S") & ""
	Else
		strVal = strVal + ""
	End If

	If Len(Request("cboSpMonth")) Then
		strVal = strVal + " AND SP.SP_Month = " & FilterVar(Request("cboSpMonth"), "''", "S") & ""
	Else
		strVal = strVal + ""
	End If
	
	If Len(Request("txtConPlanTypeCd")) Then
		strVal = strVal + " AND SP.PLAN_FLAG = " & FilterVar(UCase(Request("txtConPlanTypeCd")), "''", "S") & " "
		arrVal(1) = Trim(Request("txtConPlanTypeCd"))
		Majorcd(0) = FilterVar("S4089", " " , "S")
	Else
		strVal = strVal + ""
		arrVal(1) = ""
		Majorcd(0) = FilterVar("", " " , "S")
	End If
	
	If Len(Request("txtConDealTypeCd")) Then
		strVal = strVal + " AND SP.EXPORT_FLAG = " & FilterVar(UCase(Request("txtConDealTypeCd")), "''", "S") & " "
		arrVal(2) = Trim(Request("txtConDealTypeCd"))
		Majorcd(1) = FilterVar("S4225", " " , "S")
	Else
		strVal = strVal + ""
		arrVal(2) = ""
		Majorcd(1) = FilterVar("", " " , "S")
	End If
    
    If Len(Request("txtConSalesItem")) Then
		strVal = strVal + " AND ITEM.ITEM_CD = " & FilterVar(UCase(Request("txtConSalesItem")), "''", "S") & " "
		arrVal(3) = Trim(Request("txtConSalesItem"))
	End If
	
    
	If Len(Request("txtConPlanNum")) Then
		strVal = strVal + " AND SP.PLAN_SEQ = " & FilterVar(UCase(Request("txtConPlanNum")), "''", "S") & " "
	Else
		strVal = strVal + ""
	End If

    strVal = strVal + " AND SP.ITEM_CD = ITEM.ITEM_CD AND SP.PLAN_FLAG = PLANFLAG.MINOR_CD AND SP.EXPORT_FLAG = EXPORTFLAG.MINOR_CD "
    UNIValue(0,1) = strVal + " AND SP.PLAN_SEQ = PLANSEQ.MINOR_CD "
        
    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
'==    UNIValue(0,UBound(UNIValue,2)) = UCase(Trim(lgTailList))
    UNIValue(0,2) = "GROUP BY SP.SALES_ORG, SP.SP_YEAR, PLANFLAG.MINOR_NM, EXPORTFLAG.MINOR_NM, " _
		& " SP.CUR, PLANSEQ.MINOR_NM, SP.SP_MONTH, SP.ITEM_CD, ITEM.ITEM_NM, ITEM.SPEC " 

    UNIValue(1,0) = " " & FilterVar(arrVal(0), "''", "S") & "  AND END_ORG_FLAG=" & FilterVar("Y", "''", "S") & "  AND USAGE_FLAG=" & FilterVar("Y", "''", "S") & " "
    UNIValue(2,0) = Majorcd(0)
    UNIValue(2,1) = FilterVar(arrVal(1), " " , "S")
    UNIValue(3,0) = Majorcd(1)
    UNIValue(3,1) = FilterVar(arrVal(2), " " , "S")
    UNIValue(4,0) = FilterVar(arrVal(3), " " , "S")
	
	IF strConPlanNum = "" Then
	   Call DisplayMsgBox("202299", vbOKOnly, "", "", I_MKSCRIPT)
	   Response.End
	END IF
    
    UNIValue(5,0) = FilterVar(strConPlanNum, " ", "S")
    
    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    UNIValue(0,UBound(UNIValue,2)) = UCase(Trim(lgTailList))
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '☜: set ADO read mode
End Sub
'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub QueryData()
    Dim iStr
	FalsechkFlg = False
    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")

    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3,rs4,rs5)
		
    If  rs1.EOF And rs1.BOF Then
        rs1.Close
        Set rs1 = Nothing

		If Len(Request("txtConSalesOrg")) And FalsechkFlg = False Then
		   Call DisplayMsgBox("970000", vbInformation, "영업조직", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code
				%>
				<Script Language=vbscript>
				    parent.frm1.txtConSalesOrg.focus
				</Script>	
				<%
			   FalsechkFlg = True	
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
		
		If Len(Request("txtConSalesItem")) And FalsechkFlg = False Then
			Call DisplayMsgBox("970000", vbInformation, "품목", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code
			%>
			<Script Language=vbscript>
			    parent.frm1.txtConSalesItem.focus
			</Script>	
			<%
		   FalsechkFlg = True		
		End If
    Else    
		arrRsVal(6) = rs4(0)
		arrRsVal(7) = rs4(1)
        rs4.Close
        Set rs4 = Nothing
    End If   
   
   
    If  rs2.EOF And rs2.BOF Then
        rs2.Close
        Set rs2 = Nothing
		If Len(Request("txtConPlanTypeCd")) And FalsechkFlg = False Then
			Call DisplayMsgBox("970000", vbInformation, "계획구분", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code
			%>
			<Script Language=vbscript>
			    parent.frm1.txtConPlanTypeCd.focus
			</Script>	
			<%
		   FalsechkFlg = True	
		End if
    Else    
		arrRsVal(2) = rs2(0)
		arrRsVal(3) = rs2(1)
        rs2.Close
        Set rs2 = Nothing
    End If

    If  rs3.EOF And rs3.BOF Then
        rs3.Close
        Set rs3 = Nothing
		If Len(Request("txtConDealTypeCd")) And FalsechkFlg = False Then
			Call DisplayMsgBox("970000", vbInformation, "거래구분", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code
			%>
			<Script Language=vbscript>
			    parent.frm1.txtConDealTypeCd.focus
			</Script>	
			<%
		   FalsechkFlg = True	
		End if
    Else    
		arrRsVal(4) = rs3(0)
		arrRsVal(5) = rs3(1)
        rs3.Close
        Set rs3 = Nothing
    End If
    
	If (rs5.EOF And rs5.BOF) Then
		Call DisplayMsgBox("202299", vbOKOnly, "", "", I_MKSCRIPT)
		rs5.Close
		Set rs5 = Nothing
		%>
		<Script Language=vbscript>
			parent.frm1.txtConPlanNum.value = ""
			parent.frm1.txtConPlanNumNm.value = ""
			parent.frm1.txtConPlanNum.Focus()
		</Script>	
		<%
		Response.End
	Else
		%>
		<Script Language=vbscript>
			parent.frm1.txtConPlanNumNm.value = "<%=ConvSPChars(rs5("MINOR_NM"))%>"
		</Script>	
		<%    	
		rs5.Close
		Set rs5 = Nothing
	End If
    
    iStr = Split(lgstrRetMsg,gColSep)
    
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If    
    
    If  rs0.EOF And rs0.BOF And FalsechkFlg = False Then
        Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
        rs0.Close
        Set rs0 = Nothing
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
         .ggoSpread.SSShowDataByClip "<%=lgstrData%>"                            '☜: Display data 
         .lgStrPrevKey        =  "<%=lgStrPrevKey%>"                       '☜: set next data tag

  		.frm1.txtConSalesOrgNm.value		=  "<%=ConvSPChars(arrRsVal(1))%>" 
  		.frm1.txtConPlanTypeNm.value		=  "<%=ConvSPChars(arrRsVal(3))%>" 	
  		.frm1.txtConDealTypeNm.value		=  "<%=ConvSPChars(arrRsVal(5))%>"
        .frm1.txtConSalesItemNm.value		=  "<%=ConvSPChars(arrRsVal(7))%>"
        .frm1.txtHConSalesOrg.value	        =  "<%=ConvSPChars(Request("txtConSalesOrg"))%>"
		.frm1.txtHConSalesItem.value	    =  "<%=ConvSPChars(Request("txtConSalesItem"))%>"
		.frm1.txtHConSpYear.value	        =  "<%=ConvSPChars(Request("txtConSpYear"))%>"
		.frm1.cboHSpMonth.value		        =  "<%=ConvSPChars(Request("cboSpMonth"))%>"
		.frm1.txtHConPlanTypeCd.value	    =  "<%=ConvSPChars(Request("txtConPlanTypeCd"))%>"
		.frm1.txtHConDealTypeCd.value	    =  "<%=ConvSPChars(Request("txtConDealTypeCd"))%>"
        .frm1.txtHConCurr.value	            =  "<%=ConvSPChars(Request("txtConCurr"))%>"       
        .frm1.txtHConPlanNum.value	        =  "<%=ConvSPChars(Request("txtConPlanNum"))%>" 
        If "<%=FalsechkFlg%>" = False Then
         .DbQueryOk
        End If
	End with
</Script>	

