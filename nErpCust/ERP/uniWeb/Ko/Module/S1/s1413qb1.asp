<%
'**********************************************************************************************
'*  1. Module Name          : Sales
'*  2. Function Name        : 
'*  3. Program ID           : S1413QB1
'*  4. Program Name         : 담보현황조회 
'*  5. Program Desc         : 담보현황조회 
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
<!-- #Include file="../../inc/incSvrDBAgent.inc" -->
<%                                                                         '☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 

On Error Resume Next

Dim lgADF                                                                  '☜ : ActiveX Data Factory 지정 변수선언 
Dim lgstrRetMsg                                                            '☜ : Record Set Return Message 변수선언 
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0                              '☜ : DBAgent Parameter 선언 
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
Dim arrRsVal(7)
Dim BlankchkFlg

Dim iFrPoint
iFrPoint=0
'--------------- 개발자 coding part(변수선언,End)----------------------------------------------------------
  
Call LoadBasisGlobalInf()
Call LoadInfTB19029B("Q", "S", "NOCOOKIE", "QB")
Call LoadBNumericFormatB("Q", "S", "NOCOOKIE", "QB")

Call HideStatusWnd 


lgStrPrevKey   = Request("lgStrPrevKey")                               '☜ : Next key flag
lgMaxCount     = CInt(100)                           '☜ : 한번에 가져올수 있는 데이타 건수 
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
		iFrPoint	= iCnt  *  lgMaxCount
    Next

    iRCnt = -1
    
    Do while Not (rs0.EOF Or rs0.BOF)
        iRCnt =  iRCnt + 1
        iStr = ""

		For ColCnt = 0 To UBound(lgSelectListDT) - 1
            iStr = iStr & Chr(11) & FormatRsString(lgSelectListDT(ColCnt),rs0(ColCnt))
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
    Dim majorcd(1)
    Redim UNISqlId(4)                                                     '☜: SQL ID 저장을 위한 영역확보 
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------

    Redim UNIValue(4,2)

    UNISqlId(0) = "S1413QA101"
    UNISqlId(1) = "S0000QA002"
    UNISqlId(2) = "S0000QA005"
    UNISqlId(3) = "S0000QA000"
    UNISqlId(4) = "S0000QA000"


    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    UNIValue(0,0) = lgSelectList                                          '☜: Select list
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------

	strVal = " "
		
	If Len(Request("txtBpCd")) Then
		strVal = " AND A.BP_CD =  " & FilterVar(Trim(UCase(Request("txtBpcd"))), "" , "S") & " "
		arrVal(0) = FilterVar(Trim(Request("txtBpCd")), "", "S")
	Else
		strVal = ""
		arrVal(0) = FilterVar("", "", "S")
	End If
	
	If Len(Request("txtSalesGroup")) Then
		strVal = strVal & " AND A.SALES_GRP = " & FilterVar(Trim(UCase(Request("txtSalesGroup"))), "" , "S") & " "		
		arrVal(1) = FilterVar(Trim(Request("txtSalesGroup")), "", "S")
	Else
		arrVal(1) = FilterVar("", "", "S")
	End If		

    If Len (Request("txtAsignFrDt")) Then
		strVal = strVal & " AND A.ASGN_DT >= " & FilterVar(UNIConvDate(Request("txtAsignFrDt")), "''", "S") & ""		
	End If		

    If Len (Request("txtAsignToDt")) Then
		strVal = strVal & " AND A.ASGN_DT <= " & FilterVar(UNIConvDate(Request("txtAsignToDt")), "''", "S") & ""		
	End If
    
    If Request("txtRadio") = "N" Then
		strVal = strVal & " AND Len(A.DEL_TYPE) = 0"
	ElseIf Request("txtRadio") = "Y" Then
		strVal = strVal & " AND Len(A.DEL_TYPE) > 0"			
	End If
    
    If Len(Request("txtDelType") ) Then
		strVal = strVal & " AND A.DEL_TYPE = " & FilterVar(Trim(UCase(Request("txtDelType"))), "" , "S") & " "	
		arrVal(2) = FilterVar(Trim(Request("txtDelType")), "", "S")
		majorcd(0) =  FilterVar("S0003", "", "S")
    Else
		arrVal(2) = FilterVar("", "", "S")
		majorcd(0) = FilterVar("", "", "S")
    End If
    
    If Len(Request("txtColletralType") ) Then
		strVal = strVal & " AND B.MINOR_CD = " & FilterVar(Trim(UCase(Request("txtColletralType"))), "" , "S") & " "	
		arrVal(3) = FilterVar(Trim(Request("txtColletralType")), "", "S")
		majorcd(1) =  FilterVar("S0002", "", "S")
    Else
		arrVal(3) = FilterVar("", "", "S")
		majorcd(1) = FilterVar("", "", "S")
    End If
    
    UNIValue(0,1) = strVal   '---발주일 
    UNIValue(1,0) = arrVal(0)
    UNIValue(2,0) = arrVal(1)
    UNIValue(3,0) = majorcd(0)
    UNIValue(3,1) = arrVal(2)
    UNIValue(4,0) = majorcd(1)
    UNIValue(4,1) = arrVal(3)
   
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
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3, rs4)
    
    Set lgADF   = Nothing

    iStr = Split(lgstrRetMsg,gColSep)
    
    If  rs1.EOF And rs1.BOF Then
        rs1.Close
        Set rs1 = Nothing

		If Len(Request("txtBpCd")) And BlankchkFlg = False Then
		   Call DisplayMsgBox("970000", vbInformation, "고객", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code
	       BlankchkFlg = True
            ' Modify Focus Events    
            %>
                <Script language=vbs>
                Parent.frm1.txtBpCd.focus    
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

		If Len(Request("txtSalesGroup")) And BlankchkFlg = False Then
		   Call DisplayMsgBox("970000", vbInformation, "영업그룹", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code
	       BlankchkFlg = True
            ' Modify Focus Events    
            %>
                <Script language=vbs>
                Parent.frm1.txtSalesGroup.focus    
                </Script>
            <%	       	
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

		If Len(Request("txtColletralType")) And BlankchkFlg = False Then
		   Call DisplayMsgBox("970000", vbInformation, "담보유형", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code
	       BlankchkFlg = True
            ' Modify Focus Events    
            %>
                <Script language=vbs>
                Parent.frm1.txtColletralType.focus    
                </Script>
            <%	       	
		End If	
    Else    
		arrRsVal(6) = rs4(0)
		arrRsVal(7) = rs4(1)
        rs4.Close
        Set rs4 = Nothing
    End If

	If  rs3.EOF And rs3.BOF Then
        rs3.Close
        Set rs3 = Nothing

		If Len(Request("txtDelType")) And BlankchkFlg = False Then
		   Call DisplayMsgBox("970000", vbInformation, "해지구분", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code
	       BlankchkFlg = True
            ' Modify Focus Events    
            %>
                <Script language=vbs>
                Parent.frm1.txtDelType.focus    
                </Script>
            <%	       	
		End If	
    Else    
		arrRsVal(4) = rs3(0)
		arrRsVal(5) = rs3(1)
        rs3.Close
        Set rs3 = Nothing
    End If
   
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If    

    If BlankchkFlg = False Then	
		If  rs0.EOF And rs0.BOF And BlankchkFlg =False Then
		    Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
		    rs0.Close
		    Set rs0 = Nothing
            ' Modify Focus Events    
            %>
                <Script language=vbs>
                Parent.frm1.txtBpCd.focus    
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
		Call .ReFormatSpreadCellByCellByCurrency(.frm1.vspdData,"<%=iFrPoint+1%>",.frm1.vspdData.MaxRows,.GetKeyPos("A",6),.GetKeyPos("A",7),"A","Q","X","X")
		Call .ReFormatSpreadCellByCellByCurrency(.frm1.vspdData,"<%=iFrPoint+1%>",.frm1.vspdData.MaxRows,.GetKeyPos("A",6),.GetKeyPos("A",10),"A","Q","X","X")
		
        .lgStrPrevKey        =  "<%=lgStrPrevKey%>"                       '☜: set next data tag
  		
  		.frm1.txtBpNm.value				=  "<%=ConvSPChars(arrRsVal(1))%>" 
  		.frm1.txtSalesGroupNm.value		=  "<%=ConvSPChars(arrRsVal(3))%>" 	
  		.frm1.txtDelTypeNm.value		=  "<%=ConvSPChars(arrRsVal(5))%>"
  		.frm1.txtColletralTypeNm.value	=  "<%=ConvSPChars(arrRsVal(7))%>" 	
  		
  		.frm1.HBpCd.value				= "<%=ConvSPChars(Request("txtBpCd"))%>"
		.frm1.HSalesGroup.value			= "<%=ConvSPChars(Request("txtSalesGroup"))%>"
		.frm1.HColletralType.value		= "<%=ConvSPChars(Request("txtColletralType"))%>"
		.frm1.HAsignFrDt.value			= "<%=Request("txtAsignFrDt")%>"
		.frm1.HAsignToDt.value			= "<%=Request("txtAsignToDt")%>"
		.frm1.HRadio.value				= "<%=Request("txtRadio")%>"
		.frm1.HDelType.value			= "<%=ConvSPChars(Request("txtDelType"))%>"
    
        .DbQueryOk
        .frm1.vspdData.Redraw = True                
	End with
</Script>	
<%
    Response.End													'☜: 비지니스 로직 처리를 종료함 
%>

