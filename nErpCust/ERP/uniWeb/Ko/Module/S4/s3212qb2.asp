<%
'**********************************************************************************************
'*  1. Module Name          : 영업관리 
'*  2. Function Name        : 
'*  3. Program ID           : s3212qb2
'*  4. Program Name         : Local L/C 상세조회 
'*  5. Program Desc         : Local L/C 상세조회 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 
'*  8. Modified date(Last)  : 2002/08/23
'*  9. Modifier (First)     : Byun Jee Hyun
'* 10. Modifier (Last)      : Park in sik
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*                            2000/12/09
'**********************************************************************************************
%>
<!-- #Include file="../../inc/incSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDBAgent.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%

Call LoadBasisGlobalInf()
Call LoadInfTB19029B("Q", "S", "NOCOOKIE", "QB")
Call LoadBNumericFormatB("Q","S","NOCOOKIE","QB")

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
Dim arrRsVal(5)
Dim BlankchkFlg 

Dim lgDataExist
Dim lgPageNo

Dim iFrPoint
iFrPoint=0 
Const C_SHEETMAXROWS_D  = 100                                   '☆: Server에서 한번에 fetch할 최대 데이타 건수 
'--------------- 개발자 coding part(변수선언,End)----------------------------------------------------------
  
    Call HideStatusWnd 

	lgPageNo       = UNICInt(Trim(Request("lgPageNo")),0)    	
	lgDataExist    = "No"    
    lgSelectList   = Request("lgSelectList")                               '☜ : select 대상목록 
    lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)             '☜ : 각 필드의 데이타 타입 
    lgTailList     = Request("lgTailList")                                 '☜ : Orderby value

    Call TrimData()
    Call FixUNISQLData()
    Call QueryData()
    
'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------

Private Sub MakeSpreadSheetData()
    Dim  RecordCnt
    Dim  ColCnt
    Dim  iRCnt
    Dim  iStr

    lgstrData = ""
    lgDataExist    = "Yes"
	
	If CLng(lgPageNo) > 0 Then
       rs0.Move = CLng(C_SHEETMAXROWS_D) * CLng(lgPageNo)                  'C_SHEETMAXROWS_D:Max Fetched Count at once , lgStrPrevKeyIndex : Previous PageNo
       iFrPoint	= CLng(C_SHEETMAXROWS_D) * CLng(lgPageNo)
    End If

    iRCnt = -1
    
    Do while Not (rs0.EOF Or rs0.BOF)
        iRCnt =  iRCnt + 1
        iStr = ""
		For ColCnt = 0 To UBound(lgSelectListDT) - 1 
            iStr = iStr & Chr(11) & FormatRsString(lgSelectListDT(ColCnt),rs0(ColCnt))
		Next
 
        If  iRCnt < C_SHEETMAXROWS_D Then
            lgstrData      = lgstrData      & iStr & Chr(11) & Chr(12)
        Else
            lgPageNo = lgPageNo + 1
            Exit Do
        End If
        rs0.MoveNext
	Loop

    If  iRCnt < C_SHEETMAXROWS_D Then                                            '☜: Check if next data exists        
		lgPageNo = ""
    End If
  	
	rs0.Close
    Set rs0 = Nothing 
    Set lgADF = Nothing                                                    '☜: ActiveX Data Factory Object Nothing
End Sub
'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Private Sub FixUNISQLData()

    Dim strVal
    Dim arrVal(2)
    Redim UNISqlId(3)                                                     '☜: SQL ID 저장을 위한 영역확보 
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------

    Redim UNIValue(3,2)

    UNISqlId(0) = "S3212QA201"
    UNISqlId(1) = "S0000QA002"
    UNISqlId(2) = "S0000QA005"
    UNISqlId(3) = "S0000QA001"

    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    UNIValue(0,0) = lgSelectList                                          '☜: Select list
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------

	strVal = " "
	
	If Len(Request("txtconBp_cd")) Then
		strVal = "AND A.APPLICANT = " & FilterVar(UCase(Request("txtconBp_cd")), "''", "S") & " "
		arrVal(0) = Trim(Request("txtconBp_cd"))
	Else
		strVal = ""
	End If

	If Len(Request("txtSalesGroup")) Then
		strVal = strVal & " AND A.SALES_GRP = " & FilterVar(UCase(Request("txtSalesGroup")), "''", "S") & " "		
		arrVal(1) = Trim(Request("txtSalesGroup"))
	End If		
    
 	If Len(Request("txtItem_cd")) Then
		strVal = strVal & " AND B.ITEM_CD = " & FilterVar(UCase(Request("txtItem_cd")), "''", "S") & " "		
		arrVal(2) = Trim(Request("txtItem_cd"))
	End If
	    
    If Len(Request("txtOpenFrDt")) Then
		strVal = strVal & " AND A.OPEN_DT >= " & FilterVar(UNIConvDate(Trim(Request("txtOpenFrDt"))), "''", "S") & ""		
	End If		
	
	If Len(Request("txtOpenToDt")) Then
		strVal = strVal & " AND A.OPEN_DT <= " & FilterVar(UNIConvDate(Trim(Request("txtOpenToDt"))), "''", "S") & ""		
	End If

    UNIValue(0,1) = strVal   '---발주일 
    UNIValue(1,0) = FilterVar(arrVal(0), " " , "S")
    UNIValue(2,0) = FilterVar(arrVal(1), " " , "S")
    UNIValue(3,0) = FilterVar(arrVal(2), " " , "S")


    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    UNIValue(0,UBound(UNIValue,2)) = UCase(Trim(lgTailList))
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '☜: set ADO read mode
 
End Sub
'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Private Sub QueryData()
    Dim lgstrRetMsg                                             '☜ : Record Set Return Message 변수선언 
    Dim lgADF                                                   '☜ : ActiveX Data Factory 지정 변수선언 
    Dim iStr
    BlankchkFlg = False
	
	
    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3)
    
    Set lgADF   = Nothing

    iStr = Split(lgstrRetMsg,gColSep)


    
    
    If  rs1.EOF And rs1.BOF Then
        rs1.Close
        Set rs1 = Nothing

		If Len(Request("txtconBp_cd")) And BlankchkFlg = False Then
		   Call DisplayMsgBox("970000", vbInformation, "개설신청인", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code
	       BlankchkFlg = True
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

    If  rs3.EOF And rs3.BOF Then
        rs3.Close
        Set rs3 = Nothing
		If Len(Request("txtItem_cd")) And BlankchkFlg = False Then
		   Call DisplayMsgBox("970000", vbInformation, "품목", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code
	       BlankchkFlg = True	
            ' Modify Focus Events    
            %>
                <Script language=vbs>
                Parent.frm1.txtItem_cd.focus    
                </Script>
            <%	       
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
        
    If BlankchkFlg = False Then	
		If  rs0.EOF And rs0.BOF And BlankchkFlg =  False Then
		    Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
		    rs0.Close
		    Set rs0 = Nothing
            ' Modify Focus Events    
            %>
                <Script language=vbs>
                Parent.frm1.txtconBp_cd.focus    
                </Script>
            <%		    
   
		Else    
		    Call  MakeSpreadSheetData()
		End If
	End If	

End Sub

'----------------------------------------------------------------------------------------------------------
' Set default value or preset value
'----------------------------------------------------------------------------------------------------------
Private Sub TrimData()
End Sub
%>

<Script Language=vbscript>
    With parent
        .frm1.txtconBp_Nm.value		= "<%=ConvSPChars(arrRsVal(1))%>" 
        .frm1.txtSalesGroupNm.value	= "<%=ConvSPChars(arrRsVal(3))%>" 
        .frm1.txtItem_Nm.value		= "<%=ConvSPChars(arrRsVal(5))%>" 
        
         If "<%=lgDataExist%>" = "Yes" Then
			'Set condition data to hidden area
			If "<%=lgPageNo%>" = "1" Then   ' "1" means that this query is first and next data exists
				
				.frm1.HconBp_cd.value		= "<%=ConvSPChars(Request("txtconBp_cd"))%>"
				.frm1.HSalesGroup.value	= "<%=ConvSPChars(Request("txtSalesGroup"))%>"
				.frm1.HItem_cd.value	= "<%=ConvSPChars(Request("txtItem_cd"))%>"
				.frm1.HOpenFrDt.value	= "<%=ConvSPChars(Request("txtOpenFrDt"))%>"
				.frm1.HOpenToDt.value		= "<%=ConvSPChars(Request("txtOpenToDt"))%>"
				
			End If  	
			
			.ggoSpread.Source    = .frm1.vspdData 
			.frm1.vspdData.Redraw = False
			.ggoSpread.SSShowDataByClip  "<%=lgstrData%>", "F"
			
			Call Parent.ReFormatSpreadCellByCellByCurrency(Parent.frm1.vspdData,"<%=iFrPoint+1%>",parent.frm1.vspddata.maxrows,Parent.GetKeyPos("A",8),Parent.GetKeyPos("A",7),"C", "Q" ,"X","X")		
			Call Parent.ReFormatSpreadCellByCellByCurrency(Parent.frm1.vspdData,"<%=iFrPoint+1%>",parent.frm1.vspddata.maxrows,Parent.GetKeyPos("A",8),Parent.GetKeyPos("A",9),"A", "Q" ,"X","X")		
						
			.lgPageNo			 = "<%=lgPageNo%>"							  '☜: Next next data tag
  			.DbQueryOk
  			.frm1.vspdData.Redraw = True
  			
  		End If
        
	End With
</Script>	
<%
    Response.End													'☜: 비지니스 로직 처리를 종료함 
%>
