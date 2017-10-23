<%
'**********************************************************************************************
'*  1. Module Name          : 영업 
'*  2. Function Name        : 
'*  3. Program ID           : s4212qb1
'*  4. Program Name         : 통관상세조회 
'*  5. Program Desc         : 통관상세조회 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 
'*  8. Modified date(Last)  : 2000/12/09
'*  9. Modifier (First)     : Byun Jee Hyun
'* 10. Modifier (Last)      : Kim Hyungsuk
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
Call LoadBNumericFormatB("Q", "S", "NOCOOKIE", "QB")

On Error Resume Next

Dim lgADF                                                                  
Dim lgstrRetMsg                                                            
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0                              
Dim lgstrData                                                              
Dim lgStrPrevKey                                                                                                                
Dim lgTailList                                                             
Dim lgSelectList
Dim lgSelectListDT
Dim strPoType	                                                           '⊙ : 발주형태 
Dim strPoFrDt	                                                           '⊙ : 발주일 
Dim strPoToDt	                                                           '⊙ :
Dim strSpplCd	                                                           '⊙ : 공급처 
Dim strPurGrpCd	                                                           '⊙ : 구매그룹 
Dim strItemCd	                                                           '⊙ : 품목 
Dim strTrackNo	                                                           '⊙ : Tracking No
Dim arrRsVal(5)
Const C_SHEETMAXROWS_D  = 100                                   '☆: Server에서 한번에 fetch할 최대 데이타 건수 
Dim FalsechkFlg     

    Call HideStatusWnd 
    lgStrPrevKey   = Request("lgStrPrevKey")                                   
    lgSelectList   = Request("lgSelectList")                               
    lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)             
    lgTailList     = Request("lgTailList")                                 

    Call TrimData()
    Call FixUNISQLData()
    Call QueryData()
'----------------------------------------------------------------------------------------------------------

Private Sub MakeSpreadSheetData()
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

    For iRCnt = 1 to iCnt  *  C_SHEETMAXROWS_D                                   '☜ : Discard previous data
        rs0.MoveNext
    Next

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
            iCnt = iCnt + 1
            lgStrPrevKey = CStr(iCnt)
            Exit Do
        End If
        rs0.MoveNext
	Loop

    If  iRCnt < C_SHEETMAXROWS_D Then                                            
        lgStrPrevKey = ""                                                  '☜: 다음 데이타 없다.
    End If
  	
	rs0.Close
    Set rs0 = Nothing 
    Set lgADF = Nothing                                                    '☜: ActiveX Data Factory Object Nothing
End Sub
'----------------------------------------------------------------------------------------------------------
Private Sub FixUNISQLData()

    Dim strVal
    Dim arrVal(2)
    Redim UNISqlId(3)                                                     '☜: SQL ID 저장을 위한 영역확보 

    Redim UNIValue(3,2)

    UNISqlId(0) = "S4212QA101"
    UNISqlId(1) = "S0000QA002"
    UNISqlId(2) = "S0000QA005"
    UNISqlId(3) = "S0000QA001"

    UNIValue(0,0) = lgSelectList                                          '☜: Select list

	strVal = " "
	
	If Len(Request("txtconBp_cd")) Then
		strVal = "AND A.APPLICANT = " & FilterVar(Request("txtconBp_cd"), "''", "S") & " "
		arrVal(0) = Trim(Request("txtconBp_cd"))
	Else
		strVal = ""
	End If

	If Len(Request("txtSalesGroup")) Then
		strVal = strVal & " AND A.SALES_GRP = " & FilterVar(Request("txtSalesGroup"), "''", "S") & " "		
		arrVal(1) = Trim(Request("txtSalesGroup"))
	End If		
    
 	If Len(Request("txtItem_cd")) Then
		strVal = strVal & " AND B.ITEM_CD = " & FilterVar(Request("txtItem_cd"), "''", "S") & " "		
		arrVal(2) = Trim(Request("txtItem_cd"))
	End If		
    
    If Len(Request("txtOpenFrDt")) Then
		strVal = strVal & " AND A.IV_DT >= " & FilterVar(UNIConvDate(Trim(Request("txtOpenFrDt"))), "''", "S") & ""		
	End If		
	
	If Len(Request("txtOpenToDt")) Then
		strVal = strVal & " AND A.IV_DT <= " & FilterVar(UNIConvDate(Trim(Request("txtOpenToDt"))), "''", "S") & ""		
	End If

	If Len(Request("txtIvNo")) Then
		strVal = strVal & " AND A.IV_NO = " & FilterVar(Request("txtIvNo"), "''", "S") & " "		
	End If

	If Len(Request("gBizArea")) Then
		strVal = strVal & " AND B.BIZ_AREA = " & FilterVar(Request("gBizArea"), "''", "S") & " "			
	End If

	If Len(Request("gPlant")) Then
		strVal = strVal & " AND B.PLANT_CD = " & FilterVar(Request("gPlant"), "''", "S") & " "			
	End If

	If Len(Request("gSalesGrp")) Then
		strVal = strVal & " AND A.SALES_GRP = " & FilterVar(Request("gSalesGrp"), "''", "S") & " "			
	End If

	If Len(Request("gSalesOrg")) Then
		strVal = strVal & " AND A.SALES_ORG = " & FilterVar(Request("gSalesOrg"), "''", "S") & " "			
	End If

    UNIValue(0,1) = strVal   '---발주일 
    UNIValue(1,0) = FilterVar(Trim(Request("txtconBp_cd")), " " , "S")
    UNIValue(2,0) = FilterVar(Trim(Request("txtSalesGroup")), " " , "S")
    UNIValue(3,0) = FilterVar(Trim(Request("txtItem_cd")), " " , "S")
    
    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    UNIValue(0,UBound(UNIValue,2)) = UCase(Trim(lgTailList))
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '☜: set ADO read mode
End Sub
'----------------------------------------------------------------------------------------------------------
Private Sub QueryData()
    Dim lgstrRetMsg                                             
    Dim lgADF                                                   
    Dim iStr
	
	FalsechkFlg = False

    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3)
    
    If  rs1.EOF And rs1.BOF Then
        rs1.Close
        Set rs1 = Nothing

   		If Len(Trim(Request("txtconBp_cd"))) And FalsechkFlg = False Then
		   Call DisplayMsgBox("970000", vbInformation, "수입자", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code
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
    
    If  rs2.EOF And rs2.BOF Then
        rs2.Close
        Set rs2 = Nothing

   		If Len(Trim(Request("txtSalesGroup"))) And FalsechkFlg = False Then
		   Call DisplayMsgBox("970000", vbInformation, "영업그룹", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code
		   FalsechkFlg = True
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

   		If Len(Trim(Request("txtItem_cd"))) And FalsechkFlg = False Then
		   Call DisplayMsgBox("970000", vbInformation, "품목", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code
		   FalsechkFlg = True
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
        
    If rs0.EOF And rs0.BOF And FalsechkFlg = False Then
        Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
        rs0.Close
        Set rs0 = Nothing
        ' Modify Focus Events
        %>
            <Script language=vbs>
            Parent.frm1.txtSalesGroup.focus    
            </Script>
        <%
    Else    
        Call  MakeSpreadSheetData()
    End If

End Sub
%>

<Script Language=vbscript>
    With parent
        .ggoSpread.Source    = .frm1.vspdData 
        
        .frm1.vspdData.Redraw = False
			.ggoSpread.SSShowDataByClip "<%=lgstrData%>","F"						'☜ : Display data
			Call .ReFormatSpreadCellByCellByCurrency(.frm1.vspdData,-1,-1,.GetKeyPos("A",2),.GetKeyPos("A",3),"A","Q","X","X")
			Call .ReFormatSpreadCellByCellByCurrency(.frm1.vspdData,-1,-1,.GetKeyPos("A",2),.GetKeyPos("A",4),"C","Q","X","X")
			        
        .lgStrPrevKey        =  "<%=lgStrPrevKey%>"                       '☜: set next data tag
        .frm1.txtconBp_Nm.value		= "<%=ConvSPChars(arrRsVal(1))%>" 
        .frm1.txtSalesGroupNm.value	= "<%=ConvSPChars(arrRsVal(3))%>" 
        .frm1.txtItem_Nm.value		= "<%=ConvSPChars(arrRsVal(5))%>"    
             	
        .frm1.txtHconBp_cd.value	= "<%=ConvSPChars(Request("txtconBp_cd"))%>"
		.frm1.txtHSalesGroup.value	= "<%=ConvSPChars(Request("txtSalesGroup"))%>"
		.frm1.txtHItem_cd.value		= "<%=ConvSPChars(Request("txtItem_cd"))%>"
		.frm1.txtHIvNo.value		= "<%=ConvSPChars(Request("txtIvNo"))%>"
		.frm1.txtHOpenFrDt.value	= "<%=Request("txtOpenFrDt")%>"
		.frm1.txtHOpenToDt.value	= "<%=Request("txtOpenToDt")%>"
         .DbQueryOk
         
         .frm1.vspdData.Redraw = True 
	End with
</Script>	
<%
    Response.End													'☜: 비지니스 로직 처리를 종료함 
%>

