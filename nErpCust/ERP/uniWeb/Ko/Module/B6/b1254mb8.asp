<%
'************************************************************************************
'*  1. Module Name          : 영업 
'*  2. Function Name        : 
'*  3. Program ID           : b1254mb8
'*  4. Program Name         : 영업그룹조회 
'*  5. Program Desc         : 영업그룹조회 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2000/03/21
'*  8. Modified date(Last)  : 2002/04/21
'*  9. Modifier (First)     : Mr Cho
'* 10. Modifier (Last)      : Park in sik
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**************************************************************************************
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDBAgent.inc" -->
<%													

On Error Resume Next

Dim lgDataExist
Dim lgPageNo

Dim lgADF                                                                  '☜ : ActiveX Data Factory 지정 변수선언 
Dim lgstrRetMsg                                                            '☜ : Record Set Return Message 변수선언 
Dim UNISqlId, UNIValue, UNILock, UNIFlag,rs0			                   '☜ : DBAgent Parameter 선언 
Dim rs1, rs2 ,rs3
Dim lgstrData                                                              '☜ : data for spreadsheet data
Dim lgStrPrevKey                                                           '☜ : 이전 값 
Dim lgTailList                                                             '☜ : Orderby절에 사용될 field 리스트 
Dim lgSelectList
Dim lgSelectListDT
Dim strSalesGrp1	                                                       
Dim strSalesOrg	                                                           
Dim strCostcenter
Dim BlankchkFlg	                                                           
Const C_SHEETMAXROWS_D  = 30               

Dim arrRsVal(5)								
  
	Call LoadBasisGlobalInf()
    Call HideStatusWnd 

	lgPageNo       = UNICInt(Trim(Request("lgPageNo")),0)					'☜:
    lgStrPrevKey   = Request("lgStrPrevKey")                               '☜ : Next key flag    
    lgSelectList   = Request("lgSelectList")                               '☜ : select 대상목록 
    lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)             '☜ : 각 필드의 데이타 타입 
    lgTailList     = Request("lgTailList")                                 '☜ : Orderby value
	lgDataExist    = "No"

    Call TrimData()
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
       rs0.Move     = CLng(C_SHEETMAXROWS_D) * CLng(lgPageNo)                  'C_SHEETMAXROWS_D:Max Fetched Count at once , lgStrPrevKeyIndex : Previous PageNo
    End If
    
    iLoopCount = -1
    
   Do while Not (rs0.EOF Or rs0.BOF)
   
        iLoopCount =  iLoopCount + 1
        iRowStr = ""
        
		For ColCnt = 0 To UBound(lgSelectListDT) - 1 
            iRowStr = iRowStr & Chr(11) & FormatRsString(lgSelectListDT(ColCnt),rs0(ColCnt))
		Next
 
        If iLoopCount < C_SHEETMAXROWS_D Then
           lgstrData = lgstrData & iRowStr & Chr(11) & Chr(12)
        Else
           lgPageNo = lgPageNo + 1
           Exit Do
        End If
        
        rs0.MoveNext
	Loop

    If iLoopCount < C_SHEETMAXROWS_D Then                                      '☜: Check if next data exists
       lgPageNo = ""
    End If
    rs0.Close                                                       '☜: Close recordset object
    Set rs0 = Nothing	                                            '☜: Release ADF
End Sub

'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()
    Dim strVal
    Redim UNISqlId(4)           '☜: SQL ID 저장을 위한 영역확보 
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
    Redim UNIValue(4,2)

    UNISqlId(0) = "B1254MA801"			'* : 데이터 조회를 위한 SQL문 만듬 
	
	UNISqlId(1) = "B1254MA802"			'영업그룹 
	UNISqlId(2) = "B1254MA803"			'영업조직 
	UNISqlId(3) = "B1254MA804"			'비용집계처 
	
	'--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    UNIValue(0,0) = lgSelectList                                          '☜: Select list
     
	UNIValue(1,0)  = UCase(Trim(strSalesGrp1))
    UNIValue(2,0)  = UCase(Trim(strSalesOrg))
    UNIValue(3,0)  = UCase(Trim(strCostcenter))
    
	'--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
	
	strVal = " a.sales_grp is not null "
    
	If Trim(Request("txtSales_Grp1")) <> "" Then
		strVal = strVal& " And A.SALES_GRP >= " & FilterVar(Trim(UCase(Request("txtSales_Grp1"))), " " , "S") & "  AND A.SALES_GRP <= " & FilterVar("zzzzzzzzz", "''", "S") & "  "
	Else
		strVal = strVal& "" 
	End If
	
	If Trim(Request("txtSales_Org")) <> "" Then
		strVal = strVal& " AND A.SALES_ORG >= " & FilterVar(Trim(UCase(Request("txtSales_Org"))), " " , "S") & "  AND A.SALES_ORG <=  " & FilterVar(Trim(UCase(Request("txtSales_Org"))), " " , "S") & " "
	Else
		strVal = strVal& "" 
	End If
	
	If Trim(Request("txtCost_center")) <> "" Then
		strVal = strVal& " AND A.COST_CD >= " & FilterVar(Trim(UCase(Request("txtCost_center"))), " " , "S") & "  AND A.COST_CD <=  " & FilterVar(Trim(UCase(Request("txtCost_center"))), " " , "S") & " "
	Else
		strVal = strVal& "" 
	End If

	If Trim(Request("txtRadio")) <> "" Then
		strVal = strVal& " AND A.USAGE_FLAG >= " & FilterVar(UCase(Request("txtRadio")), "''", "S") & " AND A.USAGE_FLAG <=  " & FilterVar(UCase(Request("txtRadio")), "''", "S") & ""
	Else
		strVal = strVal& ""
	End If
  		
    UNIValue(0,1) = strVal   
	
    UNIValue(0,UBound(UNIValue,2)) = " " & UCase(Trim(lgTailList))
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
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag,rs0,rs1,rs2,rs3) '* : Record Set 의 갯수 조정 
    
    Set lgADF   = Nothing
    
    iStr = Split(lgstrRetMsg,gColSep)

    
	
	'============================= 추가된 부분 =====================================================================
    If  rs1.EOF And rs1.BOF Then
        rs1.Close
        Set rs1 = Nothing
        If Len(Request("txtSales_Grp1")) And BlankchkFlg = False Then
		   Call DisplayMsgBox("970000", vbInformation, "영업그룹", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code
	       BlankchkFlg = True
            ' Modify Focus Events    
            %>
                <Script language=vbs>
                Parent.frm1.txtSales_Grp1.focus    
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
        If Len(Request("txtSales_Org")) And BlankchkFlg = False Then
		   Call DisplayMsgBox("970000", vbInformation, "영업조직", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code
	       BlankchkFlg = True
            ' Modify Focus Events    
            %>
                <Script language=vbs>
                Parent.frm1.txtSales_Org.focus    
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
        If Len(Request("txtCost_center")) And BlankchkFlg = False Then
		   Call DisplayMsgBox("970000", vbInformation, "비용집계처", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code
	       BlankchkFlg = True
            ' Modify Focus Events    
            %>
                <Script language=vbs>
                Parent.frm1.txtCost_center.focus    
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
		If  rs0.EOF And rs0.BOF And BlankchkFlg = False Then
		    Call DisplayMsgBox("125400", vbOKOnly, "", "", I_MKSCRIPT)
		    rs0.Close
		    Set rs0 = Nothing
            ' Modify Focus Events    
            %>
                <Script language=vbs>
                Parent.frm1.txtSales_Grp1.focus    
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
	
	'영업그룹 
    If Len(Trim(Request("txtSales_Grp1"))) Then
    	strSalesGrp1 = " " & FilterVar(Trim(Request("txtSales_Grp1")), " " , "S") & " "
    	
    Else
    	strSalesGrp1 = "''"
    End If
    '영업조직 
    If Len(Trim(Request("txtSales_Org"))) Then
    	strSalesOrg = " " & FilterVar(Trim(Request("txtSales_Org")), " " , "S") & " "
    Else
    	strSalesOrg = "''"
    End If
		
	'비용집계처 
    If Len(Trim(Request("txtCost_center"))) Then
    	strCostcenter = " " & FilterVar(Trim(Request("txtCost_center")), " " , "S") & " "
    Else
    	strCostcenter = "''"
    End If

End Sub


%>
<Script Language=vbscript>
    parent.frm1.txtSales_Grp_nm1.value		=  "<%=ConvSPChars(arrRsVal(1))%>" 	
  	parent.frm1.txtSales_Org_Nm.value		=  "<%=ConvSPChars(arrRsVal(3))%>" 	
  	parent.frm1.txtCost_center_nm.value		=  "<%=ConvSPChars(arrRsVal(5))%>" 	
	If "<%=lgDataExist%>" = "Yes" Then
		With parent
			If "<%=lgPageNo%>" = "1" Then   ' "1" means that this query is first and next data exists
				.frm1.HSales_Grp.value		= "<%=ConvSPChars(Request("txtSales_Grp1"))%>"
				.frm1.HSales_Org.value		= "<%=ConvSPChars(Request("txtSales_Org"))%>"			
				.frm1.HCost_Center.value	= "<%=ConvSPChars(Request("txtCost_center"))%>"
				.frm1.HRadio.value			= "<%=Request("txtRadio")%>"
			End If
			.ggoSpread.Source    = .frm1.vspdData 
			.ggoSpread.SSShowDataByClip "<%=lgstrData%>"          '☜: Display data 
			.lgPageNo      =  "<%=lgPageNo%>"               '☜ : Next next data tag		
			.DbQueryOk
		
		End with
	End If   
</Script>	
<%
    Response.End													'☜: 비지니스 로직 처리를 종료함 
%>
