<%
'********************************************************************************************************
'*  1. Module Name          : 영업 
'*  2. Function Name        : 매출채권관리 
'*  3. Program ID           : S3111BB1
'*  4. Program Name         : 수주참조 
'*  5. Program Desc         : 매출채권등록 참조화면 
'*  6. Comproxy List        :
'*  7. Modified date(First) : 2000/04/20
'*  8. Modified date(Last)  : 2002/04/17
'*  9. Modifier (First)     : Cho song hyon
'* 10. Modifier (Last)      : Ahn Tae Hee
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*                            -2000/04/20 : 3rd 화면 layout & ASP Coding
'*                            -2000/08/11 : 4th 화면 layout
'*                            -2001/12/18 : Date 표준적용 
'*                            -2002/04/17 : ADO변환 
'********************************************************************************************************
%>
<!-- #Include file="../../inc/incSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDBAgent.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%													

On Error Resume Next

Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3, rs4                           '☜ : DBAgent Parameter 선언 
Dim lgstrData                                                              '☜ : data for spreadsheet data
Dim lgMaxCount                                                             '☜ : 한번에 가져올수 있는 데이타 건수 
Dim lgTailList                                                             '☜ : Orderby절에 사용될 field 리스트 
Dim lgSelectList													       '☜ : select 대상목록 
Dim lgSelectListDT														   '☜ : 각 필드의 데이타 타입	
Dim lgDataExist
Dim lgPageNo
   
'--------------- 개발자 coding part(변수선언,Start)--------------------------------------------------------
'--------------- 개발자 coding part(변수선언,End)----------------------------------------------------------
Call LoadBasisGlobalInf()
Call LoadInfTB19029B("I", "*", "NOCOOKIE", "RB")
'	Call LoadBNumericFormatB("Q", "*", "NOCOOKIE", "PB")
Call HideStatusWnd 

lgPageNo       = UNICInt(Trim(Request("lgPageNo")),0)              '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
lgMaxCount     = 30						                          '☜ : 한번에 가져올수 있는 데이타 건수 
lgSelectList   = Request("lgSelectList")                               '☜ : select 대상목록 
lgTailList     = Request("lgTailList")                                 '☜ : Order by value
lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)             '☜ : 각 필드의 데이타 타입 
lgDataExist      = "No"

Call  FixUNISQLData()                                                '☜ : DB-Agent로 보낼 parameter 데이타 set
call  QueryData()                                                    '☜ : DB-Agent를 통한 ADO query


'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
'QueryData()에 의해서 Query가 되면 MakeSpreadSheetData()에 의해서 데이터를 스프레드시트에 뿌려주는 프로시져 
'----------------------------------------------------------------------------------------------------------
Sub MakeSpreadSheetData()

    Dim iLoopCount                                                                     
    Dim iRowStr
    Dim ColCnt
    
    lgDataExist    = "Yes"
    lgstrData      = ""
  
    If CLng(lgPageNo) > 0 Then
       rs0.Move     = CLng(lgMaxCount) * CLng(lgPageNo)                  'lgMaxCount:Max Fetched Count at once , lgStrPrevKeyIndex : Previous PageNo
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
Sub SetConditionData()
End Sub

'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
' Query하기 전에  DB Agent 배열을 이용하여 Query문을 만드는 프로시져 
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()

    Dim strVal															  '☜:UNISqlId(0)에 들어가는 입력변수 
																		  '아래에 보면 화면단에서 넣어 주는 query시 where조건임을 알 수 있다.	
    Redim UNISqlId(0)                                                        '☜: SQL ID 저장을 위한 영역확보 
    Redim UNIValue(0,2)													  '☜: 각각의 SQL ID와 입력될 where 조건의 쌍으로 된 2차원 배열 

    UNISqlId(0) = "S3111BA101"  ' main query(spread sheet에 뿌려지는 query statement)
	    
    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    UNIValue(0,0) = lgSelectList                                          '☜: Select list
																		  '	UNISqlId(0)의 첫번째 ?에 입력됨				
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------

	strVal = ""

	If Len(Trim(Request("txtSoldtoParty"))) Then
		strVal = " AND SH.sold_to_party  =  " & FilterVar(Request("txtSoldtoParty"), "''", "S") & ""
	End If
		
	If Len(Trim(Request("txtSalesGrp"))) Then
		strVal = strVal & " AND SH.SALES_GRP =  " & FilterVar(Request("txtSalesGrp"), "''", "S") & ""
	End If
		
	If Len(Trim(Request("txtBillType"))) Then
		strVal = strVal & " AND BT.BILL_TYPE =  " & FilterVar(Request("txtBillType"), "''", "S") & ""
	End If
		
	If Len(Trim(Request("txtCurrency"))) Then
		strVal = strVal & " AND SH.CUR =  " & FilterVar(Request("txtCurrency"), "''", "S") & ""
	End If
		
	If Len(Trim(Request("txtSOType"))) Then
		strVal = strVal & " AND SH.SO_TYPE = " & FilterVar(Request("txtSOType"), "''", "S") & ""
	End If
		
	If Len(Trim(Request("txtFromDt"))) Then
		strVal = strVal & " AND SH.SO_DT >=  " & FilterVar(UNIConvDate(Request("txtFromDt")), "''", "S") & ""
	End If
		
	If Len(Trim(Request("txtToDt"))) Then
		strVal = strVal & " AND SH.SO_DT <=  " & FilterVar(UNIConvDate(Request("txtToDt")), "''", "S") & ""
	End If
		
	 UNIValue(0,1) = strVal				'UNISqlId(0)의 두번째 ?에 입력됨	
	  
     '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    UNIValue(0,UBound(UNIValue, 2)) = UCase(Trim(lgTailList))			  '	UNISqlId(0)의 마지막 ?에 입력됨	
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '☜: set ADO read mode
	 
End Sub
 
'----------------------------------------------------------------------------------------------------------
' Query Data
' ADO의 Record Set이용하여 Query를 하고 Record Set을 넘겨서 MakeSpreadSheetData()으로 Spreadsheet에 데이터를 
' 뿌림 
' ADO 객체를 생성할때 prjPublic.dll파일을 이용한다.(상세내용은 vb로 작성된 prjPublic.dll 소스 참조)
'----------------------------------------------------------------------------------------------------------
	Sub QueryData()
	    Dim lgstrRetMsg                                             '☜ : Record Set Return Message 변수선언 
		Dim lgADF                                                   '☜ : ActiveX Data Factory 지정 변수선언 
	    Dim iStr
		
		Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")									'☜:ADO 객체를 생성 
	    
	    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0,rs1,rs2,rs3,rs4)

	    Set lgADF   = Nothing    
	    iStr = Split(lgstrRetMsg,gColSep)
		
		
	    If iStr(0) <> "0" Then
	        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
	    End If    
	    
	    If  rs0.EOF And rs0.BOF And FalsechkFlg = False Then
	        rs0.Close
	        Set rs0 = Nothing
	        Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)	'No Data Found!!
			%>
			<Script Language=vbscript>
			parent.frm1.txtSoldToParty.focus
			</Script>	
			<%
	    Else   
	        Call  MakeSpreadSheetData()
			Call  SetConditionData()
	    End If
	   
	End Sub
%>   

<Script Language=vbscript>
	With parent.frm1
	
		<%If lgDataExist = "Yes" Then%>
		'Set condition data to hidden area
		' "1" means that this query is first and next data exists	
			<%If lgPageNo = "1" Then %>
			.txtHSoldToParty.value	= "<%=Request("txtSoldToParty")%>"
			.txtHFromDt.value		= "<%=Request("txtFromDT")%>"
			.txtHToDt.value			= "<%=Request("txtToDT")%>"
			.txtHSalesGrp.value		= "<%=Request("txtSalesGrp")%>"
			.txtHBillType.value		= "<%=Request("txtBillType")%>"
			.txtHCurrency.value		= "<%=Request("txtCurrency")%>"
			.txtHSOType.value		= "<%=Request("txtSOType")%>"
			<%End If%>
	    'Show multi spreadsheet data from this line
		.vspdData.Redraw = False                
		parent.ggoSpread.Source = .vspdData 
		parent.ggoSpread.SSShowDataByClip "<%=lgstrData%>", "F"
		parent.lgPageNo      =  "<%=lgPageNo%>"               '☜ : Next next data tag
        parent.DbQueryOk
		.vspdData.Redraw = True                
		<%End If%>
	End With
</Script>	
<%
Response.End
%>
