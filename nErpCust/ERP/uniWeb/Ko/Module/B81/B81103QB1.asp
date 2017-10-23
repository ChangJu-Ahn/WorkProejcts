<%@ LANGUAGE="VBSCRIPT" %>
<%Option Explicit    %>
<!--
'**********************************************************************************************
'*  1. Module Name          : CIS
'*  2. Function Name        : 
'*  3. Program ID           : B81103MB1
'*  4. Program Name         : 품목구성코드조회 
'*  5. Program Desc         : 품목구성코드등록()
'*  6. Component List       : PM1G121.cMMntSpplItemPriceS
'*  7. Modified date(First) : 2005/01/23
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : lee wol san
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change" 
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************
-->
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<%	
Const C_SHEETMAXROWS_D  = 100     
call LoadBasisGlobalInf()
call LoadInfTB19029B("I", "*","NOCOOKIE","MB") 
call LoadBNumericFormatB("I","*","NOCOOKIE","MB")

    Dim lgOpModeCRUD
    Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0                 '☜ : DBAgent Parameter 선언 
    Dim rs1, rs2, rs3, rs4,rs5
	Dim istrData
    Dim lgPageNo
	Dim iErrorPosition
	Dim strSpread
	Dim FromReqDt , ToReqDt
	
	
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    Call HideStatusWnd                                                               '☜: Hide Processing message
	
    lgOpModeCRUD  = Request("txtMode")
    FromReqDt     = Request("txtFromReqDt")
    ToReqDt       = Request("txtToReqDt")
	strSpread = Request("txtSpread")
	
	Call SubBizQueryMulti()
	  										                                              '☜: Read Operation 
   
'============================================================================================================
' Name : SubBizQueryMulti
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
   On Error Resume Next

	lgPageNo       = UNICInt(Trim(Request("lgPageNo")),1)    '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
	
	Call FixUNISQLData()
	Call QueryData()	
	

End Sub    
	    

'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
' Query하기 전에  DB Agent 배열을 이용하여 Query문을 만드는 프로시져 
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()
    Dim strVal
	Redim UNISqlId(3)                                                     '☜: SQL ID 저장을 위한 영역확보 
    Redim UNIValue(2,4)                                                 '⊙: DB-Agent로 전송될 parameter를 위한 변수 
                                                              
     UNISqlId(0) = "B81102MA101" 											' header
     UNISqlId(1) = "B81QB_MINOR" 	
     UNISqlId(2) = "B81QB_MINOR" 
     
     UNIValue(0,0)=" TOP " & lgPageNo * C_SHEETMAXROWS_D & " A.ITEM_ACCT,B.MINOR_NM,A.ITEM_KIND,C.MINOR_NM,CASE A.ITEM_LVL WHEN 'L1' THEN '대분류' WHEN 'L2' THEN '중분류' "
     UNIValue(0,0)= UNIValue(0,0) & " WHEN 'L3' THEN '소분류' END, A.CLASS_CD,A.CLASS_NAME,A.PARENT_CLASS_CD,"
     UNIValue(0,0)= UNIValue(0,0) & " dbo.ufn_s_CIS_GetParentNn(A.ITEM_ACCT,A.ITEM_KIND,A.PARENT_CLASS_CD,'', A.ITEM_LVL),A.REMARK "
     UNIValue(0,1)="LIKE " & FilterVar(Request("txtItem_acct")&"%", "''", "S") & ""
     UNIValue(0,2)="LIKE " & FilterVar(Request("txtItem_kind")&"%", "''", "S") & ""
     
     UNIValue(0,3)="AND convert(varchar(12),A.INSRT_DT ,112) BETWEEN '"&uniConvDate(FromReqDt)&"' AND '"&uniConvDate(ToReqDt)&"' ORDER BY A.ITEM_ACCT,A.ITEM_KIND,A.ITEM_LVL, A.PARENT_CLASS_CD,A.CLASS_CD "
     
     UNIValue(1,0) ="'P1001'"
     UNIValue(1,1) ="" & FilterVar(Request("txtItem_acct"), "''", "S") & ""
     
     UNIValue(2,0) ="'Y1001'"
     UNIValue(2,1) ="" & FilterVar(Request("txtItem_kind"), "''", "S") & ""
  
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
    
    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1,rs2)

	Set lgADF   = Nothing
	iStr = Split(lgstrRetMsg,gColSep)

	If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
		Response.end
    End If 
 
    '----- UI 각 항목 체크 ----
    'if trim( Request("txtItem_acct")) <> "" then     call fnCheckItem (rs1,"txtItem_acct","품목계정"   ) 
    'if trim( Request("txtItem_kind")) <> "" then call fnCheckItem (rs2,"txtItem_kind","품목구분"   ) 
    
  If  rs0.EOF And rs0.BOF  Then
       
        Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
        
        rs0.Close
        Set rs0 = Nothing
		Response.Write "<Script Language=VBScript>" & vbCrLF
		Response.Write "Call parent.SetToolBar(""11000000000011"") " & vbCrLF
		Response.Write "</Script>" & vbCrLF
		Response.end
    ELSE
		
         call ListupDataGrid (rs0.getRows,"")
         
    End If  
End Sub


'----------------------------------------------------------------------------------------------------------
'ListupDataGrid
'----------------------------------------------------------------------------------------------------------

 Sub ListupDataGrid(pArr,dataFormatCol)
	Dim strData
	Dim i,j,moveLine,RowCnt
	RowCnt=0
	moveLine = (lgPageNo - 1) * C_SHEETMAXROWS_D
		for i=moveLine to uBound(pArr,2)
		RowCnt=RowCnt+1
			for j=0 to uBound(pArr,1)
			
			if inStr(dataFormatCol,"," & j&",") > 0 then
				strData = strData & Chr(11) & UniConvDateDbToCompany(pArr(j,i),"")
			else
				strData = strData & Chr(11) & trim(ConvSPChars(pArr(j,i)))
			end if	
			
		
			next 
			strData =  strData & Chr(11) & i &  Chr(11) & Chr(12) 
		next 
		
		Response.Write "<Script Language=vbscript>" & vbCr
		Response.Write "With parent" & vbCr
		Response.Write "	.ggoSpread.Source       = .frm1.vspdData "			& vbCr
		Response.Write "    .frm1.vspdData.Redraw = False   "                  & vbCr   
		Response.Write "	.ggoSpread.SSShowData     """ & strData	 & """" & ",""F""" & vbCr
		Response.Write "	.DbQueryOk " & vbCr 
		Response.Write  "   .frm1.vspdData.Redraw = True " & vbCr
		Response.Write "	.lgPageNo  = """ & lgPageNo + 1 & """" & vbCr 
		if RowCnt<C_SHEETMAXROWS_D then
			Response.Write "    .lgPageNo= """"  "                  & vbCr 
		end if
		Response.Write "End With"		& vbCr
		Response.Write "</Script>"		& vbCr
		
End Sub	





%>
