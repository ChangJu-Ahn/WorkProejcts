<%
'======================================================================================================
'*  1. Module Name          : CIS
'*  2. Function Name        : 
'*  3. Program ID           : B81107QB1.asp
'*  4. Program Name         : B81107QB1.asp
'*  5. Program Desc         : 의뢰현황조회
'*  6. Modified date(First) :  2005/01/23
'*  7. Modified date(Last)  : 
'*  8. Modifier (First)     : Lee Wol san
'*  15. Modifier (Last)      :
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================
%>
<% Option Explicit %>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgSvrVariables.inc" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->

<%													'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 

Call HideStatusWnd									'☜: 모든 작업 완료후 작업진행중 표시창을 Hide

Dim strMode											'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 
Dim strSpread
Dim StrNextKey		' 다음 값 
Dim lgStrPrevKey	' 이전 값 
Dim LngMaxRow		' 현재 그리드의 최대Row
Dim LngRow         
Dim RowData(5)
Dim RowDataPre
Dim lgSelectList
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0 ,RS1

Dim rtnReq_no
DIM lgStrToKey
Const C_SHEETMAXROWS_D =100
call LoadBasisGlobalInf()


lgStrToKey = UNICInt(Trim(Request("lgStrToKey")),1)   


    '---------------------------------------Common-----------------------------------------------------------
    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '☜: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    'Multi SpreadSheet
    strSpread = Request("txtSpread")
    lgLngMaxRow       = Request("txtMaxRows")  
   
  
    
    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)    
             Call SubBizQueryMulti()
             

    End Select


'============================================================================================================
' Name : SubBizQueryMulti
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
  
	Call FixUNISQLData()
	Call QueryData()	
	
	
End Sub    
	    
	    
'============================================================================================================
' Set DB Agent arg
'============================================================================================================

Sub FixUNISQLData()
    Dim strVal
	Redim UNISqlId(0)                                                     
    Redim UNIValue(0,5)    
    dim txtItemAcct ,rbo_gbn,txtStatus
    
    Dim FromReqDt,ToReqDt
    
      rbo_gbn=  Request("rbo_gbn")
      txtStatus	= trim(request("rbo_status") )  
      txtItemAcct	= trim(request("txtItemAcct") )        
      FromReqDt     = filterVar(UNIConvDate(Request("txtFromReqDt")),"''","S")
      ToReqDt       = filterVar(UNIConvDate(Request("txtToReqDt")),"''","S")
     
        
     IF rbo_gbn="N" then
            '---------------------   
            '신규의뢰 
            '---------------------
            UNISqlId(0) = "B81107QA1" 
			UNIValue(0,0)="TOP " & lgStrToKey * C_SHEETMAXROWS_D & " "
			UNIValue(0,0)= UNIValue(0,0) & "REQ_NO, 'N',ITEM_CD,ITEM_NM,ITEM_SPEC,dbo.ufn_s_CIS_GetStatus(STATUS),dbo.ufn_GetCodeName('Y1007' ,R_GRADE),dbo.ufn_GetCodeName('Y1008' ,T_GRADE),  dbo.ufn_GetCodeName('Y1008' ,P_GRADE) ,dbo.ufn_GetCodeName('Y1008' ,Q_GRADE),REQ_DT,REMARK"
			UNIValue(0,1)= "AND B.REQ_DT BETWEEN "&FromReqDt&" AND " & ToReqDt
			if txtStatus<>"'*'" then UNIValue(0,1)= UNIValue(0,1) & "AND STATUS IN("&txtStatus&")"
			if txtItemAcct<>"" then  UNIValue(0,2)= "AND ITEM_ACCT like " & filterVar( txtItemAcct&"%","''","S")
			UNIValue(0,2) = UNIValue(0,2) & " ORDER BY 1"
     
     
     elseif rbo_gbn="C" then
		    '---------------------   
            '품목변경 
            '---------------------
            UNISqlId(0) = "B81107QA2" 
			UNIValue(0,0)="TOP " & lgStrToKey * C_SHEETMAXROWS_D & " "
			UNIValue(0,0)= UNIValue(0,0) & "B.REQ_NO,'C', A.ITEM_CD,A.ITEM_NM,A.ITEM_SPEC,B.STATUS, dbo.ufn_GetCodeName('Y1007' ,R_GRADE),dbo.ufn_GetCodeName('Y1008' ,T_GRADE),  dbo.ufn_GetCodeName('Y1008' ,P_GRADE) ,dbo.ufn_GetCodeName('Y1008' ,Q_GRADE),B.REQ_DT,B.REMARK"
			UNIValue(0,1)= "AND B.REQ_DT BETWEEN "&FromReqDt&" AND " & ToReqDt
			if txtStatus<>"'*'" then UNIValue(0,1)= UNIValue(0,1) & "AND STATUS IN("&txtStatus&")"
			if txtItemAcct<>"" then  UNIValue(0,2)= "AND A.ITEM_ACCT like " & filterVar( txtItemAcct&"%","''","S")
			UNIValue(0,2) = UNIValue(0,2) & " ORDER BY 1"
     else 'P
			'---------------------   
            '품명/규격변경 
            '---------------------
            UNISqlId(0) = "B81107QA3" 
			UNIValue(0,0)="TOP " & lgStrToKey * C_SHEETMAXROWS_D & " "
			UNIValue(0,0)= UNIValue(0,0) & "B.REQ_NO,'P', A.ITEM_CD,A.ITEM_NM,A.ITEM_SPEC,B.STATUS,dbo.ufn_GetCodeName('Y1007' ,R_GRADE),dbo.ufn_GetCodeName('Y1008' ,T_GRADE),  dbo.ufn_GetCodeName('Y1008' ,P_GRADE) ,dbo.ufn_GetCodeName('Y1008' ,Q_GRADE),B.REQ_DT,B.REMARK"
			UNIValue(0,1)= "AND B.REQ_DT BETWEEN "&FromReqDt&" AND " & ToReqDt
			if txtStatus<>"'*'" then UNIValue(0,1)= UNIValue(0,1) & "AND STATUS IN("&txtStatus&")"
			if txtItemAcct<>"" then  UNIValue(0,2)= "AND ITEM_ACCT like " & filterVar( txtItemAcct&"%","''","S")
			UNIValue(0,2) = UNIValue(0,2) & " ORDER BY 1"
     
     end if
                                       

   
     UNILock = DISCONNREAD :	UNIFlag = "1"                                 '☜: set ADO read mode
   
End Sub




'============================================================================================================
' Query Data
' ADO의 Record Set이용하여 Query를 하고 Record Set을 넘겨서 MakeSpreadSheetData()으로 Spreadsheet에 데이터를 
' 뿌림 
' ADO 객체를 생성할때 prjPublic.dll파일을 이용한다.(상세내용은 vb로 작성된 prjPublic.dll 소스 참조)
'============================================================================================================
Sub QueryData()
    Dim lgstrRetMsg                                             
    Dim lgADF                                                  
    Dim iStr
    
    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1)

	Set lgADF   = Nothing
	iStr = Split(lgstrRetMsg,gColSep)

	If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
		Response.end
    End If 
 
 
  If  rs0.EOF And rs0.BOF  Then
        Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
        rs0.Close
        Set rs0 = Nothing
	
		Response.end
    ELSE

	  call ListupDataGrid (rs0.getRows,",10,")
    End If  
End Sub
'============================================================================================================
' ListupDataGrid
'============================================================================================================

 Sub ListupDataGrid(pArr,dataFormatCol)
	Dim strData
	Dim i,j,moveLine,RowCnt
	RowCnt=0
	moveLine = (lgStrToKey - 1) * C_SHEETMAXROWS_D
	
		for i=moveLine to uBound(pArr,2)
		RowCnt=RowCnt+1
			for j=0 to uBound(pArr,1)
			
			if inStr(dataFormatCol,"," & j&",") > 0 then
				strData = strData & Chr(11) & UNIDateClientFormat(pArr(j,i))
			else
				
				strData = strData & Chr(11) & replace(ConvSPChars(pArr(j,i)),chr(13)&chr(10),"  ")
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
		Response.Write "	.lgStrToKey  = """ & lgStrToKey + 1 & """" & vbCr 
		if RowCnt<C_SHEETMAXROWS_D then
			Response.Write "    .lgStrToKey= """"  "                  & vbCr 
		end if
		Response.Write "End With"		& vbCr
		Response.Write "</Script>"		& vbCr
		
End Sub	



%>











