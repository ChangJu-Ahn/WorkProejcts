<%
'======================================================================================================
'*  1. Module Name          : CIS
'*  2. Function Name        : 
'*  3. Program ID           : B81108QB1.asp
'*  4. Program Name         : B81108QB1.asp
'*  5. Program Desc         : 통합코드조회
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
    
    Dim FromReqDt,ToReqDt,req_user,item_kind,item_cd
    Dim FromEndDt,ToEndDt,item_spec
    
    
    rbo_gbn		=  Request("rbo_gbn")
    txtStatus		= trim(request("rbo_status") )  
    txtItemAcct	= trim(request("txtItemAcct") )   

     
     FromReqDt     = trim(Request("txtFromReqDt"))
     ToReqDt       = trim(Request("txtToReqDt") )
     FromEndDt     = trim(request("txtFromEndDt"))
     ToEndDt       = trim(request("txtToEndDt"))
     req_user	   = trim(Request("txtreq_user"))
     item_kind     = trim(Request("txtitem_kind"))
     item_cd       = trim(Request("txtitem_cd")) 
     item_spec     = trim(Request("txtitem_spec")) 
    
    if FromEndDt="" then FromEndDt="1900-01-01"
    if ToEndDt="" then ToEndDt="2999-12-31"
    

    UNISqlId(0) = "B81108QA1" 
	UNIValue(0,0)="TOP " & lgStrToKey * C_SHEETMAXROWS_D & " "
	UNIValue(0,0)= UNIValue(0,0) & "dbo.ufn_GetCodeName('Y1001' ,  A.ITEM_KIND ) KIND_NM, A.ITEM_CD, A.ITEM_NM,A.ITEM_SPEC,dbo.ufn_s_CIS_GetStatus(B.STATUS),"
	UNIValue(0,0)= UNIValue(0,0) & "dbo.ufn_GetCodeName('Y1006' , B.REQ_ID ) ,B.REQ_DT,B.END_DT, A.REMARK"
			
	UNIValue(0,1)= "AND B.END_DT BETWEEN "&filterVar(FromEndDt,"''","S")&" AND " & filterVar(ToEndDt,"''","S")
	UNIValue(0,1)= UNIValue(0,1) & " AND B.REQ_DT BETWEEN "&filterVar(FromReqDt,"''","S")&" AND " & filterVar(ToReqDt,"''","S")
	
	if txtStatus<>"'*'" then UNIValue(0,1)= UNIValue(0,1) & "AND STATUS IN("&txtStatus&")"
	if req_user<>"" then  UNIValue(0,2)= "AND B.REQ_ID like " & filterVar( req_user&"%","''","S")
	if txtItemAcct<>"" then  UNIValue(0,3)= "AND A.ITEM_ACCT like " & filterVar( txtItemAcct&"%","''","S")
	if item_kind<>"" then  UNIValue(0,4)= "AND A.ITEM_KIND like " & filterVar( item_kind&"%","''","S")
	if item_cd<>"" then  UNIValue(0,5)= "AND A.ITEM_CD like " & filterVar( item_cd&"%","''","S")
	
	if item_spec<>"" then  UNIValue(0,5)=  UNIValue(0,5) & "AND A.ITEM_SPEC like " & filterVar( item_spec&"%","''","S")
	
	
	 
	 		
	 UNIValue(0,5) = UNIValue(0,5) & " ORDER BY 1"
     
		 
    
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

	  call ListupDataGrid (rs0.getRows,",6,7,")
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











