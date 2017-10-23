<%@LANGUAGE = VBScript%>
<% Option Explicit%>
<%
'***********************************************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 일자별AS클레임집계(HB)
'*  3. Program ID           : P5225MA1
'*  4. Program Name         : 일자별AS클레임집계(HB)
'*  5. Program Desc         :
'*  6. Comproxy List        : +S31111MaintSoHdrSvr
'*  7. Modified date(First) : 2005/05/03
'*  8. Modified date(Last)  : 2005/06/02
'*  9. Modifier (First)     : Yoo Myung Sik
'* 10. Modifier (Last)      : Chen, Jae Hyun
'* 11. Comment              :
'************************************************************************************************************************

%>
<!-- #Include file="../../inc/IncServer.asp" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->	
<!-- #Include file="../../inc/adovbs.inc" -->	
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->		

<%													'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 

On Error Resume Next														'☜: 

    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '☜: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)

    lgLngMaxRow       = Request("txtMaxRows")                                        '☜: Read Operation Mode (CRUD)
    lgMaxCount        = CInt(Request("lgMaxCount"))                                  '☜: Fetch count at a time for VspdData
    lgStrPrevKeyIndex = UNICInt(Trim(Request("lgStrPrevKeyIndex")),0)                '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)



Call HideStatusWnd

Dim strMode		
Dim LngMaxRow							' 현재 그리드의 최대Row
Dim LngRow

Dim lsReqdlvyFromDt
Dim lsReqdlvyToDt
Dim seltype, txtCastCd

strMode = Request("txtMode")												'☜ : 현재 상태를 받음 

Call SubOpenDB(lgObjConn)                                                        '☜: Make a DB Connection

Select Case strMode

	Case CStr(UID_M0001)														'☜: 현재 조회/Prev/Next 요청을 받음 

	    Err.Clear                                                                '☜: Protect system from crashing
	    
	    lgMaxCount=100
		iSelCount = lgMaxCount + lgMaxCount *  lgStrPrevKeyIndex + 1
		 
		lsReqdlvyFromDt = trim(Request("txtReqdlvyFromDt"))
		lsReqdlvyToDt = trim(Request("txtReqdlvyToDT"))	
		seltype = trim(Request("seltype"))
		txtCastCd = trim(Request("txtCastCd"))
		
	    '설비명 조회 
	    If Trim(txtCastCd) <> "" Then
			lgStrSQL = "SELECT FACILITY_NM FROM Y_FACILITY (NOLOCK) WHERE FACILITY_CD = " & Filtervar(txtCastCd, "''", "S")      
			If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
				Response.Write " <Script Language=""VBScript""> " & vbCrLf
				Response.Write " Parent.frm1.txtCastNm.value =  """"" & vbCrLf
				Response.Write " </Script> " & vbCrLf	
			Else
				Response.Write " <Script Language=""VBScript""> " & vbCrLf
				Response.Write " Parent.frm1.txtCastNm.value =  """ & ConvSPChars(lgObjRs("FACILITY_NM")) & """" & vbCrLf
				Response.Write " </Script> " & vbCrLf
			End if   
		Else
			Response.Write " <Script Language=""VBScript""> " & vbCrLf
			Response.Write " Parent.frm1.txtCastNm.value =  """"" & vbCrLf
			Response.Write " </Script> " & vbCrLf	   
	    End If
	    
	    '메인 조회 
	    lgStrSQL = "Select a.fac_cast_cd as ab,b.facility_nm as bc, dbo.ufn_GetCodeName('z410', facility_accnt) as cd,a.work_dt as de,a.insp_text as ef,isnull(insp_hour,0) as fg  "                    
	    lgStrSQL = lgStrSQL & ",isnull(insp_min,0) as gh,insp_dept as hi "  
	    lgStrSQL = lgStrSQL & " from	y_fac_cast_plan as a,y_facility as b "  
	    lgStrSQL = lgStrSQL & " where	 a.fac_cast_cd=b.facility_cd "  
	    lgStrSQL = lgStrSQL & " and a.gubun_cd='10' "  
	    lgStrSQL = lgStrSQL & " and plan_gubun='1' "   
	       
		if Trim(lsReqdlvyFromDt) <> "" Then 
			lgStrSQL = lgStrSQL & " and a.work_dt >= '" & lsReqdlvyFromDt & "' "  
		end if

		if Trim(lsReqdlvyToDt) <> "" Then 
			lgStrSQL = lgStrSQL & " and a.work_dt <=  '" & lsReqdlvyToDt & "' " 		
		end if
		
		if Trim(seltype) <> "" Then 
			lgStrSQL = lgStrSQL & " and facility_accnt = '" & seltype & "' "  
		end if

		if Trim(txtCastCd) <> "" Then 
			lgStrSQL = lgStrSQL & " and a.fac_cast_cd = '" & txtCastCd & "' " 		
		end if
	   
	    lgStrSQL = lgStrSQL & " order by ab"
	     
	    'call svrmsgbox(lgstrsql,vbinformation,i_mkscript)
	   
	    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
			
	        lgStrPrevKeyIndex = ""
	        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
	    Else 
	        Do While Not lgObjRs.EOF
						
						lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ab")) 
						lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("bc")) 
						lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("cd")) 
						lgstrData = lgstrData & Chr(11) & UNIDateClientFormat(lgObjRs("de")) 
						lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ef")) 
						lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("fg"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
						lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("gh"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
						lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("hi"))	
						lgstrData = lgstrData & Chr(11) & Chr(12)

			
		'------ Developer Coding part (End   ) ------------------------------------------------------------------

			    lgObjRs.MoveNext

	        Loop 


	    End if 
 	
    Call SubCloseRs(lgObjRs)                                                              '☜: Release RecordSSet

%>
    
<Script Language="VBScript">

    With Parent
       .ggoSpread.Source     = .frm1.vspdData
       .ggoSpread.SSShowData "<%=lgstrData%>"

       .DBQueryOk                  
	End with
       
</Script>	

<%
	Call SubCloseDB(lgObjConn)                                                       '☜: Close DB Connection
	Response.End																				'☜: Process End
End Select
%>
