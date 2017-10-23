<%@LANGUAGE = VBScript%>
<% Option Explicit%>
<%
'***********************************************************************************************************************
'*  1. Module Name          : 생산관리 
'*  2. Function Name        : 설비관리대장(HB)
'*  3. Program ID           : XP011OA_KO321
'*  4. Program Name         : 설비관리대장(HB)
'*  5. Program Desc         : 
'*  6. Comproxy List        :
'*  7. Modified date(First) : 2001/04/18
'*  8. Modified date(Last)  : 2005/07/29
'*  9. Modifier (First)     : Joo Young Hoon
'* 10. Modifier (Last)      : Joo Young Hoon
'* 11. Comment      
'************************************************************************************************************************
'☜ : 항상 서버 사이드 구문의 시작점인 좌꺽쇠(<)% 와 %우꺽쇠(>)는 New Line에 위치하여 
'	  서버 사이드 구문과 클라이언트 사이드 구문의 위치를 가늠할 수 있도록 한다.
'☜ : 아래 HTML 구문은 변경되어서는 안된다. 

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
Dim StrNextKey							' 다음 값 
Dim LngMaxRow							' 현재 그리드의 최대Row
Dim LngRow
Dim intGroupCount															'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 
Dim StrNext1
Dim StrNext2
Dim StrSeq
Dim lsSEQ


strMode = Request("txtMode")												'☜ : 현재 상태를 받음 

Call SubOpenDB(lgObjConn)                                                        '☜: Make a DB Connection

Select Case strMode

Case CStr(UID_M0001)														'☜: 현재 조회/Prev/Next 요청을 받음 
    
	Dim TNum
	Dim lsReqdlvyFromDt
	Dim lsReqdlvyToDt
	Dim txtFacilityCd, txtFacilityNM
	Dim JFlag, seltype
	
	JFlag=false
	

    Err.Clear                                                                '☜: Protect system from crashing

	lgMaxCount=100
	iSelCount = lgMaxCount + lgMaxCount *  lgStrPrevKeyIndex + 1
	 
	lsReqdlvyFromDt = trim(Request("txtReqdlvyFromDt"))
	lsReqdlvyToDt = trim(Request("txtReqdlvyToDT"))
	txtFacilityCd = trim(Request("txtFacilityCd"))
	txtFacilityNM = trim(Request("txtFacilityNM"))
	seltype = trim(Request("seltype"))

		
    lgStrSQL = " SELECT FACILITY_CD AS AB,FACILITY_NM AS BC, dbo.ufn_GetCodeName('Y6002', A.EMP_NO) AS CD, SET_DT AS DE,PROD_CO AS EF,PROD_AMT AS FG,PM_DT AS GH        "                    
    lgStrSQL = lgStrSQL & " FROM Y_FACILITY A "  
          
          
	if Trim(lsReqdlvyFromDt) <> "" Then 
		lgStrSQL = lgStrSQL & " where a.set_dt >= '" & lsReqdlvyFromDt & "' "  
		JFlag=true
	end if


	if Trim(lsReqdlvyToDt) <> "" Then
		if JFlag then
			lgStrSQL = lgStrSQL & " and a.set_dt <= '" & lsReqdlvyToDt & "' "	
		else
			lgStrSQL = lgStrSQL & " where a.set_dt <= '" & lsReqdlvyToDt & "' "
			JFlag=true
		end if
	end if
		
	if Trim(txtFacilityCd) <> "" Then
		if JFlag then
			lgStrSQL = lgStrSQL & " and   a.facility_Cd = '" & txtFacilityCd & "' " 
		else
			lgStrSQL = lgStrSQL & " where   a.facility_Cd = '" & txtFacilityCd & "' " 
			JFlag=true
		end if
	end if
	
	
	if Trim(txtFacilityNM) <> "" Then
		if JFlag then
			lgStrSQL = lgStrSQL & " and  a.facility_NM like '%" & txtFacilityNM & "%' " 		
		else
			lgStrSQL = lgStrSQL & " where  a.facility_NM like '%" & txtFacilityNM & "%' " 
			JFlag=true
		end if
	end if  
	
	
	if Trim(seltype) <> "" Then
		if JFlag then
			lgStrSQL = lgStrSQL & " and  a.emp_no='" & seltype & "' " 		
		else
			lgStrSQL = lgStrSQL & " where  a.emp_no='" & seltype & "' " 
			JFlag=true
		end if
	end if  
	   
    lgStrSQL = lgStrSQL & " order by ab "	

    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
		
        lgStrPrevKeyIndex = ""
        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
        Call SetErrorStatus()
    else 
		
		TNum=0		
		
		
		
        Do While Not lgObjRs.EOF
						
			lgstrData = lgstrData & Chr(11) & ConvSpChars(lgObjRs("ab")) 
			lgstrData = lgstrData & Chr(11) & ConvSpChars(lgObjRs("bc"))
			lgstrData = lgstrData & Chr(11) & ConvSpChars(lgObjRs("cd")) 
			lgstrData = lgstrData & Chr(11) & UNIDateClientFormat(lgObjRs("de")) 
			lgstrData = lgstrData & Chr(11) & ConvSpChars(lgObjRs("ef")) 
			lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("fg"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
			lgstrData = lgstrData & Chr(11) & UNIDateClientFormat(lgObjRs("gh"))	
			lgstrData = lgstrData & Chr(11) & Chr(12)					
		
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
			
		    lgObjRs.MoveNext
             
        Loop 

    End if 
 	
 	Call SubHandleError("MR",lgObjConn,lgObjRs,Err)
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
	Response.End																				 '☜: Process End

End Select


%>
