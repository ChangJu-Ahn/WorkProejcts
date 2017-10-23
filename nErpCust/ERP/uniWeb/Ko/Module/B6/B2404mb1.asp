<%
'**********************************************************************************************
'*  1. Module Name          : Basis Architect
'*  2. Function Name        : Organization(horg_his 부서변경History)
'*  3. Program ID           : B2404mb1.asp
'*  4. Program Name         : B2404mb1.asp
'*  5. Program Desc         :
'*  6. Comproxy List        : +B24041ControlHorgHis
'                             +B24048ListHorgHis
'*  7. Modified date(First) : 2000/10/27
'*  8. Modified date(Last)  : 2002/12/16
'*  9. Modifier (First)     : Hwang Jeong-won
'* 10. Modifier (Last)      : Sim Hae Young
'* 11. Comment              :
'**********************************************************************************************
Option Explicit		
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/lgsvrvariables.inc"  -->
<!-- #Include file="../../inc/adovbs.inc"  -->
<!-- #Include file="../../inc/incServeradodb.asp"  -->
<%
Dim strSpread

On Error Resume Next                                                             '☜: Protect system from crashing
Err.Clear                                                                        '☜: Clear Error status

Call LoadBasisGlobalInf()

Call HideStatusWnd                                                               '☜: Hide Processing message                                                
lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)   
strSpread         = Request("txtSpread")
  
Select Case lgOpModeCRUD    
    Case CStr(UID_M0001)                                                         '☜: Query
         Call SubBizQueryMulti()
    Case CStr(UID_M0002)                                                         '☜: Save,Update
         Call SubBizSaveMulti()
    Case CStr("Gen")        
         Call SubBatch()     
End Select
Sub SubBizQueryMulti() 	
 	
 	On Error Resume Next 
	Dim PB6G071		
	Dim importorgid
	Dim importoldorgid	
    
    importorgid = Request("importorgid")
    importoldorgid = Request("importoldorgid") 	
 	
    Set PB6G071 = server.CreateObject ("PB6G071.cBListHorgHis")     
    
    If CheckSYSTEMError(Err,True) = True Then
        set PB6G071 = nothing
        Response.End  
    End If	
	on error goto 0
	
	On Error Resume Next	
    lgstrData = PB6G071.B_READ_HORG_HIS(gStrGlobalCollection,importorgid,importoldorgid)      
    If CheckSYSTEMError(Err,True) = True Then
        set PB6G071 = nothing
        Response.End  
    End If	
	on error goto 0  
	
    Set PB6G071 = nothing  
 
End Sub

Sub SubBizSaveMulti()

	On Error Resume Next  
    Dim PB6G071  
    
    Set PB6G071 = server.CreateObject("PB6G071.cBControlHorgHis")  
	If CheckSYSTEMError(Err,True) = True Then
        set PB6G071 = nothing
        Response.End  
    End If	
	on error goto 0 
	 
    On Error Resume Next  
    
    call PB6G071.CONTROL_HORG_HIS(gStrGlobalCollection,strSpread)
    If CheckSYSTEMError(Err,True) = True Then
        set PB6G071 = nothing
        Response.End  
    End If	
	on error goto 0 
    
    Set  PB6G071 = nothing  

End Sub

Sub SubBatch()

    Dim iDx
    Dim iLoopMax
    Dim iKey1

	lgStrSQL = "Select DEPT, LDEPTNM "
    lgStrSQL = lgStrSQL & " From horg_mas "
    lgStrSQL = lgStrSQL & " Where orgid =  " & FilterVar(Request("txtOldOrgId"), "''", "S") & " "
    
    Call SubOpenDB(lgObjConn)                               '☜: Make a DB Connection        
    
        
    If FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
       lgStrPrevKeyIndex = ""
       Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
    Else
       Call SubSkipRs(lgObjRs,lgMaxCount * lgStrPrevKeyIndex)
       lgstrData = ""
       iDx       = 1
       Do While Not lgObjRs.EOF
            lgstrData = lgstrData & Chr(11) & lgObjRs("DEPT")
            lgstrData = lgstrData & Chr(11) & ""
            lgstrData = lgstrData & Chr(11) & lgObjRs("LDEPTNM")                  
            lgstrData = lgstrData & Chr(11) & ""
            lgstrData = lgstrData & Chr(11) & ""
            lgstrData = lgstrData & Chr(11) & ""
            lgstrData = lgstrData & Chr(11) & ""
            lgstrData = lgstrData & Chr(11) & ""
            lgstrData = lgstrData & Chr(11) & lgLngMaxRow + iDx
            lgstrData = lgstrData & Chr(11) & Chr(12)

    	    lgObjRs.MoveNext          
 
            iDx =  iDx + 1                       
        Loop     
    End If      
      
    Call SubCloseRs(lgObjRs)  
 %>
<Script Language=vbscript> 
        With Parent
			.frm1.vspdData.MaxRows = 0
            .ggoSpread.Source     = .frm1.vspdData      
            .ggoSpread.SSShowData "<%=ConvSPChars(lgstrData)%>"
			.Batch_OK()        
		End with   
</Script>
<%
Call SubCloseDB(lgObjConn)                          
End Sub
%>

<Script Language="VBScript">
    Select Case "<%=lgOpModeCRUD %>"
       Case "<%=UID_M0001%>"       
             With Parent
                .ggoSpread.Source  = .frm1.vspdData                
                .ggoSpread.SSShowData "<%=ConvSPChars(lgstrData)%>"
                .DBQueryOk
	         End with       
       Case "<%=UID_M0002%>"      
                 Parent.DBSaveOk                 
    End Select       
</Script>	
