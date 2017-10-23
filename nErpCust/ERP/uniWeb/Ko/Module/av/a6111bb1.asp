<%@ LANGUAGE=VBSCript %>
<%Option Explicit%>
<!-- #Include file="../../inc/incSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/incServerAdoDb.asp"  -->
<!-- #Include file="../../inc/lgsvrvariables.inc"  -->
<% 
	Call LoadBasisGlobalInf() 

    On Error Resume Next														'☜: 
    Err.Clear                                                                        '☜: Clear Error status    Dim strMode

    Call HideStatusWnd                                                               '☜: Hide Processing message
    Dim strMode
    strMode      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
 
 
 Dim strFilePath
 
    Select Case strMode
        Case CStr(UID_M0001)                                                         '☜: Query 
            Call SubMakeDisk()

        Case CStr(UID_M0002)            
			Call SubFileDownLoad()          
          
        Case CStr(UID_M0003)                                                         '☜: Delete
            Call SubFileDownLoad2()
          
    End Select
   
%>    
    <script language="vbscript">
		Dim SF
		
		On Error Resume Next
		
		Set SF = CreateObject("uni2kCM.SaveFile")
		Call SF.SaveTextFile("<%=strFilePath %>")
	
		Set SF = Nothing
		
</script>

<%    
'-----------------------------------------------------------------------------------------------
'            File DownLoad(With B.A)
'-----------------------------------------------------------------------------------------------
Sub SubFileDownLoad()
    On Error Resume Next    
    Err.Clear                                                               '☜: Protect system from crashing
    
    Call HideStatusWnd

    strFilePath = "http://" & Request.ServerVariables("LOCAL_ADDR") & ":" _
               & Request.ServerVariables("SERVER_PORT")
    If Instr(1, Request.ServerVariables("URL"), "Module") <> 0 Then
        strFilePath = strFilePath & Mid(Request.ServerVariables("URL"), 1, InStr(1, Request.ServerVariables("URL"), "Module") - 1)     
    End If

    strFilePath = strFilePath  & "files/u2000/"
    strFilePath = strFilePath & Request("txtFileName")


End Sub

'--------------------------------------------------------------------------------------------------------
'					            File DownLoad(With xinSoft solution)
'--------------------------------------------------------------------------------------------------------
Sub SubFileDownLoad2()
    On Error Resume Next    
    Err.Clear                                                               '☜: Protect system from crashing

    Dim xdn

    Call HideStatusWnd
    set xdn = Server.CreateObject("Xionsoft.XionFileDownLoad")


    'strFilePath = "/" & Mid(Request.ServerVariables("URL"), 2, instr(2, Request.ServerVariables("URL"), "/") - 2) & "/" & gLang & "/files/" & gCompany & "/"
    'strFilePath = "/" & Mid(Request.ServerVariables("URL"), 2, instr(2, Request.ServerVariables("URL"), "/") - 2) & "/"
    strFilePath = "/" & Mid(Request.ServerVariables("URL"), 2, instr(2, Request.ServerVariables("URL"), "/") - 2) & "/template/files/" & gCompany & "/"

    'Call ServerMesgBox(strFilePath, vbInformation, I_MKSCRIPT)

    '다음 두 줄 임시 COMMENT 2001/1/27
    xdn.DownFromFile strFilePath & Request("txtFileName")
    xdn.OnEndPage

    set xdn = nothing

    Response.End


End Sub


'--------------------------------------------------------------------------------------------------------
'					         
'--------------------------------------------------------------------------------------------------------

Sub SubMakeDisk()
    On Error Resume Next    
    Err.Clear   
		
    Dim IArrData 
    Dim iPAVG010


    Dim ex1_file_name
    Dim ex2_return_code

    Const I1_ief_supplied   = 0
    Const I1_biz_area_cd	= 1
    Const I1_start_issued_dt =2
    Const I1_end_issued_dt =3
    Const I1_report_issued_dt =4
    Const I1_file_name =5
    Const I1_file_path =6

    Redim IArrData(I1_file_path)

     IArrData(I1_ief_supplied) = "B" '누락분 디스켓생성 구분자 

     IArrData(I1_biz_area_cd)	=	UCase(Trim(Request("txtBizAreaCD")))
     IArrData(I1_start_issued_dt) =	UNIConvDate(Request("txtIssueDt1"))
     IArrData(I1_end_issued_dt)	=	UNIConvDate(Request("txtIssueDt2"))

     IArrData(I1_report_issued_dt)	=	UNIConvDate(Request("txtReportDt"))
     IArrData(I1_file_name)			=	Request("txtFileName") & ""

     strFilePath = Server.MapPath("../../files/u2000") & "\"

     IArrData(I1_file_path)			=	 strFilePath


    Set iPAVG010 = Server.CreateObject("PAVG010.cbAVatDiskSvrEab")
            
    If CheckSYSTEMError(Err,True) = True Then
        Exit Sub
    End If


    Call iPAVG010.EAB_A_VAT_DISK_SVR(gStrGlobalCollection,IArrData,ex1_file_name,ex2_return_code)


    If CheckSYSTEMError(Err, True) = True Then					
           Set iPAVG010 = Nothing
           Exit Sub
    End If    

            
            DIM FileName
            
            FileName=ex1_file_name
            
            Call HideStatusWnd
            Set iPAVG010 = Nothing
             lgStrSQL =  " 		select " 
			lgStrSQL = lgStrSQL & " 	TAX_BIZ_AREA_CD,"
			lgStrSQL = lgStrSQL & " 	TAX_BIZ_AREA_NM "
			lgStrSQL = lgStrSQL & " From B_TAX_BIZ_AREA "
			lgStrSQL = lgStrSQL & " Where TAX_BIZ_AREA_CD = " & FilterVar(Trim(Request("txtBizAreaCD")), "''","S") 

			Call SubOpenDB(lgObjConn)                                                        '☜: Make a DB Connection
			If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = True Then                    'If data not exists
%>
				<Script language="VBScript">
					With Parent.frm1 
						.txtBizAreaNM.value = "<%=ConvSPChars(lgObjRs("TAX_BIZ_AREA_NM"))%>"
					End With                                                       

				</Script>  	
<%
			end if
			Call SubCloseRs(lgObjRs) 


            Response.Write " <SCRIPT LANGUAGE=VBSCRIPT>" & vbCr
            Response.write " parent.subVatDiskOK(""" & FileName & """)" & vbCr
            response.write "</SCRIPT>" & vbCr

    Call HideStatusWnd
    Set iPAVG010 = Nothing														'☜: Unload Comproxy
Response.End
		
End Sub
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
%>


