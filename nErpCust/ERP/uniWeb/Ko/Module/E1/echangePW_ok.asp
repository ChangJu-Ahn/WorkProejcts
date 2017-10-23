<%@ LANGUAGE=VBSCript%>
<!-- #Include file="../../inc/Adovbs.inc"  -->
<!-- #Include file="../../inc/incServerAdoDb.asp" -->
<!-- #Include file="../../inc/incServer.asp" -->
<!-- #Include file="../../inc/incSvrFuncSims.inc" -->
<!-- #Include file="../../inc/incSvrVarSims.inc"  -->

<%
Dim strPass
Dim Enc1, Enc2

    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

        Set Enc2 = Server.CreateObject("EDCodeCom.EDCodeObj.1")

        if err.number <> 0 then
%>
            <script language=vbscript>
                msgbox "CreateObject error - EDCodeCom", vbExclamation , _
                    "<%=gLogoName%>" & " login error 6"
            </script>
<%
            response.end
        end if

	'------ Developer Coding part (End   ) ------------------------------------------------------------------ 

    strPass = Request("txtPassword3")
'    Response.Write strPass
    strPass = Enc2.Encode(strPass)
    Call SubOpenDB(lgObjConn)                                                        '¢Ð: Make a DB Connection

    '--------------------------------------------------------------------------------------------------------
    lgStrSQL = "UPDATE  E11002T"
    lgStrSQL = lgStrSQL & " SET " 
    lgStrSQL = lgStrSQL & " password =  " & FilterVar(strpass , "''", "S") & ","
    lgStrSQL = lgStrSQL & " updt_emp_no =  " & FilterVar(gEmpNo , "''", "S") & ""
    lgStrSQL = lgStrSQL & " WHERE UID = '"   & gUsrId & "'"

    '---------- Developer Coding part (End  ) ---------------------------------------------------------------

    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords

    Set Enc2 = Nothing
    Call SubCloseDB(lgObjConn)                                                       '¢Ð: Close DB Connection
%>
    <script language=vbscript>
        parent.close
    </script>
