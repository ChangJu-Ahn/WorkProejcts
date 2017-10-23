<%@ LANGUAGE=VBSCript%>
<!-- #Include file="../ESSinc/Adovbs.inc"  -->
<!-- #Include file="../ESSinc/incServerAdoDb.asp" -->
<!-- #Include file="../ESSinc/incServer.asp" -->
<!-- #Include file="../ESSinc/incSvrFuncSims.inc" -->
<!-- #Include file="../ESSinc/incSvrVarSims.inc"  -->

<%
Dim strPass

    On Error Resume Next                                                             '¢Ð: Protect system from crashing

    strPass = Request("txtPassword3")
    Call SubOpenDB(lgObjConn)                                                        '¢Ð: Make a DB Connection

    lgStrSQL = "UPDATE  E11002T"
    lgStrSQL = lgStrSQL & " SET " 
    lgStrSQL = lgStrSQL & " password =  " & FilterVar(strpass , "''", "S") & ","
    lgStrSQL = lgStrSQL & " updt_emp_no =  " & FilterVar(gEmpNo , "''", "S") & ""
    lgStrSQL = lgStrSQL & " WHERE UID = '"   & gUsrId & "'"

    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords

    Call SubCloseDB(lgObjConn)                                                       '¢Ð: Close DB Connection
%>
    <script language=vbscript>
		Sub Window_OnLoad()
		    Call parent.SaveOk()
		End Sub

        parent.close
    </script>
