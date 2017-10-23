<%@ page language="java" contentType="text/html; charset=utf-8" pageEncoding="utf-8"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<%@ include file="/WEB-INF/include/include-header.jspf"%>
</head>
<body>
	<jsp:include page="../sample/top.jsp" flush="false"></jsp:include>
	<H2>시스템 기준정보</H2>
	<form id="frm" name="frm" enctype="multipart/form-data">
		<table class="board_view">
			<tr>
				<th>시스템 구분</th>
				<td colspan="3">
					<input type="radio" id="Major" name="GUNUN" onclick="fn_view(this)" style="display: inline;"> Major
					<input type="radio" id="Minor" name="GUNUN" onclick="fn_view(this)" style="display: inline;"> Minor
				</td>
			</tr>
		</table>
	</form>
	<div class="grid_Layer">
		<br />
		<table id="grid"></table>
		<div id="pager"></div>
	</div>
	<%@ include file="/WEB-INF/include/include-body.jspf"%>
	<script type="text/javascript">
		
		function fn_view(input){
			switch(input.id.toUpperCase())
			{
			 	case "MAJOR" : 
			 		alert("major");
			 		break;
			 		
			 	case "MINOR" : 
			 		alert("minor");
			 		break;
			}
		}
	
    </script>
</body>
</html>