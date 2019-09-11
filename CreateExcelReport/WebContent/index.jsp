<%@ page language="java" contentType="text/html; charset=ISO-8859-1"
	pageEncoding="ISO-8859-1"%>
<!DOCTYPE html>
<html>
<head>
<meta charset="ISO-8859-1">
<title>ExcelGenerator</title>
<link rel="stylesheet"
	href="${pageContext.request.contextPath}/css/excel.css" />
<script
	src="https://ajax.googleapis.com/ajax/libs/jquery/3.4.1/jquery.min.js"></script>
</head>
<body>
	<h2>Print staffs information in excel based on department!</h2>

	<form action="ExcelGenerator" method="Post">

		<div>Select Department: &nbsp; &nbsp;</div>

		<div>
			<select id="select-department" name="selDepartment">
				<option value="0">--Select Department--</option>
				<option value="1">HR</option>
				<option value="2">IT</option>
				<option value="3">Account</option>
				<option value="4">Administration</option>
				<option value="5">Marketing</option>
				<option value="5">Finance</option>
			</select>
		</div>

		<div class="errorMessage">&nbsp; &nbsp; ${message}</div>
		<br> <br> <input type="submit" value="Print" name="btSubmit">

		<script>
			$(document).ready(function() {
				$('#select-department').change(function() {
					var departValue = this.value;
					var departmentId = $('#select-department').val();

					if (departValue == departmentId) {
						$('.errorMessage').hide();
					}
				});

			});
		</script>

	</form>
</body>
</html>