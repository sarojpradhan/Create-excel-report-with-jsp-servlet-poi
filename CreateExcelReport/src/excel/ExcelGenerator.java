package excel;

import java.io.IOException;
import java.io.PrintWriter;
import java.text.SimpleDateFormat;
import java.util.List;
import java.util.Map;
import java.util.ArrayList;
import java.util.HashMap;

import javax.servlet.ServletException;
import javax.servlet.ServletOutputStream;
import javax.servlet.annotation.WebServlet;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;

@WebServlet("/ExcelGenerator")
public class ExcelGenerator extends HttpServlet {
	private static final long serialVersionUID = 1L;

	protected void doPost(HttpServletRequest request, HttpServletResponse response)
			throws ServletException, IOException {

		String departmentId = request.getParameter("selDepartment"); // get user selected department id
		System.out.println("Department id: " + departmentId); // debug point

		request.setAttribute("message", "Please select proper department!"); // set message for end users if they don't select proper department
		
		//if end user selects proper department, then call printStaffInformationInExcel () method to print excel report otherwise display the message
		if (departmentId.equals("0")) {
			request.getRequestDispatcher("index.jsp").forward(request, response);
		} else {
			printStaffInformationInExcel(request, response, departmentId);
		}

	}

	public Map<String, List<Staff>> getStaffInDepartment() {

		// create HashMap to add department id and staffs associated in it
		Map<String, List<Staff>> staffInDepartment = new HashMap<String, List<Staff>>();

		// Create staffs
		Staff staff1 = new Staff("John", "Walker", "xyx1234", "+358478512364", "john.walker@xyz.com");
		Staff staff2 = new Staff("David", "Parera", "xyx7854", "+358477952364", "david.parera@xyz.com");
		Staff staff3 = new Staff("Katty", "Kit", "xyx7235", "+358478963364", "katty.kit@xyz.com");
		Staff staff4 = new Staff("Sam", "Rickshow", "xyx8954", "+358478365643", "sam.rickshow@xyz.com");
		Staff staff5 = new Staff("Michel", "Perl", "xyx7452", "+35847837843", "michel.perl@xyz.com");
		Staff staff6 = new Staff("Harry", "porter", "xyx7854", "+358478365741", "harry.porter@xyz.com");
		Staff staff7 = new Staff("Angelina", "jolly", "xyx8856", "+358478312643", "angelina.jolly@xyz.com");

		// Add staffs in HR department
		List<Staff> staffInHR = new ArrayList<Staff>();
		staffInHR.add(staff1);
		staffInHR.add(staff2);

		// Add staffs in IT department
		List<Staff> staffInIT = new ArrayList<Staff>();
		staffInIT.add(staff3);
		staffInIT.add(staff4);

		// Add staffs in Account department
		List<Staff> staffInAccount = new ArrayList<Staff>();
		staffInAccount.add(staff5);

		// Add staffs in Administration department
		List<Staff> staffInAdministration = new ArrayList<Staff>();
		staffInAdministration.add(staff6);
		staffInAdministration.add(staff7);

		// There is no staffs for now in Marketing and Finance departments
		List<Staff> staffInMarketing = new ArrayList<Staff>();
		List<Staff> staffInFinance = new ArrayList<Staff>();

		// Add department id and list of staffs associated to it
		staffInDepartment.put("1", staffInHR);
		staffInDepartment.put("2", staffInIT);
		staffInDepartment.put("3", staffInAccount);
		staffInDepartment.put("4", staffInAdministration);
		staffInDepartment.put("5", staffInMarketing);
		staffInDepartment.put("6", staffInFinance);

		System.out.println("The size of HashMap1 is: " + staffInDepartment.size());
		return staffInDepartment;
	}

	public void printStaffInformationInExcel(HttpServletRequest request, HttpServletResponse response,
			String departmentId) throws ServletException, IOException {

		Map<String, List<Staff>> staffInDepartment = new HashMap<String, List<Staff>>();
		staffInDepartment = getStaffInDepartment();

		List<Staff> listOfStaffInDepartment = new ArrayList<Staff>();

		// add staffs in the list based on department selected by user
		for (String departId : staffInDepartment.keySet()) {

			if (departId.equals(departmentId)) {
				System.out.println("DEPARTMENT ID:" + departId); // debug point
				listOfStaffInDepartment = staffInDepartment.get(departId);
			}
		}

		// if there is staff in the department then create excel report otherwise print
		// message to end users
		if (listOfStaffInDepartment.size() != 0) {

			String timestamp = new SimpleDateFormat("yyyy.MM.dd.HH.mm.ss").format(new java.util.Date()); // get
																											// current
																											// timestamp
			String excelFileName = "STAFF-INFO" + timestamp + ".xls"; // create excel file name and add current
																		// timestamp so that each file will be
																		// unique
			// create excel and make it ready for the download
			response.setContentType("application/vnd.ms-excel");
			response.setCharacterEncoding("UTF-8");
			response.setHeader("Content-Disposition", "attachment; filename=\"" + excelFileName + "\"");

			String[] titles = { "First Name", "Last Name", "Staff Number", "Phone", "Email" };
			try {

				// Creating a Workbook
				HSSFWorkbook workbook = new HSSFWorkbook();

				// Creating spreadsheet
				HSSFSheet spreadsheet = workbook.createSheet(" STAFF-INFO ");

				// For creating font bold
				CellStyle boldStyleForCell = workbook.createCellStyle();
				Font font = workbook.createFont();// Create font
				font.setBold(true);// Make font bold
				boldStyleForCell.setFont(font);

				// Print data in the left side of the excel cell
				// CellStyle dateCellStyle = workbook.createCellStyle();
				// dateCellStyle.setAlignment(HorizontalAlignment.LEFT);

				// In excel row number 1 print info
				HSSFRow excelRow1 = spreadsheet.createRow(1);
				Cell cellForReportInfo = excelRow1.createCell(0);
				cellForReportInfo.setCellStyle(boldStyleForCell);
				cellForReportInfo.setCellValue("Staff Information");

				Cell cellForDateTitle = excelRow1.createCell(1);
				cellForDateTitle.setCellValue("Print date and time:");

				Cell cellForDate = excelRow1.createCell(2);
				cellForDate.setCellValue(timestamp);

				// Setting column name at the top of excel file
				HSSFRow headerRow = spreadsheet.createRow(3);

				for (int i = 0; i < titles.length; i++) {
					Cell cell = headerRow.createCell(i);
					cell.setCellStyle(boldStyleForCell);
					cell.setCellValue(titles[i]);
				}

				int rowCount = 4;
				for (Staff staff : listOfStaffInDepartment) {
					HSSFRow excelRow4 = spreadsheet.createRow(rowCount++);

					excelRow4.createCell(0).setCellValue(staff.getFirst_name());
					excelRow4.createCell(1).setCellValue(staff.getLast_name());
					excelRow4.createCell(2).setCellValue(staff.getStaff_number());
					excelRow4.createCell(3).setCellValue(staff.getPhone());
					excelRow4.createCell(4).setCellValue(staff.getEmail());

				}

				// Auto adjusting of the columns of excel file
				for (int i = 0; i < titles.length; i++) {
					spreadsheet.autoSizeColumn(i);
				}

				// Write data in the excel
				ServletOutputStream out = response.getOutputStream();
				workbook.write(out);

				// Close output stream and workbook
				workbook.close();

				out.flush();
				out.close();

			} catch (Exception e) {
				e.printStackTrace();
			} finally {
			}
		} else {
			// print message to end users when there is no staff in selected department
			request.setAttribute("message", "Staff not available!");
			request.getRequestDispatcher("index.jsp").forward(request, response);
		}
	}

}
