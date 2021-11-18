package com.gmoz.utils;

import java.io.IOException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;

import org.apache.poi.hssf.usermodel.HSSFHeader;
import org.apache.poi.hssf.usermodel.HeaderFooter;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.PrintSetup;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFPrintSetup;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.extensions.XSSFHeaderFooter;
import org.springframework.stereotype.Component;

import com.gmoz.entity.StudentEntity;

@Component
public class FileExporter {

	private XSSFWorkbook workbook;
	private XSSFSheet sheet;
	private int rowIndex;

	/**
	 * 
	 * @param sheetName		Tên Sheet
	 * @param tableName		Tên tiêu đề bảng
	 * @param limit			Số hàng mỗi bảng
	 * @param list			Danh sách dữ liệu bảng
	 * @param response		response from servlet
	 * @throws IOException
	 */
	public void exportToExcel(String sheetName, String tableName, int limit, List<StudentEntity> list,
			HttpServletResponse response) throws IOException {
		rowIndex = 0;
		workbook = new XSSFWorkbook();
		setResponseHeader(response, "application/octet-stream", ".xlsx", sheetName + "_");
		sheet = workbook.createSheet(sheetName);

		// Set the Page margins
		sheet.setMargin(Sheet.LeftMargin, 0.25);
		sheet.setMargin(Sheet.RightMargin, 0.25);
		sheet.setMargin(Sheet.TopMargin, 0.75);
		sheet.setMargin(Sheet.BottomMargin, 0.75);
//		sheet.setAutobreaks(true);
		sheet.setFitToPage(true);

		// Set the Header and Footer Margins
		sheet.setMargin(Sheet.HeaderMargin, 0.25);
		sheet.setMargin(Sheet.FooterMargin, 0.25);

		// Setup print layout settings
		XSSFPrintSetup layout = sheet.getPrintSetup();
		layout.setLandscape(false);
		layout.setFitWidth((short) 1);
		layout.setFitHeight((short) 0);
//		layout.setPaperSize(PrintSetup.A4_PAPERSIZE);
		layout.setPaperSize(PrintSetup.LEGAL_PAPERSIZE);
		layout.setFooterMargin(0.25);

		// Write Header and Footer
		XSSFHeaderFooter header = (XSSFHeaderFooter) sheet.getHeader();
//		header.setCenter(HSSFHeader.font("Tahoma", "Bold") + HSSFHeader.fontSize((short) 20) + "TITLE");
		header.setRight(HSSFHeader.font("Stencil-Normal", "Normal") + HSSFHeader.fontSize((short) 15) + "Trang "
				+ HeaderFooter.page());

		System.out.println("export to excel row = " + rowIndex);

		// Write data rows
		writeDataLine(tableName, list, limit);
		ServletOutputStream outStream = response.getOutputStream();
		workbook.write(outStream);

		// Close file
		workbook.close();
		outStream.close();
	}

	/**
	 * Thiết lập tên file trả về
	 * 
	 * @param response		Phản hổi trả về từ HttpServlet
	 * @param contentType	Kiểu dữ liệu trả về
	 * @param extension		Đuôi dữ liệu trả về
	 * @param prefix		Tiền tố file trả về
	 * @throws IOException
	 */
	public void setResponseHeader(HttpServletResponse response, String contentType, String extension, String prefix)
			throws IOException {
		DateFormat dateFormat = new SimpleDateFormat("dd-MM-yyyy_HH-mm-ss");
		String timeStamp = dateFormat.format(new Date());
		String fileName = prefix + timeStamp + extension;

		response.setContentType(contentType);

		String headerKey = "Content-Disposition";
		String headerValue = "attachment; filename=" + fileName;
		response.setHeader(headerKey, headerValue);
	}

	/**
	 * Tạo tiêu đề sheet
	 * 
	 * @param title	Tên tiêu đề
	 * @param clazz	Đối tượng cần lấy tên
	 */
	private void writeTitle(String title, Class<?> clazz) {
		sheet.addMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 0, clazz.getDeclaredFields().length - 1));
		XSSFRow row = sheet.createRow(rowIndex);
		XSSFCellStyle titleStyle = workbook.createCellStyle();
		XSSFFont font = workbook.createFont();
		font.setBold(true);
		font.setFontHeight(20);
		font.setFontName("Tahoma");
		titleStyle.setFont(font);
		titleStyle.setAlignment(HorizontalAlignment.CENTER);
		createCell(row, 0, title, titleStyle);

		rowIndex += 2;
	}
	
	/**
	 * Write Header of table
	 */
	private void writeHeaderLine() {
		System.out.println("create header row = " + rowIndex);
		XSSFRow row = sheet.createRow(rowIndex);

		// Set Font style
		XSSFFont font = workbook.createFont();
		font.setBold(true);
		font.setFontHeight(13);
		font.setFontName("Tahoma");

		// Set Cell style
		XSSFCellStyle cellStyle = workbook.createCellStyle();
		cellStyle.setFont(font);
		cellStyle.setAlignment(HorizontalAlignment.CENTER);
		cellStyle.setBorderLeft(BorderStyle.THIN);
		cellStyle.setBorderRight(BorderStyle.THIN);
		cellStyle.setBorderTop(BorderStyle.THIN);
		cellStyle.setBorderBottom(BorderStyle.THIN);

		createCell(row, 0, "STT", cellStyle);
		createCell(row, 1, "Mã Sinh viên", cellStyle);
		createCell(row, 2, "Họ và tên", cellStyle);
		createCell(row, 3, "Giới tính", cellStyle);
		createCell(row, 4, "Ngày sinh", cellStyle);
		createCell(row, 5, "Số điện thoại", cellStyle);
		createCell(row, 6, "Tuổi", cellStyle);

		rowIndex++;
	}

	/**
	 * Create a cell
	 * 
	 * @param row			Hàng hiện tại
	 * @param columnIndex	Chỉ số cột
	 * @param value			Giá trị đối tượng
	 * @param style			Định dạng cell
	 */
	private void createCell(XSSFRow row, int columnIndex, Object value, CellStyle style) {
		XSSFCell cell = row.createCell(columnIndex);
		sheet.autoSizeColumn(columnIndex);
		if (value instanceof Integer) {
			cell.setCellValue((Integer) value);
		} else if (value instanceof Boolean) {
			cell.setCellValue((Boolean) value);
		} else if (value instanceof String) {
			cell.setCellValue((String) value);
		}
		cell.setCellStyle(style);
	}

	private static final SimpleDateFormat sdf = new SimpleDateFormat("dd/MM/yyyy");

	/**
	 * 
	 * @param tableName Tên tiêu đề của bảng
	 * @param list 		Danh sách dữ liệu bảng
	 * @param limit		Số hàng mỗi bảng
	 */
	private void writeDataLine(String tableName, List<StudentEntity> list, int limit) {
		// Set Font style
		XSSFFont font = workbook.createFont();
		font.setFontHeight(13);
		font.setFontName("Tahoma");

		// Set Cell style
		XSSFCellStyle cellStyle = workbook.createCellStyle();
		cellStyle.setFont(font);
		cellStyle.setBorderLeft(BorderStyle.THIN);
		cellStyle.setBorderRight(BorderStyle.THIN);
		cellStyle.setBorderTop(BorderStyle.THIN);
		cellStyle.setBorderBottom(BorderStyle.THIN);

		writeTitle(tableName, StudentEntity.class);
		writeHeaderLine();

		for (int i = 0; i < list.size(); i++) {
			System.out.println("create data row = " + rowIndex);
			XSSFRow row = sheet.createRow(rowIndex);
			int columnIndex = 0;
			createCell(row, columnIndex++, i + 1, cellStyle);
			createCell(row, columnIndex++, list.get(i).getId(), cellStyle);
			createCell(row, columnIndex++, list.get(i).getName(), cellStyle);
			createCell(row, columnIndex++, String.valueOf(list.get(i).getGender() ? "Nam" : "Nữ"), cellStyle);
			createCell(row, columnIndex++, sdf.format(list.get(i).getBirthdate()).toString(), cellStyle);
			createCell(row, columnIndex++, list.get(i).getPhone(), cellStyle);
			createCell(row, columnIndex++, list.get(i).getAge(), cellStyle);

			if ((i + 1) % limit == 0) {
				sheet.setRowBreak(rowIndex++);
				writeTitle(tableName, StudentEntity.class);
				writeHeaderLine();
			} else {
				rowIndex++;
			}
		}
	}
}
