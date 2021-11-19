package com.gmoz.utils;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Component;

import com.gmoz.entity.ClassEntity;
import com.gmoz.entity.StudentEntity;

/**
 * Mô tả: Class hỗ trợ thêm dữ liệu vào file Excel có sẵn
 * 
 * @author DucBV
 * @version 1.0
 * @since 18/11/2021
 * 
 */
@Component
public class JExcelHelper {

	private File file;
	private FileInputStream inputStream;
	private XSSFWorkbook workbook;
	private XSSFSheet sheet;
	private int rowIndex;
	private List<Row> headers;

	/**
	 * Mô tả: Cập nhật Cell từ excel có sẵn
	 * 
	 * @param path  Đường dẫn file
	 * @param limit Giới hạn số lượng hàng trên một bảng
	 * @throws IOException
	 */
	public void updateExcel(String path, int limit, List<ClassEntity> list) throws IOException {
		file = new File(path);

		// Đọc một file XSL.
		inputStream = new FileInputStream(file);

		// Đối tượng workbook cho file XSL.
		workbook = new XSSFWorkbook(inputStream);

		// Lấy ra sheet đầu tiên từ workbook
		sheet = workbook.getSheetAt(0);

		// Init rowIndex = last row of sheet
		rowIndex = sheet.getLastRowNum() + 1;

		// Get Rows of header
		headers = getHeaderRows();

//		copyRow(workbook, sheet, 2, rowIndex++);
//		writeHeaderLine();
		writeDataLine(list, limit);

		// Auto size column
		for (int i = 0; i < sheet.getRow(0).getPhysicalNumberOfCells(); i++) {
			sheet.autoSizeColumn(i);
//			sheet.setColumnWidth(i, 9000);
//			sheet.autoSizeColumn(i);
//			System.out.println(sheet.getColumnWidth(i));
		}

		inputStream.close();

		// Ghi file
		FileOutputStream out = new FileOutputStream(file);
		workbook.write(out);
		workbook.close();
		out.close();
	}

	/**
	 * Mô tả: Thu thập danh sách các hàng
	 * 
	 * @return trả về giá trị các hàng
	 */
	private List<Row> getHeaderRows() {
		List<Row> rows = new ArrayList<>();
		Sheet sheet = workbook.getSheetAt(0);
		for (Row row : sheet) {
			rows.add(row);
		}
		rows.remove(rows.size() - 1);
		return rows;
	}

	/**
	 * Mô tả: Thu thập danh sách kiểu dáng của một hàng
	 * 
	 * @param row dòng/hàng cần thu thập
	 * @return trả về danh sách kiểu dáng các ô trong một hàng
	 */
	private List<CellStyle> getCellStyles(Row row) {
		List<CellStyle> cellStyles = new ArrayList<>();
		for (Cell cell : row) {
			cellStyles.add(cell.getCellStyle());
		}
		return cellStyles;
	}

	/**
	 * Mô tả: Create a cell
	 * 
	 * @param row         Hàng hiện tại
	 * @param columnIndex Chỉ số cột
	 * @param value       Giá trị đối tượng
	 * @param style       Định dạng cell
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

	/**
	 * Mô tả: Nhập toàn bộ đầu template của excel
	 */
	private void writeHeaderLine() {
		System.out.println("VE TIEU DE");
		for (int i = 0; i < headers.size(); i++) {
			copyRow(workbook, sheet, i, rowIndex++);
		}
	}

	private static final SimpleDateFormat sdf = new SimpleDateFormat("dd/MM/yyyy");

	/**
	 * Mô tả: Nhập dữ liệu từ danh sách với số lượng tự động ngắt trang
	 * 
	 * @param tableName Tên tiêu đề của bảng
	 * @param list      Danh sách dữ liệu bảng
	 * @param limit     Giới hạn số lượng dòng trên một bảng
	 */
	private void writeDataLine(List<ClassEntity> list, int limit) {
		// Nạp dữ liệu từ dòng cuối có sẵn
		rowIndex -= 1;

		// Get CellStyles of row body
		List<CellStyle> cellStyles = getCellStyles(sheet.getRow(sheet.getLastRowNum()));

		// Số thứ tự
		int count = 1;
		int countBreak = 0;
		for (ClassEntity clas : list) {

			// Get cell style of class name property
			CellStyle cellStyle = cellStyles.get(1);
			cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);

			// Merge Cell class name
			int countStudent = clas.getStudents().size();
			// Số hàng còn lại của bảng bị tách
			int restCount = 0;

			// Đếm số hàng sẽ tạo trong list
			countBreak += countStudent;
			if (countBreak >= limit) {
				countBreak -= limit;
				System.out.println("\nTach bang...");
				System.out.println("from = " + rowIndex);
				System.out.println("limit = " + limit);
				System.out.println("count = " + count);
				System.out.println("to = " + (rowIndex + limit - count % limit));

				restCount = countStudent - limit + count % limit - 1;

				sheet.addMergedRegion(new CellRangeAddress(rowIndex, rowIndex + limit - count % limit, 1, 1));
				createCell(getCurrentRow(rowIndex), 1, clas.getName(), cellStyle);
			} else {
				System.out.println("\nNguyen bang...");
				System.out.println("from = " + rowIndex);
				System.out.println("to = " + (rowIndex + countStudent - 1));
				sheet.addMergedRegion(new CellRangeAddress(rowIndex, rowIndex + countStudent - 1, 1, 1));
				createCell(getCurrentRow(rowIndex), 1, clas.getName(), cellStyle);
			}

			// Create other cells
			for (StudentEntity student : clas.getStudents()) {
				int columnIndex = 0;
				int cellStyleIndex = 0;

				createCell(getCurrentRow(rowIndex), columnIndex++, count, cellStyles.get(cellStyleIndex++));
				// Skip column class name
				columnIndex++;
				cellStyleIndex++;
				createCell(getCurrentRow(rowIndex), columnIndex++, student.getId(), cellStyles.get(cellStyleIndex++));
				createCell(getCurrentRow(rowIndex), columnIndex++, student.getName(), cellStyles.get(cellStyleIndex++));
				createCell(getCurrentRow(rowIndex), columnIndex++, String.valueOf(student.getGender() ? "Nam" : "Nữ"),
						cellStyles.get(cellStyleIndex++));
				createCell(getCurrentRow(rowIndex), columnIndex++, sdf.format(student.getBirthdate()).toString(),
						cellStyles.get(cellStyleIndex++));
				createCell(getCurrentRow(rowIndex), columnIndex++, student.getPhone(),
						cellStyles.get(cellStyleIndex++));
				createCell(getCurrentRow(rowIndex), columnIndex++, student.getAge(), cellStyles.get(cellStyleIndex));

				if (count % limit == 0 && restCount != 0) {
					System.out.println("New table-------");
					sheet.setRowBreak(rowIndex++);
					writeHeaderLine();

					System.out.println("from = " + rowIndex);
					System.out.println("to = " + (rowIndex + restCount - 1));
					sheet.addMergedRegion(new CellRangeAddress(rowIndex, rowIndex + restCount - 1, 1, 1));

					createCell(sheet.createRow(rowIndex), 1, clas.getName(), cellStyle);
				} else {
					rowIndex++;
				}
				System.out.println("count = " + count);
				count++;
			}
		}
	}

	/**
	 * Get current row object or create if its empty
	 * 
	 * @param rowIndex
	 * @return
	 */
	private XSSFRow getCurrentRow(int rowIndex) {
		if (sheet.getRow(rowIndex) == null) {
			return sheet.createRow(rowIndex);
		}
		return sheet.getRow(rowIndex);
	}

	/**
	 * Mô tả: Sao chép hàng tới hàng
	 * 
	 * @param workbook  File excel
	 * @param worksheet Sheet excel
	 * @param from      chỉ số dòng/hàng cần sao chép
	 * @param to        chỉ số dòng/hàng sao chép tới
	 */
	public void copyRow(Workbook workbook, Sheet worksheet, int from, int to) {
		Row sourceRow = worksheet.getRow(from);
		Row newRow = worksheet.getRow(to);

		if (sourceRow == null) {
			newRow = worksheet.createRow(to);
			return;
		}

		if (alreadyExists(newRow))
			worksheet.shiftRows(to, worksheet.getLastRowNum(), 1);
		else
			newRow = worksheet.createRow(to);

		for (int i = 0; i < sourceRow.getLastCellNum(); i++) {
			Cell oldCell = sourceRow.getCell(i);
			Cell newCell = newRow.createCell(i);
			if (oldCell != null) {
				copyCellStyle(workbook, oldCell, newCell);
				copyCellComment(oldCell, newCell);
				copyCellHyperlink(oldCell, newCell);
				copyCellDataTypeAndValue(oldCell, newCell);
			}
		}

		copyAnyMergedRegions(worksheet, sourceRow, newRow);
	}

	/**
	 * Mô tả: Sao chép kiểu, định dạng một ô tới ô khác
	 * 
	 * @param workbook File excel
	 * @param oldCell  ô excel cũ
	 * @param newCell  ô excel mới
	 */
	private void copyCellStyle(Workbook workbook, Cell oldCell, Cell newCell) {
		newCell.setCellStyle(oldCell.getCellStyle());
	}

	/**
	 * Mô tả: Sao chép comment tại ô excel tới ô khác
	 * 
	 * @param oldCell ô cũ
	 * @param newCell ô mới
	 */
	private void copyCellComment(Cell oldCell, Cell newCell) {
		if (newCell.getCellComment() != null)
			newCell.setCellComment(oldCell.getCellComment());
	}

	/**
	 * Mô tả: Sao chép đường dẫn siêu liên kết của một ô excel tới ô khác
	 * 
	 * @param oldCell ô cũ
	 * @param newCell ô mới
	 */
	private void copyCellHyperlink(Cell oldCell, Cell newCell) {
		if (oldCell.getHyperlink() != null)
			newCell.setHyperlink(oldCell.getHyperlink());
	}

	/**
	 * Mô tả: Sao chép giá trị và kiểu định dạng dữ liệu của một ô excel tới ô khác
	 * 
	 * @param oldCell ô cũ
	 * @param newCell ô mới
	 */
	private void copyCellDataTypeAndValue(Cell oldCell, Cell newCell) {
		setCellDataType(oldCell, newCell);
		setCellDataValue(oldCell, newCell);
	}

	/**
	 * Mô tả: Sao chép định dạng kiểu dữ liệu của một ô excel tới ô khác
	 * 
	 * @param oldCell
	 * @param newCell
	 */
	@SuppressWarnings("deprecation")
	private static void setCellDataType(Cell oldCell, Cell newCell) {
		newCell.setCellType(oldCell.getCellType());
	}

	/**
	 * Mô tả: Sao chép giá trị của một ô excel tới ô khác
	 * 
	 * @param oldCell
	 * @param newCell
	 */
	private void setCellDataValue(Cell oldCell, Cell newCell) {
		switch (oldCell.getCellType()) {
		case BLANK:
			newCell.setCellValue(oldCell.getStringCellValue());
			break;
		case BOOLEAN:
			newCell.setCellValue(oldCell.getBooleanCellValue());
			break;
		case ERROR:
			newCell.setCellErrorValue(oldCell.getErrorCellValue());
			break;
		case FORMULA:
			newCell.setCellFormula(oldCell.getCellFormula());
			break;
		case NUMERIC:
			newCell.setCellValue(oldCell.getNumericCellValue());
			break;
		case STRING:
			newCell.setCellValue(oldCell.getRichStringCellValue());
			break;
		case _NONE:
			break;
		default:
			break;
		}
	}

	/**
	 * Mô tả: Kiểm trả dòng/hàng đã tồn tại hay chưa?
	 * 
	 * @param newRow chỉ số dòng/hàng cần kiểm tra
	 * @return trả về true nếu đã tồn tại và ngược lại
	 */
	private boolean alreadyExists(Row newRow) {
		return newRow != null;
	}

	/**
	 * Mô tả: Sao chép toàn bộ hàng/dòng có merge
	 * 
	 * @param worksheet File excel
	 * @param sourceRow nguồn của hàng/dòng
	 * @param newRow    nguồn của hàng/dòng
	 */
	private void copyAnyMergedRegions(Sheet worksheet, Row sourceRow, Row newRow) {
		for (int i = 0; i < worksheet.getNumMergedRegions(); i++)
			copyMergeRegion(worksheet, sourceRow, newRow, worksheet.getMergedRegion(i));
	}

	/**
	 * Mô tả: Sao chép toàn bộ hàng/dòng theo khoảng có merge
	 * 
	 * @param worksheet    File excel
	 * @param sourceRow    nguồn của hàng/dòng
	 * @param newRow       đích của hàng/dòng
	 * @param mergedRegion số khoảng
	 */
	private void copyMergeRegion(Sheet worksheet, Row sourceRow, Row newRow, CellRangeAddress mergedRegion) {
		CellRangeAddress range = mergedRegion;
		if (range.getFirstRow() == sourceRow.getRowNum()) {
			int lastRow = newRow.getRowNum() + (range.getLastRow() - range.getFirstRow());
			CellRangeAddress newCellRangeAddress = new CellRangeAddress(newRow.getRowNum(), lastRow,
					range.getFirstColumn(), range.getLastColumn());
			worksheet.addMergedRegion(newCellRangeAddress);
		}
	}
}