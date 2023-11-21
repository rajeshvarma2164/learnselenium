
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class pdfexce {

	private static final String FILE_NAME = "C:\\Users\\rajes\\learn\\test.xlsx";

	public static synchronized void writeToExcel(int row, int column, String value) {
		try (FileInputStream inputStream = new FileInputStream(FILE_NAME);
				Workbook workbook = WorkbookFactory.create(inputStream)) {
			
			short greenColor = IndexedColors.GREEN.getIndex();
	        short redColor = IndexedColors.RED.getIndex();

	     // Create CellStyle for highlighting
            CellStyle greenHighlightStyle = createHighlightStyle(workbook, greenColor);
            CellStyle redHighlightStyle = createHighlightStyle(workbook, redColor);	        
			

			Sheet sheet = workbook.getSheet("Sheet1");
			if (sheet == null) {
				sheet = workbook.createSheet("Sheet1");
			}

			Row excelRow = sheet.getRow(row);
			if (excelRow == null) {
				excelRow = sheet.createRow(row);
			}

			Cell cell = excelRow.createCell(column);
			cell.setCellValue(value);
			cell.setCellStyle(greenHighlightStyle); 


			try (FileOutputStream outputStream = new FileOutputStream(FILE_NAME)) {
				workbook.write(outputStream);
			}

			System.out.println("Excel file written successfully!");

		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	// based on input
	private static void excelcolor() {
	        // Sample data
	        String[] data = {"A", "B", "C", "D"};
	        String[] inputCodes = {"A", "C"};
	        
	        try (Workbook workbook = new XSSFWorkbook()) {
	            Sheet sheet = workbook.createSheet("Sheet1");

	        // Highlight colors
	        short greenColor = IndexedColors.GREEN.getIndex();
	        short redColor = IndexedColors.RED.getIndex();

	     // Create CellStyle for highlighting
            CellStyle greenHighlightStyle = createHighlightStyle(workbook, greenColor);
            CellStyle redHighlightStyle = createHighlightStyle(workbook, redColor);	      

	            // Write data to Excel and highlight cells based on input codes
	            for (int i = 0; i < data.length; i++) {
	                Row row = sheet.createRow(i);
	                Cell cell = row.createCell(0);
	                cell.setCellValue(data[i]);

	                // Highlight cell based on input code
	                if (contains(inputCodes, data[i])) {
	                    cell.setCellStyle(greenHighlightStyle); // Change this line to redHighlightStyle for red color
	                }
	            }

	            // Write the Excel file
	            try (FileOutputStream fileOut = new FileOutputStream(FILE_NAME)) {
	                workbook.write(fileOut);
	            }

	            System.out.println("Excel file created successfully!");
	        } catch (IOException e) {
	            e.printStackTrace();
	        }
	    }

	    private static boolean contains(String[] array, String target) {
	        for (String s : array) {
	            if (s.equals(target)) {
	                return true;
	            }
	        }
	        return false;
	    }
	
	    private static CellStyle createHighlightStyle(Workbook workbook, short fillColor) {
	        CellStyle highlightStyle = workbook.createCellStyle();
	        highlightStyle.setFillForegroundColor(fillColor);
	        highlightStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
	        return highlightStyle;
	    }
	    
	    // row column 
	    
	    public static synchronized void writeToExcel(String rowName, String columnName, String value) {
	        try (FileInputStream inputStream = new FileInputStream(FILE_NAME);
	             Workbook workbook = WorkbookFactory.create(inputStream)) {

	            Sheet sheet = workbook.getSheet("Sheet1");
	            if (sheet == null) {
	                sheet = workbook.createSheet("Sheet1");
	            }

	            int rowIndex = getRowIndex(sheet, rowName);
	            Row excelRow = sheet.getRow(rowIndex);
	            if (excelRow == null) {
	                excelRow = sheet.createRow(rowIndex);
	            }

	            int columnIndex = getColumnIndex(sheet, columnName);
	            Cell cell = excelRow.createCell(columnIndex);
	            cell.setCellValue(value);

	            try (FileOutputStream outputStream = new FileOutputStream(FILE_NAME)) {
	                workbook.write(outputStream);
	            }

	            System.out.println("Excel file written successfully!");

	        } catch (IOException e) {
	            e.printStackTrace();
	        }
	    }

	    private static int getRowIndex(Sheet sheet, String rowName) {
	        for (Row row : sheet) {
	            Cell cell = row.getCell(2); // Assuming the first column contains row names
	            if (cell != null && cell.getStringCellValue().equals(rowName)) {
	                return row.getRowNum();
	            }
	        }
	        // If the row with the specified name is not found, create a new row
	        return sheet.getPhysicalNumberOfRows();
	    }

	    private static int getColumnIndex(Sheet sheet, String columnName) {
	        Row headerRow = sheet.getRow(0);
	        for (Cell cell : headerRow) {
	            if (cell.getStringCellValue().equals(columnName)) {
	                return cell.getColumnIndex();
	            }
	        }
	        // If the column with the specified name is not found, create a new column
	        return headerRow.getPhysicalNumberOfCells();
	    }
	    
	    
	public static void main(String[] args) {
		writeToExcel(1, 2, "test, Excel!");
		writeToExcel("test", "hello", "looo");
					
	}

}
