import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * This program illustrates how to update an existing Microsoft Excel document.
 * Append new rows to an existing sheet.
 * 
 * @author www.codejava.net
 *
 */
public class ExcelFileUpdateExample1 {


	public static void main(String[] args) {
		String excelFilePath = "Inventario.xlsx";
		File archivoExcel = new File(excelFilePath);
		//Creación y formato del nuevo archivo
		if (!archivoExcel.exists()) {
			try {
				System.out.println("Creación de nuevo archivo ya que no existe\n");
				archivoExcel.createNewFile();
				XSSFWorkbook workbook = new XSSFWorkbook();
				workbook.createSheet();
				
				Sheet sheet = workbook.getSheetAt(0);

				// Create file system using specific name
				FileOutputStream out = new FileOutputStream(archivoExcel);

				// Formatenado celdas
				FileInputStream inputStream = new FileInputStream(archivoExcel);

				Object[][] bookData = { { "No", "BookTitle", "Author", "Price" }, };

				for (Object[] aBook : bookData) {
					Row row = sheet.createRow(0);

					int columnCount = -1;
					Cell cell;
					for (Object field : aBook) {
						cell = row.createCell(++columnCount);
						if (field instanceof String) {
							cell.setCellValue((String) field);
						} else if (field instanceof Integer) {
							cell.setCellValue((Integer) field);
						}
					}

				}

				inputStream.close();

				FileOutputStream outputStream = new FileOutputStream(excelFilePath);

				workbook.write(out);
				out.close();
				System.out.println("Inventario.xlsx creado correctamente");
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		}
		try {
			FileInputStream inputStream = new FileInputStream(archivoExcel);
			Workbook workbook = WorkbookFactory.create(inputStream);

			Sheet sheet = workbook.getSheetAt(0);

			Object[][] bookData = { { "El que se duerme pierde", "Tom Peter", 16 },
					{ "Sin lugar a duda", "Ana Gutierrez", 26 }, { "El arte de dormir", "Nico", 32 },
					{ "Buscando a Nemo", "Humble Po", 41 }, };

			int rowCount = sheet.getLastRowNum();

			for (Object[] aBook : bookData) {
				Row row = sheet.createRow(++rowCount);

				int columnCount = 0;

				Cell cell = row.createCell(columnCount);
				cell.setCellValue(rowCount);

				for (Object field : aBook) {
					cell = row.createCell(++columnCount);
					if (field instanceof String) {
						cell.setCellValue((String) field);
					} else if (field instanceof Integer) {
						cell.setCellValue((Integer) field);
					}
				}

			}

			inputStream.close();

			FileOutputStream outputStream = new FileOutputStream(excelFilePath);
			workbook.write(outputStream);
			sheet = workbook.getSheetAt(0);
			FormulaEvaluator formulaEvaluator= workbook.getCreationHelper().createFormulaEvaluator(); 
			for (Row row : sheet) // iteration over row using for each loop
			{
				for (Cell cell : row) // iteration over cell using for each loop
				{
					switch (formulaEvaluator.evaluateInCell(cell).getCellType()) {
					case Cell.CELL_TYPE_NUMERIC: // field that represents numeric cell type
						// getting the value of the cell as a number
						System.out.print(cell.getNumericCellValue() + "\t\t");
						break;
					case Cell.CELL_TYPE_STRING: // field that represents string cell type
						// getting the value of the cell as a string
						System.out.print(cell.getStringCellValue() + "\t\t");
						break;
					}
				}
				System.out.println();
			} 
			workbook.close();
			outputStream.close();
			
		} catch (IOException | EncryptedDocumentException | InvalidFormatException ex) {
			ex.printStackTrace();
		}
	}

}
