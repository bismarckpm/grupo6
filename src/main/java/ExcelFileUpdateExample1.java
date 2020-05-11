import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Scanner;

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
	
	/*
	 * METODO: creacionArchivoInexistente(File archivoExcel)
	 * RETORNA: void
	 * DESCRIPCION: El metodo es llamado en caso de que el archivo "Inventario.xlsx" no esta
	 * 		creado en el directorio. Crea el archivo Excele "Inventario.xlsx" y agrega la 
	 * 		1era hoja con encabezado correspondiente.
	 */
	private static void creacionArchivoInexistente(File archivoExcel) {
		try {
			System.out.println("Creando el archivo Excel...\n");
			archivoExcel.createNewFile();
			XSSFWorkbook workbook = new XSSFWorkbook(); // Se crea el libro de Excel
			workbook.createSheet("Java Books 1"); // Se crea la 1era hoja
			
			Sheet sheet = workbook.getSheetAt(0); // Obtengo el objeto de la 1era hoja

			// Create file system using specific name
			FileOutputStream out = new FileOutputStream(archivoExcel);

			// Formatenado celdas
			FileInputStream inputStream = new FileInputStream(archivoExcel);
			
			addHeading(sheet); // Agrego el encabezado de la tabla

			inputStream.close();

			workbook.write(out);
			workbook.close(); // JAR - Cierro el workbook
			out.close();
			System.out.println("Inventario.xlsx creado correctamente");
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
	
	
	/*
	 * METODO: getSheetSlot(Workbook workbook)
	 * RETORNA: INT - Indice de la hoja disponible con menos de 30 registros
	 * DESCRIPCION: Retorna el indice de la hoja correspondiente del workbook que aun
	 * 		tienes menos de 30 registros.
	 */
	private static int getSheetSlot(Workbook workbook) {
		
		int availableSheet = 0; // Inicializando la variable
		
		// Verifico cada hoja en el work 
		for (Sheet sheet : workbook) {
			int rowCount = sheet.getLastRowNum();
			if (rowCount < 30) {
				availableSheet = workbook.getSheetIndex(sheet);
			}
		}
				
		return availableSheet;
	}
	
	
	/*
	 * METODO: addHeading(Sheet sheet)
	 * RETORNA: void
	 * DESCRIPCION: Agrega el encabezado de la tabla en la hoja pasada por parametro.
	 */
	private static void addHeading(Sheet sheet) {
		
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
		
	}
	
	
	/*
	 * METODO: addRecord(Workbook workbook, Sheet sheet, int rowCount, Object[][] bookData)
	 * RETORNA: void
	 * DESCRIPCION: Agrega los registros pasados por parametro en la hoja correspondiente del
	 * 		workbook.
	 * 		-- Si se execeden los 30 registros en la hoja en la que se esta agregando, se crea
	 * 		   una hoja adicional para agregar el resto.
	 */
 	private static void addRecords(Workbook workbook, Sheet sheet, int rowCount, Object[][] bookData, String excelFilePath) {
				
		int columnCount = 0;
		
		for (Object[] aBook : bookData) {
			
			if (rowCount + 1 <= 30) {		
				columnCount = 0;
				Row row = sheet.createRow(++rowCount);
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
			} else {
				int sheetsQuantity = workbook.getNumberOfSheets();				
				sheet = workbook.createSheet("Java Books " + ++sheetsQuantity);
				addHeading(sheet);
				
				rowCount = sheet.getFirstRowNum();
				
				columnCount = 0;
				Row row = sheet.createRow(++rowCount);
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

		}
		FileOutputStream outputStream = null;;
		try {
			outputStream = new FileOutputStream(excelFilePath);
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		try {
			workbook.write(outputStream);
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		try {
			outputStream.close();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

 	
 	/*
	 * METODO: viewResults(Workbook workbook, FormulaEvaluator formulaEvaluator)
	 * RETORNA: void
	 * DESCRIPCION: Muestra los resultados de la operacion por consola.
	 */
 	private static void viewResults(Workbook workbook, FormulaEvaluator formulaEvaluator, String excelFilePath) {

		int sheetCounter = 0;
		for (Sheet hoja : workbook) {

			System.out.format("%n" + hoja.getSheetName() + " - Hoja (%d/%d)%n", ++sheetCounter, workbook.getNumberOfSheets());
			
			for (Row row : hoja) // iteration over row using for each loop
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
			try {
				workbook.close();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		}
 	}
 	
	private static void menu(File archivoExcel) {
		while (1 == 1) {
			FileInputStream inputStream = null;
			try {
				inputStream = new FileInputStream(archivoExcel);
			} catch (FileNotFoundException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
			Workbook workbook = null;
			try {
				workbook = WorkbookFactory.create(inputStream);
			} catch (EncryptedDocumentException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			} catch (InvalidFormatException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
			int availableSheet = getSheetSlot(workbook); // Obtengo el indice de la hoja que aun tiene menos de 30
															// registros
			Sheet sheet = workbook.getSheetAt(availableSheet); // Luego de obtener el indice, obtengo el objeto de la
																// hoja
			System.out.println("____________MENU____________\n");
			System.out.println("1. Visualizar Archivo\n");
			System.out.println("2. Agregar Registro\n");
			System.out.println("3. Modificar Registro\n");
			System.out.println("4. Salir");
			int rowCount = sheet.getLastRowNum();
			Scanner input = new Scanner(System.in);
			int selection = input.nextInt();
			String libro;
			String autor;
			int precio;
			FormulaEvaluator formulaEvaluator;
			switch (selection) {
			case 1:
				if (!archivoExcel.exists()) {
					System.out.println("No existe");
				}
				else {
					sheet = workbook.getSheetAt(0); // Obtengo 1era hoja para mostrar resultados desde el inicio
					formulaEvaluator = workbook.getCreationHelper().createFormulaEvaluator(); 
					viewResults(workbook, formulaEvaluator, archivoExcel.getPath());
				}
				break;
			case 2:
				if (!archivoExcel.exists()) {
					creacionArchivoInexistente(archivoExcel);
				}
				System.out.println("\n\n\n\n\n\n\n");
				System.out.println("Inserte título del libro: ");
				libro = input.nextLine();
				System.out.println("\n");
				System.out.println("Inserte autor del libro: ");
				autor = input.nextLine();
				System.out.println("\n");
				System.out.println("Inserte precio del libro: ");
				precio = input.nextInt();
				Object dataBook2[][] = {{ libro, autor, precio},};
				addRecords(workbook, sheet, rowCount, dataBook2, archivoExcel.getPath());
				sheet = workbook.getSheetAt(0); // Obtengo 1era hoja para mostrar resultados desde el inicio
				formulaEvaluator = workbook.getCreationHelper().createFormulaEvaluator(); 
				viewResults(workbook, formulaEvaluator, archivoExcel.getPath());
				break;
			case 3:
				System.out.println("\n\n\n\n\n\n\n");
				System.out.println("Inserte título del libro: ");
				int rowModificationNumb = input.nextInt();
				System.out.println("Inserte título del libro: ");
				libro = input.nextLine();
				System.out.println("\n");
				System.out.println("Inserte autor del libro: ");
				autor = input.nextLine();
				System.out.println("\n");
				System.out.println("Inserte precio del libro: ");
				precio = input.nextInt();
				Object[][] dataBook= {{libro, autor, precio}};
				//AGREGA ACÁ MARCO EL MÉTODO DE MODIFICAR;
				//----
				//ARRIBA
				sheet = workbook.getSheetAt(0); // Obtengo 1era hoja para mostrar resultados desde el inicio
				formulaEvaluator = workbook.getCreationHelper().createFormulaEvaluator(); 
				viewResults(workbook, formulaEvaluator, archivoExcel.getPath());
				break;
			}
			if (selection == 4)
				break;
		}
	}
 	
 	private static void codigo_principal() {
 		String excelFilePath = "Inventario.xlsx";
		File archivoExcel = new File(excelFilePath);
		// Creacion y formato del nuevo archivo en caso de que no exista en el directorio
		if (!archivoExcel.exists()) {
			creacionArchivoInexistente(archivoExcel);
		}
		try {
			// Escritura en el archivo Excel
			FileInputStream inputStream = new FileInputStream(archivoExcel);
			Workbook workbook = WorkbookFactory.create(inputStream);
			
			int availableSheet = getSheetSlot(workbook); // Obtengo el indice de la hoja que aun tiene menos de 30 registros
			Sheet sheet = workbook.getSheetAt(availableSheet); // Luego de obtener el indice, obtengo el objeto de la hoja

			// Hard coded data
			Object[][] bookData = { 
					{ "El que se duerme pierde", "Tom Peter", 16 },
					{ "Sin lugar a duda", "Ana Gutierrez", 26 }, 
					{ "El arte de dormir", "Nico", 32 },
					{ "Buscando a Nemo", "Humble Po", 41 }
			};

			int rowCount = sheet.getLastRowNum(); // Obtengo el No del ultimo registro de la hoja correspondiente

			addRecords(workbook ,sheet, rowCount, bookData, archivoExcel.getPath()); // Agrego los registros

			// Escritura en la consola
			FileOutputStream outputStream = new FileOutputStream(excelFilePath);
			workbook.write(outputStream);
			sheet = workbook.getSheetAt(0); // Obtengo 1era hoja para mostrar resultados desde el inicio
			FormulaEvaluator formulaEvaluator = workbook.getCreationHelper().createFormulaEvaluator(); 
			
			viewResults(workbook, formulaEvaluator, archivoExcel.getPath()); // Mostrar resultados por consola 
			
		} catch (IOException | EncryptedDocumentException | InvalidFormatException ex) {
			ex.printStackTrace();
		}
	}
 	
 	
	public static void main(String[] args) {
		String excelFilePath = "Inventario.xlsx";
		File archivoExcel = new File(excelFilePath);
		menu(archivoExcel);
	}

}
