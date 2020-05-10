import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.net.CacheResponse;
import java.util.Scanner;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.net.CacheResponse;
import java.util.Scanner;

import javax.swing.JOptionPane;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.DataFormatter;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.net.CacheResponse;
import java.util.Scanner;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

/**
 * This program illustrates how to update an existing Microsoft Excel document.
 * Append new rows to an existing sheet.
 * 
 * @author www.codejava.net
 *
 */
public class ExcelFileUpdateExample1 {

	public static void ModuloB (){

				String excelFilePath = "Inventario.xlsx";

				try {
					FileInputStream inputStream = new FileInputStream(new File(excelFilePath));
					Workbook workbook = WorkbookFactory.create(inputStream);
					Sheet sheet = workbook.getSheetAt(0);

					int bookCount = 1;
					int rowCount = sheet.getLastRowNum();
					int cellCount = 0;
					int i = 0;
					int j =0;
					int k = 0;
					int m = 0;

					System.out.println();
					System.out.println();
					System.out.println();
					System.out.println();
					System.out.println ("...MODULO B...");
					System.out.println();
					System.out.println();
					System.out.println();

					Object[][] bookData = {
							{"El que se duerme pierde", "Tom Peter", 16},
							{"Sin lugar a duda", "Ana Gutierrez", 26},
							{"El arte de dormir", "Nico", 32},
							{"Buscando a Nemo", "Humble Po", 41},
							{"El que se duerme pierde", "Tom Peter", 16},
							{"Sin lugar a duda", "Ana Gutierrez", 26},
							{"El arte de dormir", "Nico", 32},
							{"Buscando a Nemo", "Humble Po", 41},
							{"El que se duerme pierde", "Tom Peter", 16},
							{"Sin lugar a duda", "Ana Gutierrez", 26},
							{"El arte de dormir", "Nico", 32},
							{"Buscando a Nemo", "Humble Po", 41},
							{"El que se duerme pierde", "Tom Peter", 16},
							{"Sin lugar a duda", "Ana Gutierrez", 26},
							{"El arte de dormir", "Nico", 32},
							{"Buscando a Nemo", "Humble Po", 41},
							{"El que se duerme pierde", "Tom Peter", 16},
							{"Sin lugar a duda", "Ana Gutierrez", 26},
							{"El arte de dormir", "Nico", 32},
							{"Buscando a Nemo", "Humble Po", 41},
							{"El que se duerme pierde", "Tom Peter", 16},
							{"Sin lugar a duda", "Ana Gutierrez", 26},
							{"El arte de dormir", "Nico", 32},
							{"Buscando a Nemo", "Humble Po", 41},
							{"El que se duerme pierde", "Tom Peter", 16},
							{"Sin lugar a duda", "Ana Gutierrez", 26},
							{"El arte de dormir", "Nico", 32},
							{"Buscando a Nemo", "Humble Po", 41},
							{"El que se duerme pierde", "Tom Peter", 16},
							{"Sin lugar a duda", "Ana Gutierrez", 26},
							{"El arte de dormir", "Nico", 32},
							{"Buscando a Nemo", "Humble Po", 41},
							{"El que se duerme pierde", "Tom Peter", 16},
							{"Sin lugar a duda", "Ana Gutierrez", 26},
							{"El arte de dormir", "Nico", 32},
							{"Buscando a Nemo", "Humble Po", 41},
							{"El que se duerme pierde", "Tom Peter", 16},
							{"Sin lugar a duda", "Ana Gutierrez", 26},
							{"El arte de dormir", "Nico", 32},
							{"Buscando a Nemo", "Humble Po", 41},
							{"El que se duerme pierde", "Tom Peter", 16},
							{"Sin lugar a duda", "Ana Gutierrez", 26},
							{"El arte de dormir", "Nico", 32},
							{"Buscando a Nemo", "Humble Po", 41},
							{"El que se duerme pierde", "Tom Peter", 16},
							{"Sin lugar a duda", "Ana Gutierrez", 26},
							{"El arte de dormir", "Nico", 32},
							{"Buscando a Nemo", "Humble Po", 41},
							{"El que se duerme pierde", "Tom Peter", 16},
							{"Sin lugar a duda", "Ana Gutierrez", 26},
							{"El arte de dormir", "Nico", 32},
							{"Buscando a Nemo", "Humble Po", 41},
							{"El que se duerme pierde", "Tom Peter", 16},
							{"Sin lugar a duda", "Ana Gutierrez", 26},
							{"El arte de dormir", "Nico", 32},
							{"Buscando a Nemo", "Humble Po", 41},
							{"El que se duerme pierde", "Tom Peter", 16},
							{"Sin lugar a duda", "Ana Gutierrez", 26},
							{"El arte de dormir", "Nico", 32},
							{"Buscando a Nemo", "Humble Po", 41}
					};



					System.out.println ("Se Procedera a Insertar 60 registros para probar el modulo.");
					//comienza if



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

						if (rowCount == 30){
							String bookBombre = "Java Books" + ++bookCount;
							sheet = workbook.createSheet(bookBombre);
							rowCount = 0;
						}
					}

					//termina if

					int sheetCount = workbook.getNumberOfSheets();
					Scanner entradaEscaner5 = new Scanner (System.in);
					String x = entradaEscaner5.nextLine(); //wait enter
					System.out.println ("Insercion exitosa.");

					for ( j = 0 ; j < sheetCount; ++j){
						x = entradaEscaner5.nextLine();
						sheet = workbook.getSheetAt(j);
						System.out.println();

						System.out.println ("Hoja: " + sheet.getSheetName());
						x = entradaEscaner5.nextLine();
						rowCount = sheet.getLastRowNum();

						for ( k = 1 ; k <= rowCount; ++k){

							Row row = sheet.getRow(k);
							System.out.println(' ');
							System.out.println();
							System.out.println ("Fila: " + k);
							System.out.print(' ');
							cellCount = row.getLastCellNum();

							for ( m = 1 ; m < cellCount; ++m){
								System.out.print(' ');
								Cell nro = row.getCell(m);

								if (m == 1 || m == 2){
									System.out.print (nro.getStringCellValue());
									System.out.print(' ');
								}

								if(m == 3){
									System.out.print (nro.getNumericCellValue());
									System.out.print(' ');
								}

							}

						}

					};

					inputStream.close();

					FileOutputStream outputStream = new FileOutputStream(excelFilePath);
					workbook.write(outputStream);
					workbook.close();
					outputStream.close();

				} catch (IOException | EncryptedDocumentException
						| InvalidFormatException ex) {
					ex.printStackTrace();
				}
				System.out.println();
				System.out.println();
				System.out.println();
				System.out.println();
				System.out.println ("...FIN MODULO B...");
				System.out.println();
				System.out.println();
				System.out.println();
			}


			public static int contarHojas(Object[][] bd){

				int cantHojas = 0;

				for (Object[] aBook : bd) {
					++cantHojas;
				}

				return cantHojas;
			}



	public static void ModuloC (){

				String excelFilePath = "Inventario.xlsx";

				try {
					FileInputStream inputStream = new FileInputStream(new File(excelFilePath));
					Workbook workbook = WorkbookFactory.create(inputStream);

					Sheet sheet = workbook.getSheetAt(0);

					int bookCount = 1;
					int cellCount = 0;
					int rowCount = sheet.getLastRowNum();
					int i = 0;
					int j =0;
					int k = 0;
					int m = 0;
					int entradaTeclado = 0;
					String entradaTeclado2 = "";
					String entradaTeclado3 = "";
					int entradaTeclado4 = 0;
					String nombreAutor = "";
					String Author = "Author";
					String Price = "Price";

					System.out.println();
					System.out.println();
					System.out.println();
					System.out.println ("...MODULO C...");
					///// CICLO PARA SUSTITUIR EN UNA CELDA ESPECIFICA
					System.out.println();
					System.out.println();
					System.out.println ("Ingrese el numero identificador del registro");
					Scanner entradaEscaner = new Scanner (System.in);
					entradaTeclado = entradaEscaner.nextInt ();

					for ( i = 1 ; i < rowCount; ++i){
						Row row = sheet.getRow(i);
						Cell nro = row.getCell(0);
						Double doubleComparar = nro.getNumericCellValue();
						int nroComparar = doubleComparar.intValue();

						if (entradaTeclado == nroComparar){
							System.out.println ("Ingrese el nombre del atributo a modficar('Price' o 'Author')");
							Scanner entradaEscaner2 = new Scanner (System.in);
							entradaTeclado2 = entradaEscaner2.nextLine();

							if (entradaTeclado2.equalsIgnoreCase(Author)){

								nro = row.getCell(2);

								System.out.println ("Ingrese el nuevo nombre del Autor a modficar");
								Scanner entradaEscaner3 = new Scanner (System.in);
								entradaTeclado3 = entradaEscaner3.nextLine();
								System.out.println ("AUTOR ACTUALIZADO");

								nro.setCellValue(entradaTeclado3);

							}
							else if (entradaTeclado2.equalsIgnoreCase(Price)){

								nro = row.getCell(3);

								System.out.println ("Ingrese el Precio nuevo");
								Scanner entradaEscaner4 = new Scanner (System.in);
								entradaTeclado4 = entradaEscaner4.nextInt();
								System.out.println ("PRECIO ACTUALIZADO");
								nro.setCellValue(entradaTeclado4);

							}
							else{
								System.out.println();
								System.out.println ("ERROR: Nombre de atributo incorrecto");
								System.out.println ("NO OCURRIO NINGUN CAMBIO");
								System.out.println();
							}


						}
						else if (entradaTeclado<1 || entradaTeclado>rowCount){
							if(i==1) {
								System.out.println();
								System.out.println("ERROR: El registro que ingreso no existe");
								System.out.println("NO OCURRIO NINGUN CAMBIO");
								System.out.println();
							}
						}
					};

					///CICLO PARA MOSTRAR POR PANTALLA

					System.out.println();
					System.out.println ("A continuacion se muestran todos los registros que estan dentro del archivo de Excel");
					Scanner entradaEscaner5 = new Scanner (System.in);
					String x = entradaEscaner5.nextLine(); //wait enter

					//entradaTeclado = entradaEscaner5.nextInt ();
					System.out.println();

					int sheetCount = workbook.getNumberOfSheets();

					for ( j = 0 ; j < sheetCount; ++j){

						sheet = workbook.getSheetAt(j);
						System.out.println();
						System.out.println ("Hoja: " + sheet.getSheetName());
						rowCount = sheet.getLastRowNum();

						for ( k = 1 ; k <= rowCount; ++k){
							x = entradaEscaner5.nextLine();//wait enter
							Row row = sheet.getRow(k);
							System.out.println();
							System.out.println ("Fila: " + k);
							cellCount = row.getLastCellNum();

							for ( m = 1 ; m < cellCount; ++m){

								Cell nro = row.getCell(m);

								if (m == 1 || m == 2){
									System.out.println (nro.getStringCellValue());
								}

								if(m == 3){
									System.out.println (nro.getNumericCellValue());
								}

							}

						}

					};

					inputStream.close();

					FileOutputStream outputStream = new FileOutputStream(excelFilePath);
					workbook.write(outputStream);
					workbook.close();
					outputStream.close();

				} catch (IOException | EncryptedDocumentException
						| InvalidFormatException ex) {
					ex.printStackTrace();
				}
				System.out.println();
				System.out.println();
				System.out.println();
				System.out.println ("...FIN DEL MODULO C...");
				System.out.println();
				System.out.println();
				System.out.println();

	}

	public static void main(String[] args) {
		System.out.println();
		System.out.println();
		System.out.println ("...Grupo 5...");
		System.out.println();
		System.out.println ("Jose Romero");
		System.out.println ("Moises Escudero");
		System.out.println();
		System.out.println();


		int ciclo=1;
		int menu=0;
		while (ciclo==1) {
			System.out.println();
			System.out.println ("Menu de inicio:");
			System.out.println();
			System.out.println ("Presione opcion '2' para el modulo B");
			System.out.println();
			System.out.println ("Presione opcion '3' para el modulo C");
			System.out.println();
			System.out.println ("Presione cualquier otro numero para salir..");
			Scanner entradaEscaner = new Scanner(System.in);
			menu = entradaEscaner.nextInt();
			if(menu==2){
				ModuloB();
			}
			else if(menu==3){
				ModuloC();
			}
			else{
				ciclo=0;
				System.out.println();
				System.out.println();
				System.out.println ("Programa cerrado exitosamente");
				System.out.println();
				System.out.println();
			}
			
		}
	}
}
