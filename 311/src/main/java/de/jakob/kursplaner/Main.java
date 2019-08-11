package de.jakob.kursplaner;

import java.io.File;
import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Iterator;
import java.util.Scanner;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Main {

	public static void main(String[] args) throws Exception {

		ArrayList<File> files = new ArrayList<File>();
		File folder = new File("C:\\Users\\Jakob\\Desktop\\Schule\\Oberstufe\\Q1");

		for (File f : folder.listFiles()) {
			files.add(new File(f.getAbsolutePath() + File.separator + "Kurs.xlsx"));
		}

		String faecher = "";
		boolean trigger = false;
		boolean charIndicator = false;

		ArrayList<String> namen = new ArrayList<String>();
		String remember = "";

		for (File excelFile : files) {
			FileInputStream fis = new FileInputStream(excelFile);

			// we create an XSSF Workbook object for our XLSX Excel File
			XSSFWorkbook workbook = new XSSFWorkbook(fis);
			// we get first sheet
			XSSFSheet sheet = workbook.getSheetAt(0);

			// we iterate on rows
			Iterator<Row> rowIt = sheet.iterator();
			while (rowIt.hasNext()) {
				Row row = rowIt.next();

				// iterate on cells for the current row
				Iterator<Cell> cellIterator = row.cellIterator();

				while (cellIterator.hasNext()) {
					Cell cell = cellIterator.next();
					if (!trigger && !cell.toString().equals("Schüler") && !cell.toString().equals("Vorname")
							&& !cell.toString().equals("Name") && !cell.toString().equals(" ")) {
						trigger = true;
						remember = cell.toString();
					} else if (trigger) {
						trigger = false;
						remember += ", " + cell.toString();
						if (!namen.contains(remember)) {
							namen.add(remember);
						}
					}
				}

			}
			workbook.close();
		}

		Collections.sort(namen);
		namen.remove(0);
		String nachname = "";
		String vorname = "";
		for (String s : namen) {
			charIndicator = false;
			nachname = s.split(", ")[0];
			vorname = s.split(", ")[1];

			/*
			 * Scanner myObj = new Scanner(System.in); // Create a Scanner object
			 * System.out.println("Vor- und Nachname eingeben"); vorname = myObj.next();
			 * nachname = myObj.next(); vorname = "Jan Oliver"; System.out.println(nachname
			 * + ", " + vorname); myObj.close();
			 */
			for (File excelFile : files) {
				FileInputStream fis = new FileInputStream(excelFile);

				// we create an XSSF Workbook object for our XLSX Excel File
				XSSFWorkbook workbook = new XSSFWorkbook(fis);
				// we get first sheet
				XSSFSheet sheet = workbook.getSheetAt(0);

				// we iterate on rows
				Iterator<Row> rowIt = sheet.iterator();

				while (rowIt.hasNext()) {
					Row row = rowIt.next();

					// iterate on cells for the current row
					Iterator<Cell> cellIterator = row.cellIterator();

					while (cellIterator.hasNext()) {

						Cell cell = cellIterator.next();
						if (!trigger && cell.toString().equalsIgnoreCase(nachname)) {
							trigger = true;
							
						} else if (trigger && cell.toString().equalsIgnoreCase(vorname)) {
							trigger = false;
							faecher += ((charIndicator) ? ", " : " ") + excelFile.getParentFile().getName();
							charIndicator = true;
						} else if (trigger) {
							trigger = false;
							// faecher += " (" + excelFile.getParentFile().getName() + ")";
						}
					}

				}

				workbook.close();
				fis.close();
			}
			if (charIndicator) {
				System.out.println(nachname + ", " + vorname
						+ (faecher.contains(",")
								? " hast du in den Fächern" + faecher.substring(0, faecher.lastIndexOf(",")) + " und"
										+ faecher.substring(faecher.lastIndexOf(",") + 1)
								: " hast du nur im Fach" + faecher));
				if (faecher.contains("(")) {
					System.out.println("Klammern meinen Facher in denen nur der Nachname übereinstimmt");
				}
			} else {
				System.out.println(vorname + " " + nachname + " hast du in keinem Fach!");
			}
			faecher = "";
		}
	}
}
