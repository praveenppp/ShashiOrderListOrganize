package ExcelSheet;

import java.io.File;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class CreateNewAssetsExcel {
public ArrayList<CustomerDataModel> readExcel(String filePath, String fileName, String sheetName) {
		
		/// create array of model object set to return 100 rows
		ArrayList<CustomerDataModel> model = new ArrayList<CustomerDataModel>();
		CustomerDataModel modeler = new CustomerDataModel();

		// Create a object of File class to open xlsx file
		File file = new File(filePath + "\\" + fileName);

		// Create an object of FileInputStream class to read excel file
		InputStream inputStream = null;
		try {
			inputStream = new FileInputStream(file);
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

		Workbook workbook = null;
//		System.out.println("coing 1");
		// Find the file extension by spliting file name in substring and
		// getting only extension name

		String fileExtensionName = fileName.substring(fileName.indexOf("."));
		DataFormatter objDefaultFormat = new DataFormatter();
		// Check condition if the file is xlsx file

		if (fileExtensionName.equals(".xlsx")) {

			// If it is xlsx file then create object of XSSFWorkbook class
//			System.out.println("coing 1.1");
			try {
				workbook = new XSSFWorkbook(inputStream);

			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			} catch (Exception e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}

		}

		// Check condition if the file is xls file

		else if (fileExtensionName.equals(".xls")) {

			// If it is xls file then create object of XSSFWorkbook class

			try {
				workbook = new HSSFWorkbook(inputStream);
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}

		}

		// Read sheet inside the workbook by its name
//		System.out.println("coing 2");
		Sheet sheet = workbook.getSheet(sheetName);

		// Find number of rows in excel file

		int rowCount = sheet.getLastRowNum() - sheet.getFirstRowNum();

		// Create a loop over all the rows of excel file to read it

		// int count=0;

		for (int i = 1; i <rowCount+1 ; i++) {// made i=1 to start from 2nd
			// row

			Row row = sheet.getRow(i);
//			System.out.println("No. of rows: "+rowCount);
			// Create a loop to print cell values in a row

			for (int j = 0; j < row.getLastCellNum(); j++) {
//				System.out.println("No. of columns: "+j+"/"+row.getLastCellNum());
				
				// Print excel data in console
//				System.out.println("number: " + j + " last cell no: " + row.getLastCellNum());

//				System.out.println("J - "+j);
				// This will evaluate the cell, And any type of cell will return
				// string value
				String cellString = objDefaultFormat.formatCellValue(row.getCell(j));
//				System.out.println("Value - "+cellString);
//				String cellString=row.getCell(j).toString();

				// getStringCellValue().toString();
//				System.out.println("outer: " + cellString + " j value: " + j);

				// logic to populate data
				if (j == 0) {

					modeler.setTimestamp(cellString);

				} else if (j == 1) {

					modeler.setName(cellString);
				}
				else if (j == 2) {

					modeler.setApartment(cellString);
				}
				else if (j == 3) {

					modeler.setFlatNumber(cellString);
				}
				else if (j == 4) {

					modeler.setAmount(cellString);
				}
				else if (j == 5) {

					modeler.setAgaseSopu(cellString);
				}
				else if (j == 6) {

					modeler.setKothambri(cellString);
				}
				else if (j == 7) {

					modeler.setKaribevu(cellString);
				}
				else if (j == 8) {

					modeler.setNugge(cellString);
				}
				else if (j == 9) {

					modeler.setGanike(cellString);
				}
				else if (j == 10) {

					modeler.setHarive(cellString);
				}
				else if (j == 11) {

					modeler.setPundi(cellString);
				}
				else if (j == 12) {

					modeler.setGreenAmaranthus(cellString);
				}
				else if (j == 13) {

					modeler.setMenthaya(cellString);
				}
				else if (j == 14) {

					modeler.setPalak(cellString);
				}
				else if (j ==15) {

					modeler.setKempDhantu(cellString);
				}
				else if (j == 16) {

					modeler.setSabsige(cellString);
				}
				else if (j == 17) {

					modeler.setVandelaga(cellString);
				}
				else if (j == 18) {

					modeler.setDoddPatra(cellString);
				}
				else if (j == 19) {

					modeler.setGreenLettuce(cellString);
				}
				else if (j == 20) {

					modeler.setAsparagus(cellString);
				}
				else if (j == 21) {

					modeler.setPakChoi(cellString);
				}
				else if (j == 22) {

					modeler.setParsley(cellString);
				}
				else if (j == 23) {

					modeler.setLemonGrass(cellString);
				}
				else if (j == 24) {

					modeler.setSpringOnion(cellString);
				}
				else if (j == 25) {

					modeler.setMint(cellString);
				}
				else if (j == 26) {

					modeler.setAmrutBalli(cellString);
				}
				else if (j == 27) {

					modeler.setChakota(cellString);
				}
				else if (j == 28) {

					modeler.setAmla(cellString);
				}
				else if (j == 29) {

					modeler.setBananaFlower(cellString);
				}
				else if (j == 30) {

					modeler.setBananaStem(cellString);
				}
				else if (j == 31) {

					modeler.setCoconut(cellString);
				}
				else if (j == 32) {

					modeler.setOnion(cellString);
				}
				else if (j == 33) {

					modeler.setPotato(cellString);
				}
				else if (j == 34) {

					modeler.setGarlic(cellString);
				}
				else if (j == 35) {

					modeler.setGinger(cellString);
				}
				else if (j == 36) {

					modeler.setBeetRoot(cellString);
				}
				else if (j == 37) {

					modeler.setCucumberGreen(cellString);
				}
				else if (j == 38) {

					modeler.setCucumberWhite(cellString);
				}
				else if (j == 39) {

					modeler.setFrenchCucumber(cellString);
				}
				else if (j == 40) {

					modeler.setCabbage(cellString);
				}
				else if (j == 41) {

					modeler.setBitterGourd(cellString);
				}
				else if (j == 42) {

					modeler.setNatiBitterGourd (cellString);
				}
				else if (j == 43) {

					modeler.setBottleGourd(cellString);
				}
				else if (j == 44) {

					modeler.setBottleBrinjal(cellString);
				}
				else if (j == 45) {

					modeler.setBrinjalGreen (cellString);
				}
				else if (j == 46) {

					modeler.setBrinjalRound(cellString);
				}
				else if (j == 47) {

					modeler.setChowChow(cellString);
				}
				else if (j == 48) {

					modeler.setCapsicumGreen(cellString);
				}
				else if (j == 49) {

					modeler.setColourCapsicum(cellString);
				}
				else if (j == 50) {

					modeler.setCarrotOoty(cellString);
				}
				else if (j == 51) {

					modeler.setCauliFlower(cellString);
				}
				else if (j == 52) {

					modeler.setChilliGreen(cellString);
				}
				else if (j == 53) {

					modeler.setCorn(cellString);
				}
				else if (j == 54) {

					modeler.setKnolKhol(cellString);
				}
				else if (j == 55) {

					modeler.setLadysFinger(cellString);
				}
				else if (j == 56) {

					modeler.setLemon(cellString);
				}
				else if (j == 57) {

					modeler.setRawBanana(cellString);
				}
				else if (j == 58) {

					modeler.setRidgeGourd(cellString);
				}
				else if (j == 59) {

					modeler.setSambharCucumbe(cellString);
				}
				else if (j == 60) {

					modeler.setSnakeGourd (cellString);
				}
				else if (j == 61) {

					modeler.setSweetCorn(cellString);
				}
				else if (j == 62) {

					modeler.setSweetPotato(cellString);
				}
				else if (j == 63) {

					modeler.setBroccoli(cellString);
				}
				else if (j == 64) {

					modeler.setFarmTomato(cellString);
				}
				else if (j == 65) {

					modeler.setClusteredBeans(cellString);
				}
				else if (j == 66) {

					modeler.setBeansFarm(cellString);
				}
				else if (j == 67) {

					modeler.setBeansFlat(cellString);
				}
				else if (j == 68) {

					modeler.setBeansNati(cellString);
				}
				else if (j == 69) {

					modeler.setBeansRing3(cellString);
				}
				else if (j == 70) {

					modeler.setNatiBattani(cellString);
				}
				else if (j == 71) {

					modeler.setBattani(cellString);
				}
				else if (j == 72) {

					modeler.setRadish(cellString);
				}
				else if (j == 73) {

					modeler.setSweetPumpkin(cellString);
				}
				else if (j == 74) {

					modeler.setTomatoNati(cellString);
				}
				else if (j == 75) {

					modeler.setCelery(cellString);
				}
				else if (j == 76) {

					modeler.setWhitePumpkin(cellString);
				}
				else if (j == 77) {

					modeler.setYamRoot(cellString);
				}
				else if (j == 78) {

					modeler.setLongBeans(cellString);
				}
				else if (j == 79) {

					modeler.setChineseCabbage (cellString);
				}
				else if (j == 80) {

					modeler.setBabyCorn(cellString);
				}
				else if (j == 81) {

					modeler.setBajjiChilli(cellString);
				}
				else if (j == 82) {

					modeler.setIcebergLettuce(cellString);
				}
				else if (j == 83) {

					modeler.setCherryTomato(cellString);
				}
				else if (j == 84) {

					modeler.setSoakedAvarekalu(cellString);
				}
				else if (j == 85) {

					modeler.setBananaPachabale(cellString);
				}
				else if (j == 86) {

					modeler.setPomegranate(cellString);
				}
				else if (j == 87) {

					modeler.setBananaYellaki(cellString);
				}
				else if (j == 88) {

					modeler.setBananaRed(cellString);
				}
				else if (j == 89) {

					modeler.setMuskMellon(cellString);
				}
				else if (j == 90) {

					modeler.setPapayaredlady(cellString);
				}
				else if (j == 91) {

					modeler.setAppleRed(cellString);
				}
				else if (j == 92) {

					modeler.setAppleGreen(cellString);
				}
				else if (j == 93) {

					modeler.setGrapesGreen(cellString);
				}
				else if (j == 94) {

					modeler.setGrapesBlack(cellString);
				}
				else if (j == 95) {

					modeler.setMangoTotapuriRaw(cellString);
				}
				else if (j == 96) {

					modeler.setPineapple(cellString);
				}
				else if (j == 97) {

					modeler.setSapota(cellString);
				}
				else if (j == 98) {

					modeler.setOrangeNati(cellString);
				}
				else if (j == 99) {

					modeler.setOrangeMalta(cellString);
				}
				else if (j == 100) {

					modeler.setMusambi(cellString);
				}
				else if (j == 101) {

					modeler.setButterfruit(cellString);
				}
				else if (j == 102) {

					modeler.setKiwi(cellString);
				}
				else if (j == 103) {

					modeler.setMushroom(cellString);
					model.add(modeler);
					modeler = new CustomerDataModel();
					
				}
				
				
			}
		}
				
//               System.out.println("count of objects: ");


		return model;
	}
}






