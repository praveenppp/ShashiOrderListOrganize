package TestClasses;

import java.io.IOException;
import java.util.ArrayList;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import ExcelSheet.CreateNewAssetsExcel;
import ExcelSheet.CustomerDataModel;
import ExcelSheet.LoginExcel;

public class OrganizeOrder 
{
public static void main(String[] args) throws InvalidFormatException, InterruptedException, IOException {

	CreateNewAssetsExcel excel = new CreateNewAssetsExcel();
	ArrayList<CustomerDataModel> excelData=excel.readExcel("C:\\Orders", "orders.xlsx", "DailyOrders");
	LoginExcel excelHeader = new LoginExcel();
	int count = 1;
	int coulumn =0;
//	excelHeader.createPage("Organized");    
	ArrayList<String> datas=new ArrayList<String>();
	for(CustomerDataModel data:excelData)
	{
		count++;
		System.out.println("Name - "+data.getName());
		System.out.println("Apartment - "+data.getApartment());
		System.out.println("FlatNumber - "+data.getFlatNumber());
		System.out.println("Amount - "+data.getAmount());
		
		
		datas.add((excelHeader.getExcelData("DailyOrders", 0, 1)));datas.add(data.getName());
		excelHeader.setExcelData1("Organized", count++, 1, datas);
		datas.clear();
		
		datas.add((excelHeader.getExcelData("DailyOrders", 0, 2)));datas.add(data.getApartment());
		excelHeader.setExcelData1("Organized", count++, 1, datas);
		datas.clear();
		
		datas.add((excelHeader.getExcelData("DailyOrders", 0, 3)));datas.add(data.getFlatNumber());
		excelHeader.setExcelData1("Organized", count++, 1, datas);
		datas.clear();
		
		datas.add((excelHeader.getExcelData("DailyOrders", 0, 4)));datas.add(data.getAmount());
		excelHeader.setExcelData1("Organized", count++, 1, datas);
		datas.clear();
		
		if(!data.getAgaseSopu().isEmpty())
		{
			datas.add((excelHeader.getExcelData("DailyOrders", 0, 5)));datas.add(data.getAgaseSopu());
			excelHeader.setExcelData1("Organized", count++, 1, datas);
			datas.clear();
		}
		if(!data.getKothambri().isEmpty())
		{
			datas.add((excelHeader.getExcelData("DailyOrders", 0, 6)));datas.add(data.getKothambri());
			excelHeader.setExcelData1("Organized", count++, 1, datas);
			datas.clear();
		}
		if(!data.getKaribevu().isEmpty())
		{
			datas.add((excelHeader.getExcelData("DailyOrders", 0, 7)));datas.add(data.getKaribevu());
			excelHeader.setExcelData1("Organized", count++, 1, datas);
			datas.clear();
		}
		if(!data.getNugge().isEmpty())
		{
			datas.add((excelHeader.getExcelData("DailyOrders", 0, 8)));datas.add(data.getNugge());
			excelHeader.setExcelData1("Organized", count++, 1, datas);
			datas.clear();
		}
		if(!data.getGanike().isEmpty())
		{
			datas.add((excelHeader.getExcelData("DailyOrders", 0, 9)));datas.add(data.getGanike());
			excelHeader.setExcelData1("Organized", count++, 1, datas);
			datas.clear();
		}
		if(!data.getHarive().isEmpty())
		{
			datas.add((excelHeader.getExcelData("DailyOrders", 0, 10)));datas.add(data.getHarive());
			excelHeader.setExcelData1("Organized", count++, 1, datas);
			datas.clear();
		}
		if(!data.getPundi().isEmpty())
		{
			datas.add((excelHeader.getExcelData("DailyOrders", 0, 11)));datas.add(data.getPundi());
			excelHeader.setExcelData1("Organized", count++, 1, datas);
			datas.clear();
		}
		if(!data.getGreenAmaranthus().isEmpty())
		{
			datas.add((excelHeader.getExcelData("DailyOrders", 0, 12)));datas.add(data.getGreenAmaranthus());
			excelHeader.setExcelData1("Organized", count++, 1, datas);
			datas.clear();
		}
		if(!data.getMenthaya().isEmpty())
		{
			datas.add((excelHeader.getExcelData("DailyOrders", 0, 13)));datas.add(data.getMenthaya());
			excelHeader.setExcelData1("Organized", count++, 1, datas);
			datas.clear();
		}
		if(!data.getPalak().isEmpty())
		{
			datas.add((excelHeader.getExcelData("DailyOrders", 0, 14)));datas.add(data.getPalak());
			excelHeader.setExcelData1("Organized", count++, 1, datas);
			datas.clear();
		}
		if(!data.getKempDhantu().isEmpty())
		{
			datas.add((excelHeader.getExcelData("DailyOrders", 0, 15)));datas.add(data.getKempDhantu());
			excelHeader.setExcelData1("Organized", count++, 1, datas);
			datas.clear();
		}
		if(!data.getSabsige().isEmpty())
		{
			datas.add((excelHeader.getExcelData("DailyOrders", 0, 16)));datas.add(data.getSabsige());
			excelHeader.setExcelData1("Organized", count++, 1, datas);
			datas.clear();
		}
		if(!data.getVandelaga().isEmpty())
		{
			datas.add((excelHeader.getExcelData("DailyOrders", 0, 17)));datas.add(data.getVandelaga());
			excelHeader.setExcelData1("Organized", count++, 1, datas);
			datas.clear();
		}
		if(!data.getDoddPatra().isEmpty())
		{
			datas.add((excelHeader.getExcelData("DailyOrders", 0, 18)));datas.add(data.getDoddPatra());
			excelHeader.setExcelData1("Organized", count++, 1, datas);
			datas.clear();
		}
		if(!data.getGreenLettuce().isEmpty())
		{
			datas.add((excelHeader.getExcelData("DailyOrders", 0, 19)));datas.add(data.getGreenLettuce());
			excelHeader.setExcelData1("Organized", count++, 1, datas);
			datas.clear();
		}
		if(!data.getAsparagus().isEmpty())
		{
			datas.add((excelHeader.getExcelData("DailyOrders", 0, 20)));datas.add(data.getAsparagus());
			excelHeader.setExcelData1("Organized", count++, 1, datas);
			datas.clear();
		}
		if(!data.getPakChoi().isEmpty())
		{
			datas.add((excelHeader.getExcelData("DailyOrders", 0, 21)));datas.add(data.getPakChoi());
			excelHeader.setExcelData1("Organized", count++, 1, datas);
			datas.clear();
		}
		if(!data.getParsley().isEmpty())
		{
			datas.add((excelHeader.getExcelData("DailyOrders", 0, 22)));datas.add(data.getParsley());
			excelHeader.setExcelData1("Organized", count++, 1, datas);
			datas.clear();
		}
		if(!data.getLemonGrass().isEmpty())
		{
			datas.add((excelHeader.getExcelData("DailyOrders", 0, 23)));datas.add(data.getLemonGrass());
			excelHeader.setExcelData1("Organized", count++, 1, datas);
			datas.clear();
		}
		if(!data.getSpringOnion().isEmpty())
		{
			datas.add((excelHeader.getExcelData("DailyOrders", 0, 24)));datas.add(data.getSpringOnion());
			excelHeader.setExcelData1("Organized", count++, 1, datas);
			datas.clear();
		}
		if(!data.getMint().isEmpty())
		{
			datas.add((excelHeader.getExcelData("DailyOrders", 0, 25)));datas.add(data.getMint());
			excelHeader.setExcelData1("Organized", count++, 1, datas);
			datas.clear();
		}
		if(!data.getAmrutBalli().isEmpty())
		{
			datas.add((excelHeader.getExcelData("DailyOrders", 0, 26)));datas.add(data.getAmrutBalli());
			excelHeader.setExcelData1("Organized", count++, 1, datas);
			datas.clear();
		}
		if(!data.getChakota().isEmpty())
		{
			datas.add((excelHeader.getExcelData("DailyOrders", 0, 27)));datas.add(data.getChakota());
			excelHeader.setExcelData1("Organized", count++, 1, datas);
			datas.clear();
		}
		if(!data.getAmla().isEmpty())
		{
			datas.add((excelHeader.getExcelData("DailyOrders", 0, 28)));datas.add(data.getAmla());
			excelHeader.setExcelData1("Organized", count++, 1, datas);
			datas.clear();
		}
		if(!data.getBananaFlower().isEmpty())
		{
			datas.add((excelHeader.getExcelData("DailyOrders", 0, 29)));datas.add(data.getBananaFlower());
			excelHeader.setExcelData1("Organized", count++, 1, datas);
			datas.clear();
		}
		if(!data.getBananaStem().isEmpty())
		{
			datas.add((excelHeader.getExcelData("DailyOrders", 0, 30)));datas.add(data.getBananaStem());
			excelHeader.setExcelData1("Organized", count++, 1, datas);
			datas.clear();
		}
		if(!data.getCoconut().isEmpty())
		{
			datas.add((excelHeader.getExcelData("DailyOrders", 0, 31)));datas.add(data.getCoconut());
			excelHeader.setExcelData1("Organized", count++, 1, datas);
			datas.clear();
		}
		if(!data.getOnion().isEmpty())
		{
			datas.add((excelHeader.getExcelData("DailyOrders", 0, 32)));datas.add(data.getOnion());
			excelHeader.setExcelData1("Organized", count++, 1, datas);
			datas.clear();
		}
		if(!data.getPotato().isEmpty())
		{
			datas.add((excelHeader.getExcelData("DailyOrders", 0, 33)));datas.add(data.getPotato());
			excelHeader.setExcelData1("Organized", count++, 1, datas);
			datas.clear();
		}
		if(!data.getGarlic().isEmpty())
		{
			datas.add((excelHeader.getExcelData("DailyOrders", 0, 34)));datas.add(data.getGarlic());
			excelHeader.setExcelData1("Organized", count++, 1, datas);
			datas.clear();
		}
		if(!data.getGinger().isEmpty())
		{
			datas.add((excelHeader.getExcelData("DailyOrders", 0, 35)));datas.add(data.getGinger());
			excelHeader.setExcelData1("Organized", count++, 1, datas);
			datas.clear();
		}
		if(!data.getBeetRoot().isEmpty())
		{
			datas.add((excelHeader.getExcelData("DailyOrders", 0, 36)));datas.add(data.getBeetRoot());
			excelHeader.setExcelData1("Organized", count++, 1, datas);
			datas.clear();
		}
		if(!data.getCucumberGreen().isEmpty())
		{
			datas.add((excelHeader.getExcelData("DailyOrders", 0, 37)));datas.add(data.getCucumberGreen());
			excelHeader.setExcelData1("Organized", count++, 1, datas);
			datas.clear();
		}
		if(!data.getCucumberWhite().isEmpty())
		{
			datas.add((excelHeader.getExcelData("DailyOrders", 0, 38)));datas.add(data.getCucumberWhite());
			excelHeader.setExcelData1("Organized", count++, 1, datas);
			datas.clear();
		}
		if(!data.getFrenchCucumber().isEmpty())
		{
			datas.add((excelHeader.getExcelData("DailyOrders", 0, 39)));datas.add(data.getFrenchCucumber());
			excelHeader.setExcelData1("Organized", count++, 1, datas);
			datas.clear();
		}
		if(!data.getCabbage().isEmpty())
		{
			datas.add((excelHeader.getExcelData("DailyOrders", 0, 40)));datas.add(data.getCabbage());
			excelHeader.setExcelData1("Organized", count++, 1, datas);
			datas.clear();
		}
		if(!data.getBitterGourd().isEmpty())
		{
			datas.add((excelHeader.getExcelData("DailyOrders", 0, 41)));datas.add(data.getBitterGourd());
			excelHeader.setExcelData1("Organized", count++, 1, datas);
			datas.clear();
		}
		if(!data.getNatiBitterGourd().isEmpty())
		{
			datas.add((excelHeader.getExcelData("DailyOrders", 0, 42)));datas.add(data.getNatiBitterGourd());
			excelHeader.setExcelData1("Organized", count++, 1, datas);
			datas.clear();
		}
		if(!data.getBottleGourd().isEmpty())
		{
			datas.add((excelHeader.getExcelData("DailyOrders", 0, 43)));datas.add(data.getBottleGourd());
			excelHeader.setExcelData1("Organized", count++, 1, datas);
			datas.clear();
		}
		if(!data.getBottleBrinjal().isEmpty())
		{
			datas.add((excelHeader.getExcelData("DailyOrders", 0, 44)));datas.add(data.getBottleBrinjal());
			excelHeader.setExcelData1("Organized", count++, 1, datas);
			datas.clear();
		}
		if(!data.getBrinjalGreen().isEmpty())
		{
			datas.add((excelHeader.getExcelData("DailyOrders", 0, 45)));datas.add(data.getBrinjalGreen());
			excelHeader.setExcelData1("Organized", count++, 1, datas);
			datas.clear();
		}
		if(!data.getBrinjalRound().isEmpty())
		{
			datas.add((excelHeader.getExcelData("DailyOrders", 0, 46)));datas.add(data.getBrinjalRound());
			excelHeader.setExcelData1("Organized", count++, 1, datas);
			datas.clear();
		}
		if(!data.getChowChow().isEmpty())
		{
			datas.add((excelHeader.getExcelData("DailyOrders", 0, 47)));datas.add(data.getChowChow());
			excelHeader.setExcelData1("Organized", count++, 1, datas);
			datas.clear();
		}
		if(!data.getCapsicumGreen().isEmpty())
		{
			datas.add((excelHeader.getExcelData("DailyOrders", 0, 48)));datas.add(data.getCapsicumGreen());
			excelHeader.setExcelData1("Organized", count++, 1, datas);
			datas.clear();
		}
		if(!data.getColourCapsicum().isEmpty())
		{
			datas.add((excelHeader.getExcelData("DailyOrders", 0, 49)));datas.add(data.getColourCapsicum());
			excelHeader.setExcelData1("Organized", count++, 1, datas);
			datas.clear();
		}
		if(!data.getCarrotOoty().isEmpty())
		{
			datas.add((excelHeader.getExcelData("DailyOrders", 0, 50)));datas.add(data.getCarrotOoty());
			excelHeader.setExcelData1("Organized", count++, 1, datas);
			datas.clear();
		}
		if(!data.getCauliFlower().isEmpty())
		{
			datas.add((excelHeader.getExcelData("DailyOrders", 0, 51)));datas.add(data.getCauliFlower());
			excelHeader.setExcelData1("Organized", count++, 1, datas);
			datas.clear();
		}
		if(!data.getChilliGreen().isEmpty())
		{
			datas.add((excelHeader.getExcelData("DailyOrders", 0, 52)));datas.add(data.getChilliGreen());
			excelHeader.setExcelData1("Organized", count++, 1, datas);
			datas.clear();
		}
		if(!data.getCorn().isEmpty())
		{
			datas.add((excelHeader.getExcelData("DailyOrders", 0, 53)));datas.add(data.getCorn());
			excelHeader.setExcelData1("Organized", count++, 1, datas);
			datas.clear();
		}
		if(!data.getKnolKhol().isEmpty())
		{
			datas.add((excelHeader.getExcelData("DailyOrders", 0, 54)));datas.add(data.getKnolKhol());
			excelHeader.setExcelData1("Organized", count++, 1, datas);
			datas.clear();
		}
		if(!data.getLadysFinger().isEmpty())
		{
			datas.add((excelHeader.getExcelData("DailyOrders", 0, 55)));datas.add(data.getLadysFinger());
			excelHeader.setExcelData1("Organized", count++, 1, datas);
			datas.clear();
		}
		if(!data.getLemon().isEmpty())
		{
			datas.add((excelHeader.getExcelData("DailyOrders", 0, 56)));datas.add(data.getLemon());
			excelHeader.setExcelData1("Organized", count++, 1, datas);
			datas.clear();
		}
		if(!data.getRawBanana().isEmpty())
		{
			datas.add((excelHeader.getExcelData("DailyOrders", 0, 57)));datas.add(data.getRawBanana());
			excelHeader.setExcelData1("Organized", count++, 1, datas);
			datas.clear();
		}
		if(!data.getRidgeGourd().isEmpty())
		{
			datas.add((excelHeader.getExcelData("DailyOrders", 0, 58)));datas.add(data.getRidgeGourd());
			excelHeader.setExcelData1("Organized", count++, 1, datas);
			datas.clear();
		}
		if(!data.getSambharCucumbe().isEmpty())
		{
			datas.add((excelHeader.getExcelData("DailyOrders", 0, 59)));datas.add(data.getSambharCucumbe());
			excelHeader.setExcelData1("Organized", count++, 1, datas);
			datas.clear();
		}
		if(!data.getSnakeGourd().isEmpty())
		{
			datas.add((excelHeader.getExcelData("DailyOrders", 0, 60)));datas.add(data.getSnakeGourd());
			excelHeader.setExcelData1("Organized", count++, 1, datas);
			datas.clear();
		}
		if(!data.getSweetCorn().isEmpty())
		{
			datas.add((excelHeader.getExcelData("DailyOrders", 0, 61)));datas.add(data.getSweetCorn());
			excelHeader.setExcelData1("Organized", count++, 1, datas);
			datas.clear();
		}
		
		if(!data.getSweetPotato().isEmpty())
		{
			datas.add((excelHeader.getExcelData("DailyOrders", 0, 62)));datas.add(data.getSweetPotato());
			excelHeader.setExcelData1("Organized", count++, 1, datas);
			datas.clear();
		}
		if(!data.getBroccoli().isEmpty())
		{
			datas.add((excelHeader.getExcelData("DailyOrders", 0, 63)));datas.add(data.getBroccoli());
			excelHeader.setExcelData1("Organized", count++, 1, datas);
			datas.clear();
		}
		if(!data.getFarmTomato().isEmpty())
		{
			datas.add((excelHeader.getExcelData("DailyOrders", 0, 64)));datas.add(data.getFarmTomato());
			excelHeader.setExcelData1("Organized", count++, 1, datas);
			datas.clear();
		}
		if(!data.getClusteredBeans().isEmpty())
		{
			datas.add((excelHeader.getExcelData("DailyOrders", 0, 65)));datas.add(data.getClusteredBeans());
			excelHeader.setExcelData1("Organized", count++, 1, datas);
			datas.clear();
		}
		if(!data.getBeansFarm().isEmpty())
		{
			datas.add((excelHeader.getExcelData("DailyOrders", 0, 66)));datas.add(data.getBeansFarm());
			excelHeader.setExcelData1("Organized", count++, 1, datas);
			datas.clear();
		}
		if(!data.getBeansFlat().isEmpty())
		{
			datas.add((excelHeader.getExcelData("DailyOrders", 0, 67)));datas.add(data.getBeansFlat());
			excelHeader.setExcelData1("Organized", count++, 1, datas);
			datas.clear();
		}
		if(!data.getBeansNati().isEmpty())
		{
			datas.add((excelHeader.getExcelData("DailyOrders", 0, 68)));datas.add(data.getBeansNati());
			excelHeader.setExcelData1("Organized", count++, 1, datas);
			datas.clear();
		}
		if(!data.getBeansRing3().isEmpty())
		{
			datas.add((excelHeader.getExcelData("DailyOrders", 0, 69)));datas.add(data.getBeansRing3());
			excelHeader.setExcelData1("Organized", count++, 1, datas);
			datas.clear();
		}
		if(!data.getNatiBattani().isEmpty())
		{
			datas.add((excelHeader.getExcelData("DailyOrders", 0, 70)));datas.add(data.getNatiBattani());
			excelHeader.setExcelData1("Organized", count++, 1, datas);
			datas.clear();
		}
		if(!data.getBattani().isEmpty())
		{
			datas.add((excelHeader.getExcelData("DailyOrders", 0, 71)));datas.add(data.getBattani());
			excelHeader.setExcelData1("Organized", count++, 1, datas);
			datas.clear();
		}
		if(!data.getRadish().isEmpty())
		{
			datas.add((excelHeader.getExcelData("DailyOrders", 0, 72)));datas.add(data.getRadish());
			excelHeader.setExcelData1("Organized", count++, 1, datas);
			datas.clear();
		}
		if(!data.getSweetPumpkin().isEmpty())
		{
			datas.add((excelHeader.getExcelData("DailyOrders", 0, 73)));datas.add(data.getSweetPumpkin());
			excelHeader.setExcelData1("Organized", count++, 1, datas);
			datas.clear();
		}
		if(!data.getTomatoNati().isEmpty())
		{
			datas.add((excelHeader.getExcelData("DailyOrders", 0, 74)));datas.add(data.getTomatoNati());
			excelHeader.setExcelData1("Organized", count++, 1, datas);
			datas.clear();
		}
		if(!data.getCelery().isEmpty())
		{
			datas.add((excelHeader.getExcelData("DailyOrders", 0, 75)));datas.add(data.getCelery());
			excelHeader.setExcelData1("Organized", count++, 1, datas);
			datas.clear();
		}
		if(!data.getWhitePumpkin().isEmpty())
		{
			datas.add((excelHeader.getExcelData("DailyOrders", 0, 76)));datas.add(data.getWhitePumpkin());
			excelHeader.setExcelData1("Organized", count++, 1, datas);
			datas.clear();
		}
		if(!data.getYamRoot().isEmpty())
		{
			datas.add((excelHeader.getExcelData("DailyOrders", 0, 77)));datas.add(data.getYamRoot());
			excelHeader.setExcelData1("Organized", count++, 1, datas);
			datas.clear();
		}
		if(!data.getLongBeans().isEmpty())
		{
			datas.add((excelHeader.getExcelData("DailyOrders", 0, 78)));datas.add(data.getLongBeans());
			excelHeader.setExcelData1("Organized", count++, 1, datas);
			datas.clear();
		}
		if(!data.getChineseCabbage().isEmpty())
		{
			datas.add((excelHeader.getExcelData("DailyOrders", 0, 79)));datas.add(data.getChineseCabbage());
			excelHeader.setExcelData1("Organized", count++, 1, datas);
			datas.clear();
		}
		if(!data.getBabyCorn().isEmpty())
		{
			datas.add((excelHeader.getExcelData("DailyOrders", 0, 80)));datas.add(data.getBabyCorn());
			excelHeader.setExcelData1("Organized", count++, 1, datas);
			datas.clear();
		}
		if(!data.getBajjiChilli().isEmpty())
		{
			datas.add((excelHeader.getExcelData("DailyOrders", 0, 81)));datas.add(data.getBajjiChilli());
			excelHeader.setExcelData1("Organized", count++, 1, datas);
			datas.clear();
		}
		if(!data.getIcebergLettuce().isEmpty())
		{
			datas.add((excelHeader.getExcelData("DailyOrders", 0, 82)));datas.add(data.getIcebergLettuce());
			excelHeader.setExcelData1("Organized", count++, 1, datas);
			datas.clear();
		}
		if(!data.getCherryTomato().isEmpty())
		{
			datas.add((excelHeader.getExcelData("DailyOrders", 0, 83)));datas.add(data.getCherryTomato());
			excelHeader.setExcelData1("Organized", count++, 1, datas);
			datas.clear();
		}
		if(!data.getSoakedAvarekalu().isEmpty())
		{
			datas.add((excelHeader.getExcelData("DailyOrders", 0, 84)));datas.add(data.getSoakedAvarekalu());
			excelHeader.setExcelData1("Organized", count++, 1, datas);
			datas.clear();
		}
		if(!data.getBananaPachabale().isEmpty())
		{
			datas.add((excelHeader.getExcelData("DailyOrders", 0, 85)));datas.add(data.getBananaPachabale());
			excelHeader.setExcelData1("Organized", count++, 1, datas);
			datas.clear();
		}
		if(!data.getPomegranate().isEmpty())
		{
			datas.add((excelHeader.getExcelData("DailyOrders", 0, 86)));datas.add(data.getPomegranate());
			excelHeader.setExcelData1("Organized", count++, 1, datas);
			datas.clear();
		}
		if(!data.getBananaYellaki().isEmpty())
		{
			datas.add((excelHeader.getExcelData("DailyOrders", 0, 87)));datas.add(data.getBananaYellaki());
			excelHeader.setExcelData1("Organized", count++, 1, datas);
			datas.clear();
		}
		if(!data.getBananaRed().isEmpty())
		{
			datas.add((excelHeader.getExcelData("DailyOrders", 0, 88)));datas.add(data.getBananaRed());
			excelHeader.setExcelData1("Organized", count++, 1, datas);
			datas.clear();
		}
		if(!data.getMuskMellon().isEmpty())
		{
			datas.add((excelHeader.getExcelData("DailyOrders", 0, 89)));datas.add(data.getMuskMellon());
			excelHeader.setExcelData1("Organized", count++, 1, datas);
			datas.clear();
		}
		if(!data.getPapayaredlady().isEmpty())
		{
			datas.add((excelHeader.getExcelData("DailyOrders", 0, 90)));datas.add(data.getPapayaredlady());
			excelHeader.setExcelData1("Organized", count++, 1, datas);
			datas.clear();
		}
		if(!data.getAppleRed().isEmpty())
		{
			datas.add((excelHeader.getExcelData("DailyOrders", 0, 91)));datas.add(data.getAppleRed());
			excelHeader.setExcelData1("Organized", count++, 1, datas);
			datas.clear();
		}
		if(!data.getAppleGreen().isEmpty())
		{
			datas.add((excelHeader.getExcelData("DailyOrders", 0, 92)));datas.add(data.getAppleGreen());
			excelHeader.setExcelData1("Organized", count++, 1, datas);
			datas.clear();
		}
		if(!data.getGrapesGreen().isEmpty())
		{
			datas.add((excelHeader.getExcelData("DailyOrders", 0, 93)));datas.add(data.getGrapesGreen());
			excelHeader.setExcelData1("Organized", count++, 1, datas);
			datas.clear();
		}
		if(!data.getGrapesBlack().isEmpty())
		{
			datas.add((excelHeader.getExcelData("DailyOrders", 0, 94)));datas.add(data.getGrapesBlack());
			excelHeader.setExcelData1("Organized", count++, 1, datas);
			datas.clear();
		}
		if(!data.getMangoTotapuriRaw().isEmpty())
		{
			datas.add((excelHeader.getExcelData("DailyOrders", 0, 95)));datas.add(data.getMangoTotapuriRaw());
			excelHeader.setExcelData1("Organized", count++, 1, datas);
			datas.clear();
		}
		if(!data.getPineapple().isEmpty())
		{
			datas.add((excelHeader.getExcelData("DailyOrders", 0, 96)));datas.add(data.getPineapple());
			excelHeader.setExcelData1("Organized", count++, 1, datas);
			datas.clear();
		}
		if(!data.getSapota().isEmpty())
		{
			datas.add((excelHeader.getExcelData("DailyOrders", 0, 97)));datas.add(data.getSapota());
			excelHeader.setExcelData1("Organized", count++, 1, datas);
			datas.clear();
		}
		if(!data.getOrangeNati().isEmpty())
		{
			datas.add((excelHeader.getExcelData("DailyOrders", 0, 98)));datas.add(data.getOrangeNati());
			excelHeader.setExcelData1("Organized", count++, 1, datas);
			datas.clear();
		}
		if(!data.getOrangeMalta().isEmpty())
		{
			datas.add((excelHeader.getExcelData("DailyOrders", 0, 99)));datas.add(data.getOrangeMalta());
			excelHeader.setExcelData1("Organized", count++, 1, datas);
			datas.clear();
		}
		if(!data.getMusambi().isEmpty())
		{
			datas.add((excelHeader.getExcelData("DailyOrders", 0, 100)));datas.add(data.getMusambi());
			excelHeader.setExcelData1("Organized", count++, 1, datas);
			datas.clear();
		}
		if(!data.getButterfruit().isEmpty())
		{
			datas.add((excelHeader.getExcelData("DailyOrders", 0, 101)));datas.add(data.getButterfruit());
			excelHeader.setExcelData1("Organized", count++, 1, datas);
			datas.clear();
		}
		if(!data.getKiwi().isEmpty())
		{
			datas.add((excelHeader.getExcelData("DailyOrders", 0, 102)));datas.add(data.getKiwi());
			excelHeader.setExcelData1("Organized", count++, 1, datas);
			datas.clear();
		}
		if(!data.getMushroom().isEmpty())
		{
			datas.add((excelHeader.getExcelData("DailyOrders", 0, 103)));datas.add(data.getMushroom());
			excelHeader.setExcelData1("Organized", count++, 1, datas);
			datas.clear();
		}
		
		count++;
		System.out.println("------------------------------------------");
	}
//	for(int i=7;i<103;i++)
//	{
//		System.out.println("if(!data.getKothambri().isEmpty())");
//			System.out.println("\tSystem.out.println(excelHeader.getExcelData(\"DailyOrders\", 0, "+i+")+\" - \"+ data.getKothambri());");
//		
//	}
	

}
}
