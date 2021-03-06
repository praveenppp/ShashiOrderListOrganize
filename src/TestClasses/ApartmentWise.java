package TestClasses;

import java.util.ArrayList;

import ExcelSheet.CreateNewAssetsExcel;
import ExcelSheet.CustomerDataModel;
import ExcelSheet.LoginExcel;

public class ApartmentWise
{
public static void main(String[] args) {
	CreateNewAssetsExcel excel = new CreateNewAssetsExcel();
	ArrayList<CustomerDataModel> excelData=excel.readExcel("C:\\Orders", "orders.xlsx", "DailyOrders");
	LoginExcel excelHeader = new LoginExcel();
	
	for(CustomerDataModel data:excelData)
	{
		System.out.println(data.getApartment());
	}
}
}
