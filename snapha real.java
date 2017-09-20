

import java.io.File;
import java.io.IOException;
import java.util.Arrays;
import java.util.Date;
import java.util.Scanner;

import jxl.Cell;
import jxl.CellType;
import jxl.DateCell;
import jxl.LabelCell;
import jxl.NumberCell;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;
import jxl.write.Number;
import jxl.write.WritableCell; 

public class excel1 {

	public static void main (String[] args) throws WriteException, IOException, BiffException {
		Scanner console1 = new Scanner(System.in);
		Workbook workbook = Workbook.getWorkbook(new File ("D:\\snapha.xls"));
		WritableWorkbook copy = Workbook.createWorkbook(new File("D:\\temp.xls"), workbook);
		WritableSheet sheet = copy.getSheet(0); 
		
		//checks if name is new member or not 
		int nm = 0; 
		System.out.println("Let me check if member is new. Member name?");
		String name = console1.nextLine();
		if (nameExists(sheet, name)) {
			System.out.println("Member is already in system");
			nm = 1; 
		} else {
			System.out.println("New member!");
			nm = 2; 
		}

	//main functions of adding new member or adding mtg date
		
		//add new member, first mtg date, and first total mtg 
		if (nm == 2){
			Cell[] fill = sheet.getColumn(0);
			int lastempty = fill.length;
			label (sheet, 0, lastempty, name);
			System.out.println("Member added"); 
		
			int totalmtg = 1;
			addnmMtg(console1, sheet, name, totalmtg); 
			
		//add mtg date to existing member 
		} else if (nm == 1){
			int namerow = getIndex(0,sheet,name);
			int totalmtg = Integer.parseInt(printCell (sheet, 2, namerow));
			addMtg (console1, sheet, name, totalmtg);
			
		} else {
			System.out.println("invalid action");
		}

		copy.write(); 
		copy.close();
		//copy.importSheet(string , int , sheet);
		workbook.close();
	}
	
	//adds a cell with String contents
		public static void label(WritableSheet sheet, int column, int row, String label) throws RowsExceededException, WriteException {
			Label label1 = new Label (column,row, label);
			sheet.addCell(label1);
		}
		
	//prints cell from sheet using cell indexes 
	public static String printCell (WritableSheet sheet, int x, int y) {
		Cell a1 = sheet.getCell(x,y);
		String stringa = a1.getContents();
		return stringa;
	}
	
	//checks if member is in system already (column 0)
	public static boolean nameExists (WritableSheet sheet, String name) {
		boolean exist = false; 
		Cell[] a = sheet.getColumn(0);
		for (int i = 0; i < a.length; i++) {
			if (printCell(sheet, 0,i).equals(name)) {
				exist = true; 
			}
		}
		return exist; 
	}

	//gets row of cell only for names
	public static int getIndex (int column, WritableSheet sheet, String name) {
		Cell[] a = sheet.getColumn(column);
		int count = 0; 
		for (int i=0; i < a.length; i++){
			if (printCell(sheet, column,i).equals(name)) {
				return count; 
			} else {
				count++;
			}
		}
		return count;
	}
	
	public static void addMtg (Scanner console, WritableSheet sheet, String name, int totalmtg) throws RowsExceededException, WriteException {
		System.out.println("Enter new mtg date (x/xx/xxxx)");
		String date = console.nextLine();
		sheet.findCell(name);

		//adds row for a new date
		int namerow = getIndex(0,sheet,name);
		sheet.insertRow(namerow+1);
		label (sheet, 1, namerow + 1, date);

		//increases total meetings attended by 1
		totalmtg++;
		Number total = new Number (2, namerow, totalmtg); 
		sheet.addCell(total);

		System.out.println("Mtg added for " + name);
	} 
	public static void addnmMtg (Scanner console, WritableSheet sheet, String name, int totalmtg) throws RowsExceededException, WriteException {
		System.out.println("Enter first mtg date (x/xx/xxxx)");
		String date = console.nextLine();
		sheet.findCell(name);

		//adds cell for a new date next to name
		int namerow = getIndex(0,sheet,name);
		label (sheet, 1, namerow, date);

		//creates totalmtg cell for new member 
		label(sheet, 2, (getIndex (0, sheet,name)), "1");
		
		System.out.println("First mtg added for new member " + name);
	} 
}