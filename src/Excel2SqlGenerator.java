import java.io.BufferedWriter;
import java.io.File;
import java.io.FileWriter;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.WorkbookSettings;
import jxl.read.biff.BiffException;

public class Excel2SqlGenerator {
	private static final String EXCELPATH = "/Path/to/Source/excelfile.xls";
	private static final String OUTPUTNAME = "InsertScriptsGeneratedFromExcel.sql";
	private static final String TABLENAME = "tuik.egitim_mahalle";
	private static List<Column> columnList;
	
	public static void main(String[] args) {
		try {
			configureColumns();
			generateSqlInsertScripts(0);
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
	
	private static void configureColumns(){
		columnList = new ArrayList<>();
		columnList.add(new Column("col1", Column.TYPE_NUMBER));
		columnList.add(new Column("col2", Column.TYPE_STRING));
	}
	
	private static void generateSqlInsertScripts(int sheetNum) throws IOException{
		File inputWorkbook = new File(EXCELPATH);
		Workbook w;
		try {
			WorkbookSettings workbookSettings = new WorkbookSettings();
			workbookSettings.setEncoding("Cp1252");
			w = Workbook.getWorkbook(inputWorkbook, workbookSettings);
			Sheet sheet = w.getSheet(sheetNum);
			String inserts[] = new String[sheet.getRows()];
			for (int i = 1; i < sheet.getRows(); i++) {
				//inserts[i] = "INSERT INTO Location (pocellkod,istasyonkod,oraclekod,bayi_ad,adres,responsible,tel,mintika,bolge,priority,kanal,longitude,latitude) values (";
				inserts[i] = "INSERT INTO "+TABLENAME+" (";
				for (Column column : columnList) {
					inserts[i]+=column.name+",";
				}
				inserts[i] = inserts[i].substring(0, inserts[i].length()-1);
				inserts[i]+=") values (";
				for (int j = 0; j < sheet.getColumns(); j++) {
					Cell cell = sheet.getCell(j, i);
					/*
					 * You can use cell data type instead of my Column values
					 * cell.getType();
					 */
					Column currColumn = columnList.get(j);
					if(currColumn.type.equals(Column.TYPE_NUMBER)){
						if(cell.getContents()==null||cell.getContents().equals(""))
							inserts[i] += 0;
						else
							inserts[i] += cell.getContents();
					}else if(currColumn.type.equals(Column.TYPE_STRING)){
						inserts[i] += "'" + cell.getContents() + "'";
					}else{
						//unknown type
					}
					if (j != sheet.getColumns() - 1)
						inserts[i] += ",";
				}
				inserts[i] += ");";
				System.out.println(inserts[i]);
			}
			String sql = "";
			for (int i = 1; i < inserts.length; i++)
				sql += inserts[i] + "\n";
			try {
				System.out.println("Creating file...");
				// Create file
				FileWriter fstream = new FileWriter(OUTPUTNAME);
				BufferedWriter out = new BufferedWriter(fstream);
				out.write(sql);
				// Close the output stream
				out.close();
				System.out.println(OUTPUTNAME+ " was created succesfully!");
			} catch (Exception e) {// Catch exception if any
				System.err.println("Error: " + e.getMessage());
			}

		} catch (BiffException e) {
			e.printStackTrace();
		}
	}
	
	static class Column{
		public static final String TYPE_STRING = "string";
		public static final String TYPE_NUMBER = "number";
		String name;
		String type;
		public Column(String name, String type) {
			super();
			this.name = name;
			this.type = type;
		}
	}
	
	
}
