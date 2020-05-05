import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.CellType;

import java.io.*;

public class Main {
    public static void main(String[] args) throws IOException {
        BufferedReader reader = new BufferedReader(new InputStreamReader(System.in));

        System.out.println("Provide path to excel file");
        String path = reader.readLine();
        HSSFWorkbook hssfWorkbook = new HSSFWorkbook(new FileInputStream(path));

        System.out.println("Choose sheet, starts from 0");
        int sheetNum = Integer.parseInt(reader.readLine());
        HSSFSheet hssfSheet = hssfWorkbook.getSheetAt(sheetNum);

        System.out.println("Choose column");
        int columnName = Integer.parseInt(reader.readLine());
        addBreak(hssfSheet, columnName);

//        deleteAllNonNumeric(hssfSheet, 0);

        FileOutputStream fileOutputStream = new FileOutputStream("result.xls");
        hssfWorkbook.write(fileOutputStream);
        fileOutputStream.close();

        hssfWorkbook.close();
    }

    public static void addBreak(HSSFSheet sheet, int cellNumber) {
        int counter = 0;
        HSSFRow row = sheet.getRow(counter);
        while (row != null && !isCellEmpty(row.getCell(cellNumber))) {
            if (row.getCell(cellNumber).getCellType().equals(CellType.STRING)) {
                HSSFCell cell = row.getCell(cellNumber);
                String value = cell.getStringCellValue().trim();

                cell.setCellValue(addDash(value));
            } else if (row.getCell(cellNumber).getCellType().equals(CellType.NUMERIC)) {
                HSSFCell cell = row.getCell(cellNumber);
                String value = Double.valueOf(cell.getNumericCellValue()).toString();
                if (value.endsWith(".0")) {
                    value = value.substring(0, value.length() - 2);
                }

                cell.setCellValue(addDash(value));
            }
            row = sheet.getRow(++counter);
        }
    }

    private static String addDash(String value) {
        if (value.length() == 1) {
            System.out.println(value);
            return value;
        } else if (value.length() % 2 == 0) {
            int start = 0;
            StringBuilder sb = new StringBuilder();
            for (int i = 2; i <= value.length(); i += 2) {
                if (i != 2) {
                    sb.append("-");
                }
                sb.append(value, start, i);
                start = i;
            }
            System.out.println(value + " " + sb.toString());
            return sb.toString();
        } else {
            int start = 0;
            StringBuilder sb = new StringBuilder();
            for (int i = 2; i <= value.length(); i += 2) {
                if (i != 2) {
                    sb.append("-");
                }
                sb.append(value, start, i);
                start = i;
            }
            sb.append("-");
            sb.append(value.substring(value.length() - 1));
            System.out.println(value + " " + sb.toString());
            return sb.toString();
        }
    }

    public static void deleteAllNonNumeric(HSSFSheet sheet, int cellNumber) {
        int counter = 0;
        HSSFRow row = sheet.getRow(counter);
        while (row != null && !isCellEmpty(row.getCell(cellNumber))) {
            if (row.getCell(cellNumber).getCellType().equals(CellType.STRING)) {
                HSSFCell cell = row.getCell(cellNumber);
                String value = cell.getStringCellValue();
                String digits = value.replaceAll("[^\\d.]", "");
                cell.setCellValue(digits);
                System.out.println(value + " " + digits);
            }
            row = sheet.getRow(++counter);
            System.out.println(counter);
        }
    }

    public static boolean isCellEmpty(final HSSFCell cell) {
        if (cell == null) { // use row.getCell(x, Row.CREATE_NULL_AS_BLANK) to avoid null cells
            return true;
        }

        if (cell.getCellType().equals(CellType.BLANK)) {
            return true;
        }

        if (cell.getCellType().equals(CellType.STRING) && cell.getStringCellValue().trim().isEmpty()) {
            return true;
        }

        return false;
    }
}
