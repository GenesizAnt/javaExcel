package com.mypackage;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFCell;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.List;
import java.util.Scanner;

public class Main implements AutoFilter {

    static final String FORMAT_XLSX = ".xlsx";

    public static void main(String[] args) throws IOException {

        Scanner scanner = new Scanner(System.in);
        System.out.println("Укажите наименование файла из которого копируете (строго только название!)");
        String nameFileIn = scanner.nextLine();
        System.out.println("Укажите наименование файла в который копируете (строго только название!)");
        String nameFileOut = scanner.nextLine();
        System.out.println("Укажи код Партнера - только 5 цифр! Если первые нули не указывай их");
        String codePartner = scanner.nextLine();

        System.out.println("Данные получены! Хоббиты уже делают перепись...");

        FileInputStream streamIn = new FileInputStream(nameFileIn + FORMAT_XLSX);
        XSSFWorkbook inPutFileName = new XSSFWorkbook(streamIn);

        FileOutputStream streamOut = new FileOutputStream(nameFileOut + FORMAT_XLSX);
        XSSFWorkbook outPutFileName = new XSSFWorkbook();
        XSSFSheet sheetPartner = outPutFileName.createSheet("Лист1");

        try {
            copyTitleSub(inPutFileName, outPutFileName);
            copyTableSub(inPutFileName, outPutFileName, codePartner);
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            System.out.println("Вроде все переписали, можно проверять");
        }

        outPutFileName.write(streamOut);
        streamOut.close();

        streamIn.close();

    }

    public static void copyTitleSub(XSSFWorkbook inPutFileName, XSSFWorkbook outPutFileName) throws IOException {
        XSSFSheet sheetSub = inPutFileName.getSheetAt(0);
        XSSFRow row0 = sheetSub.getRow(0);
        List<String> copyTitle = new ArrayList<>();
        for (int i = 0; i < 8; i++) {
            XSSFCell cells = row0.getCell(i);
            copyTitle.add(cells.getStringCellValue());
        }

        XSSFSheet sheetPartner = outPutFileName.getSheetAt(0);
        XSSFRow row1 = sheetPartner.createRow(0);

        for (int i = 0; i < 8; i++) {
            XSSFCell cells = row1.createCell(i);
            cells.setCellValue(String.valueOf(copyTitle.get(i)));
        }

    }

    public static void copyTableSub(XSSFWorkbook inPutFileName, XSSFWorkbook outPutFileName, String codePartner) throws IOException {
        XSSFSheet sheetSub = inPutFileName.getSheetAt(0);

        int newFileRow = 1;
        int oldFileRow = 1;

        XSSFRow row0 = sheetSub.getRow(oldFileRow);
        XSSFCell cell0 = row0.getCell(1);
        XSSFSheet sheetPartner = outPutFileName.getSheetAt(0);

        while (!getCellText(cell0).equals("")) {
            if (getCellText(cell0).equals(codePartner + ".0")) {
                XSSFRow row1 = sheetPartner.createRow(newFileRow++);
                for (int i = 0; i < 8; i++) {
                    XSSFCell cell1 = row1.createCell(i);
                    if (i <= 5) {
                        cell1.setCellValue(getCellText(row0.getCell(i)).replace(".0", ""));
                    } else {
                        cell1.setCellValue(getCellText(row0.getCell(i)));
                    }
                }
            }
            row0 = sheetSub.getRow(++oldFileRow);
            if (row0 == null) {
                break;
            }
            cell0 = row0.getCell(1);
        }

    }

    public static String getCellText(Cell cell) {

        String result = "";

        switch (cell.getCellType()) {
            case STRING:
                result = cell.getRichStringCellValue().getString();
                break;
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    result = sdf.format(cell.getDateCellValue());
                } else {
                    result = String.valueOf(cell.getNumericCellValue());
                }
                break;
            case BOOLEAN:
                result = String.valueOf(cell.getBooleanCellValue());
                break;
            case FORMULA:
                result = cell.getCellFormula();
                break;
            case BLANK:
                System.out.println();
                break;
            default:
                break;
        }
        return result;
    }

    public static SimpleDateFormat sdf = new SimpleDateFormat("dd.MM.yyyy");
}


