package com.medication.tracker;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.time.format.DateTimeParseException;
import java.time.format.ResolverStyle;
import java.util.ArrayList;
import java.util.List;
import java.util.Scanner;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class MedicationTracker {

    private static final String CUSTOMER_DETAILS = "CustomerDetails";
    private static final String MEDICINES = "Medicines";
    private static final String NAME = "Name";
    private static final String DATE = "Date";

    public static void main(String[] args) {

        Scanner in = new Scanner(System.in);
        System.out.println("Enter your name");
        String userName = in.next();
        MedicationTracker medicationTracker = new MedicationTracker();
        medicationTracker.customerOptions(in, userName);

    }

    private void customerOptions(Scanner in, String login) {
        int totalAttempts = 3;

        while (totalAttempts != 0) {
            System.out.println("Please choose from the following.");
            System.out.println("1. To Enter your medication details");
            System.out.println("2. Download your medical history");
            while (!in.hasNextInt()) {
                if (totalAttempts == 0) {
                    maximumAttempt(in);
                }
                System.out.println("Please enter a number");
                in.next();
                totalAttempts--;
            }
            int value = in.nextInt();
            if (value == 1) {
                captureCustomerDetails(in, login);
            } else if (value == 2) {
                createCustomerFile(in, login);
            } else {
                System.out.println("Invalid number. Please enter either 1 or 2 ");
                totalAttempts--;
            }
        }
        if (totalAttempts == 0) {
            maximumAttempt(in);
        }

    }

    private void createCustomerFile(Scanner in, String login) {

        createFileIfNotPresent(login + ".xls");
        String fileName = login + ".xls";
        String customerFileName = "CustomerMedication.xls";

        FileInputStream inputStream;
        try {
            File customerFile = new File(customerFileName);
            if (!customerFile.exists()) {
                noRecords(in);
            }

            inputStream = new FileInputStream(customerFile);

            Workbook workbook = WorkbookFactory.create(inputStream);
            Sheet sheet = workbook.getSheet(CUSTOMER_DETAILS);
            int lastRowNum = 0;
            List<Integer> matchingRows = findMatchingRowsForCustomer(sheet, login);
            if (matchingRows.isEmpty()) {
                noRecords(in);
            } else {
                List<Row> customerRows = getCustomerRows(matchingRows, sheet);
                Workbook workbook1 = null;
                FileOutputStream fileOut = null;
                for (Row row : customerRows) {
                    lastRowNum++;
                    FileInputStream inputStream1 = new FileInputStream(new File(fileName));
                    workbook1 = WorkbookFactory.create(inputStream1);
                    Sheet sheet1 = workbook1.getSheet(CUSTOMER_DETAILS);
                    HSSFRow rowhead = (HSSFRow) sheet1.createRow((short) lastRowNum);
                    rowhead.createCell(0).setCellValue(row.getCell(0).getRichStringCellValue().getString().trim());
                    rowhead.createCell(1).setCellValue(row.getCell(1).getRichStringCellValue().getString().trim());
                    rowhead.createCell(2).setCellValue(row.getCell(2).getRichStringCellValue().getString().trim());
                    fileOut = new FileOutputStream(fileName);
                    workbook1.write(fileOut);
                }

                fileOut.close();
                System.out.println("File created. Have a nice day!!!");
                in.close();
                System.exit(0);

            }
            System.out.println("Thanks for entering the details. Have a nice day !!!");
            in.close();
            System.exit(0);
        } catch (FileNotFoundException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        } catch (EncryptedDocumentException | IOException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }

    }

    private void noRecords(Scanner in) {
        System.out.println("No records found for the customer");
        in.close();
        System.exit(0);
    }

    private void maximumAttempt(Scanner in) {
        System.out.println("Maximum attempts reached. Please try again later");
        in.close();
        System.exit(0);
    }

    private void captureCustomerDetails(Scanner in, String login) {
        int totalAttempts = 3;

        while (totalAttempts != 0) {
            System.out.println("Please choose from the following.");
            System.out.println("1. Do you want to enter details for today");
            System.out.println("2. Enter details for a previous day");
            while (!in.hasNextInt()) {
                if (totalAttempts == 0) {
                    maximumAttempt(in);
                }
                System.out.println("Please enter a number");
                in.next();
                totalAttempts--;
            }
            int value = in.nextInt();
            if (value == 1) {
                captureCustomerInformation(in, login, LocalDate.now());
            } else if (value == 2) {
                captureCustomerInformationSomeDay(in, login);
            } else {
                System.out.println("Invalid number. Please enter either 1 or 2 ");
                totalAttempts--;
            }
        }
        if (totalAttempts == 0) {
            maximumAttempt(in);
        }

    }

    private void captureCustomerInformationSomeDay(Scanner in, String login) {
        int totalAttempts = 3;

        while (totalAttempts != 0) {
            System.out.println("Please ennter which day you want to enter details in yyyy-mm-dd format");
            String date = in.next();
            boolean isValidDate = isValidDate(date);
            if (isValidDate) {
                captureCustomerInformation(in, login, LocalDate.parse(date));
            } else {
                System.out.println("Invalid Date. Please enter correct Date ");
                totalAttempts--;
            }
        }
        if (totalAttempts == 0) {
            maximumAttempt(in);
        }

    }

    private boolean isValidDate(String date) {
        try {
            LocalDate.parse(date, DateTimeFormatter.ofPattern("uuuu-M-d").withResolverStyle(ResolverStyle.STRICT));
            return true;

        } catch (DateTimeParseException e) {
            return false;
        }
    }

    private void captureCustomerInformation(Scanner in, String login, LocalDate date) {

        String fileName = "CustomerMedication.xls";
        createFileIfNotPresent(fileName);
        FileInputStream inputStream;
        try {
            inputStream = new FileInputStream(new File(fileName));
            Workbook workbook = WorkbookFactory.create(inputStream);
            Sheet sheet = workbook.getSheet(CUSTOMER_DETAILS);
            int lastRowNum = sheet.getLastRowNum();

            DateTimeFormatter formatter = DateTimeFormatter.ofPattern("dd-MM-yyyy");
            List<Integer> matchingRows = findMatchingRowsForCustomer(sheet, login);
            if (matchingRows.isEmpty()) {
                System.out.println("Please enter the medicines with ',' in between");
                String medicines = in.next();
                createMedicineEntry(login, fileName, workbook, sheet, lastRowNum, date, formatter, medicines);
            } else {
                List<Row> customerRows = getCustomerRows(matchingRows, sheet);
                Row rowForToday = getRowForToday(customerRows, date.format(formatter));
                if (rowForToday == null) {
                    System.out.println("Please enter the medicines with ',' in between");
                    String medicines = in.next();
                    createMedicineEntry(login, fileName, workbook, sheet, lastRowNum, date, formatter, medicines);
                } else {
                    System.out.println("This will override the already entered value");
                    System.out.println("Please enter the medicines with ',' in between");
                    String medicines = in.next();
                    createMedicineEntry(login, fileName, workbook, sheet, rowForToday.getRowNum() - 1, date, formatter,
                            medicines);
                }
            }
            System.out.println("Thanks for entering the details. Have a nice day !!!");
            in.close();
            System.exit(0);
        } catch (FileNotFoundException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        } catch (EncryptedDocumentException | IOException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }

    }

    private void createMedicineEntry(String login, String fileName, Workbook workbook, Sheet sheet, int lastRowNum,
            LocalDate date, DateTimeFormatter formatter, String medicines) throws FileNotFoundException, IOException {
        HSSFRow rowhead = (HSSFRow) sheet.createRow((short) ++lastRowNum);
        rowhead.createCell(0).setCellValue(date.format(formatter));
        rowhead.createCell(1).setCellValue(login);
        rowhead.createCell(2).setCellValue(medicines);
        FileOutputStream fileOut = new FileOutputStream(fileName);
        workbook.write(fileOut);
        fileOut.close();
        workbook.close();
    }

    private Row getRowForToday(List<Row> customerRows, String date) {
        for (Row row : customerRows) {
            if (row.getCell(0).getRichStringCellValue().getString().trim().contains(date)) {
                return row;
            }
        }
        return null;

    }

    private List<Row> getCustomerRows(List<Integer> matchingRows, Sheet sheet) {
        List<Row> customerRows = new ArrayList<Row>();
        for (Integer rowNum : matchingRows) {
            customerRows.add(sheet.getRow(rowNum));
        }
        return customerRows;
    }

    public List<Integer> findMatchingRowsForCustomer(Sheet sheet, String cellContent) {
        List<Integer> matchedRows = new ArrayList<>();
        for (Row row : sheet) {
            for (Cell cell : row) {
                if (cell.getCellType() == CellType.STRING) {
                    if (cell.getRichStringCellValue().getString().trim().equalsIgnoreCase(cellContent)) {
                        matchedRows.add(row.getRowNum());
                    }
                }
            }
        }
        return matchedRows;
    }

    private void createFileIfNotPresent(String fileName) {
        File customerFile = new File(fileName);
        if (customerFile.exists()) {
            return;
        } else {
            try {
                HSSFWorkbook workbook = new HSSFWorkbook();
                HSSFSheet sheet = workbook.createSheet(CUSTOMER_DETAILS);

                HSSFRow rowhead = sheet.createRow((short) 0);
                rowhead.createCell(0).setCellValue(DATE);
                rowhead.createCell(1).setCellValue(NAME);
                rowhead.createCell(2).setCellValue(MEDICINES);
                FileOutputStream fileOut = new FileOutputStream(fileName);
                workbook.write(fileOut);
                fileOut.close();
                workbook.close();
                return;
            } catch (IOException e) {
                // TODO Auto-generated catch block
                e.printStackTrace();
            }

        }

    }

}
