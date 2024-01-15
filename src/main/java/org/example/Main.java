package org.example;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Scanner;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Main {
    public static void main(String[] args) throws IOException {
        List<String[]> excelTable = new ArrayList<String[]>();
        int rowIterator = 1;
        Scanner scanner = new Scanner(System.in);
        //Поменять путь к папке
        //System.out.println("Вставьте путь до файла Excel");
        //String filepath = scanner.nextLine();
        FileInputStream fis = new FileInputStream("C:\\Users\\Vasilev\\Desktop\\VasilevAD_sechenia.xlsm"); //"C:\\Users\\Vasilev\\Desktop\\Формула_Выбора_сечения_ВасильевАД.xlsm"
        XSSFWorkbook wb = new XSSFWorkbook(fis);
        XSSFSheet sheet = wb.getSheet("Лист1");
        boolean lineStillNeed = true;
        while (lineStillNeed) {
            Row row = sheet.getRow(rowIterator);
            excelTable.add(new String[5]);
            for (int i = 0; i < 5; i++) {
                Cell cell = row.getCell(i);
                cell.setCellType(Cell.CELL_TYPE_STRING);
                excelTable.get(rowIterator - 1)[i] = cell.getStringCellValue();
            }
            Row rowCheck = sheet.getRow(rowIterator + 1);
            if (rowCheck == null) {
                lineStillNeed = false;
            }
            rowIterator++;
        }
        System.out.printf("|----------------------------------------------------------------------------------------------------|\n");
        System.out.printf("| %-30s | %-4s | %-10s | %-27s | %-15s |\n", "Марка", "№п\\п", "Тип кабеля", "Ном.ток кабеля (проверка 1)", "Сечение кабеля");
        System.out.printf("|--------------------------------|------|------------|-----------------------------|-----------------|\n");
        for (String[] row : excelTable) {
            System.out.printf("| %-30s | %-4s | %-10s | %-27s | %-15s | ", row[0], row[1], row[2], row[3], row[4]);
            System.out.println();
        }
        System.out.printf("|----------------------------------------------------------------------------------------------------|\n");
        System.out.println("Значения типов кабелей считались из Экселя");

        int moshnost = 0;
        String unomFaza = "";
        double cosf = 0;
        int cabelLength = 0;
        String indent = "                      "; // 24 spaces.
        System.out.println("Взять входные значения из экселя? (Да/Нет)");
        String answer = scanner.nextLine();
        boolean isCorrentInput = false;
        if (isRightAnswer(answer)) {
            answer = "";
            moshnost = Integer.parseInt(sheet.getRow(1).getCell(7).getRawValue());
            unomFaza = sheet.getRow(2).getCell(7).getStringCellValue();
            cosf = sheet.getRow(3).getCell(7).getNumericCellValue();
            cabelLength = Integer.parseInt(sheet.getRow(5).getCell(7).getRawValue());
            System.out.println("Из экселя считались следующие значения:");
            System.out.printf("%s %-10d \n", indent.substring(0, indent.length() - "Мощность, Рном, кВт = ".length()) + "Мощность, Рном, кВт =", moshnost);
            System.out.printf("%s %-3s \n", indent.substring(0, indent.length() - "Unom,  В (Фаза) = ".length()) + "Unom,  В (Фаза) =", unomFaza);
            System.out.printf("%s %.2f \n", indent.substring(0, indent.length() - "cosf = ".length()) + "cosf =", cosf);
            System.out.printf("%s %-10d \n\n", indent.substring(0, indent.length() - "Длина кабеля, м = ".length()) + "Длина кабеля, м =", cabelLength);
            System.out.println("Вас все устраивает? (Да/Нет)");
            answer = scanner.nextLine();
            if (isRightAnswer(answer)) {
                isCorrentInput = true;
            }
        }
        if (!isCorrentInput) {
            System.out.println("Вы выбрали не брать значение из экселя, поэтому придется их ввести");
            try {
                while (!isCorrentInput) {
                    System.out.print("Мощность, Рном, кВт = ");
                    moshnost = scanner.nextInt();
                    System.out.println();
                    System.out.println("Unom,  В (Фаза)\n1 - A\n2 - B\n3 - C\n4 - ABС");
                    switch (scanner.nextInt()) {
                        case 1:
                            unomFaza = "A";
                            break;
                        case 2:
                            unomFaza = "B";
                            break;
                        case 3:
                            unomFaza = "C";
                            break;
                        case 4:
                            unomFaza = "ABC";
                            break;
                        default:
                            unomFaza = "A";
                    }
                    System.out.println();
                    System.out.print("cosf = ");
                    cosf = scanner.nextDouble();
                    System.out.println();
                    System.out.print("Длина кабеля, м = ");
                    cabelLength = scanner.nextInt();
                    System.out.println();
                    answer = "";
                    System.out.println("Вы ввели следующие значения:");
                    System.out.printf("%s %-10d \n", indent.substring(0, indent.length() - "Мощность, Рном, кВт = ".length()) + "Мощность, Рном, кВт =", moshnost);
                    System.out.printf("%s %-3s \n", indent.substring(0, indent.length() - "Unom,  В (Фаза) = ".length()) + "Unom,  В (Фаза) =", unomFaza);
                    System.out.printf("%s %.2f \n", indent.substring(0, indent.length() - "cosf = ".length()) + "cosf =", cosf);
                    System.out.printf("%s %-10d \n\n", indent.substring(0, indent.length() - "Длина кабеля, м = ".length()) + "Длина кабеля, м =", cabelLength);
                    System.out.println("Вас все устраивает? (Да/Нет)");
                    answer = scanner.nextLine();
                    answer = scanner.nextLine();
                    if (isRightAnswer(answer)) {
                        isCorrentInput = true;
                    }
                }
            } catch (Exception e) {
                System.out.println("Что-то пошло не так, попробуйте снова\n" + e.getLocalizedMessage() + "\n");
                System.exit(130);
            }
        }

        double tok_nagruzki = sila_toka(moshnost, unomFaza, cosf);
        double sechenie = sechenie4(excelTable, cabelLength, cosf, tok_nagruzki, unomFaza, 4);
        double dU = Up(unomFaza, cosf, tok_nagruzki, cabelLength, sechenie);

        System.out.println("Вычисленные значения:");
        System.out.printf("%s %.2f \n", indent.substring(0, indent.length() - "Ток нагрузки, А = ".length()) + "Ток нагрузки, А =", tok_nagruzki);
        if (sechenie != 0) {
            System.out.printf("%s %.2f \n", indent.substring(0, indent.length() - "Сечение кабеля = ".length()) + "Сечение кабеля =", sechenie);
            System.out.printf("%s %.2f \n\n", indent.substring(0, indent.length() - "dU = ".length()) + "dU =", dU);


            System.out.println("Подходят кабели под следующими марками:");
            System.out.printf("|----------------------------------------------------------------------------------------------------|\n");
            System.out.printf("| %-30s | %-4s | %-10s | %-27s | %-15s |\n", "Марка", "№п\\п", "Тип кабеля", "Ном.ток кабеля (проверка 1)", "Сечение кабеля");
            System.out.printf("|--------------------------------|------|------------|-----------------------------|-----------------|\n");
            for (String[] row : excelTable) {
                if (Double.parseDouble(row[3]) > tok_nagruzki && Double.parseDouble(row[4]) >= sechenie) {
                    System.out.printf("| %-30s | %-4s | %-10s | %-27s | %-15s | ", row[0], row[1], row[2], row[3], row[4]);
                    System.out.println();
                }
            }
            System.out.printf("|----------------------------------------------------------------------------------------------------|\n");
        } else { // Ошибка
            System.out.printf("%s \n", indent.substring(0, indent.length() - "Сечение кабеля = ".length()) + "Сечение кабеля = -");
            System.out.printf("%s \n\n", indent.substring(0, indent.length() - "dU = ".length()) + "dU = -");
        }

    }

    public static double sila_toka(int moshnost, String unomFaza, double cosf) {
        switch (unomFaza) {
            case "A":
            case "B":
            case "C":
                return 1000 * moshnost / 220.0 / cosf;
            case "ABC":
                return 1000 * moshnost / 380.0 / cosf / 1.73;
        }
        return 0;
    }

    public static double Up(String unomFaza, double cosf, double tok_nagruzki, int cabelLength, double sechenie) {
        switch (unomFaza) {
            case "A":
            case "B":
            case "C":
                return 2 * (0.0225 * cabelLength * cosf / sechenie + 0.00008 * cabelLength * Math.sqrt(1 - Math.pow(cosf, 2))) * tok_nagruzki * 100 / 220.0;
            case "ABC":
                return 1 * (0.0225 * cabelLength * cosf / sechenie + 0.00008 * cabelLength * Math.sqrt(1 - Math.pow(cosf, 2))) * tok_nagruzki * 100 / 380.0;
        }
        return 0;
    }

    public static double sechenie4(List<String[]> excelTable, int cabelLength, double cosf, double tok_nagruzki, String unomFaza, double dUnom) {
        double dU = 0;
        int i = 0;
        double Inom = 0;
        double sechenie;
        do {
            if (excelTable.size() == i) {
                if (dU > dUnom && tok_nagruzki > Inom) {
                    System.out.println("Невозможно подобрать кабель по сечению и току");
                } else if (dU > dUnom) {
                    System.out.println("Невозможно подобрать кабель по сечению");
                } else if (tok_nagruzki > Inom) {
                    System.out.println("Невозможно подобрать кабель по току");
                }
                sechenie = 0;
                break;
            }
            sechenie = Double.parseDouble(excelTable.get(i)[4]);
            Inom = Double.parseDouble(excelTable.get(i)[3]);
            i++;
            dU = Up(unomFaza, cosf, tok_nagruzki, cabelLength, sechenie);
        } while (dU >= dUnom || tok_nagruzki >= Inom);
        return sechenie;
    }

    public static boolean isRightAnswer(String answer) {
        return answer.equals("Да") || answer.equals("да") || answer.equals("д") || answer.equals("Д");
    }

}