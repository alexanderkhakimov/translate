package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class ExcelTranslator {
    public static void main(String[] args) throws FileNotFoundException {
//        Scanner scanner = new Scanner(System.in);
//        System.out.print("Введите путь к папке с файлами в формате Excel: ");
//        String src = scanner.nextLine();
//        System.out.print("Введите путь к папке, где будут сохранены файлы: ");
//        String dest = scanner.nextLine();
//        System.out.print("Введите путь к словарю для перевода: ");
//        String excelDirectionary = scanner.nextLine();
        String excelDirectionary="C:\\Users\\SurfaceBook\\Desktop\\translate\\dict.xlsx";
                String dest ="C:\\Users\\SurfaceBook\\Desktop\\translate\\dest";
                String src ="C:\\Users\\SurfaceBook\\Desktop\\translate\\src";

        List<File> listFile = readFilesFromDir(new File(src), ".xlsx");
        Map<String, String> exDic = readDictionary(excelDirectionary);
        exDic.put("тУ","TU");
        translate(listFile, exDic, dest);


    }

    private static Map<String, String> readDictionary(String excelDirectionary) throws FileNotFoundException {
        Map<String, String> map = new HashMap<>();
        try (FileInputStream fis = new FileInputStream(excelDirectionary);
             Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheetAt(0);
            Iterator<Row> rowIterator = sheet.rowIterator();
            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();
                Cell keyCell = row.getCell(0);
                Cell valueCell = row.getCell(1);
                if (keyCell != null && valueCell != null) {
                    String key =keyCell.getStringCellValue();
                    if(!(key.equals("ТУ") || key.equals("ГОСТ")|| key.equals("ОСТ")||
                            key.equals("Ц-ОСТ")|| key.equals("Ц.фос-ОСТ")|| key.equals("Ц.фос.окс-ОСТ")) ){
                        key = key.toLowerCase();
                    }
                    String value = valueCell.getStringCellValue();
                    map.put(key, value);
                }
            }

        } catch (IOException e) {
            throw new RuntimeException(e);
        }
        return map;
    }

    private static void translate(List<File> listFile, Map<String, String> excelDirectionary, String dest) {
        for (File file : listFile) {
            String nameFIle = file.getName();
            try (FileInputStream fis = new FileInputStream(file);
                 Workbook workbook = new XSSFWorkbook(fis)) {

                Sheet sheet = workbook.getSheetAt(0); // Получаем первый лист

                for (Row row : sheet) {
                    for (Cell cell : row) {
                        if (cell.getCellType() == CellType.STRING) {
                            String cellValue = cell.getStringCellValue();
                            String translatedValue = translateText(cellValue, excelDirectionary);
                            if (translatedValue.length() > 0) {
                                translatedValue = translatedValue.substring(0, 1).toUpperCase() + translatedValue.substring(1);
                            }
                            cell.setCellValue(translatedValue);
                        }
                    }
                }

                // Сохраните изменения в файл
                try (FileOutputStream fos = new FileOutputStream(dest + "\\" + nameFIle)) {
                    workbook.write(fos);
                }

            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

    private static String translateText(String text, Map<String, String> dictionary) {
        String temp = text;
        if (temp.length() > 0) {
            temp = temp.substring(0, 1).toLowerCase() + temp.substring(1);
        }
        for (Map.Entry<String, String> entry : dictionary.entrySet()) {
            String key = entry.getKey();
            Pattern pattern = Pattern.compile(key, Pattern.CASE_INSENSITIVE);
            Matcher matcher = pattern.matcher(temp);
            while (matcher.find()) {
                String translateWord = temp.replaceAll(key, entry.getValue());
                return translateText(translateWord, dictionary);
            }
        }
        return temp;
    }


    private static List<File> readFilesFromDir(File excelDirectory, String ext) {
        List<File> files = new ArrayList<>();
        for (File f : Objects.requireNonNull(excelDirectory.listFiles())) {
            if (f.isFile() && f.getName().toLowerCase().endsWith(ext)) files.add(f);
            else if (f.isDirectory()) {
                List<File> dirFiles = readFilesFromDir(f, ext);
                files.addAll(dirFiles);
            }
        }
        return files;

    }
}