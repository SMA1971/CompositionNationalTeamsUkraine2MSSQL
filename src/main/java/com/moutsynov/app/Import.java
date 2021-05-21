package com.moutsynov.app;

import org.apache.commons.compress.archivers.sevenz.SevenZArchiveEntry;
import org.apache.commons.compress.archivers.sevenz.SevenZFile;
import org.apache.commons.io.FileUtils;
import org.apache.commons.io.FilenameUtils;

import org.apache.poi.ss.usermodel.*;

import java.io.*;
import java.net.MalformedURLException;
import java.net.URI;
import java.net.URL;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.SQLException;
import java.sql.Statement;
import java.text.SimpleDateFormat;
import java.util.*;

public class Import {

    private static final String URL7Z = "https://data.gov.ua/dataset/abd6229a-4cd8-4cc2-b6ca-793978e42b10/resource/d0eeb170-5c29-491f-8e1e-f9ab0123f188/download/spiski-zbirnikh-2020.7z";
    private static final String CONNECTION_URL = "jdbc:sqlserver://CITRIX:1433;databaseName=KOE_2;user=SMA1971;password=aB231955";

    private static ArrayList<Sportsman> sportsmans = new ArrayList<Sportsman>();

    public static void main(String[] args) {
        // Определяем временную директорию на данном компьютере
        String tmpdir = System.getProperty("java.io.tmpdir");
        // формируем полный путь к файлу архива, который будет скачан из Интернета
        String filename = tmpdir + getFilename7z();
        // формируем полный путь к директории, где будет "развернут" архив, после скачивания
        String tmpdir7z = "Import." + new SimpleDateFormat("yyyymmdd_hhmmss").format(Calendar.getInstance().getTime());

        // Если данный файл, с таким названием и по данному пути существует, то удаляем его
        File f = new File(filename);
        if (f.exists())
            f.delete();

        // скачиваем файл из Интернета
        if (downloadFileFromURL(URL7Z, filename) == true) {
            // Извлекаем файлы из архива
            decompress(filename, tmpdir + tmpdir7z);

            // подбираем файлы формата (XLS и XLSX), которые необходимо просмотреть на предмет информации
            ArrayList<File> fileList = new ArrayList<>();
            searchFiles(new File(tmpdir + tmpdir7z), fileList);

            // читаем файлы и извлекаем данные из них
            for (File file : fileList) {
                try {
                    readFromExcel(file);
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }

            // сохраняем данные в MSSQL
            saveToMssql();

            // удаляем директорию где были распакованы файлы из архива
            try {
                FileUtils.deleteDirectory(new File(tmpdir + tmpdir7z));
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

    // Определяем название файла из ссылки
    private static String getFilename7z() {
        URL url = null;
        try {
            url = new URL(URL7Z);
        } catch (MalformedURLException e) {
            e.printStackTrace();
        }
        return FilenameUtils.getName(url.getPath());
    }

    // Скачивание файла из Интеренет в указанное место
    // Входные параметры:
    //      url - ссылка на файл в Интернете
    //      filename - где будет сохранен скаченный файл
    private static boolean downloadFileFromURL(String url, String filename) {
        try (InputStream inputStream = URI.create(url).toURL().openStream()) {
            Files.copy(inputStream, Paths.get(filename));
            return true;
        } catch (MalformedURLException e) {
            e.printStackTrace();
            return false;
        } catch (IOException e) {
            e.printStackTrace();
            return false;
        }
    }

    // Извлечение файлов из архива формата 7z
    // Входные параметры:
    //      filename - полный путь к архиву
    //      dir - где распаковывать данный архив
    private static void decompress(String filename, String dir) {
        try {
            SevenZFile sevenZFile = new SevenZFile(new File(filename));
            SevenZArchiveEntry entry = sevenZFile.getNextEntry();
            while (entry != null) {
                File file = new File(dir + File.separator + entry.getName());
                if (entry.isDirectory()) {
                    if (!file.exists()) {
                        file.mkdirs();
                    }
                    entry = sevenZFile.getNextEntry();
                    continue;
                }
                if (!file.getParentFile().exists()) {
                    file.getParentFile().mkdirs();
                }
                FileOutputStream out = new FileOutputStream(file);
                byte[] content = new byte[(int) entry.getSize()];
                sevenZFile.read(content, 0, content.length);
                out.write(content);
                out.close();
                entry = sevenZFile.getNextEntry();
            }
            sevenZFile.close();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    // Поиск файлов формата (XLS и XLSX) для извлечения информации
    // Входные параметры:
    //      rootDir - корневая директория где производится поиск файлов, включая поддиректории в ней
    //      fileList - список найденных файлов, удовлетворяющий критерию поиска
    private static void searchFiles(File rootDir, List<File> fileList) {
        if (rootDir.isDirectory()) {
            File[] directoryFiles = rootDir.listFiles();
            if (directoryFiles != null) {
                for (File file : directoryFiles) {
                    if (file.isDirectory()) {
                        searchFiles(file, fileList);
                    } else {
                        if (file.getName().toLowerCase().endsWith(".xls") || file.getName().toLowerCase().endsWith(".xlsx")) {
                            fileList.add(file);
                        }
                    }
                }
            }
        }
    }

    // Открываем книгу Excel и читаем данные из всех страниц книги
    // Входные параметры:
    //    file - файл формата Excel стандарта (XLS и XLSX)
    private static void readFromExcel(File file) throws IOException {
        Workbook workbook = WorkbookFactory.create(file);
        int count = workbook.getNumberOfSheets();

        for (int i = 0; i < count; i++)
            readSheetAt(workbook, i);

        workbook.close();
    }

    // Если в ячейке Excel хранится день рождения в формате текста, то пытаемся ее преобразовать в Date.
    private static Date strToDate(String value) {
        Date result = null;

        String s = value.trim();
        s = s.replace(',', '.');
        s = s.replace(' ', '.');
        s = s.replace("..", ".");
        s = s.replace("\u00a0","");
        s = s.replace("\uFEFF","").trim();

        String[] sss = s.split("\\.");

        if ((sss.length == 3) && (s.indexOf("р.н.") == -1) && (Integer.parseInt(sss[2]) < 2021)){
            Calendar calendar = Calendar.getInstance();
            calendar.set(Integer.parseInt(sss[2]), Integer.parseInt(sss[1]), Integer.parseInt(sss[0]));
            result = calendar.getTime();
        }

        return result;
    }

    // Прочитываем данные из листа Excel книги формата (XLS и XLSX) в ArrayList<Sportsman>
    // Входные параметры:
    //    workbook - Excel книга формата
    //    index - индекс листа книги
    private static void readSheetAt(Workbook workbook, int index) {
        Sheet sheet = workbook.getSheetAt(index);
        int rowsCount = sheet.getPhysicalNumberOfRows();
        int columnFIO = -1;
        int columnBirthday = -1;
        // Определяем какие колонки на листе отвечают за ФИО и ДР
        for (int rowNum = 0; rowNum < rowsCount; rowNum++) {
            Row row = sheet.getRow(rowNum);
            if (row != null) {
                int count = row.getPhysicalNumberOfCells();
                for (int cellNum = 0; cellNum < count; cellNum++) {

                    if ((row.getCell(cellNum) != null) && (row.getCell(cellNum).getCellType() == CellType.STRING)) {
                        String s = row.getCell(cellNum).getStringCellValue();
                        if (s.toLowerCase().indexOf("прізвище") != -1)
                            columnFIO = cellNum;
                        if (s.toLowerCase().indexOf("народж") != -1)
                            columnBirthday = cellNum;
                        if ((columnFIO != -1) && (columnBirthday != -1))
                            break;
                    }
                }
            }
            if ((columnFIO != -1) && (columnBirthday != -1))
                break;
        }
        // "Пробегаем" все строки листа в поиске информации. Если информация соотв. данному алгоритму, тогда помещаем ее в ArrayList<Sportsman>
        if ((columnFIO != -1) && (columnBirthday != -1)) {
            for (int rowNum = 1; rowNum < rowsCount; rowNum++) {
                Row row = sheet.getRow(rowNum);
                if ((row != null) && (row.getCell(columnFIO) != null) && (row.getCell(columnBirthday) != null)) {
                    if ((row.getCell(columnFIO).getCellType() == CellType.STRING) && (row.getCell(columnFIO).getStringCellValue().trim() != "")) {
                        String fio = row.getCell(columnFIO).getStringCellValue().trim();
                        if ((row.getCell(columnBirthday).getCellType() == CellType.STRING)
                                && (row.getCell(columnBirthday).getStringCellValue().trim() != "")
                                && (row.getCell(columnBirthday).getStringCellValue().trim().indexOf("народж") == -1)) {

                            Date birthday = strToDate( row.getCell(columnBirthday).getStringCellValue() );
                            if (birthday != null)
                                sportsmans.add(new Sportsman(fio, birthday));
                        } else if ((row.getCell(columnBirthday).getCellType() == CellType.NUMERIC)) {
                            Date birthday = DateUtil.getJavaDate((double) row.getCell(columnBirthday).getNumericCellValue());
                            sportsmans.add(new Sportsman(fio, birthday));
                        }
                    }
                }
            }
        }
    }

    // сохраняем данные из ArrayList<Sportsman> в MSSQL сервер
    private static void saveToMssql() {
        try {
            Connection con = DriverManager.getConnection(CONNECTION_URL);
            Statement stmt = con.createStatement();
            stmt.execute("DELETE FROM sportsmans");

            for(Sportsman sportsman : sportsmans)
                stmt.execute(sportsman.getInsert());

        } catch (SQLException throwables) {
            throwables.printStackTrace();
        }
    }

    // сохраняем данные из ArrayList<Sportsman> в CSV файл
    public static void saveToCSV(String filename) {
        byte[] bom = { (byte) 0xFF, (byte) 0xFE};
        try {
            FileOutputStream fout = new FileOutputStream(filename);
            fout.write(bom);
            OutputStreamWriter out = new OutputStreamWriter(fout, "UTF-16LE");
            out.write("lastname;firstname;patronymic;birthday\r\n");

            for(Sportsman sportsman : sportsmans)
                out.write(sportsman.getInfo() + "\r\n");

            out.close();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }

    }

}
