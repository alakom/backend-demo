package com.example.demo.helpers.example;

import com.example.demo.models.WorkloadQuery;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.List;
import java.util.Map;

public class ExampleExcelProcessor {
    public static Workbook getExampleFile() throws IOException {
        // Получение полного пути к файлу шаблона
        File templateFile = new File("templates/Example.xlsx");
        return createReport(templateFile);
    }
    private static Workbook createReport(File templateFile) throws IOException {
        try {
            // Загрузка шаблонного файла Excel
            FileInputStream fis = new FileInputStream(templateFile);
            Workbook workbook = WorkbookFactory.create(fis); // Автоматическое определение формата (xls или xlsx)
            fis.close();
            return workbook;
        }catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } catch (Exception e) {
            e.printStackTrace();
        }
        return null;
    }
}
