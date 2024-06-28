package com.example.demo.helpers.workload;

import com.example.demo.models.Employee;
import com.example.demo.models.WorkloadQuery;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.*;

import java.io.*;
import java.lang.reflect.Field;
import java.util.HashMap;
import java.util.Map;
import java.util.ArrayList;
import java.util.List;

public class WorkloadExcelProcessor {
    public static Map<String, List<WorkloadQuery>> disassembleData(List<WorkloadQuery> workloadQueries, List<Employee> employees) {
        Map<String, List<WorkloadQuery>> resultMap = new HashMap<>();

        for (WorkloadQuery data : workloadQueries) {
            String teacher = findById(employees, data.getTeacher());
            // Получаем или создаем список для данного года
            List<WorkloadQuery> list = resultMap.computeIfAbsent(teacher, k -> new ArrayList<WorkloadQuery>());

            // Добавляем новый объект WorkloadQuery в список
            list.add(data);
        }
        return resultMap;
    }

    private static String findById(List<Employee> employees, Integer id) {
        for (Employee employee : employees) {
            if (employee.getId() == id) {
                return employee.getEmployeeName();
            }
        }
        return "";
    }

    public static Workbook createReportWorkload(Map<String, List<WorkloadQuery>> mapWorkLoad, List<WorkloadQuery> data) throws IOException {
        // Получение полного пути к файлу шаблона
        File templateFile = new File("templates/WorkloadTemplate.xlsx");
        return createReport(templateFile, mapWorkLoad, data);
    }

    private static Workbook createReport(File templateFile, Map<String, List<WorkloadQuery>> mapWorkLoad, List<WorkloadQuery> allData) throws IOException {
        try {
            // Загрузка шаблонного файла Excel
            FileInputStream fis = new FileInputStream(templateFile);
            Workbook workbook = WorkbookFactory.create(fis); // Автоматическое определение формата (xls или xlsx)
            fis.close();

            // Создание новых листов на основе списка
            for (Map.Entry<String, List<WorkloadQuery>> data : mapWorkLoad.entrySet()) {
                String sheetName = data.getKey();
                List<WorkloadQuery> workloadQueries = data.getValue();

                // Создаем копию листа с шаблоном и переименовываем ее по имени преподавателя (по id)
                Sheet teacherSheet = workbook.cloneSheet(0);
                workbook.setSheetName(workbook.getSheetIndex(teacherSheet), sheetName);

                // Заменяем данные на листе копии
                replaceDataInSheet(teacherSheet, workloadQueries, sheetName);
            }
            //вставляем список всех бакалавров
            Sheet BSheet = workbook.cloneSheet(0);
            workbook.setSheetName(workbook.getSheetIndex(BSheet), "Бакалавры");
            replaceAllDataInSheet(BSheet, allData, "Бакалавры", true);

            //вставляем список всех магистров
            Sheet MSheet = workbook.cloneSheet(0);
            workbook.setSheetName(workbook.getSheetIndex(MSheet), "Магистры");
            replaceAllDataInSheet(MSheet, allData, "Магистры", false);

            workbook.removeSheetAt(0);
            workbook.getCreationHelper().createFormulaEvaluator().evaluateAll();
            return workbook;
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } catch (Exception e) {
            e.printStackTrace();
        }
        return null;
    }

    private static void replaceAllDataInSheet(Sheet teacherSheet, List<WorkloadQuery> workloadQueries, String sheetName, Boolean isBac) throws NoSuchFieldException, IllegalAccessException {
        int numRow = 5;
        int count = 1;
        Map<Integer, String> cells = createCells();
        for (WorkloadQuery res : workloadQueries) {
            if ((res.getSemesterDescr() < 9) == isBac) {
                Row row = teacherSheet.getRow(numRow);
                if (row != null) {
                    for (Map.Entry<Integer, String> currCell : cells.entrySet()) {
                        Integer numCell = currCell.getKey();
                        String valCell = currCell.getValue();
                        Cell cell = row.getCell(numCell);
                        if (cell != null) {
                            if (numCell == 0) {
                                cell.setCellValue(count);
                                continue;
                            }
                            writeValueInCell(cell, valCell, res);
                        }
                    }
                    Cell cellGroup = row.getCell(8);
                    if (cellGroup == null) {
                        cellGroup = row.createCell(8);
                    }
                    cellGroup.setCellValue(1);
                    if (res.getPrPr() > 0 || res.getPredDipPr() > 0 || res.getUchPr() > 0) {
                        Cell weekCount = row.getCell(7);
                        if (weekCount == null) {
                            weekCount = row.createCell(7);
                        }
                        if (res.getQualificationDescr().contains("B")) {
                            weekCount.setCellValue(6);
                        } else {
                            if(res.getUchPr() > 0){
                                weekCount.setCellValue(6);
                            }
                            else {
                                weekCount.setCellValue(4);
                            }
                        }
                    }
                }
                numRow++;
                count++;
            }
        }
        Row rowName = teacherSheet.getRow(3);
        Cell cellName = rowName.getCell(57);
        cellName.setCellValue(sheetName);

        Row rowYear = teacherSheet.getRow(0);
        Cell cellYear = rowYear.getCell(16);
        Integer year = workloadQueries.get(0).getStudyYear();
        Integer nextYear = year + 1;
        cellYear.setCellValue(year + "/" + nextYear);
    }

    private static void replaceDataInSheet(Sheet teacherSheet, List<WorkloadQuery> workloadQueries, String sheetName) throws NoSuchFieldException, IllegalAccessException {
        int numRow = 5;
        int count = 1;
        Map<Integer, String> cells = createCells();
        for (WorkloadQuery res : workloadQueries) {
            Row row = teacherSheet.getRow(numRow);
            if (row != null) {
                for (Map.Entry<Integer, String> currCell : cells.entrySet()) {
                    Integer numCell = currCell.getKey();
                    String valCell = currCell.getValue();
                    Cell cell = row.getCell(numCell);
                    if (cell != null) {
                        if (numCell == 0) {
                            cell.setCellValue(count);
                            continue;
                        }
                        writeValueInCell(cell, valCell, res);
                    }
                }
                Cell cellGroup = row.getCell(8);
                if (cellGroup == null) {
                    cellGroup = row.createCell(8);
                }
                cellGroup.setCellValue(1);
                if (res.getPrPr() > 0 || res.getPredDipPr() > 0 || res.getUchPr() > 0) {
                    Cell weekCount = row.getCell(7);
                    if (weekCount == null) {
                        weekCount = row.createCell(7);
                    }
                    if (res.getQualificationDescr().contains("B")) {
                        weekCount.setCellValue(6);
                    } else {
                        weekCount.setCellValue(4);
                    }
                }
            }
            numRow++;
            count++;
        }
        Row rowName = teacherSheet.getRow(3);
        Cell cellName = rowName.getCell(57);
        cellName.setCellValue(sheetName);

        Row rowYear = teacherSheet.getRow(0);
        Cell cellYear = rowYear.getCell(16);
        Integer year = workloadQueries.get(0).getStudyYear();
        Integer nextYear = year + 1;
        cellYear.setCellValue(year + "/" + nextYear);
    }

    private static Map<Integer, String> createCells() {
        Map<Integer, String> res = new HashMap<>();
        res.put(0, "#");
        res.put(1, "descr");
        res.put(2, "specialityDescr");
        res.put(3, "facultyDescr");
        res.put(4, "studyForm");
        res.put(5, "semesterDescr");
        res.put(6, "studentCount");
        res.put(7, "weekCount");
//        res.put(8,"групп"); = 1
        res.put(9, "subgroupCount");
        res.put(10, "lectureCount");
        res.put(11, "practiceCount");
        res.put(12, "labCount");
        res.put(13, "ekz");
        res.put(14, "zach");
        res.put(15, "rgr");
        res.put(16, "kr");
        res.put(17, "kp");
        res.put(18, "uchPr");
        res.put(19, "prPr");
        res.put(20, "predDipPr");
        res.put(21, "konsZaoch");
        res.put(22, "gek");
        res.put(23, "gak");
        res.put(24, "gakPred");
        res.put(25, "dpRuk");
        res.put(28, "dpRetz");
        res.put(29, "aspRuk");
        res.put(30, "magRuk");
        res.put(31, "magRetz");
        res.put(32, "rukKaf");
        return res;
    }

    private static void writeValueInCell(Cell cell, String valCell, WorkloadQuery row) throws IllegalAccessException, NoSuchFieldException {
        // Получаем класс объекта
        Class<?> clazz = row.getClass();
        // Получаем поле по его имени
        Field field = clazz.getDeclaredField(valCell);
        // Разрешаем доступ к полю, если оно private или protected
        field.setAccessible(true);
        // Получаем значение поля из объекта
        Object valueFromObject = field.get(row);
        //Приводим объект к нужному типу и записываем в ячейку
        if (valueFromObject instanceof String) {
            String str = valueFromObject.toString();
            if (str.equals("Очно")) {
                cell.setCellValue("о");
            } else if (str.equals("Заочно")) {
                cell.setCellValue("з");
            } else if (str.equals("Очно-заочно")) {
                cell.setCellValue("оз");
            } else {
                cell.setCellValue(str.replaceAll("\"", ""));
            }
        } else if (valueFromObject instanceof Integer) {
            cell.setCellValue((Integer) valueFromObject);
        } else if (valueFromObject instanceof Boolean) {
            boolean bool = (boolean) valueFromObject;
            if (bool) {
                cell.setCellValue("+");
            }
        }
    }
}

