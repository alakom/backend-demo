    package com.example.demo.helpers.semester;

    import com.example.demo.models.WorkloadQuery;
    import org.apache.poi.ss.usermodel.*;

    import java.io.File;
    import java.io.FileInputStream;
    import java.io.FileNotFoundException;
    import java.io.IOException;
    import java.util.List;

    import static com.example.demo.helpers.WorkbookHelpers.*;

    public class SemesterExcelProcessor {

        static double sumPlan = 0;
        static double sumDogovor = 0;

        public static Workbook createSemesterReport(List<WorkloadQuery> workloadQueries, Boolean isAutumn) {
            sumDogovor = 0;
            sumPlan = 0;
            // Получение полного пути к файлу шаблона
            File templateFile = new File("templates/SemesterTemplate.xlsx");
            return createReport(templateFile, workloadQueries, isAutumn);
        }

        private static Workbook createReport(File templateFile, List<WorkloadQuery> workload, Boolean isAutumn) {
            try {
                // Загрузка шаблонного файла Excel
                FileInputStream fis = new FileInputStream(templateFile);
                Workbook workbook = WorkbookFactory.create(fis); // Автоматическое определение формата (xls или xlsx)
                fis.close();
                // Создание списков с бакалаврами и магистрами
                List<WorkloadQuery> bachelors = selectionBachelors(workload, isAutumn);
                List<WorkloadQuery> magisters = selectionMagisters(workload, isAutumn);
                //Создание нового листа
                Sheet resultSheet = workbook.cloneSheet(0);
                String sheetName = generateSheetName(workload, isAutumn);
                workbook.setSheetName(workbook.getSheetIndex(resultSheet), sheetName);
                // Заполнение шаблонных фраз
                GenerateTitle(resultSheet, workload, isAutumn);
                //Заполнение магистров
                writeWorkloadSemester(resultSheet, magisters, 31);
                //Заполнение бакалавров
                writeWorkloadSemester(resultSheet, bachelors, 8);
                //Добавление итоговых сумм
                setCellSheet(resultSheet, 35 + bachelors.size() + magisters.size(), 4, sumPlan);
                setCellSheet(resultSheet, 36 + bachelors.size() + magisters.size(), 4, sumDogovor);
                //Удаление листа-шаблона
                workbook.removeSheetAt(0);
                //Активация формул
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

        private static String generateSheetName(List<WorkloadQuery> workload, Boolean isAutumn) {
            Integer year = workload.get(0).getStudyYear();
            return (isAutumn ? "1 семестер " : "2 семестер ") + year + "-" + (year + 1);
        }

        private static void GenerateTitle(Sheet sheet, List<WorkloadQuery> workload, Boolean isAutumn) {
            Integer year = workload.get(0).getStudyYear();
            String yearValue = year + " / " + (year + 1);
            String semesterValue = isAutumn ? "Осенний" : "Весенний";
            String strValue = "экзаменационной сессии  за  " + semesterValue + " семестр " + yearValue + " учебного года";
            String strValue2 = "Итого по кафедре ПО за " + semesterValue + " семестр " + yearValue + " уч. год  = ";
            setCellSheet(sheet, 2, 0, strValue);
            setCellSheet(sheet, 34, 1, strValue2);

        }

        private static void writeWorkloadSemester(Sheet sheet, List<WorkloadQuery> workload, int startRow) {
            int numRow = startRow, count = 1;
            for (WorkloadQuery data : workload) {
                if (numRow != startRow) {
                    sheet.shiftRows(numRow, sheet.getLastRowNum(), 1, true, true); // Сдвиг строк ниже
                    Row newRow = sheet.getRow(numRow);
                    if (newRow == null) {
                        newRow = sheet.createRow(numRow);
                    }
                    newRow.setHeightInPoints(45f);
                }
                List<Double> LaborCosts = CalculationLaborCosts(data);
                setCellSheetWithStyle(sheet, numRow, 0, count);
                setCellSheetWithStyle(sheet, numRow, 1, data.getSpecialityDescr());
                setCellSheetWithStyle(sheet, numRow, 2, getCourse(data.getSemesterDescr()));
                setCellSheetWithStyle(sheet, numRow, 3, data.getDescr().replaceAll("\"", ""));
                // Нахождение суммы с использованием стрима и метода sum()
                double sum = LaborCosts.stream().mapToDouble(Double::doubleValue).sum();
                setCellSheetWithStyle(sheet, numRow, 4, sum);
                setCellSheetWithStyle(sheet, numRow, 5, LaborCosts.get(0));
                setCellSheetWithStyle(sheet, numRow, 6, "");
                setCellSheetWithStyle(sheet, numRow, 7, "");
                setCellSheetWithStyle(sheet, numRow, 8, data.getStudentCount());
                setCellSheetWithStyle(sheet, numRow, 9, data.getContract() ? "Договор" : "Инд.план");
                if (data.getContract()) {
                    sumDogovor += sum;
                } else {
                    sumPlan += sum;
                }
                count++;
                numRow++;
            }
            setFormulaCell(sheet, numRow, 4, "SUM(E" + (startRow + 1) + ":E" + numRow + ")");
            setFormulaCell(sheet, numRow, 5, "SUM(F" + (startRow + 1) + ":F" + numRow + ")");
            setFormulaCell(sheet, numRow, 6, "AVERAGE(G" + (startRow + 1) + ":G" + numRow + ")");
            setFormulaCell(sheet, numRow, 7, "AVERAGE(H" + (startRow + 1) + ":H" + numRow + ")");
            setFormulaCell(sheet, numRow, 8, "SUM(I" + (startRow + 1) + ":I" + numRow + ")");
        }

    }
