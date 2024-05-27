package com.example.demo.controllers;

import com.example.demo.helpers.writer.WriteDataBase;
import com.example.demo.services.*;
import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import static com.example.demo.helpers.WorkbookHelpers.multipartFileToWorkbook;

@RestController
@RequestMapping("/api/database")
public class FillDataBaseController {
    private final WriteDataBase writeDataBase;


    public FillDataBaseController(StudyYearService studyYearService, StudentsGroupService studentsGroupService, DisciplineService disciplineService, WorkloadService workloadService, WorkloadAssignService workloadAssignService, EmployeeService employeeService, DeleteDataService deleteDataService) {
        writeDataBase = new WriteDataBase(studyYearService, studentsGroupService, disciplineService, workloadService, workloadAssignService, employeeService, deleteDataService);
    }

    @PostMapping("/fill")
    public ResponseEntity<String> fillDataBase(@RequestParam("year") Integer year, @RequestParam("file") MultipartFile file) {
        // Проверяем, что параметры переданы корректно
        if (file.isEmpty() || file == null) {
            return ResponseEntity.badRequest().body("Файл с данными не передан");
        }
        try {
            Workbook workbook = multipartFileToWorkbook(file);
            writeDataBase.fillDataBase(workbook, year);
            return ResponseEntity.ok().body("Успех");
        } catch (Exception e) {
            e.printStackTrace();
            return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR).body(e.getMessage());
        }
    }
}