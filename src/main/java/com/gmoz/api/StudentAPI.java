package com.gmoz.api;

import java.io.IOException;
import java.util.List;

import javax.servlet.http.HttpServletResponse;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;

import com.gmoz.entity.StudentEntity;
import com.gmoz.repository.StudentRepository;
import com.gmoz.utils.FileExporter;
import com.gmoz.utils.JExcelHelper;

@RestController
public class StudentAPI {

	@Autowired
	private StudentRepository studentRepository;

	@Autowired
	private FileExporter export;
	
	@Autowired
	private JExcelHelper excelHelper;

	@GetMapping("/export")
	public void exportToExcel(HttpServletResponse response) throws IOException {
		List<StudentEntity> list = studentRepository.findAll();
		export.exportToExcel("Students", "DANH SÁCH SINH VIÊN", 50, list, response);
	}

	@GetMapping("/go")
	public void updateExcel(HttpServletResponse response) throws IOException {
		List<StudentEntity> list = studentRepository.findAll();
		excelHelper.updateExcel("C:/Users/holid/Desktop/Students_18-11-2021_10-16-55.xlsx", 50, list);
//		for (ClassEntity classEntity : list) {
//			System.out.println(classEntity.getName());
//			for (StudentEntity student : classEntity.getStudents()) {
//				System.out.println(student.toString());
//			}
//		}
	}
}
