package com.nilesh.springExcelExport.controller;

import java.io.IOException;
import java.util.List;

import javax.servlet.http.HttpServletResponse;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.ResponseBody;

import com.nilesh.springExcelExport.excel.UserExcelExporter;
import com.nilesh.springExcelExport.excel.UserExcelImporter;
import com.nilesh.springExcelExport.model.Student;
import com.nilesh.springExcelExport.repository.StudentRepository;

@Controller
@RequestMapping("/web")
public class HomeController {
	
	@RequestMapping("/home")
	public String homePage() {
		return "HomePage";
	}

	@Autowired
	private StudentRepository studentRepo;

	@RequestMapping("/export/excel")
	@ResponseBody
	public String exportToExcel(HttpServletResponse response) throws IOException {
		response.setContentType("application/octet-stream");
		String headerKey = "Content-Disposition";
		String headervalue = "attachment; filename=Student_info.xlsx";

		response.setHeader(headerKey, headervalue);
		List<Student> listStudent = studentRepo.findAll();
		UserExcelExporter exp = new UserExcelExporter(listStudent);
		exp.export(response);
		return "Export Successfully";
	}

	@RequestMapping("/import/excel")
	@ResponseBody
	public String importFromExcel(@RequestBody String url) {
		UserExcelImporter excelImporter=new UserExcelImporter();
		List<Student> listStudent= excelImporter.excelImport(url);
		studentRepo.saveAll(listStudent);
		return "Import Successfully";
	}

}
