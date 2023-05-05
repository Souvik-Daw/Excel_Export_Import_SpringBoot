package com.nilesh.springExcelExport;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.CommandLineRunner;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

import com.nilesh.springExcelExport.model.Student;
import com.nilesh.springExcelExport.repository.StudentRepository;

@SpringBootApplication
public class SpringExcelExportApplication implements CommandLineRunner {

	public static void main(String[] args) {
		SpringApplication.run(SpringExcelExportApplication.class, args);
	}
	
	@Autowired
	private StudentRepository studentRepo;

	@Override
	public void run(String... args) throws Exception {
		Student student1=new Student(1,"Nilesh","My Address line 1,My Address line 2","Rajkot","360001");
		Student student2=new Student(2,"Mahesh","My Address line A,My Address line B","Ahmedabad","360002");
		
		studentRepo.save(student1);
		studentRepo.save(student2);
		
	}

}
