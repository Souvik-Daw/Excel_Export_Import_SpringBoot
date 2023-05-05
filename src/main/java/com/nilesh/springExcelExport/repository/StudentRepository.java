package com.nilesh.springExcelExport.repository;

import org.springframework.data.jpa.repository.JpaRepository;

import com.nilesh.springExcelExport.model.Student;

public interface StudentRepository extends JpaRepository<Student, Long> {

}
