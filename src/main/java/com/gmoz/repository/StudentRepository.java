package com.gmoz.repository;

import org.springframework.data.jpa.repository.JpaRepository;

import com.gmoz.entity.StudentEntity;

public interface StudentRepository extends JpaRepository<StudentEntity, String> {

}
