package com.gmoz.entity;

import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import javax.persistence.Column;
import javax.persistence.Entity;
import javax.persistence.JoinColumn;
import javax.persistence.JoinTable;
import javax.persistence.ManyToMany;
import javax.persistence.Table;

import lombok.Getter;
import lombok.Setter;

@Getter
@Setter
@Entity
@Table(name = "students")
public class StudentEntity extends BaseEntity {

	private String fullname;

	private Date birthdate;

	private String address;

	private Boolean gender;

	@Column(unique = true)
	private String phone;

	@Column(unique = true)
	private String email;

	private String note;

	@ManyToMany
	@JoinTable(	name = "student_class", 
				joinColumns = @JoinColumn(name = "student_id"), 
				inverseJoinColumns = @JoinColumn(name = "class_id"))
	private List<ClassEntity> classes = new ArrayList<>();

}
