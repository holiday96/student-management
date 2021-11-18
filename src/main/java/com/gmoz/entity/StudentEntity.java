package com.gmoz.entity;

import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import javax.persistence.Column;
import javax.persistence.Entity;
import javax.persistence.Id;
import javax.persistence.JoinColumn;
import javax.persistence.JoinTable;
import javax.persistence.ManyToMany;
import javax.persistence.Table;

import lombok.Getter;
import lombok.Setter;
import lombok.ToString;

@Getter
@Setter
@ToString
@Entity
@Table(name = "students")
public class StudentEntity {

	@Id
	private String id;

	private String name;

	private Boolean gender;

	private Date birthdate;

	@Column(unique = true)
	private String phone;

	private Integer age;

	@ManyToMany
	@JoinTable(name = "student_class", joinColumns = @JoinColumn(name = "student_id"), inverseJoinColumns = @JoinColumn(name = "class_id"))
	private List<ClassEntity> classes = new ArrayList<>();

}
