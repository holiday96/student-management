package com.gmoz.entity;

import java.util.ArrayList;
import java.util.List;

import javax.persistence.Entity;
import javax.persistence.ManyToMany;
import javax.persistence.Table;

import lombok.Getter;
import lombok.Setter;

@Getter
@Setter
@Entity
@Table(name = "classes")
public class ClassEntity extends BaseEntity {

	private String name;

	@ManyToMany(mappedBy = "classes")
	private List<StudentEntity> students = new ArrayList<>();
}
