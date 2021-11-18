package com.gmoz.controller;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.GetMapping;

import com.gmoz.repository.StudentRepository;

@Controller
public class HomeController {

	@Autowired
	private StudentRepository studentRepository;

	@GetMapping("/")
	public String listStudent(Model model) {
		model.addAttribute("students", studentRepository.findAll());
		return "index";
	}
	
}
