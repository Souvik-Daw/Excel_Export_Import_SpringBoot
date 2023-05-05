package com.nilesh.springExcelExport.model;

import javax.persistence.Entity;
import javax.persistence.GeneratedValue;
import javax.persistence.GenerationType;
import javax.persistence.Id;

@Entity
public class Student {
	
	@Id
	@GeneratedValue(strategy = GenerationType.IDENTITY)
	private long sid;
	private String sname;
	private String saddress;
	private String scity;
	private String spin;
	
	
	public Student(long sid, String sname, String saddress, String scity, String spin) {
		super();
		this.sid = sid;
		this.sname = sname;
		this.saddress = saddress;
		this.scity = scity;
		this.spin = spin;
	}
	public Student() {
		super();
		// TODO Auto-generated constructor stub
	}
	public long getSid() {
		return sid;
	}
	public void setSid(long sid) {
		this.sid = sid;
	}
	public String getSname() {
		return sname;
	}
	public void setSname(String sname) {
		this.sname = sname;
	}
	public String getSaddress() {
		return saddress;
	}
	public void setSaddress(String saddress) {
		this.saddress = saddress;
	}
	public String getScity() {
		return scity;
	}
	public void setScity(String scity) {
		this.scity = scity;
	}
	public String getSpin() {
		return spin;
	}
	public void setSpin(String spin) {
		this.spin = spin;
	}
	
	
	
}
