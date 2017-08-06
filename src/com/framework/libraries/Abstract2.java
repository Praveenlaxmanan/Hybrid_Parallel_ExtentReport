package com.framework.libraries;


public class Abstract2 extends AbstractClass {

	@Override
	void test1() {
		// TODO Auto-generated method stub
		System.out.println("Implemented");
		
	}

	
	public static void main(String[] args) {
		
		AbstractClass classes = new Abstract2();
		classes.test();
		
	}
}
