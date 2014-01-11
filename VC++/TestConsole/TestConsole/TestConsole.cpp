// TestConsole.cpp : Defines the entry point for the console application.
//

#include "stdafx.h"
#include "TestClass.h"

#include <iostream>
#include <string>
#include <vector>

using namespace std;

enum Color {Red, Blue, White, Green, Yellow};

int addTwoNumbers(int a, int b);

int _tmain(int argc, _TCHAR* argv[])
{
	int x;
	string str_name;
	
	cout << "Hello World!" << endl;
	cout << "Welcome to C++ programming" << endl;

	cout << "Enter Name:";

	getline(cin, str_name);

	cout << str_name << endl;

	x = addTwoNumbers(5, 10);

	cout << x << endl;

	cout << "Enter YES" << endl;

	getline(cin, str_name);

	TestClass *t = new TestClass();
	t->setAge(11);
	
	cout << "The set age is: ";
	cout << t->getAge() << endl;
	
	TestClass tc;
	tc.setAge(15);
	
	cout << "The set age is: ";
	cout << tc.getAge() << endl;
	
	std::vector<TestClass*> animals;
	animals.push_back(new TestClass());
	for(std::vector<TestClass*>::const_iterator it = animals.begin(); it != animals.end(); ++it)
	{
		(*it)->setAge(12);
		cout << "the age is: ";
		cout << (*it)->getAge() << endl;
		delete *it;
	}
	
	if(str_name == "Yes")
	{
		return 0;
	}
	else
	{
		cout << "You must enter YES to exit" << endl;
		cout << "Enter YES:";

		getline(cin, str_name);
	}
}

int addTwoNumbers(int a, int b)
{
	int x;

	x = a + b;

	return x;
}