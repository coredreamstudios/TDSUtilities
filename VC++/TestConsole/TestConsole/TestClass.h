#pragma once

class TestClass
{
public:
	TestClass(void);
	~TestClass(void);

	void setAge(int age);
	int getAge();

private:
	int _age;
};
