#include "StdAfx.h"
#include "TestClass.h"

TestClass::TestClass(void)
{
}

TestClass::~TestClass(void)
{
}

void TestClass::setAge(int age)
{
	_age = age;
}

int TestClass::getAge()
{
	return _age;
}