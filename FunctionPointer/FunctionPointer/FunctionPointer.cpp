//================================================================================================================================================
// FunctionPointer.cpp : Defines the entry point for the console application.
//
//
//
//
//
//================================================================================================================================================

#include "stdafx.h"
#include <iostream>

using namespace std;

//------------------------------------------------------------------------------------
// 1.2 Introductory Example or How to Replace a Switch-Statement
// Task: Perform one of the four basic arithmetic operations specified by the
//       characters '+', '-', '*' or '/'.

// The four arithmetic operations ... one of these functions is selected
// at runtime with a swicth or a function pointer
float Plus    (float a, float b) { return a+b; }
float Minus   (float a, float b) { return a-b; }
float Multiply(float a, float b) { return a*b; }
float Divide  (float a, float b) { return a/b; }

void anEventExample();
void anEventExample_2();

//int sayHello(int a, int b){cout << "Hello World from the Event Receiver!\n";return 1;}
int sayHello(int a, int b);
void sayHello_2();

int (*arFunc[1])(int a, int b);
void (*arFunc_2[1])();

int intControl = 1;

//================================================================================================================================================
int _tmain(int argc, _TCHAR* argv[])
{
	arFunc[0] = sayHello;
	arFunc_2[0] = sayHello_2;
	
	while(intControl != 0)
	{
		cout << "Enter operation: "; 
		cin >> intControl;

		if(intControl == 8)
		{
			anEventExample();
		}
		else if(intControl == 6)
		{
			anEventExample_2();
		}
	}
	
	return 0;
}

//================================================================================================================================================
void anEventExample()
{
	cout << "Some interesting event has happened...!!\n";

	int result;

	int intSize = sizeof(arFunc) / sizeof(int);

	for(int ctr = 0; ctr < intSize; ctr++)
	//for(ctr = 0; ctr < 5; ctr++)
	{
		//printf("Hello from the loop....%i\n", ctr); 
		result = (*arFunc[ctr])(1, 2);
	}
}

//================================================================================================================================================
void anEventExample_2()
{
	int intSize = sizeof(arFunc) / sizeof(int);
	
	for(int ctr = 0; ctr < intSize; ctr++)
	{
		(*arFunc_2[ctr])();
	}
}

//================================================================================================================================================
int sayHello(int a, int b)
{
	cout << "Hello from the event......\n";

	return a + b;
}

//================================================================================================================================================
void sayHello_2()
{
	cout << "This is the void event...\n";
}

//================================================================================================================================================
// Solution with a switch-statement - <opCode> specifies which operation to execute
void Switch(float a, float b, char opCode)
{
   float result;
   
   // execute operation
   switch(opCode)
   {
      case '+' : result = Plus     (a, b); break;
      case '-' : result = Minus    (a, b); break;
      case '*' : result = Multiply (a, b); break;
      case '/' : result = Divide   (a, b); break;
   }
   
   cout << "Switch: 2+5=" << result << endl;         // display result
}

//================================================================================================================================================
// Solution with a function pointer - <pt2Func> is a function pointer and points to
// a function which takes two floats and returns a float. The function pointer
// "specifies" which operation shall be executed.
void Switch_With_Function_Pointer(float a, float b, float(*pt2Func)(float, float))
{
   float result = pt2Func(a, b);    // call using function pointer

   cout << "Switch replaced by function pointer: 2-5=";  // display result
   cout << result << endl;
}

//================================================================================================================================================
// Execute example code
void Replace_A_Switch()
{
   cout << endl << "Executing function 'Replace_A_Switch'" << endl;

   Switch(2, 5, /* '+' specifies function 'Plus' to be executed */ '+');
   Switch_With_Function_Pointer(2, 5, /* pointer to function 'Minus' */ &Minus);
}

//================================================================================================================================================
//	end of file
//================================================================================================================================================