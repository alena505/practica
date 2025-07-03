#include <Windows.h>
#include <windows.h>
#include <oleauto.h>  // Для работы с BSTR и SysAllocString
#include <stdio.h>   // Для sprintf (если нужно форматирование)



 __declspec(dllexport) BSTR __stdcall coolSub(double* marks, BSTR * names, int rows, int cols)
{
	double maxsred = 0;
	int bestStudent = 0;
	for (int i = 0; i < rows; i++)
	{
		double sum = 0;
		for (int j = 0; j < cols; j++)
		{
			sum += marks[i * cols + j];
		}
		double sred = sum / cols;
		if (sred > maxsred)
		{
			maxsred = sred;
			bestStudent = i;
		}

	}
	return SysAllocString(names[bestStudent]);
}

BOOL APIENTRY DllMain(HMODULE hModule, DWORD ul_reason_for_call, LPVOID lpReserved)
{
	switch (ul_reason_for_call)
	{
	case DLL_PROCESS_ATTACH:
	case DLL_THREAD_ATTACH:
	case DLL_THREAD_DETACH:
	case DLL_PROCESS_DETACH:
		break;
	}
	return TRUE;
}
