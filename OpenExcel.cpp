#include <windows.h>
#include <tchar.h>
#include <stdio.h>

int _tmain(int argc, TCHAR *argv[]) {
    // Путь к Excel
    TCHAR excelPath[] = TEXT("C:\\Program Files\\Microsoft Office\\root\\Office16\\EXCEL.EXE");

    STARTUPINFO si;
    PROCESS_INFORMATION pi;

    ZeroMemory(&si, sizeof(si));
    si.cb = sizeof(si);
    ZeroMemory(&pi, sizeof(pi));

    // Запускаем Excel без указания файла
    if (!CreateProcess(NULL, excelPath, NULL, NULL, FALSE, 0, NULL, NULL, &si, &pi)) {
        printf("CreateProcess failed (%d).\n", GetLastError());
        return 1;
    }

    // Закрываем дескрипторы процесса и потока
    // CloseHandle(pi.hProcess);
    // CloseHandle(pi.hThread);

    return 0;
}