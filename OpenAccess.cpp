#include <windows.h>
#include <tchar.h>
#include <stdio.h>

int _tmain(int argc, TCHAR *argv[]) {
    // Путь к Access
    TCHAR accessPath[] = TEXT("C:\\Program Files\\Microsoft Office\\root\\Office16\\MSACCESS.EXE");

    STARTUPINFO si;
    PROCESS_INFORMATION pi;

    ZeroMemory(&si, sizeof(si));
    si.cb = sizeof(si);
    ZeroMemory(&pi, sizeof(pi));

    // Запускаем Acess без указания файла
    if (!CreateProcess(NULL, accessPath, NULL, NULL, FALSE, 0, NULL, NULL, &si, &pi)) {
        printf("CreateProcess failed (%d).\n", GetLastError());
        return 1;
    }

    // Закрываем дескрипторы процесса и потока
    // CloseHandle(pi.hProcess);
    // CloseHandle(pi.hThread);

    return 0;
}