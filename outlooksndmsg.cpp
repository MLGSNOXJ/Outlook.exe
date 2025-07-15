#include <windows.h>
#include <string>
#include <ctime>
#include <chrono>
#include <thread>
#include <iostream>
#include <vector>
#include <io.h>
#include <fcntl.h>
#include <iomanip>
#include <sstream>
#include <conio.h>
#include <algorithm>

std::wstring EscapeForPowerShell(const std::wstring& input) {
    std::wstring result;
    for (wchar_t c : input) {
        if (c == L'\'') result += L"''";
        else if (c == L'`') result += L"``";
        else if (c == L'$') result += L"`$";
        else result += c;
    }
    return result;
}

bool RunPowerShellCommand(const std::wstring& command, std::wstring& output) {
    SECURITY_ATTRIBUTES sa = { sizeof(sa), NULL, TRUE };
    HANDLE hReadPipe, hWritePipe;

    if (!CreatePipe(&hReadPipe, &hWritePipe, &sa, 0)) {
        output = L"Ошибка создания канала";
        return false;
    }

    STARTUPINFO si = { sizeof(si) };
    PROCESS_INFORMATION pi;
    si.dwFlags = STARTF_USESHOWWINDOW | STARTF_USESTDHANDLES;
    si.wShowWindow = SW_HIDE;
    si.hStdOutput = hWritePipe;
    si.hStdError = hWritePipe;

    std::wstring fullCommand = L"powershell.exe -WindowStyle Hidden -NoProfile -ExecutionPolicy Bypass -Command \"" + command + L"\"";

    if (!CreateProcessW(NULL, const_cast<wchar_t*>(fullCommand.c_str()), NULL, NULL, TRUE, CREATE_NO_WINDOW, NULL, NULL, &si, &pi)) {
        CloseHandle(hReadPipe);
        CloseHandle(hWritePipe);
        output = L"Ошибка создания процесса PowerShell";
        return false;
    }

    CloseHandle(hWritePipe);

    const int BUFFER_SIZE = 4096;
    std::vector<char> buffer(BUFFER_SIZE);
    DWORD bytesRead;
    std::string result;

    while (ReadFile(hReadPipe, buffer.data(), BUFFER_SIZE - 1, &bytesRead, NULL) && bytesRead != 0) {
        buffer[bytesRead] = '\0';
        result += buffer.data();
    }

    WaitForSingleObject(pi.hProcess, INFINITE);
    DWORD exitCode;
    GetExitCodeProcess(pi.hProcess, &exitCode);

    CloseHandle(hReadPipe);
    CloseHandle(pi.hProcess);
    CloseHandle(pi.hThread);

    if (!result.empty()) {
        int size = MultiByteToWideChar(CP_UTF8, 0, result.c_str(), -1, NULL, 0);
        if (size > 0) {
            std::vector<wchar_t> wbuffer(size);
            MultiByteToWideChar(CP_UTF8, 0, result.c_str(), -1, wbuffer.data(), size);
            output = wbuffer.data();
        }
    }

    return exitCode == 0;
}

bool SendOutlookEmail(const std::wstring& recipient, const std::wstring& ccRecipient, const std::wstring& bccRecipient,
    const std::wstring& subject, const std::wstring& htmlBody, const std::wstring& fromEmail,
    std::wstring& statusMessage) {

    std::wstring psCommand =
        L"try {\n"
        L"  $outlook = New-Object -ComObject Outlook.Application\n"
        L"  $mail = $outlook.CreateItem(0)\n"
        L"  $mail.To = '" + EscapeForPowerShell(recipient) + L"'\n"
        L"  $mail.CC = '" + EscapeForPowerShell(ccRecipient) + L"'\n"
        L"  $mail.BCC = '" + EscapeForPowerShell(bccRecipient) + L"'\n"
        L"  $mail.Subject = '" + EscapeForPowerShell(subject) + L"'\n"
        L"  $htmlContent = @'\n" +
        htmlBody + L"\n" +
        L"'@\n"
        L"  $mail.HTMLBody = $htmlContent\n"
        L"  $namespace = $outlook.GetNamespace('MAPI')\n"
        L"  $account = $namespace.Accounts | Where-Object { $_.SmtpAddress -eq '" + EscapeForPowerShell(fromEmail) + L"' } | Select-Object -First 1\n"
        L"  if ($account -ne $null) {\n"
        L"    $mail.SendUsingAccount = $account\n"
        L"  } else {\n"
        L"    throw 'Учетная запись не найдена'\n"
        L"  }\n"
        L"  $mail.Send()\n"
        L"  Write-Output 'Письмо успешно отправлено'\n"
        L"  exit 0\n"
        L"} catch {\n"
        L"  Write-Output ('Ошибка: ' + $_.Exception.Message)\n"
        L"  exit 1\n"
        L"}";

    return RunPowerShellCommand(psCommand, statusMessage);
}

void InitConsole() {
    AllocConsole();

    HANDLE hConOut = CreateFile(L"CONOUT$", GENERIC_READ | GENERIC_WRITE, FILE_SHARE_READ | FILE_SHARE_WRITE, NULL, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, NULL);
    HANDLE hConIn = CreateFile(L"CONIN$", GENERIC_READ | GENERIC_WRITE, FILE_SHARE_READ | FILE_SHARE_WRITE, NULL, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, NULL);

    SetStdHandle(STD_OUTPUT_HANDLE, hConOut);
    SetStdHandle(STD_ERROR_HANDLE, hConOut);
    SetStdHandle(STD_INPUT_HANDLE, hConIn);

    _setmode(_fileno(stdout), _O_U16TEXT);
    _setmode(_fileno(stderr), _O_U16TEXT);
    _setmode(_fileno(stdin), _O_U16TEXT);

    std::wcout.clear();
    std::wcerr.clear();
    std::wcin.clear();

    SetConsoleTitleW(L"Outlook Email Sender");
}

std::wstring GetCurrentDateFormatted() {
    auto now = std::chrono::system_clock::now();
    std::time_t now_time = std::chrono::system_clock::to_time_t(now);
    std::tm localNow;
    localtime_s(&localNow, &now_time);

    std::wstringstream wss;
    wss << std::setw(2) << std::setfill(L'0') << localNow.tm_mday << L"."
        << std::setw(2) << std::setfill(L'0') << (localNow.tm_mon + 1) << L"."
        << (localNow.tm_year + 1900);

    return wss.str();
}

void PrintHeader(const std::wstring& recipient, const std::wstring& ccRecipient, const std::wstring& bccRecipient, int sendHour, int sendMinute) {
    system("cls");
    std::wcout << L"=== Служба отправки писем через Outlook ===" << std::endl;
    std::wcout << L"Отправитель: getpro0576@gmail.com" << std::endl;
    std::wcout << L"Получатели: " << std::endl;
    std::wcout << L"  Кому: " << recipient << std::endl;
    std::wcout << L"  Копия: " << (ccRecipient.empty() ? L"(не указано)" : ccRecipient) << std::endl;
    std::wcout << L"  Скрытая копия: " << (bccRecipient.empty() ? L"(не указано)" : bccRecipient) << std::endl;
    std::wcout << L"Тема письма: Ежедневный отчёт" << std::endl;
    std::wcout << L"Текст письма: Добрый день. [дата] ПАО «Россети Волга» без происшествий.\n"
        << L"С уважением,\n"
        << L"Чернов Валентин\n"
        << L"Диспетчер отдела МиЙАО\n"
        << L"Департамента информационной безопасности\n"
        << L"ПАО «Россети Волга»\n"
        << L"Улица Полиграфическая, 79а, Энгельс, 413106\n"
        << L"Тел.: +7 (8452) 30-35-02\n"
        << L"Моб.: +7 (937) 25-141-50\n"
        << L"E-mail: va.chernov@rossetivolga.ru" << std::endl;
    std::wcout << L"Время отправки: " << std::setw(2) << std::setfill(L'0') << sendHour << L":"
        << std::setw(2) << std::setfill(L'0') << sendMinute << std::endl;
    std::wcout << L"===========================================" << std::endl;
    std::wcout << L"Ожидание времени отправки..." << std::endl;
    std::wcout << L"Нажмите 'm' для изменения настроек" << std::endl << std::endl;
}

void ShowMenu(std::wstring& recipient, std::wstring& ccRecipient, std::wstring& bccRecipient, int& sendHour, int& sendMinute) {
    int choice;
    do {
        std::wcout << L"\n=== Меню настройки ===" << std::endl;
        std::wcout << L"1. Изменить адрес получателя (Кому)" << std::endl;
        std::wcout << L"2. Изменить адрес получателя копии (Копия)" << std::endl;
        std::wcout << L"3. Изменить адрес скрытого получателя (Скрытая копия)" << std::endl;
        std::wcout << L"4. Изменить время отправки" << std::endl;
        std::wcout << L"5. Просмотреть текущие настройки" << std::endl;
        std::wcout << L"6. Вернуться к ожиданию отправки" << std::endl;
        std::wcout << L"Выберите действие: ";

        std::wcin >> choice;
        std::wcin.ignore((std::numeric_limits<std::streamsize>::max)(), L'\n');

        switch (choice) {
        case 1:
            std::wcout << L"Текущий адрес получателя: " << recipient << std::endl;
            std::wcout << L"Введите новый адрес получателя: ";
            std::getline(std::wcin, recipient);
            std::wcout << L"Адрес получателя изменен на: " << recipient << std::endl;
            break;
        case 2:
            std::wcout << L"Текущий адрес получателя копии: " << (ccRecipient.empty() ? L"(не указано)" : ccRecipient) << std::endl;
            std::wcout << L"Введите новый адрес получателя копии (оставьте пустым для удаления): ";
            std::getline(std::wcin, ccRecipient);
            std::wcout << L"Адрес получателя копии изменен на: " << (ccRecipient.empty() ? L"(не указано)" : ccRecipient) << std::endl;
            break;
        case 3:
            std::wcout << L"Текущий адрес скрытого получателя: " << (bccRecipient.empty() ? L"(не указано)" : bccRecipient) << std::endl;
            std::wcout << L"Введите новый адрес скрытого получателя (оставьте пустым для удаления): ";
            std::getline(std::wcin, bccRecipient);
            std::wcout << L"Адрес скрытого получателя изменен на: " << (bccRecipient.empty() ? L"(не указано)" : bccRecipient) << std::endl;
            break;
        case 4: {
            std::wcout << L"Текущее время отправки: "
                << std::setw(2) << std::setfill(L'0') << sendHour << L":"
                << std::setw(2) << std::setfill(L'0') << sendMinute << std::endl;

            int newHour, newMinute;
            std::wcout << L"Введите новый час отправки (0-23): ";
            std::wcin >> newHour;
            std::wcout << L"Введите новую минуту отправки (0-59): ";
            std::wcin >> newMinute;

            if (newHour >= 0 && newHour <= 23 && newMinute >= 0 && newMinute <= 59) {
                sendHour = newHour;
                sendMinute = newMinute;
                std::wcout << L"Время отправки изменено на: "
                    << std::setw(2) << std::setfill(L'0') << sendHour << L":"
                    << std::setw(2) << std::setfill(L'0') << sendMinute << std::endl;
            }
            else {
                std::wcout << L"Некорректное время. Изменения не сохранены." << std::endl;
            }
            std::wcin.ignore((std::numeric_limits<std::streamsize>::max)(), L'\n');
            break;
        }
        case 5:
            std::wcout << L"\nТекущие настройки:" << std::endl;
            std::wcout << L"Кому: " << recipient << std::endl;
            std::wcout << L"Копия: " << (ccRecipient.empty() ? L"(не указано)" : ccRecipient) << std::endl;
            std::wcout << L"Скрытая копия: " << (bccRecipient.empty() ? L"(не указано)" : bccRecipient) << std::endl;
            std::wcout << L"Время отправки: "
                << std::setw(2) << std::setfill(L'0') << sendHour << L":"
                << std::setw(2) << std::setfill(L'0') << sendMinute << std::endl;
            break;
        case 6:
            std::wcout << L"Возвращаемся к ожиданию времени отправки..." << std::endl;
            break;
        default:
            std::wcout << L"Неверный выбор. Попробуйте снова." << std::endl;
            break;
        }
    } while (choice != 6);
}

void ServiceMain() {
    std::wstring recipient = L"ps.mitin@rossetivolga.ru";
    std::wstring ccRecipient = L"as.zhukov@rossetivolga.ru";
    std::wstring bccRecipient = L"xartem820@gmail.com";
    const std::wstring subject = L"Ежедневный отчёт";
    const std::wstring fromEmail = L"getpro0576@gmail.com";

    // Начальное время отправки
    int sendHour = 16;
    int sendMinute = 00;

    InitConsole();
    PrintHeader(recipient, ccRecipient, bccRecipient, sendHour, sendMinute);

    while (true) {
        // Проверка нажатия клавиши 'm' в любое время
        if (_kbhit()) {
            wchar_t ch = _getwch();
            if (ch == L'm' || ch == L'M') {
                ShowMenu(recipient, ccRecipient, bccRecipient, sendHour, sendMinute);
                PrintHeader(recipient, ccRecipient, bccRecipient, sendHour, sendMinute);
            }
        }

        auto now = std::chrono::system_clock::now();
        std::time_t now_time = std::chrono::system_clock::to_time_t(now);
        std::tm localNow;
        localtime_s(&localNow, &now_time);

        if (localNow.tm_hour == sendHour && localNow.tm_min == sendMinute) {
            std::wstring currentDate = GetCurrentDateFormatted();
            std::wstring htmlBody =
                L"<html>\n"
                L"<head>\n"
                L"<style>\n"
                L"  body { font-family: Calibri, Arial, sans-serif; font-size: 11pt; }\n"
                L"  .main-text { line-height: 1.5; margin-bottom: 12pt; }\n"
                L"  .signature { line-height: 1.0; }\n"
                L"  .signature p { margin: 4pt 0; }\n"
                L"</style>\n"
                L"</head>\n"
                L"<body>\n"
                L"<div class=\"main-text\">Добрый день. " + currentDate + L" ПАО «Россети Волга» без происшествий.</div>\n"
                L"<div class=\"signature\">\n"
                L"  <p>С уважением,</p>\n"
                L"  <p>Единый центр управления безопасности</p>\n"
                L"  <p>Департамента информационной безопасности</p>\n"
                L"  <p>ПАО «Россети Волга»</p>\n"
                L"  <p>Улица Полиграфическая, 79а, Энгельс, 413106</p>\n"
                L"  <p>Тел.: +7 (8452) 30-35-02</p>\n"
                L"  <p>E-mail: <a href=\"mailto:ecub@rossetivolga.ru\">ecub@rossetivolga.ru</a></p>\n"
                L"</div>\n"
                L"</body>\n"
                L"</html>";

            std::wstring statusMessage;
            std::wcout << L"[Попытка отправки] ";
            std::wcout.flush();

            bool success = SendOutlookEmail(recipient, ccRecipient, bccRecipient, subject, htmlBody, fromEmail, statusMessage);

            if (success) {
                std::wcout << L"УСПЕХ: " << statusMessage << std::endl;
            }
            else {
                std::wcout << L"ОШИБКА: " << statusMessage << std::endl;
            }

            // Выводим сообщение и сразу продолжаем проверять ввод
            std::wcout << L"Ожидание следующего времени отправки..." << std::endl;
            std::wcout << L"Нажмите 'm' для изменения настроек" << std::endl << std::endl;

            // Задержка, чтобы избежать повторной отправки в ту же минуту
            for (int i = 0; i < 61; i++) {
                if (_kbhit()) {
                    wchar_t ch = _getwch();
                    if (ch == L'm' || ch == L'M') {
                        ShowMenu(recipient, ccRecipient, bccRecipient, sendHour, sendMinute);
                        PrintHeader(recipient, ccRecipient, bccRecipient, sendHour, sendMinute);
                        break;
                    }
                }
                std::this_thread::sleep_for(std::chrono::seconds(1));
            }
        }
        else {
            // Короткая задержка для уменьшения нагрузки на процессор
            std::this_thread::sleep_for(std::chrono::milliseconds(100));
        }
    }
}

int main() {
    ServiceMain();
    return 0;
}
