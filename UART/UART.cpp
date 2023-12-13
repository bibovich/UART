#include <Windows.h>
#include <strsafe.h>
#include <tchar.h>


inline constexpr int BUFSIZE = 100;
const wchar_t COMPORT_Rx[] = L"COM9";
const wchar_t COMPORT_Tx[] = L"COM8";


inline constexpr UINT_PTR ID_INIT_PORT = 1000;
inline constexpr UINT_PTR ID_OUTPUT_PORT = 1001;
inline constexpr UINT_PTR ID_INPUT_PORT = 1002;
inline constexpr UINT_PTR ID_CLOSE_PORT = 1003;

// Communication variables and parameters
HANDLE hCOMRx;										// Pointer to the selected COM port (receiver)
HANDLE hCOMTx;										// Pointer to the selected COM port (transmitter)
inline constexpr int nComRate = 9600;				// Baud (Bit) rate in bits/second 
inline constexpr int nComBits = 8;					// Number of bits per frame
COMMTIMEOUTS timeout;								// Timeout between processess


void createPortFile(HWND& hwnd, HANDLE& hCOM, const wchar_t* COMPORT);
void purgePort(HANDLE& hCOM);
int setCOM(HANDLE& hCOM, int nCOMRate, int nCOMBits, COMMTIMEOUTS timeout);
void outputToPort(HWND& hwnd, HANDLE& hCOM, LPCVOID buf, DWORD szBuf);
DWORD inputFromPort(HWND& hwnd, HANDLE& hCOM, LPVOID buf, DWORD szBuf);
void initPort(HWND& hwnd, HANDLE& hCOM, const wchar_t* COMPORT, int nCOMRate, int nCOMBits, COMMTIMEOUTS timeout);

LRESULT CALLBACK WndProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam);



_Use_decl_annotations_ int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPSTR lpCmdLine, int nCmdShow) {
	const TCHAR szClassName[] = _T("SerialPortApp");
	WNDCLASS wc = { 0 };

	wc.lpfnWndProc = WndProc;
	wc.hInstance = hInstance;
	wc.hbrBackground = (HBRUSH)(COLOR_WINDOW + 1);
	wc.lpszClassName = szClassName;

	if (!RegisterClass(&wc)) {
		MessageBox(nullptr, _T("Window Registration Failed!"), _T("Error"), MB_ICONEXCLAMATION | MB_OK);
		return 0;
	}

	HWND hwnd = CreateWindow(szClassName, _T("Serial Port Interface"), WS_OVERLAPPEDWINDOW,
		CW_USEDEFAULT, CW_USEDEFAULT, 400, 300, nullptr, nullptr, hInstance, nullptr);

	if (!hwnd) {
		MessageBox(nullptr, _T("Window Creation Failed!"), _T("Error"), MB_ICONEXCLAMATION | MB_OK);
		return 0;
	}

	ShowWindow(hwnd, nCmdShow);
	UpdateWindow(hwnd);

	MSG msg;
	while (GetMessage(&msg, nullptr, 0, 0) > 0) {
		TranslateMessage(&msg);
		DispatchMessage(&msg);
	}

	return 0;
}




void initPort(HWND& hwnd, HANDLE& hCOM, const wchar_t* COMPORT, int nCOMRate, int nCOMBits, COMMTIMEOUTS timeout) {
	createPortFile(hwnd, hCOM, COMPORT);						// Initializes hCOM to point to PORT#
	purgePort(hCOM);											// Purges the COM port
	setCOM(hCOM, nCOMRate, nCOMBits, timeout);					// Setting up COM port
	purgePort(hCOM);
}

void createPortFile(HWND& hwnd, HANDLE& hCOM, const wchar_t* portName) {
	hCOM = CreateFile(portName,
		GENERIC_READ | GENERIC_WRITE,
		0,
		nullptr,
		OPEN_EXISTING,
		0,
		nullptr);

	if (hCOM == INVALID_HANDLE_VALUE)
		MessageBox(hwnd, _T("Failed to open the serial port!"), _T("Error"), MB_ICONEXCLAMATION | MB_OK);
	else
		MessageBox(hwnd, _T("Serial port opened successfully!"), portName, MB_ICONINFORMATION | MB_OK);

}

void purgePort(HANDLE& hCOM) {
	PurgeComm(hCOM, PURGE_RXABORT | PURGE_RXCLEAR | PURGE_TXABORT | PURGE_TXCLEAR);
}

int setCOM(HANDLE& hCOM, int nCOMRate, int nCOMBits, COMMTIMEOUTS timeout) {
	DCB dcb;										// Windows device control block

	memset(&dcb, 0, sizeof(dcb));
	dcb.DCBlength = sizeof(dcb);
	if (!GetCommState(hCOM, &dcb))
		return(0);

	// Set our own parameters from Globals
	dcb.BaudRate = nComRate;						// Baud (bit) rate
	dcb.ByteSize = (BYTE)nComBits;					// Number of bits(8)
	dcb.Parity = 0;									// No parity	
	dcb.StopBits = ONESTOPBIT;						// One stop bit
	if (!SetCommState(hCOM, &dcb))
		return(0);

	// Timeout
	memset((void*)&timeout, 0, sizeof(timeout));
	timeout.ReadIntervalTimeout = 500;
	timeout.ReadTotalTimeoutMultiplier = 1;
	timeout.ReadTotalTimeoutConstant = 5000;
	SetCommTimeouts(hCOM, &timeout);
	return(1);
}

void outputToPort(HWND& hwnd, HANDLE& hCOM, LPCVOID buf, DWORD szBuf) {
	bool file = 0;
	DWORD NumberofBytesTransmitted;
	LPDWORD lpErrors = 0;
	LPCOMSTAT lpStat = 0;


	file = WriteFile(
		hCOM,										// Write handle pointing to COM port
		buf,										// Buffer size
		szBuf,										// Size of buffer
		&NumberofBytesTransmitted,					// Written number of bytes
		nullptr
	);

	TCHAR msg[BUFSIZE]{};
	StringCbPrintf(msg, BUFSIZE, TEXT("%d"), NumberofBytesTransmitted);

	// Handle the timeout error
	if (!file) {
		MessageBox(hwnd, _T("Failed to write into the serial port!"), _T("Error"), MB_ICONEXCLAMATION | MB_OK);
		ClearCommError(hCOM, lpErrors, lpStat);
	}
	else
		MessageBox(hwnd, msg, _T("Successful transmission, amount of bytes were transmitted: "), MB_ICONINFORMATION | MB_OK);
}

DWORD inputFromPort(HWND& hwnd, HANDLE& hCOM, LPVOID buf, DWORD szBuf) {
	bool file = 0;
	DWORD NumberofBytesRead;
	LPDWORD lpErrors = 0;
	LPCOMSTAT lpStat = 0;

	file = ReadFile(
		hCOM,										// Read handle pointing to COM port
		buf,										// Buffer size
		szBuf,  									// Size of buffer - Maximum number of bytes to read
		&NumberofBytesRead,
		nullptr
	);

	TCHAR msg[BUFSIZE]{};
	StringCbPrintf(msg, BUFSIZE, TEXT("%d"), NumberofBytesRead);
	// Handle the timeout error

	if (!file) {
		TCHAR err[BUFSIZE]{};
		StringCbPrintf(msg, BUFSIZE, TEXT("%d"), GetLastError());
		MessageBox(hwnd, _T("Read Error:"), err, MB_ICONEXCLAMATION | MB_OK);

		ClearCommError(hCOM, lpErrors, lpStat);
	}
	else
		MessageBox(hwnd, msg, _T("Successful, Amount of bytes were read: "), MB_ICONINFORMATION | MB_OK);


	return(NumberofBytesRead);
}



LRESULT CALLBACK WndProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam) {
	switch (uMsg) {
	case WM_CREATE: {
		CreateWindow(_T("BUTTON"), _T("Open Port"), WS_TABSTOP | WS_VISIBLE | WS_CHILD | BS_DEFPUSHBUTTON,
			10, 10, 100, 30, hwnd, (HMENU)ID_INIT_PORT, nullptr, nullptr);

		CreateWindow(_T("BUTTON"), _T("Close Port"), WS_TABSTOP | WS_VISIBLE | WS_CHILD | BS_DEFPUSHBUTTON,
			120, 10, 100, 30, hwnd, (HMENU)ID_CLOSE_PORT, nullptr, nullptr);

		CreateWindow(_T("BUTTON"), _T("Send Data"), WS_TABSTOP | WS_VISIBLE | WS_CHILD | BS_DEFPUSHBUTTON,
			10, 50, 100, 30, hwnd, (HMENU)ID_INPUT_PORT, nullptr, nullptr);

		CreateWindow(_T("BUTTON"), _T("Read Data"), WS_TABSTOP | WS_VISIBLE | WS_CHILD | BS_DEFPUSHBUTTON,
			120, 50, 100, 30, hwnd, (HMENU)ID_OUTPUT_PORT, nullptr, nullptr);
		break;
	}

	case WM_COMMAND: {
		switch (LOWORD(wParam)) {
			case ID_INIT_PORT: {
				initPort(hwnd, hCOMRx, COMPORT_Rx, nComRate, nComBits, timeout);
				Sleep(500);
				initPort(hwnd, hCOMTx, COMPORT_Tx, nComRate, nComBits, timeout);
				Sleep(500);
				break;
			}
			case ID_INPUT_PORT: {
				// Transmit side
				TCHAR msgOut[] = _T("some important message");			// Sent message	
	
				outputToPort(hwnd, hCOMTx, msgOut, sizeof(msgOut));
				Sleep(500);
				break;
	
			}
			case ID_OUTPUT_PORT: {
				// Receive side  
				TCHAR msgIn[BUFSIZE]{};
				DWORD bytesRead;
				bytesRead = inputFromPort(hwnd, hCOMRx, msgIn, BUFSIZE);			// Receive string from port
	
				TCHAR msg[BUFSIZE]{};
				StringCbPrintf(msg, BUFSIZE, TEXT("%d"), bytesRead);
	
				MessageBox(nullptr, msgIn, _T("Message has been received"), MB_ICONINFORMATION | MB_OK);
				break;
			}
			case ID_CLOSE_PORT: {
				purgePort(hCOMRx);											// Purge the Rx port
				purgePort(hCOMTx);											// Purge the Tx port
	
				CloseHandle(hCOMRx);										// Close the handle to Rx port 
				MessageBox(nullptr, _T("Port is closed"), COMPORT_Rx, MB_ICONINFORMATION | MB_OK);
	
				CloseHandle(hCOMTx);										// Close the handle to Tx port 
				MessageBox(nullptr, _T("Port is closed"), COMPORT_Tx, MB_ICONINFORMATION | MB_OK);
				break;
			}
		}
		break;
	}
	case WM_DESTROY: {
		// Quit the application
		PostQuitMessage(0);
		break;
	}
	default:
		// Default window procedure for other messages
		return DefWindowProc(hwnd, uMsg, wParam, lParam);
	}
	return 0;
}
