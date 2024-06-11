#define _CRT_SECURE_NO_WARNINGS
#define SafeRelease(p) { if (p) { (p)->Release(); (p) = nullptr; } }

#pragma comment(lib, "shlwapi.lib")
#pragma comment(lib, "d2d1.lib")
#pragma comment (lib,"dwrite.lib")
#pragma comment (lib,"Comctl32.lib")

#include <shlwapi.h>
#include <shobjidl.h> // IPreviewHandler, IShellItem, IInitializeWithItem, IParentAndItem
#include <new>
#include <windows.h>
#include <commctrl.h> // Include the necessary header for subclassing
#include <objidl.h>
#include <iostream>
#include <fstream>
#include <vector>
#include <string>
#include <sstream>
#include "ATSIFIO.h"
#include <stdio.h>
#include <cstdio>
#include <algorithm>
#include <d2d1.h>
#include <dwrite.h>
#include <cmath>
using namespace std;

// Set up plot parameters
const int lmargin = 60; // Left axis box margin
const int rmargin = 30; // Right axis box margin
const int tmargin = 20; // Top axis box margin
const int bmargin = 30; // Bottom axis box margin
const int labelmargin = 10; // Distance of text to axis box
const int fontsize = 14; // Font size in pixel
const int minimumaxisbox = 50; // Minimum size of the axis box
const int maximumtextbox = 200; // Maximum width of the text box
const int maximumtextboxheight = 120; // Maximum height of the text box
const int decimalPlacesX = 3; // Define the number of decimal places for the x-axis
const int decimalPlacesY = 0; // Define the number of decimal places for the y-axis

// Create colors
const D2D1_COLOR_F whiteColor = D2D1::ColorF(D2D1::ColorF::White);
const D2D1_COLOR_F blackColor = D2D1::ColorF(D2D1::ColorF::Black);
const D2D1_COLOR_F blueColor = D2D1::ColorF(D2D1::ColorF::Blue);

// This method calculates and returns the width of a RECT structure
inline int RECTWIDTH(const RECT& rc)
{
	return (rc.right - rc.left);
}

// This method calculates and returns the heigth of a RECT structure
inline int RECTHEIGHT(const RECT& rc)
{
	return (rc.bottom - rc.top);
}

class CSIFPreviewHandler : public IObjectWithSite,
	public IPreviewHandler,
	public IOleWindow,
	public IInitializeWithFile
{

public:
	char* sz_fileName = nullptr;
	vector<double> xData, yData;

	ID2D1Factory* pFactory = nullptr;
	ID2D1HwndRenderTarget* pRenderTarget = nullptr;
	ID2D1Bitmap* pBitmap = nullptr;
	IDWriteFactory* pDWriteFactory = nullptr;
	IDWriteTextFormat* pTextFormat = nullptr;
	IDWriteTextLayout* pTextLayout = nullptr;

	ID2D1SolidColorBrush* pWhiteBrush = nullptr;
	ID2D1SolidColorBrush* pBlackBrush = nullptr;
	ID2D1SolidColorBrush* pBlueBrush = nullptr;
	ID2D1SolidColorBrush* pInfoBrush = nullptr;

	AT_U32 atu32_ret = 0, atu32_noFrames = 0, atu32_frameSize = 0, atu32_noSubImages = 0, s_width = 0, s_height = 0;
	AT_U32 atu32_left = 0, atu32_bottom = 0, atu32_right = 0, atu32_top = 0, atu32_hBin = 0, atu32_vBin = 0;
	double ExposureTime = 0, SlitWidth = 0, GratLines = 0, CenterWavelength = 0;

	// Constructor
	CSIFPreviewHandler() : _cRef(1), _hwndParent(NULL), _hwndPreview(NULL), _punkSite(NULL)
	{
		// Initialize _rcParent to an initial value (for example, zero size rectangle)
		_rcParent = RECT{ 0, 0, 0, 0 };

		// Initialize Direct2D factory
		D2D1CreateFactory(D2D1_FACTORY_TYPE_SINGLE_THREADED, &pFactory);

		// Initialize DirectWrite
		HRESULT hr = DWriteCreateFactory(
			DWRITE_FACTORY_TYPE_SHARED, // Use DWRITE_FACTORY_TYPE_SHARED for a shared factory
			__uuidof(IDWriteFactory),
			reinterpret_cast<IUnknown**>(&pDWriteFactory)
		);
	}

	// Destructor
	virtual ~CSIFPreviewHandler()
	{
		if (_hwndPreview)
		{
			DestroyWindow(_hwndPreview);
		}

		delete[] sz_fileName;
		sz_fileName = nullptr;
		xData.clear();
		yData.clear();

		SafeRelease(pFactory);
		SafeRelease(pRenderTarget);
		SafeRelease(pBitmap);
		SafeRelease(pDWriteFactory);
		SafeRelease(pTextFormat);
		SafeRelease(pTextLayout);

		SafeRelease(pWhiteBrush);
		SafeRelease(pBlackBrush);
		SafeRelease(pBlueBrush);

		SafeRelease(_punkSite);
	}

	// IUnknown
	IFACEMETHODIMP QueryInterface(REFIID riid, void** ppv)
	{
		*ppv = NULL;
		static const QITAB qit[] =
		{
			QITABENT(CSIFPreviewHandler, IObjectWithSite),
			QITABENT(CSIFPreviewHandler, IOleWindow),
			QITABENT(CSIFPreviewHandler, IInitializeWithFile),
			QITABENT(CSIFPreviewHandler, IPreviewHandler),
			{ 0 },
		};
		return QISearch(this, qit, riid, ppv);
	}

	IFACEMETHODIMP_(ULONG) AddRef()
	{
		return InterlockedIncrement(&_cRef);
	}

	IFACEMETHODIMP_(ULONG) Release()
	{
		ULONG cRef = InterlockedDecrement(&_cRef);
		if (!cRef)
		{
			delete this;
		}
		return cRef;
	}

	// IObjectWithSite
	IFACEMETHODIMP SetSite(IUnknown* punkSite);
	IFACEMETHODIMP GetSite(REFIID riid, void** ppv);

	// IPreviewHandler
	IFACEMETHODIMP SetWindow(HWND hwnd, const RECT* prc);
	IFACEMETHODIMP SetFocus();
	IFACEMETHODIMP QueryFocus(HWND* phwnd);
	IFACEMETHODIMP TranslateAccelerator(MSG* pmsg);
	IFACEMETHODIMP SetRect(const RECT* prc);
	IFACEMETHODIMP DoPreview();
	IFACEMETHODIMP Unload();

	// IOleWindow
	IFACEMETHODIMP GetWindow(HWND* phwnd);
	IFACEMETHODIMP ContextSensitiveHelp(BOOL fEnterMode);

	// IInitializeWithFile
	IFACEMETHODIMP Initialize(LPCWSTR pszFilePath, DWORD);

	// Custom
	ID2D1Bitmap* MyCreateBitmap(vector<double> cData, double lPctl, double uPctl);
	void ReadSIFData(HWND hwnd, char* filename);
	LRESULT CALLBACK PreviewWindowSubclassProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam, UINT_PTR uIdSubclass, DWORD_PTR dwRefData);
	static LRESULT CALLBACK PreviewWindowStaticSubclassProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam, UINT_PTR uIdSubclass, DWORD_PTR dwRefData);

private:
	HRESULT _CreatePreviewWindow();

	long        _cRef;            // Reference count of this object
	HWND        _hwndParent;      // parent window that hosts the previewer window; do NOT DestroyWindow this
	RECT        _rcParent;        // bounding rect of the parent window
	HWND        _hwndPreview;     // the actual previewer window
	IUnknown* _punkSite;          // site pointer from host, used to get _pFrame
};

// IPreviewHandler
// This method gets called when the previewer gets created
HRESULT CSIFPreviewHandler::SetWindow(HWND hwnd, const RECT* prc)
{
	if (hwnd && prc)
	{
		_hwndParent = hwnd; // cache the HWND for later use
		_rcParent = *prc; // cache the RECT for later use

		if (_hwndPreview)
		{
			// Update preview window parent and rect information
			SetParent(_hwndPreview, _hwndParent);
			SetWindowPos(_hwndPreview, NULL, _rcParent.left, _rcParent.top,
				RECTWIDTH(_rcParent), RECTHEIGHT(_rcParent), SWP_NOMOVE | SWP_NOZORDER | SWP_NOACTIVATE);
		}
	}
	return S_OK;
} // SetWindow

// This method is called to set the keyboard focus
HRESULT CSIFPreviewHandler::SetFocus()
{
	HRESULT hr = S_FALSE;
	if (_hwndPreview)
	{
		::SetFocus(_hwndPreview);
		hr = S_OK;
	}
	return hr;
} // SetFocus

// This method is called to query the current keyboard focus
HRESULT CSIFPreviewHandler::QueryFocus(HWND* phwnd)
{
	HRESULT hr = E_INVALIDARG;
	if (phwnd)
	{
		*phwnd = ::GetFocus();
		if (*phwnd)
		{
			hr = S_OK;
		}
		else
		{
			hr = HRESULT_FROM_WIN32(GetLastError());
		}
	}
	return hr;
} // QueryFocus

// This method forwards accelerator messages to the preview handler frame
HRESULT CSIFPreviewHandler::TranslateAccelerator(MSG* pmsg)
{
	HRESULT hr = S_FALSE;
	IPreviewHandlerFrame* pFrame = NULL;
	if (_punkSite && SUCCEEDED(_punkSite->QueryInterface(&pFrame)))
	{
		hr = pFrame->TranslateAccelerator(pmsg);
		SafeRelease(pFrame);
	}
	return hr;
} // TranslateAccelerator

// This method gets called when the size of the previewer window changes (user resizes the Reading Pane)
HRESULT CSIFPreviewHandler::SetRect(const RECT* prc)
{
	HRESULT hr = E_INVALIDARG;
	if (prc)
	{
		_rcParent = *prc;
		if (_hwndPreview)
		{
			// Preview window is already created, so set its size and position
			SetWindowPos(_hwndPreview, NULL, _rcParent.left, _rcParent.top,
				RECTWIDTH(_rcParent), RECTHEIGHT(_rcParent), SWP_NOMOVE | SWP_NOZORDER | SWP_NOACTIVATE);
		}
		hr = S_OK;
	}

	return hr;
} // SetRect

// The main method that renders graphics
HRESULT CSIFPreviewHandler::DoPreview()
{
	HRESULT hr = E_FAIL;
	if (_hwndPreview == NULL && sz_fileName)
	{
		hr = _CreatePreviewWindow();
	}
	return hr;
} // DoPreview

// This method gets called when a shell item is de-selected in the listview
HRESULT CSIFPreviewHandler::Unload()
{
	if (_hwndPreview)
	{
		// Remove the window subclass
		RemoveWindowSubclass(_hwndPreview, PreviewWindowStaticSubclassProc, 0);

		// Destroy the preview window
		DestroyWindow(_hwndPreview);
		_hwndPreview = NULL;
	}

	delete[] sz_fileName;
	sz_fileName = nullptr;
	xData.clear();
	yData.clear();

	SafeRelease(pRenderTarget);
	SafeRelease(pBitmap);
	SafeRelease(pTextFormat);
	SafeRelease(pTextLayout);

	SafeRelease(pWhiteBrush);
	SafeRelease(pBlackBrush);
	SafeRelease(pBlueBrush);

	SafeRelease(_punkSite);

	return S_OK;
} // Unload

// IObjectWithSite methods
HRESULT CSIFPreviewHandler::SetSite(IUnknown* punkSite)
{
	SafeRelease(_punkSite);
	return punkSite ? punkSite->QueryInterface(&_punkSite) : S_OK;
} // SetSite

HRESULT CSIFPreviewHandler::GetSite(REFIID riid, void** ppv)
{
	*ppv = NULL;
	return _punkSite ? _punkSite->QueryInterface(riid, ppv) : E_FAIL;
} // GetSite

// IOleWindow methods
HRESULT CSIFPreviewHandler::GetWindow(HWND* phwnd)
{
	HRESULT hr = E_INVALIDARG;
	if (phwnd)
	{
		*phwnd = _hwndParent;
		hr = S_OK;
	}
	return hr;
} // GetWindow

HRESULT CSIFPreviewHandler::ContextSensitiveHelp(BOOL)
{
	return E_NOTIMPL;
} // ContextSensitiveHelp

// IInitializeWithFile methods
// This method gets called when an item gets selected in listview
HRESULT CSIFPreviewHandler::Initialize(LPCWSTR pszFilePath, DWORD)
{
	if (pszFilePath)
	{
		// Convert wide-character path to UTF-8 and assign to sz_fileName
		char tempFileName[MAX_PATH];
		int bufferSize = WideCharToMultiByte(CP_UTF8, 0, pszFilePath, -1, tempFileName, MAX_PATH, NULL, NULL);
		if (bufferSize > 0 && bufferSize <= MAX_PATH)
		{
			sz_fileName = new char[bufferSize];
			strcpy_s(sz_fileName, bufferSize, tempFileName);
			return S_OK;
		}
	}
	return E_FAIL;
} // Initialize

vector<double> calculatePercentiles(const vector<double>& data, const vector<double>& percentiles);

/**
 * MyCreateBitmap: Creates a Direct2D bitmap based on provided data and percentiles.
 *
 * This function creates a Direct2D bitmap using the given width, height, data vector, and percentiles
 * to map the data to color intensities. The resulting bitmap reflects the intensity of data values
 * within the specified percentile range.
 *
 * @param cData The data vector containing values to be mapped to color intensities.
 * @param lPctl The lower percentile value for data mapping.
 * @param uPctl The upper percentile value for data mapping.
 *
 * @return A pointer to the created ID2D1Bitmap. The caller is responsible for releasing this resource.
 *
 * Usage:
 * Call this function to create a Direct2D bitmap based on the provided data vector and percentile
 * values for mapping. The resulting bitmap will reflect the color intensities corresponding to the
 * data values within the specified percentile range. The returned ID2D1Bitmap pointer can be used
 * for rendering. Make sure to release the bitmap resource when it is no longer needed.
 *
 * Example:
 *   AT_U32 width = 800;
 *   AT_U32 height = 600;
 *   vector<double> data = { ... }; // Replace with actual data
 *   double lowerPercentile = 10.0;
 *   double upperPercentile = 90.0;
 *   ID2D1Bitmap* bitmap = MyCreateBitmap(data, lowerPercentile, upperPercentile);
 *   // The 'bitmap' can be used for rendering.
 *   // Don't forget to release the 'bitmap' when done.
 *   SafeRelease(bitmap);
 */
ID2D1Bitmap* CSIFPreviewHandler::MyCreateBitmap(vector<double> cData, double lPctl, double uPctl) {
	// Find c data range
	//double minC = *min_element(cData.begin(), cData.end());
	//double maxC = *max_element(cData.begin(), cData.end());
	vector<double> calculatedPercentiles = calculatePercentiles(cData, { lPctl, uPctl });
	double minC = calculatedPercentiles[0];
	double maxC = calculatedPercentiles[1];

	auto mapC = [&](double c) {
		return static_cast<BYTE>(max(min((static_cast<double>(c - minC) / (maxC - minC)) * 255.0, 255), 0));
		};

	D2D1_SIZE_U bitmapSize = D2D1::SizeU(static_cast<UINT32>(s_width), static_cast<UINT32>(s_height));

	D2D1_BITMAP_PROPERTIES bitmapProps = D2D1::BitmapProperties(
		D2D1::PixelFormat(DXGI_FORMAT_B8G8R8A8_UNORM, D2D1_ALPHA_MODE_IGNORE));

	pRenderTarget->CreateBitmap(bitmapSize, bitmapProps, &pBitmap);

	D2D1_RECT_U rect = D2D1::RectU(0, 0, bitmapSize.width, bitmapSize.height);

	// Calculate pixel values and create the bitmap
	BYTE* pColors = new BYTE[bitmapSize.width * bitmapSize.height * 4];
	for (UINT y = 0; y < bitmapSize.height; ++y) {
		for (UINT x = 0; x < bitmapSize.width; ++x) {
			for (UINT i = 0; i < 3; ++i) {
				pColors[(bitmapSize.height - y - 1) * bitmapSize.width * 4 + x * 4 + i]
					= mapC(cData[y * bitmapSize.width + x]); // Calculate intensity based on cData[ind]
			}
		}
	}

	pBitmap->CopyFromMemory(&rect, pColors, bitmapSize.width * sizeof(BYTE) * 4);
	delete[] pColors;
	return pBitmap;
} // MyCreateBitmap

/**
 * ShowErrorMessage: Displays an error message in a message box.
 *
 * This function shows an error message using a message box with an error icon.
 * The provided message is displayed to the user.
 *
 * @param hwnd The handle to the parent window. The message box will be modal to this window.
 * @param message The error message text to be displayed.
 *
 * Usage:
 * Call this function to display error messages in a message box. The message box will be modal to
 * the specified parent window. The provided error message will be shown to the user with an error icon.
 *
 * Example:
 *   HWND mainWindow = GetMainWindowHandle(); // Replace with your actual window handle
 *   std::string errorMsg = "An error occurred while processing data.";
 *   ShowErrorMessage(mainWindow, errorMsg);
 *   // The message box will display the error message to the user.
 */
void ShowErrorMessage(HWND hwnd, const std::string& message) {
	MessageBoxA(hwnd, message.c_str(), "Error", MB_ICONERROR | MB_OK);
} // ShowErrorMessage

/**
 * ReadSIFData: Reads data from an SIF file and populates xData and yData vectors.
 *
 * This function reads image and wavelength data from the specified SIF file using the ATSIF library.
 * It populates the global vectors xData (for wavelength data) and yData (for image data).
 *
 * @param hwnd The handle to the parent window, used for displaying error messages.
 * @param filename The name of the SIF file to read data from.
 *
 * Usage:
 * Call this function with the parent window handle and the name of the SIF file to read data from.
 * It performs various ATSIF library calls to retrieve frame and wavelength data. In case of errors
 * at any step, error messages are shown using message boxes. The data is then converted to vectors
 * and stored in the global xData and yData vectors for further processing.
 *
 * Example:
 *   HWND hwnd = GetParentWindow(); // Replace with an appropriate function to get the parent window handle.
 *   char* sifFilename = "example.sif";
 *   ReadSIFData(hwnd, sifFilename);
 *   // After this call, xData and yData vectors will contain the wavelength and image data.
 */
void CSIFPreviewHandler::ReadSIFData(HWND hwnd, char* filename) {
	char* sz_error = new char[MAX_PATH];
	float* imageBuffer = nullptr;
	double* wavelengthBuffer = nullptr;
	char* sz_exposureTime = new char[MAX_PATH];
	char* sz_slitWidth = new char[MAX_PATH];
	char* sz_gratLines = new char[MAX_PATH];
	char* sz_centerWavelength = new char[MAX_PATH];

	// Set file access mode
	atu32_ret = ATSIF_SetFileAccessMode(ATSIF_ReadAll);
	if (atu32_ret != ATSIF_SUCCESS) {
		ShowErrorMessage(hwnd, "Could not set File access Mode.");
		goto Cleanup;
	}

	// Read from file
	atu32_ret = ATSIF_ReadFromFile(filename);
	if (atu32_ret != ATSIF_SUCCESS) {
		ShowErrorMessage(hwnd, "Could not open File: " + std::string(filename) + ". Error: " + std::to_string(atu32_ret));
		goto Cleanup;
	}

	// Get number of frames
	atu32_ret = ATSIF_GetNumberFrames(ATSIF_Signal, &atu32_noFrames);
	if (atu32_ret != ATSIF_SUCCESS) {
		ShowErrorMessage(hwnd, "Could not Get Number Frames. Error: " + std::to_string(atu32_ret));
		goto Cleanup;
	}

	// Get frame size
	atu32_ret = ATSIF_GetFrameSize(ATSIF_Signal, &atu32_frameSize);
	if (atu32_ret != ATSIF_SUCCESS) {
		ShowErrorMessage(hwnd, "Could not Get Frame Size. Error: " + std::to_string(atu32_ret));
		goto Cleanup;
	}

	// Get number of sub-images
	atu32_ret = ATSIF_GetNumberSubImages(ATSIF_Signal, &atu32_noSubImages);
	if (atu32_ret != ATSIF_SUCCESS) {
		ShowErrorMessage(hwnd, "Could not Get Number Sub Images. Error: " + std::to_string(atu32_ret));
		goto Cleanup;
	}

	// Get sub-image info
	atu32_ret = ATSIF_GetSubImageInfo(ATSIF_Signal,
		0,
		&atu32_left, &atu32_bottom,
		&atu32_right, &atu32_top,
		&atu32_hBin, &atu32_vBin);
	s_width = ((atu32_right - atu32_left) + 1) / atu32_hBin;
	s_height = ((atu32_top - atu32_bottom) + 1) / atu32_vBin;
	if (atu32_ret != ATSIF_SUCCESS) {
		ShowErrorMessage(hwnd, "Could not Get Sub Image Info. Error: " + std::to_string(atu32_ret));
		goto Cleanup;
	}

	// Allocate memory for image buffer and wavelength buffer
	imageBuffer = new float[atu32_frameSize];
	wavelengthBuffer = new double[s_width];

	// Get image data
	atu32_ret = ATSIF_GetFrame(ATSIF_Signal, 0, imageBuffer, atu32_frameSize);
	if (atu32_ret != ATSIF_SUCCESS) {
		ShowErrorMessage(hwnd, "Could not Get Frame. Error: " + std::to_string(atu32_ret));
		goto Cleanup;
	}

	// Get wavelength data
	for (unsigned int i = 0; i < s_width; ++i) {
		atu32_ret = ATSIF_GetPixelCalibration(ATSIF_Signal, ATSIF_CalibX, static_cast<AT_32>(i) * atu32_hBin + atu32_left, &wavelengthBuffer[i]);
		if (atu32_ret != ATSIF_SUCCESS) {
			ShowErrorMessage(hwnd, "Could not Get Wavelength. Error: " + std::to_string(atu32_ret));
			goto Cleanup;
		}
	}

	// Get exposure time
	atu32_ret = ATSIF_GetPropertyValue(ATSIF_Signal, "ExposureTime", sz_exposureTime, MAX_PATH);
	if (atu32_ret != ATSIF_SUCCESS) {
		ShowErrorMessage(hwnd, "Could not Get Property Value. Error: " + std::to_string(atu32_ret));
		goto Cleanup;
	}
	ExposureTime = atof(sz_exposureTime);

	// Get slit width
	atu32_ret = ATSIF_GetPropertyValue(ATSIF_Signal, "SpectrographSlit1", sz_slitWidth, MAX_PATH);
	if (atu32_ret != ATSIF_SUCCESS) {
		ShowErrorMessage(hwnd, "Could not Get Property Value. Error: " + std::to_string(atu32_ret));
		goto Cleanup;
	}
	SlitWidth = atof(sz_slitWidth);

	// Get grating lines
	atu32_ret = ATSIF_GetPropertyValue(ATSIF_Signal, "SpectrographGratLines", sz_gratLines, MAX_PATH);
	if (atu32_ret != ATSIF_SUCCESS) {
		ShowErrorMessage(hwnd, "Could not Get Property Value. Error: " + std::to_string(atu32_ret));
		goto Cleanup;
	}
	GratLines = atof(sz_gratLines);

	// Get center wavelength
	atu32_ret = ATSIF_GetPropertyValue(ATSIF_Signal, "SpectrographWavelength", sz_centerWavelength, MAX_PATH);
	if (atu32_ret != ATSIF_SUCCESS) {
		ShowErrorMessage(hwnd, "Could not Get Property Value. Error: " + std::to_string(atu32_ret));
		goto Cleanup;
	}
	CenterWavelength = atof(sz_centerWavelength);

	// Close file
	atu32_ret = ATSIF_CloseFile();
	if (atu32_ret != ATSIF_SUCCESS) {
		ShowErrorMessage(hwnd, "Could not Close File. Error: " + std::to_string(atu32_ret));
		goto Cleanup;
	}

	// Convert arrays to double vectors
	for (size_t i = 0; i < s_width; ++i) {
		xData.push_back(wavelengthBuffer[i]);
	}
	for (size_t i = 0; i < atu32_frameSize; ++i) {
		yData.push_back(imageBuffer[i]);
	}

Cleanup:
	// Cleanup resources
	delete[] imageBuffer;
	delete[] wavelengthBuffer;
	delete[] sz_error;
	delete[] sz_exposureTime;
	delete[] sz_slitWidth;
	delete[] sz_gratLines;
	delete[] sz_centerWavelength;
} // ReadSIFData

// This method forwards window messages to the member function PreviewWindowSubclassProc
LRESULT CALLBACK CSIFPreviewHandler::PreviewWindowStaticSubclassProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam, UINT_PTR uIdSubclass, DWORD_PTR dwRefData)
{
	CSIFPreviewHandler* pThis = reinterpret_cast<CSIFPreviewHandler*>(dwRefData);

	// Forward the call to the member function
	return pThis->PreviewWindowSubclassProc(hwnd, uMsg, wParam, lParam, uIdSubclass, dwRefData);
}

wstring to_wstring_custom(double value);
double roundDigits(double value, int digits);

// This method handles various window messages, including resizing and painting
LRESULT CALLBACK CSIFPreviewHandler::PreviewWindowSubclassProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam, UINT_PTR /* uIdSubclass */, DWORD_PTR dwRefData)
{
	CSIFPreviewHandler* pThis = reinterpret_cast<CSIFPreviewHandler*>(dwRefData);

	HDC          hdc;
	PAINTSTRUCT  ps;

	switch (uMsg)
	{
	case WM_SIZE:
		if (pRenderTarget)
		{
			int windowWidth = LOWORD(lParam);
			int windowHeight = HIWORD(lParam);
			pRenderTarget->Resize(D2D1::SizeU(windowWidth, windowHeight));
		}
		return 0;
	case WM_PAINT:
	{
		hdc = BeginPaint(hwnd, &ps);
		if (pRenderTarget)
		{
			// Begin drawing with Direct2D
			pRenderTarget->BeginDraw();

			// Draw the white background
			pRenderTarget->Clear(whiteColor);

			// Get the dimensions of the window to draw in
			int windowwidth = pRenderTarget->GetSize().width;
			int windowheight = pRenderTarget->GetSize().height;

			D2D1_RECT_F plotRect = D2D1::RectF(lmargin, tmargin,
				max(lmargin + minimumaxisbox, windowwidth - rmargin),
				max(tmargin + minimumaxisbox, windowheight - bmargin));

			// Use Direct2D to render the bitmap
			if (pBitmap)
				pRenderTarget->DrawBitmap(pBitmap, plotRect);
			pRenderTarget->DrawRectangle(&plotRect, pBlackBrush, 2.0f);

			// Find x data range
			double minX = *min_element(xData.begin(), xData.end());
			double maxX = *max_element(xData.begin(), xData.end());
			double minY, maxY;

			// Calculate the size of the axis box rectangle
			int width = max(minimumaxisbox, windowwidth - lmargin - rmargin);
			int height = max(minimumaxisbox, windowheight - tmargin - bmargin);

			// Map data coordinates to plot coordinates
			auto mapX = [&](double x) {
				return lmargin + (x - minX) / (maxX - minX) * width;
				};

			auto mapY = [&](double y) {
				return height + tmargin - (y - minY) / (maxY - minY) * height;
				};

			if (s_height == 1) {
				// Find y data range
				minY = *min_element(yData.begin(), yData.end());
				maxY = *max_element(yData.begin(), yData.end());

				// Draw data points and connect with lines
				for (size_t i = 0; i < xData.size() - 1; ++i) {
					int x1 = static_cast<int>(mapX(xData[i]));
					int y1 = static_cast<int>(mapY(yData[i]));
					int x2 = static_cast<int>(mapX(xData[i + 1]));
					int y2 = static_cast<int>(mapY(yData[i + 1]));
					pRenderTarget->DrawLine(D2D1::Point2F(x1, y1), D2D1::Point2F(x2, y2), pBlueBrush, 1.0f);
				}
			}
			else {
				// Find y data range
				minY = atu32_bottom;
				maxY = atu32_top;

				// White on black background for colormap
				pInfoBrush = pWhiteBrush;
			}

			// Arrays of doubles for strings and arrays of integers for x coordinates
			double Xvalues[] = { minX, maxX };
			int XxCoordinates[] = { lmargin, width + lmargin };
			int XyCoordinates[] = { height + tmargin + labelmargin, height + tmargin + labelmargin };

			// Arrays of doubles for strings and arrays of integers for y coordinates
			double Yvalues[] = { minY, maxY };
			int YxCoordinates[] = { lmargin - labelmargin, lmargin - labelmargin };
			int YyCoordinates[] = { height + tmargin, tmargin };

			pTextFormat->SetTextAlignment(DWRITE_TEXT_ALIGNMENT_CENTER);

			// Draw the x strings
			for (size_t i = 0; i < sizeof(Xvalues) / sizeof(Xvalues[0]); ++i) {
				WCHAR text[32]; // Assuming a reasonable buffer size
				double value = Xvalues[i];
				if (floor(value) == value) {
					swprintf_s(text, L"%.0lf", value); // No decimal places for integers
				}
				else {
					wchar_t formatString[10];
					swprintf_s(formatString, L"%%.%dlf", decimalPlacesX); // Create the format string
					swprintf_s(text, formatString, roundDigits(value, decimalPlacesX)); // Use the format string
				}

				// Create the text layout
				pDWriteFactory->CreateTextLayout(
					text,
					wcslen(text),
					pTextFormat,
					maximumtextbox,
					fontsize,
					&pTextLayout
				);

				// Set text alignment to center
				pTextLayout->SetTextAlignment(DWRITE_TEXT_ALIGNMENT_CENTER);

				// Draw the centered text
				pRenderTarget->DrawTextLayout(
					D2D1::Point2F(static_cast<FLOAT>(XxCoordinates[i]) - (maximumtextbox / 2.0f), static_cast<FLOAT>(XyCoordinates[i])),
					pTextLayout,
					pBlackBrush
				);

				// Release the text layout
				SafeRelease(pTextLayout);
			}

			// Draw the center wavelength
			if (CenterWavelength != -1) {
				WCHAR text[32]; // Assuming a reasonable buffer size
				double value = CenterWavelength;
				if (floor(value) == value) {
					swprintf_s(text, L"%.0lf nm", value); // No decimal places for integers
				}
				else {
					wchar_t formatString[10];
					swprintf_s(formatString, L"%%.%dlf nm", decimalPlacesX); // Create the format string
					swprintf_s(text, formatString, roundDigits(value, decimalPlacesX)); // Use the format string
				}

				// Create the text layout
				pDWriteFactory->CreateTextLayout(
					text,
					wcslen(text),
					pTextFormat,
					maximumtextbox,
					fontsize,
					&pTextLayout
				);

				// Set text alignment to center
				pTextLayout->SetTextAlignment(DWRITE_TEXT_ALIGNMENT_CENTER);

				// Draw the centered text
				pRenderTarget->DrawTextLayout(
					D2D1::Point2F(static_cast<FLOAT>(XxCoordinates[0] + XxCoordinates[1]) / 2.0f - (maximumtextbox / 2.0f), static_cast<FLOAT>(XyCoordinates[1])),
					pTextLayout,
					pBlackBrush
				);

				// Release the text layout
				SafeRelease(pTextLayout);
			}

			// Draw the y strings
			for (size_t i = 0; i < sizeof(Yvalues) / sizeof(Yvalues[0]); ++i) {
				WCHAR text[32]; // Assuming a reasonable buffer size
				double value = Yvalues[i];
				if (floor(value) == value) {
					swprintf_s(text, L"%.0lf", value); // No decimal places for integers
				}
				else {
					wchar_t formatString[10];
					swprintf_s(formatString, L"%%.%dlf", decimalPlacesY); // Create the format string
					swprintf_s(text, formatString, roundDigits(value, decimalPlacesY)); // Use the format string
				}

				// Create the text layout
				pDWriteFactory->CreateTextLayout(
					text,
					wcslen(text),
					pTextFormat,
					maximumtextbox,
					fontsize,
					&pTextLayout
				);

				// Calculate the text height
				DWRITE_TEXT_METRICS textMetrics;
				pTextLayout->GetMetrics(&textMetrics);
				float textHeight = textMetrics.height;

				// Set text alignment to center
				pTextLayout->SetTextAlignment(DWRITE_TEXT_ALIGNMENT_TRAILING);

				// Draw the centered text
				pRenderTarget->DrawTextLayout(
					D2D1::Point2F(static_cast<FLOAT>(YxCoordinates[i]) - maximumtextbox, static_cast<FLOAT>(YyCoordinates[i] - textHeight / 2.0f)),
					pTextLayout,
					pBlackBrush
				);

				// Release the text layout
				SafeRelease(pTextLayout);
			}

			// Display details only if the window is big enough
			if ((width > maximumtextbox) && (height > maximumtextboxheight)) {
				vector<wstring> labels = { L"Frames:", L"Sub Images:", L"Pixels:", L"Exposure Time:", L"Slit:", L"Grating:" };
				vector<wstring> values = {
					to_wstring(atu32_noFrames),
					to_wstring(atu32_noSubImages),
					to_wstring(s_width) + L" x " + to_wstring(s_height),
					to_wstring_custom(ExposureTime) + L" s",
					(SlitWidth == -1 ? L"N/A" : to_wstring_custom(SlitWidth) + L" �m"),
					(GratLines == -1 ? L"N/A" : to_wstring_custom(GratLines) + L" mm" + std::wstring(1, L'\u207B') + std::wstring(1, L'\u00B9'))
				};

				for (size_t i = 0; i < labels.size(); ++i) {
					// Create text layout for the label (left-aligned)
					pDWriteFactory->CreateTextLayout(
						labels[i].c_str(),
						static_cast<UINT32>(labels[i].length()),
						pTextFormat,
						maximumtextbox,
						fontsize,
						&pTextLayout
					);

					// Set text alignment to leading
					pTextLayout->SetTextAlignment(DWRITE_TEXT_ALIGNMENT_LEADING);

					// Draw the label
					pRenderTarget->DrawTextLayout(
						D2D1::Point2F(static_cast<FLOAT>(width + lmargin - maximumtextbox + labelmargin), static_cast<FLOAT>(tmargin + labelmargin + (fontsize + 2) * i)),
						pTextLayout,
						pInfoBrush
					);

					// Release the text layout
					SafeRelease(pTextLayout);

					// Create text layout for the value (right-aligned)
					pDWriteFactory->CreateTextLayout(
						values[i].c_str(),
						static_cast<UINT32>(values[i].length()),
						pTextFormat,
						maximumtextbox,
						fontsize,
						&pTextLayout
					);

					// Set text alignment to trailing (to align value to the right)
					pTextLayout->SetTextAlignment(DWRITE_TEXT_ALIGNMENT_TRAILING);

					// Draw the value
					pRenderTarget->DrawTextLayout(
						D2D1::Point2F(static_cast<FLOAT>(width + lmargin - maximumtextbox - labelmargin), static_cast<FLOAT>(tmargin + labelmargin + (fontsize + 2) * i)),
						pTextLayout,
						pInfoBrush
					);

					// Release the text layout
					SafeRelease(pTextLayout);
				}
			}

			pRenderTarget->EndDraw();
			// End drawing with Direct2D
		}
		EndPaint(hwnd, &ps);
	}
	return 0;
	case WM_DESTROY:
		return 0;
	default:
		return DefWindowProc(hwnd, uMsg, wParam, lParam);
	}
} // PreviewWindowSubclassProc

// This method creates the preview window and sets up Direct2D rendering and text formatting
HRESULT CSIFPreviewHandler::_CreatePreviewWindow()
{
	// Create the preview window
	_hwndPreview = CreateWindowExW(0, L"STATIC", NULL,
		WS_CHILD | WS_VISIBLE,
		_rcParent.left, _rcParent.top, RECTWIDTH(_rcParent), RECTHEIGHT(_rcParent),
		_hwndParent, NULL, NULL, this); // Pass the instance pointer as lParam

	if (_hwndPreview)
	{
		// Set the static subclass procedure
		SetWindowSubclass(_hwndPreview, PreviewWindowStaticSubclassProc, 0, reinterpret_cast<DWORD_PTR>(this));

		pDWriteFactory->CreateTextFormat(
			L"Segoe UI",          // Font family name
			nullptr,              // Font collection (nullptr for the system font collection)
			DWRITE_FONT_WEIGHT_NORMAL,
			DWRITE_FONT_STYLE_NORMAL,
			DWRITE_FONT_STRETCH_NORMAL,
			static_cast<FLOAT>(fontsize), // Font size
			L"",                  // Locale name (empty string for the default locale)
			&pTextFormat
		);

		RECT rc;
		GetClientRect(_hwndPreview, &rc);

		// Create render target
		pFactory->CreateHwndRenderTarget(
			D2D1::RenderTargetProperties(),
			D2D1::HwndRenderTargetProperties(_hwndPreview, D2D1::SizeU(rc.right, rc.bottom)),
			&pRenderTarget);

		// Create a SolidColorBrush
		pRenderTarget->CreateSolidColorBrush(whiteColor, &pWhiteBrush);
		pRenderTarget->CreateSolidColorBrush(blackColor, &pBlackBrush);
		pRenderTarget->CreateSolidColorBrush(blueColor, &pBlueBrush);
		pInfoBrush = pBlackBrush;

		ReadSIFData(_hwndPreview, sz_fileName);

		if (s_height != 1) {
			pBitmap = MyCreateBitmap(yData, 1, 99);
		}

		ShowWindow(_hwndPreview, SW_SHOW);
		UpdateWindow(_hwndPreview);

		return S_OK;
	}

	return E_FAIL;
} // _CreatePreviewWindow

// This method is a factory function for creating an instance of the CSIFPreviewHandler class
HRESULT CSIFPreviewHandler_CreateInstance(REFIID riid, void** ppv)
{
	*ppv = NULL;

	CSIFPreviewHandler* pNew = new (std::nothrow) CSIFPreviewHandler();
	HRESULT hr = pNew ? S_OK : E_OUTOFMEMORY;
	if (SUCCEEDED(hr))
	{
		hr = pNew->QueryInterface(riid, ppv);
		pNew->Release();
	}
	return hr;
} // CSIFPreviewHandler_CreateInstance

/**
 * calculatePercentiles: Calculates percentiles from a vector of double values.
 *
 * This function takes a vector of double values "data" and another vector "percentiles" containing
 * the desired percentiles (in percent) to calculate. It returns a vector of calculated percentiles
 * corresponding to the provided data.
 *
 * @param data The input vector of double values for which percentiles are to be calculated.
 * @param percentiles A vector of double values representing the desired percentiles (in percent)
 *                   to be calculated from the data.
 * @return A vector of double values containing the calculated percentiles corresponding to the input data.
 *
 * Example usage:
 * \code{.cpp}
 * std::vector<double> data = { 12.5, 9.7, 6.8, 8.4, 15.2, 20.0, 18.3, 10.1, 7.5 };
 * std::vector<double> percentiles = { 25.0, 50.0, 75.0 };
 * std::vector<double> calculatedPercentiles = calculatePercentiles(data, percentiles);
 * for (size_t i = 0; i < percentiles.size(); ++i) {
 *     std::cout << "Percentile " << percentiles[i] << ": " << calculatedPercentiles[i] << std::endl;
 * }
 * \endcode
 */
vector<double> calculatePercentiles(const vector<double>& data, const vector<double>& percentiles) {
	vector<double> result;

	// Sort the data array
	vector<double> sortedData = data;
	sort(sortedData.begin(), sortedData.end());

	// Calculate the index for each requested percentile
	for (double p : percentiles) {
		if (p < 0.0 || p > 100.0) {
			cerr << "Error: Invalid percentile value " << p << ". Percentile values should be in the range [0, 100]." << endl;
			continue;
		}

		double index = (p / 100.0) * (sortedData.size() - 1);
		int lowerIndex = static_cast<int>(floor(index));
		int upperIndex = static_cast<int>(ceil(index));

		if (lowerIndex == upperIndex) {
			result.push_back(sortedData[lowerIndex]);
		}
		else {
			double lowerValue = sortedData[lowerIndex];
			double upperValue = sortedData[upperIndex];
			double interpolatedValue = lowerValue + (index - lowerIndex) * (upperValue - lowerValue);
			result.push_back(interpolatedValue);
		}
	}
	return result;
} // calculatePercentiles

/**
 * to_wstring_custom: Converts a double value to a wstring with trailing zeros removed.
 *
 * This function converts a double value to a wstring representation, and it removes trailing zeros
 * from the converted string.
 *
 * @param value The double value to be converted to a wstring.
 * @return A wstring containing the converted value without trailing zeros.
 *
 * Example usage:
 * \code{.cpp}
 * double value = 12.345;
 * wstring result = to_wstring_custom(value);
 * wcout << result << endl;
 * \endcode
 */
wstring to_wstring_custom(double value) {
	wstringstream wss;
	wss << value;
	return wss.str();
}

/**
 * roundDigits: Rounds a double value to a specified number of digits after the decimal point.
 *
 * This function rounds a double value to the specified number of digits after the decimal point.
 * It multiplies the value by 10^digits, rounds the result to the nearest integer, and then divides
 * it by 10^digits to round the value to the specified number of digits.
 *
 * @param value The double value to be rounded.
 * @param digits The number of digits after the decimal point to round to.
 * @return The rounded double value.
 *
 * Example usage:
 * \code{.cpp}
 * double value = 12.3456789;
 * double roundedValue = roundDigits(value, 3); // Rounds 'value' to 3 decimal places
 * wcout << roundedValue << endl;
 * \endcode
 */
double roundDigits(double value, int digits) {
	double factor = pow(10, digits);
	return round(value * factor) / factor;
}