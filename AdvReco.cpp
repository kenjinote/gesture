// THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF
// ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO
// THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A
// PARTICULAR PURPOSE.
//
// Copyright (c) Microsoft Corporation. All rights reserved.
//
// Module:
//      AdvReco.cpp
//
// Description:
//      This sample demonstrates such advanced features of the
//      Microsoft Tablet PC Automation API used for handwriting
//      recognition, as
//          - enumerating the installed recognzers
//          - creating a recognition context with a specific language
//            recognizer
//          - setting recognition input scopes
//          - using guides to improve the recognition quality
//          - dynamic background recognition
//          - gesture recognition.
//
//      This application is discussed in the Getting Started guide.
//
//      (NOTE: For code simplicity, returned HRESULT is not checked
//             on failure in the places where failures are not critical
//             for the application or very unexpected)
//
//      The interfaces used are:
//      IInkRecognizers, IInkRecognizer, IInkRecoContext,
//      IInkRecognitionResult, IInkRecognitionGuide, IInkGesture
//      IInkCollector, IInkDisp, IInkRenderer, IInkStrokes, IInkStroke
//
// Requirements:
//      One or more handwriting recognizer must be installed on the system;
//      Appropriate Asian fonts need to be installed to output the results
//      of the Asian recognizers.
//
//--------------------------------------------------------------------------

#ifndef _WIN32_WINNT
#define _WIN32_WINNT 0x0500
#endif

// Windows header files
#include <windows.h>
#include <commctrl.h>       // need it to call CreateStatusWindow

// The following definitions may be not found in the old headers installed with VC6,
// so they're copied from the newer headers found in the Microsoft Platform SDK

#ifndef ListView_SetCheckState
#define ListView_SetCheckState(hwndLV, i, fCheck) \
ListView_SetItemState(hwndLV, i, INDEXTOSTATEIMAGEMASK((fCheck ? 2 : 1)), LVIS_STATEIMAGEMASK)
#endif

#ifndef ListView_GetCheckState
#define ListView_GetCheckState(hwndLV, i) \
((((UINT)(SNDMSG((hwndLV), LVM_GETITEMSTATE, (WPARAM)(i), LVIS_STATEIMAGEMASK))) >> 12) -1)
#endif

#ifndef MIIM_STRING
#define MIIM_STRING      0x00000040
#endif
#ifndef MIIM_FTYPE
#define MIIM_FTYPE       0x00000100
#endif

// A useful macro to determine the number of elements in the array
#define countof(array)  (sizeof(array)/sizeof(array[0]))

// ATL header files
#include <atlbase.h>        // defines CComModule, CComPtr, CComVariant
CComModule _Module;
#include <atlwin.h>         // defines CWindowImpl
#include <atlcom.h>         // defines IDispEventSimpleImpl

// Headers for Tablet PC Automation interfaces
#include <msinkaut.h>
#include <msinkaut_i.c>
#include <tpcerror.h>

// The application header files
#include "resource.h"       // main symbols, including command ID's
#include "EventSinks.h"     // defines the IInkEventsImpl and IInkRecognitionEventsImpl
#include "ChildWnds.h"      // definitions of the CInkInputWnd and CRecoOutputWnd
#include "AdvReco.h"        // contains the definition of CAddRecoApp

// The set of the single stroke gestures known to this application
const InkApplicationGesture gc_igtSingleStrokeGestures[] = {
    IAG_Scratchout, IAG_Triangle, IAG_Square, IAG_Star, IAG_Check,
    IAG_Circle, IAG_DoubleCircle, IAG_Curlicue, IAG_DoubleCurlicue,
    IAG_SemiCircleLeft, IAG_SemiCircleRight,
    IAG_ChevronUp, IAG_ChevronDown, IAG_ChevronLeft,
    IAG_ChevronRight, IAG_Up, IAG_Down, IAG_Left, IAG_Right, IAG_UpDown, IAG_DownUp,
    IAG_LeftRight, IAG_RightLeft, IAG_UpLeftLong, IAG_UpRightLong, IAG_DownLeftLong,
    IAG_DownRightLong, IAG_UpLeft, IAG_UpRight, IAG_DownLeft, IAG_DownRight, IAG_LeftUp,
    IAG_LeftDown, IAG_RightUp, IAG_RightDown, IAG_Tap
};

// The static members of the event sink templates are initialized here
// (defined in EventSinks.h)

const _ATL_FUNC_INFO IInkCollectorEventsImpl<CAdvRecoApp>::mc_AtlFuncInfo[1] = {
        {CC_STDCALL, VT_EMPTY, 4, {VT_UNKNOWN, VT_UNKNOWN, VT_VARIANT, VT_BOOL|VT_BYREF}}
};

const TCHAR gc_szAppName[] = TEXT("Advanced Recognition");

/////////////////////////////////////////////////////////
//
// WinMain
//
// The WinMain function is called by the system as the
// initial entry point for a Win32-based application.
//
// Parameters:
//        HINSTANCE hInstance,      : [in] handle to current instance
//        HINSTANCE hPrevInstance,  : [in] handle to previous instance
//        LPSTR lpCmdLine,          : [in] command line
//        int nCmdShow              : [in] show state
//
// Return Values (int):
//        0 : The function terminated before entering the message loop.
//        non zero: Value of the wParam when receiving the WM_QUIT message
//
/////////////////////////////////////////////////////////
int APIENTRY wWinMain(
        HINSTANCE hInstance,
        HINSTANCE /*hPrevInstance*/,   // not used here
        LPWSTR     /*lpCmdLine*/,       // not used here
        int       nCmdShow
        )
{
    int iRet = 0;

    // Initialize the COM library and the application module
    if (S_OK == ::CoInitializeEx(NULL, COINIT_APARTMENTTHREADED))
    {
        _Module.Init(NULL, hInstance);

        // Register the common control classes used by the application
        INITCOMMONCONTROLSEX icc;
        icc.dwSize = sizeof(icc);
        icc.dwICC = ICC_LISTVIEW_CLASSES | ICC_BAR_CLASSES;
        if (TRUE == ::InitCommonControlsEx(&icc))
        {
            // Call the boilerplate function of the application
            iRet = CAdvRecoApp::Run(nCmdShow);
        }
        else
        {
            ::MessageBox(NULL, TEXT("Error initializing the common controls."),
                         gc_szAppName, MB_ICONERROR | MB_OK);
        }

        // Release the module and the COM library
        _Module.Term();
        ::CoUninitialize();
    }

    return iRet;
}

/////////////////////////////////////////////////////////
//
// CAdvRecoApp::Run
//
// The static CAdvRecoApp::Run is the boilerplate of the application.
// It instantiates and initializes an CAdvRecoApp object and runs the
// application's message loop.
//
// Parameters:
//      int nCmdShow              : [in] show state
//
// Return Values (int):
//      0 : The function terminated before entering the message loop.
//      non zero: Value of the wParam when receiving the WM_QUIT message
//
/////////////////////////////////////////////////////////
int CAdvRecoApp::Run(
        int nCmdShow
        )
{

    CAdvRecoApp theApp;

    int iRet;

    // Load the icon from the resource and associate it with the window class
    WNDCLASSEX& wc = CAdvRecoApp::GetWndClassInfo().m_wc;
    wc.hIcon = wc.hIconSm = ::LoadIcon(_Module.GetResourceInstance(),
                                       MAKEINTRESOURCE(IDR_APPICON));

    // Create the application's main window
    if (theApp.Create(NULL, CWindow::rcDefault, gc_szAppName,
                      WS_OVERLAPPEDWINDOW, 0, (UINT)0) != NULL)
    {
        // Show and update the main window
        theApp.ShowWindow(nCmdShow);
        theApp.UpdateWindow();

        // Run the boilerplate message loop
        MSG msg;
        while (::GetMessage(&msg, NULL, 0, 0) > 0)
        {
            ::TranslateMessage(&msg);
            ::DispatchMessage(&msg);
        }
        iRet = msg.wParam;
    }
    else
    {
        ::MessageBox(NULL, TEXT("Error creating the window"),
                     gc_szAppName, MB_ICONERROR | MB_OK);
        iRet = 0;
    }

    return iRet;
}

// Window message handlers //////////////////////////////

/////////////////////////////////////////////////////////
//
// CAdvRecoApp::OnCreate
//
// This WM_CREATE message handler creates and obtains interface
// pointers to the required Automation objects, sets their
// attributes, creates the child windows and enables pen input.
//
// Parameters:
//      defined in the ATL's macro MESSAGE_HANDLER,
//      none of them is used here
//
// Return Values (LRESULT):
//      always 0
//
/////////////////////////////////////////////////////////
LRESULT CAdvRecoApp::OnCreate(
        UINT /*uMsg*/,
        WPARAM /*wParam*/,
        LPARAM /*lParam*/,
        BOOL& /*bHandled*/
        )
{
    // Create child windows for ink input and recognition output,
    // listview controls for the lists of gestures, and a status bar
    if (false == CreateChildWindows())
        return -1;

    HRESULT hr;

    // Create an ink collector object.
    hr = m_spIInkCollector.CoCreateInstance(CLSID_InkCollector);
    if (FAILED(hr))
        return -1;

    // Get a pointer to the ink object interface.
    hr = m_spIInkCollector->get_Ink(&m_spIInkDisp);
    if (FAILED(hr))
        return -1;

    // Establish a connection to the collector's event source.
    // Depending on the selected collection mode, the application will be
    // listening to either Stroke or Gesture events, or both.
    hr = IInkCollectorEventsImpl<CAdvRecoApp>::DispEventAdvise(m_spIInkCollector);
    // There is nothing interesting the application can do without events
    // from the ink collector
    if (FAILED(hr))
        return -1;

    // Set the recommended subset of gestures
    PresetGestures();

    // Enable ink input in the m_wndInput window
    hr = m_spIInkCollector->put_hWnd((long)m_wndInput.m_hWnd);
    if (FAILED(hr))
        return -1;

    hr = m_spIInkCollector->put_CollectionMode(ICM_InkAndGesture);
    if (FAILED(hr))
        return -1;

    hr = m_spIInkCollector->put_Enabled(VARIANT_TRUE);
    if (FAILED(hr))
        return -1;

    return 0;
}

/////////////////////////////////////////////////////////
//
// CAdvRecoApp::OnDestroy
//
// The WM_DESTROY message handler.
// Used for clean up and also to post a quit message to
// the application itself.
//
// Parameters:
//      defined in the ATL's macro MESSAGE_HANDLER,
//      none is used here
//
// Return Values (LRESULT):
//      always 0
//
/////////////////////////////////////////////////////////
LRESULT CAdvRecoApp::OnDestroy(
        UINT /*uMsg*/,
        WPARAM /*wParam*/,
        LPARAM /*lParam*/,
        BOOL& /*bHandled*/
        )
{
    // Disable ink input and release the InkCollector object
    if (m_spIInkCollector != NULL)
    {
        IInkCollectorEventsImpl<CAdvRecoApp>::DispEventUnadvise(m_spIInkCollector);
        m_spIInkCollector->put_Enabled(VARIANT_FALSE);
        m_spIInkCollector.Release();
    }

    // Post a WM_QUIT message to the application's message queue
    ::PostQuitMessage(0);

    return 0;
}

/////////////////////////////////////////////////////////
//
// CAdvRecoApp::OnSize
//
// The WM_SIZE message handler is needed to update
// the layout of the child windows
//
// Parameters:
//      defined in the ATL's macro MESSAGE_HANDLER,
//      wParam of the WN_SIZE message is the only used here.
//
// Return Values (LRESULT):
//      always 0
//
/////////////////////////////////////////////////////////
LRESULT CAdvRecoApp::OnSize(
        UINT /*uMsg*/,
        WPARAM wParam,
        LPARAM /*lParam*/,
        BOOL& /*bHandled*/
        )
{
    if (wParam != SIZE_MINIMIZED)
    {
        UpdateLayout();
    }
    return 0;
}

// InkCollector event handlers ///////////////////////////

/////////////////////////////////////////////////////////
//
// CAdvRecoApp::OnGesture
//
// The _IInkCollectorEvents's Gesture event handler.
// See the Tablet PC Automation API Reference for the
// detailed description of the event and its parameters.
//
// Parameters:
//      IInkCursor* pIInkCursor  : [in] not used here
//      IInkStrokes* pInkStrokes : [in] the collection
//      VARIANT vGestures        : [in] safearray of IDispatch interface pointers
//                                 of the recognized  Gesture objects
//      VARIANT_BOOL* pbCancel   : [in,out] option to cancel the gesture,
//                                 default value is FALSE
//
// Return Values (HRESULT):
//      S_OK if succeeded, E_FAIL or E_INVALIDARG otherwise
//
/////////////////////////////////////////////////////////
HRESULT CAdvRecoApp::OnGesture(
        IInkCursor* /*pIInkCursor*/,
        IInkStrokes* /*pInkStrokes*/,
        VARIANT vGestures,
        VARIANT_BOOL* pbCancel
        )
{
    if (((VT_ARRAY | VT_DISPATCH) != vGestures.vt) || (NULL == vGestures.parray))
        return E_INVALIDARG;
    if (0 == vGestures.parray->rgsabound->cElements)
        return E_INVALIDARG;

    // The gestures in the array are supposed to be ordered by their recognition
    // confidence level. This sample picks up the top one.
    // NOTE: when in the InkAndGesture collection mode, besides the gestures expected
    // by the application there also can come a gesture object with the id IAG_NoGesture
    // This application cancels the event if the object with ISG_NoGesture has
    // the top confidence level (the first item in the array).
    InkApplicationGesture idGesture = IAG_NoGesture;
    IDispatch** ppIDispatch;
    HRESULT hr = ::SafeArrayAccessData(vGestures.parray, (void HUGEP**)&ppIDispatch);
    if (SUCCEEDED(hr))
    {
        CComQIPtr<IInkGesture> spIInkGesture(ppIDispatch[0]);
        if (spIInkGesture != NULL)
        {
            hr = spIInkGesture->get_Id(&idGesture);
        }
        ::SafeArrayUnaccessData(vGestures.parray);
    }

    // Load the name of the gesture from the resource string table
    UINT idGestureName;
    bool bAccepted;     // will be true, if the gesture is known to this application
    if (IAG_NoGesture != idGesture)
    {
        bAccepted = GetGestureName(idGesture, idGestureName);
    }
    else    // ignore the event (IAG_NoGesture had the highest confidence level,
            // or something has failed
    {
        bAccepted = false;
        idGestureName = 0;
    }

    // If the current collection mode is ICM_GestureOnly or if we accept
    // the gesture, the gesture's strokes will be removed from the ink object,
    // So, the window needs to be updated in the strokes' area.
    if (true == bAccepted)
    {
        // Get the rectangle to update.
        //RECT rc;
        //m_wndInput.GetClientRect(&rc);

        //m_wndInput.InvalidateRect(&rc);
    }
    else // if something's failed,
         // or the gesture is either unknown or unchecked in the list
    {
        // Reject the gesture. The InkCollector will fire Stroke event(s)
        // for the strokes, so they'll be handled in the OnStroke method.
        *pbCancel = VARIANT_TRUE;
        idGestureName = IDS_GESTURE_UNKNOWN;        
    }

    SendMessage(WM_COMMAND, ID_CLEAR);

    // Update the results window as well
    m_wndResults.SetGestureName(idGestureName);
    m_wndResults.Invalidate();

    return hr;
}

// Command handlers /////////////////////////////////////

/////////////////////////////////////////////////////////
//
// CAdvRecoApp::OnClear
//
// This command handler is called when user clicks on "Clear"
// in the Ink menu. It's supposed to delete the collected ink,
// and update the child windows after that.
//
// Parameters:
//      defined in the ATL's macro COMMAND_ID_HANDLER
//      none of them is used here
//
// Return Values (LRESULT):
//      always 0
//
/////////////////////////////////////////////////////////
LRESULT CAdvRecoApp::OnClear(
        WORD /*wNotifyCode*/,
        WORD /*wID*/,
        HWND /*hWndCtl*/,
        BOOL& /*bHandled*/
        )
{
    if (m_spIInkDisp != NULL)
    {
        // Delete all strokes from the Ink object, ignore returned value
        m_spIInkDisp->DeleteStrokes(0);
    }

    // Update the child windows
    m_wndResults.ResetResults();    // empties the strings
    m_wndResults.Invalidate();
    m_wndInput.Invalidate();

    return 0;
}


/////////////////////////////////////////////////////////
//
// CAdvRecoApp::OnExit
//
// This command handler is called when user clicks
// on "Exit" in the Ink menu.
//
// Parameters:
//      defined in the ATL's macro COMMAND_ID_HANDLER
//      none of them is used here
//
// Return Values (LRESULT):
//      always 0
//
/////////////////////////////////////////////////////////
LRESULT CAdvRecoApp::OnExit(
        WORD /*wNotifyCode*/,
        WORD /*wID*/,
        HWND /*hWndCtl*/,
        BOOL& /*bHandled*/
        )
{
    // Close the application window
    SendMessage(WM_CLOSE);
    return 0;
}

// Helper methods //////////////////////////////

/////////////////////////////////////////////////////////
//
// CAdvRecoApp::OnLVColumnClick
//
// When user clicks on the column header in either of the gesture
// listview controls, the CAdvRecoApp object receives a WM_NOTIFY
// message that is mapped to this handler for the actual processing.
// The application checks or unchecks all the items in the control
// the notification came from.
//
// Parameters:
//      defined in the ATL's macro NOTIFY_HANDLER
//
// Return Values (LRESULT):
//      always 0
//
/////////////////////////////////////////////////////////
LRESULT CAdvRecoApp::OnLVColumnClick(
        int idCtrl,
        LPNMHDR /*pnmh*/,
        BOOL& bHandled
        )
{
    if (mc_iSSGestLVId == idCtrl)
    {
        m_bAllSSGestures = !m_bAllSSGestures;
        ListView_SetCheckState(m_hwndSSGestLV, -1, m_bAllSSGestures);
    }
    else
    {
        bHandled = FALSE;
    }

    return 0;
}

/////////////////////////////////////////////////////////
//
// CAdvRecoApp::OnLVItemChanging
//
// When user checks or unchecks an item in either of the gesture
// listview controls, the state of the item is changing and the CAdvRecoApp
// object receives a WM_NOTIFY message that is mapped to this handler
// for the actual processing.
// The application set or reset the gesture status in the InkCollector.
//
// Parameters:
//      defined in the ATL's macro NOTIFY_HANDLER
//
// Return Values (LRESULT):
//      TRUE to
//
/////////////////////////////////////////////////////////
LRESULT CAdvRecoApp::OnLVItemChanging(
        int idCtrl,
        LPNMHDR pnmh,
        BOOL& /*bHandled*/
        )
{
    if (m_spIInkCollector == NULL)
        return FALSE;

    LPNMLISTVIEW pnmv = (LPNMLISTVIEW)pnmh;

    LRESULT lRet;

    // Ignore all the changes which are not of the item state,
    // and the item state changes other then checked/unchecked (LVIS_STATEIMAGEMASK)
    if (LVIF_STATE == pnmv->uChanged && 0 != (LVIS_STATEIMAGEMASK & pnmv->uNewState))
    {
        lRet = TRUE;   // prevent the change if something is wrong
        BOOL bChecked = ((LVIS_STATEIMAGEMASK & pnmv->uNewState) >> 12) == 2;
        InkApplicationGesture igtGesture = IAG_NoGesture;
        if (mc_iSSGestLVId == idCtrl)
        {
            if (pnmv->iItem >= 0 && pnmv->iItem < countof(gc_igtSingleStrokeGestures))
                igtGesture = gc_igtSingleStrokeGestures[pnmv->iItem];
        }

        if (IAG_NoGesture != igtGesture && SUCCEEDED(
            m_spIInkCollector->SetGestureStatus(igtGesture, bChecked ? VARIANT_TRUE : VARIANT_FALSE)))
        {
            // Allow the change in the control's item state
            lRet = FALSE;
        }
    }
    else
    {
        // Allow all the other changes
        lRet = FALSE;
    }

    return lRet;
}

/////////////////////////////////////////////////////////
//
// CAdvRecoApp::CreateChildWindows
//
// This helper method is called from WM_CREATE message handler.
// The child windows and controls are created and initialized here.
//
// Parameters:
//      none
//
// Return Values (bool):
//      true if the windows have been created successfully,
//      false otherwise
//
/////////////////////////////////////////////////////////
bool CAdvRecoApp::CreateChildWindows()
{
    if ((m_wndInput.Create(m_hWnd, CWindow::rcDefault, NULL,
                           WS_CHILD, WS_EX_CLIENTEDGE, (UINT)mc_iInputWndId) == NULL)
        || (m_wndResults.Create(m_hWnd, CWindow::rcDefault, NULL,
                                WS_CHILD, WS_EX_CLIENTEDGE, (UINT)mc_iOutputWndId) == NULL))
    {
        return false;
    }


    HINSTANCE hInst = _Module.GetResourceInstance();

    // Create a listview control for the list of the single stroke gestures
    m_hwndSSGestLV = ::CreateWindowEx(WS_EX_CLIENTEDGE, WC_LISTVIEW, NULL,
                                      WS_VISIBLE | WS_CHILD | WS_BORDER | LVS_REPORT,
                                      0, 0, 1, 1,
                                      m_hWnd, (HMENU)mc_iSSGestLVId,
                                      _Module.GetModuleInstance(), NULL);
    if (NULL == m_hwndSSGestLV)
        return false;

    //
    ListView_SetExtendedListViewStyleEx(m_hwndSSGestLV, LVS_EX_CHECKBOXES, LVS_EX_CHECKBOXES);

    // Create a column
    LV_COLUMN lvC;
    lvC.mask = LVCF_FMT | LVCF_WIDTH | LVCF_TEXT | LVCF_SUBITEM;
    lvC.fmt = LVCFMT_LEFT;
    lvC.iSubItem = 0;
    lvC.cx = mc_cxGestLVWidth - 22;
    lvC.pszText = TEXT("Single Stroke Gestures");
    if (-1 == ListView_InsertColumn(m_hwndSSGestLV, lvC.iSubItem, &lvC))
        return false;

    // Insert items - the names of the single stroke gestures.
    TCHAR szText[100];  // large enough to load a gesture name into
    LV_ITEM lvItem;
    lvItem.mask = LVIF_TEXT /*| LVIF_IMAGE*/ | LVIF_STATE;
    lvItem.state = 0;
    lvItem.stateMask = 0;
    lvItem.pszText = szText;
    lvItem.iSubItem = 0;
    for (ULONG i = 0; i < mc_cNumSSGestures; i++)
    {
        lvItem.iItem = i;
        // Load the names from the application resource, there should be
        // mc_cNumSSGestures names there with sequential id's starting
        // with IDS_SSGESTURE_FIRST
        ::LoadString(hInst, IDS_SSGESTURE_FIRST + i, szText, countof(szText));
        if (-1 == ListView_InsertItem(m_hwndSSGestLV, &lvItem))
            return false;
    }

    // Update the child windows' positions and sizes so that they cover
    // entire client area of the main window.
    UpdateLayout();

    return true;
}

/////////////////////////////////////////////////////////
//
// CAdvRecoApp::UpdateLayout
//
// This helper method is called when the size of the main
// window has been changed and the child windows' positions
// need to be updated so that they cover entire client area
// of the main window.
//
// Parameters:
//      none
//
// Return Values (void):
//      none
//
/////////////////////////////////////////////////////////
void CAdvRecoApp::UpdateLayout()
{
    RECT rect;
    GetClientRect(&rect);

    // update the size and position of the gesture listviews
    if (::IsWindow(m_hwndSSGestLV))
    {
        // calculate the rectangle covered by the list views
        RECT rcGest = rect;
        if (rcGest.right < mc_cxGestLVWidth)
        {
            rcGest.left = 0;
        }
        else
        {
            rcGest.left = rcGest.right - mc_cxGestLVWidth;
        }

        rect.right = rcGest.left;

        // show the single stroke gesture listview control
        ::SetWindowPos(m_hwndSSGestLV, NULL,
                        rcGest.left, rcGest.top,
                        rcGest.right - rcGest.left, rcGest.bottom - rcGest.top,
                        SWP_NOZORDER | SWP_NOACTIVATE | SWP_SHOWWINDOW);
    }

    // update the size and position of the output window
    if (::IsWindow(m_wndResults.m_hWnd))
    {
        int cyResultsWnd = m_wndResults.GetBestHeight();
        if (cyResultsWnd > rect.bottom)
        {
            cyResultsWnd = rect.bottom;
        }
        ::SetWindowPos(m_wndResults.m_hWnd, NULL,
                       rect.left, rect.bottom - cyResultsWnd,
                       rect.right - rect.left, cyResultsWnd - rect.top,
                       SWP_NOZORDER | SWP_NOACTIVATE | SWP_SHOWWINDOW);
        rect.bottom -= cyResultsWnd;
    }

    // update the size and position of the ink input window
    if (::IsWindow(m_wndInput.m_hWnd))
    {
        ::SetWindowPos(m_wndInput.m_hWnd, NULL,
                       rect.left, rect.top,
                       rect.right - rect.left, rect.bottom - rect.top,
                       SWP_NOZORDER | SWP_NOACTIVATE | SWP_SHOWWINDOW);
    }
}

/////////////////////////////////////////////////////////
//
// CAdvRecoApp::GetGestureName
//
// This helper function returns the resource id of
// the string with the name of the given gesture.
//
// Parameters:
//      InkApplicationGesture idGesture    : [in] the gesture's id
//      UINT& idGestureName         : [out] the id of the string with the the name
//                                    of the gesturein the resource string table
// Return Values (bool):
//      true if the gesture is known to the application, false otherwise
//
/////////////////////////////////////////////////////////
bool CAdvRecoApp::GetGestureName(
        InkApplicationGesture igtGesture,
        UINT& idGestureName
        )
{
    idGestureName = IDS_GESTURE_UNKNOWN;

    // First, try to find the gesture among the single stroke ones
    ULONG iCount = countof(gc_igtSingleStrokeGestures);
    ULONG i;
    for (i = 0; i < iCount; i++)
    {
        if (gc_igtSingleStrokeGestures[i] == igtGesture)
        {
            idGestureName = IDS_SSGESTURE_FIRST + i;
            break;
        }
    }

    return (IDS_GESTURE_UNKNOWN != idGestureName);
}

/////////////////////////////////////////////////////////
//
// CAdvRecoApp::PresetGestures
//
// Sets the status of the recommended subset of gestures
// to TRUE in InkCollector
//
// Parameters:
//      none
//
// Return Values (void):
//      none
//
/////////////////////////////////////////////////////////
void CAdvRecoApp::PresetGestures()
{
    if (0 == ::IsWindow(m_hwndSSGestLV))
        return;
    for (ULONG i = 0; i < countof(gc_igtSingleStrokeGestures); i++) {
        ListView_SetCheckState(m_hwndSSGestLV, i, TRUE);
    }
}
