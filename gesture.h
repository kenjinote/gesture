// THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF
// ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO
// THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A
// PARTICULAR PURPOSE.
//
// Copyright (c) Microsoft Corporation. All rights reserved.
//
// Module:       
//      AdvRecoApp.h
//
// Description:
//      The header file for the CAdvRecoApp class - the application window 
//      class of the AdvReco sample.
//		The methods of the class are defined in the AdvReco.cpp file.
//   
//--------------------------------------------------------------------------

#pragma once

/////////////////////////////////////////////////////////////////////////////
// CAdvRecoApp

class CAdvRecoApp : 
    public CWindowImpl<CAdvRecoApp>,
    public IInkCollectorEventsImpl<CAdvRecoApp>
{
public:
    // Constants 

    enum { 
        // child windows IDs
        mc_iInputWndId = 1, 
        mc_iOutputWndId = 2, 
        mc_iSSGestLVId = 4,
        // recognition guide box data
        // the width of the gesture list views 
        mc_cxGestLVWidth = 160, 
        // the number of the gesture names in the string table
        mc_cNumSSGestures = 36,     // single stroke gestures
    };

    // Automation API interface pointers
    CComPtr<IInkCollector>          m_spIInkCollector;
    CComPtr<IInkDisp>               m_spIInkDisp;

    // Child windows
    CInkInputWnd    m_wndInput;
    CRecoOutputWnd  m_wndResults;
    HWND            m_hwndSSGestLV;     // single stroke gestures list view

    // Helper data members
    bool            m_bAllSSGestures;

    // Static method that creates an object of the class
    static int Run(int nCmdShow);

    // Constructor
    CAdvRecoApp() :
        m_hwndSSGestLV(NULL), m_bAllSSGestures(true)
    {
    }

    // Helper methods
    bool    CreateChildWindows();
    void    UpdateLayout();
    bool    GetGestureName(InkApplicationGesture idGesture, UINT& idGestureName);
    void    PresetGestures();
    

// Declare the class objects' window class with NULL background.
// There's no need to paint CAdvRecoApp window background because
// the entire client area is covered by the child windows.
DECLARE_WND_CLASS_EX(NULL, 0, -1)
    
// ATL macro's to declare which commands/messages the class is interested in.
BEGIN_MSG_MAP(CAdvRecoApp)
    MESSAGE_HANDLER(WM_CREATE, OnCreate)
    MESSAGE_HANDLER(WM_DESTROY, OnDestroy)
    MESSAGE_HANDLER(WM_SIZE, OnSize)
    COMMAND_ID_HANDLER(ID_CLEAR, OnClear)
    COMMAND_ID_HANDLER(ID_EXIT, OnExit)
    NOTIFY_HANDLER(mc_iSSGestLVId, LVN_COLUMNCLICK, OnLVColumnClick)
    NOTIFY_HANDLER(mc_iSSGestLVId, LVN_ITEMCHANGING, OnLVItemChanging)
END_MSG_MAP()

public:

    // Window message handlers
    LRESULT OnCreate(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
    LRESULT OnDestroy(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
    LRESULT OnSize(UINT, WPARAM, LPARAM, BOOL& bHandled);
    LRESULT OnLVColumnClick(int idCtrl, LPNMHDR pnmh, BOOL& bHandled);
    LRESULT OnLVItemChanging(int idCtrl, LPNMHDR pnmh, BOOL& bHandled);
    
    // Command handlers
    LRESULT OnClear(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled);
    LRESULT OnExit(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled);

    // Ink collector event handler
    HRESULT OnGesture(IInkCursor* pIInkCursor, IInkStrokes* pIInkStrokes, 
                      VARIANT vGestures, VARIANT_BOOL* pbCancel);
};

