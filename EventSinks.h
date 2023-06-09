// THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF
// ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO
// THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A
// PARTICULAR PURPOSE.
//
// Copyright (c) Microsoft Corporation. All rights reserved.
//
// Module:       
//      EventSinks.h
//
// Description:
//      This file contains the definitions of the event sink templates,
//      instantiated in the CAdvRecoApp class.
//
//      The event source interfaces used are: 
//      _IInkEvents, _IInkRecognitionEvent, _IInkCollectorEvents
//   
//--------------------------------------------------------------------------

#pragma once

// The DISPID's of the events 
#ifndef DISPID_CEStroke
    #define DISPID_CEStroke                     0x00000001
#endif
#ifndef DISPID_CEGesture
    #define DISPID_CEGesture                    0x0000000a
#endif
#ifndef DISPID_RERecognitionWithAlternates
    #define DISPID_RERecognitionWithAlternates  0x00000001
#endif

// IDispEventSimpleImpl requires a constant as a sink id; define it here
#define SINK_ID  1

/////////////////////////////////////////////////////////
//                                          
// IInkCollectorEventsImpl
// 
// The template is derived from the ATL's IDispEventSimpleImpl 
// to implement a sink for the IInkCollectorEvents, fired by 
// the InkCollector object
// Since the IDispEventSimpleImpl doesn't require to supply 
// implementation code for every event, this template has a handler
// for the Gesture event only.
//
/////////////////////////////////////////////////////////

template <class T>
class ATL_NO_VTABLE IInkCollectorEventsImpl :
	public IDispEventSimpleImpl<SINK_ID, 
                                IInkCollectorEventsImpl<T>, 
                                &DIID__IInkCollectorEvents>
{
public:
    // ATL structures with the type information for each event, 
    // handled in this template.(Initialized in the AdvReco.cpp)
    static const _ATL_FUNC_INFO mc_AtlFuncInfo[1];

BEGIN_SINK_MAP(IInkCollectorEventsImpl)
    SINK_ENTRY_INFO(SINK_ID, 
                    DIID__IInkCollectorEvents, 
                    DISPID_CEGesture, 
                    Gesture, 
                    const_cast<_ATL_FUNC_INFO*>(&mc_AtlFuncInfo[0]))
END_SINK_MAP()

    HRESULT __stdcall Gesture(IInkCursor* pIInkCursor, IInkStrokes* pInkStrokes, 
                              VARIANT vGestures, VARIANT_BOOL* pbCancel)
    {
		T* pT = static_cast<T*>(this);
        return pT->OnGesture(pIInkCursor, pInkStrokes, vGestures, pbCancel);
    }
};

