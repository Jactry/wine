/*
 * RichEdit GUIDs and OLE interface
 *
 * Copyright 2004 by Krzysztof Foltman
 * Copyright 2004 Aric Stewart
 *
 * This library is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.
 *
 * This library is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
 * Lesser General Public License for more details.
 *
 * You should have received a copy of the GNU Lesser General Public
 * License along with this library; if not, write to the Free Software
 * Foundation, Inc., 51 Franklin St, Fifth Floor, Boston, MA 02110-1301, USA
 */

#include <stdarg.h>

#define NONAMELESSUNION
#define NONAMELESSSTRUCT
#define COBJMACROS

#include "windef.h"
#include "winbase.h"
#include "wingdi.h"
#include "winuser.h"
#include "ole2.h"
#include "richole.h"
#include "editor.h"
#include "tom.h"
#include "wine/debug.h"

WINE_DEFAULT_DEBUG_CHANNEL(richedit);

/* there is no way to be consistent across different sets of headers - mingw, Wine, Win32 SDK*/

#include "initguid.h"
DEFINE_GUID(IID_ITextServices, 0x8d33f740, 0xcf58, 0x11ce, 0xa8, 0x9d, 0x00, 0xaa, 0x00, 0x6c, 0xad, 0xc5);
DEFINE_GUID(IID_ITextHost, 0x13e670f4,0x1a5a,0x11cf,0xab,0xeb,0x00,0xaa,0x00,0xb6,0x5e,0xa1);
DEFINE_GUID(IID_ITextHost2, 0x13e670f5,0x1a5a,0x11cf,0xab,0xeb,0x00,0xaa,0x00,0xb6,0x5e,0xa1);
DEFINE_GUID(IID_ITextDocument, 0x8cc497c0, 0xa1df, 0x11ce, 0x80, 0x98, 0x00, 0xaa, 0x00, 0x47, 0xbe, 0x5d);
DEFINE_GUID(IID_ITextRange, 0x8cc497c2, 0xa1df, 0x11ce, 0x80, 0x98, 0x00, 0xaa, 0x00, 0x47, 0xbe, 0x5d);
DEFINE_GUID(IID_ITextSelection, 0x8cc497c1, 0xa1df, 0x11ce, 0x80, 0x98, 0x00, 0xaa, 0x00, 0x47, 0xbe, 0x5d);

typedef struct ITextSelectionImpl ITextSelectionImpl;
typedef struct IOleClientSiteImpl IOleClientSiteImpl;
typedef struct ITextRangeImpl ITextRangeImpl;

typedef struct ME_CursorPair {
    LONG start, end;
    LONG ref;
} ME_CursorPair;

typedef struct IRichEditOleImpl {
    IRichEditOle IRichEditOle_iface;
    ITextDocument ITextDocument_iface;
    LONG ref;

    ME_TextEditor *editor;
    ITextSelectionImpl *txtSel;
    IOleClientSiteImpl *clientSite;
    struct list rangelist;
} IRichEditOleImpl;

struct ITextRangeImpl {
    ITextRange ITextRange_iface;
    LONG ref;
    ME_CursorPair *cursorPair;
    struct list entry;

    IRichEditOleImpl *reOle;
};

struct ITextSelectionImpl {
    ITextSelection ITextSelection_iface;
    LONG ref;

    IRichEditOleImpl *reOle;
};

struct IOleClientSiteImpl {
    IOleClientSite IOleClientSite_iface;
    LONG ref;

    IRichEditOleImpl *reOle;
};

static ME_CursorPair *CreateCursorPair(LONG start, LONG end)
{
    ME_CursorPair *cursorPair = heap_alloc(sizeof(ME_CursorPair));

    if (!cursorPair)
        return NULL;
    cursorPair->ref = 1;
    cursorPair->start = start;
    cursorPair->end = end;
    return cursorPair;
}

static void CursorPair_AddRef(ME_CursorPair *cursorPair)
{
    InterlockedIncrement(&cursorPair->ref);
}

static void CursorPair_Release(ME_CursorPair *cursorPair)
{
    ULONG ref = InterlockedDecrement(&cursorPair->ref);

    TRACE("%p ref = %u\n", cursorPair, ref);
    if (!ref)
        heap_free(cursorPair);
}

static inline IRichEditOleImpl *impl_from_IRichEditOle(IRichEditOle *iface)
{
    return CONTAINING_RECORD(iface, IRichEditOleImpl, IRichEditOle_iface);
}

static inline IRichEditOleImpl *impl_from_ITextDocument(ITextDocument *iface)
{
    return CONTAINING_RECORD(iface, IRichEditOleImpl, ITextDocument_iface);
}

static HRESULT WINAPI
IRichEditOle_fnQueryInterface(IRichEditOle *me, REFIID riid, LPVOID *ppvObj)
{
    IRichEditOleImpl *This = impl_from_IRichEditOle(me);

    TRACE("%p %s\n", This, debugstr_guid(riid) );

    *ppvObj = NULL;
    if (IsEqualGUID(riid, &IID_IUnknown) ||
        IsEqualGUID(riid, &IID_IRichEditOle))
        *ppvObj = &This->IRichEditOle_iface;
    else if (IsEqualGUID(riid, &IID_ITextDocument))
        *ppvObj = &This->ITextDocument_iface;
    if (*ppvObj)
    {
        IRichEditOle_AddRef(me);
        return S_OK;
    }
    FIXME("%p: unhandled interface %s\n", This, debugstr_guid(riid) );
 
    return E_NOINTERFACE;   
}

static ULONG WINAPI
IRichEditOle_fnAddRef(IRichEditOle *me)
{
    IRichEditOleImpl *This = impl_from_IRichEditOle(me);
    ULONG ref = InterlockedIncrement( &This->ref );

    TRACE("%p ref = %u\n", This, ref);

    return ref;
}

static ULONG WINAPI
IRichEditOle_fnRelease(IRichEditOle *me)
{
    IRichEditOleImpl *This = impl_from_IRichEditOle(me);
    ULONG ref = InterlockedDecrement(&This->ref);

    TRACE ("%p ref=%u\n", This, ref);

    if (!ref)
    {
        ITextRangeImpl *txtRge;
        TRACE ("Destroying %p\n", This);
        This->txtSel->reOle = NULL;
        ITextSelection_Release(&This->txtSel->ITextSelection_iface);
        IOleClientSite_Release(&This->clientSite->IOleClientSite_iface);
        LIST_FOR_EACH_ENTRY(txtRge, &This->rangelist, ITextRangeImpl, entry)
            txtRge->reOle = NULL;
        heap_free(This);
    }
    return ref;
}

static HRESULT WINAPI
IRichEditOle_fnActivateAs(IRichEditOle *me, REFCLSID rclsid, REFCLSID rclsidAs)
{
    IRichEditOleImpl *This = impl_from_IRichEditOle(me);
    FIXME("stub %p\n",This);
    return E_NOTIMPL;
}

static HRESULT WINAPI
IRichEditOle_fnContextSensitiveHelp(IRichEditOle *me, BOOL fEnterMode)
{
    IRichEditOleImpl *This = impl_from_IRichEditOle(me);
    FIXME("stub %p\n",This);
    return E_NOTIMPL;
}

static HRESULT WINAPI
IRichEditOle_fnConvertObject(IRichEditOle *me, LONG iob,
               REFCLSID rclsidNew, LPCSTR lpstrUserTypeNew)
{
    IRichEditOleImpl *This = impl_from_IRichEditOle(me);
    FIXME("stub %p\n",This);
    return E_NOTIMPL;
}

static inline IOleClientSiteImpl *impl_from_IOleClientSite(IOleClientSite *iface)
{
    return CONTAINING_RECORD(iface, IOleClientSiteImpl, IOleClientSite_iface);
}

static HRESULT WINAPI
IOleClientSite_fnQueryInterface(IOleClientSite *me, REFIID riid, LPVOID *ppvObj)
{
    TRACE("%p %s\n", me, debugstr_guid(riid) );

    *ppvObj = NULL;
    if (IsEqualGUID(riid, &IID_IUnknown) ||
        IsEqualGUID(riid, &IID_IOleClientSite))
        *ppvObj = me;
    if (*ppvObj)
    {
        IOleClientSite_AddRef(me);
        return S_OK;
    }
    FIXME("%p: unhandled interface %s\n", me, debugstr_guid(riid) );

    return E_NOINTERFACE;
}

static ULONG WINAPI IOleClientSite_fnAddRef(IOleClientSite *iface)
{
    IOleClientSiteImpl *This = impl_from_IOleClientSite(iface);
    return InterlockedIncrement(&This->ref);
}

static ULONG WINAPI IOleClientSite_fnRelease(IOleClientSite *iface)
{
    IOleClientSiteImpl *This = impl_from_IOleClientSite(iface);
    ULONG ref = InterlockedDecrement(&This->ref);
    if (ref == 0)
        heap_free(This);
    return ref;
}

static HRESULT WINAPI IOleClientSite_fnSaveObject(IOleClientSite *iface)
{
    IOleClientSiteImpl *This = impl_from_IOleClientSite(iface);
    if (!This->reOle)
        return CO_E_RELEASED;

    FIXME("stub %p\n", iface);
    return E_NOTIMPL;
}


static HRESULT WINAPI IOleClientSite_fnGetMoniker(IOleClientSite *iface, DWORD dwAssign,
        DWORD dwWhichMoniker, IMoniker **ppmk)
{
    IOleClientSiteImpl *This = impl_from_IOleClientSite(iface);
    if (!This->reOle)
        return CO_E_RELEASED;

    FIXME("stub %p\n", iface);
    return E_NOTIMPL;
}

static HRESULT WINAPI IOleClientSite_fnGetContainer(IOleClientSite *iface,
        IOleContainer **ppContainer)
{
    IOleClientSiteImpl *This = impl_from_IOleClientSite(iface);
    if (!This->reOle)
        return CO_E_RELEASED;

    FIXME("stub %p\n", iface);
    return E_NOTIMPL;
}

static HRESULT WINAPI IOleClientSite_fnShowObject(IOleClientSite *iface)
{
    IOleClientSiteImpl *This = impl_from_IOleClientSite(iface);
    if (!This->reOle)
        return CO_E_RELEASED;

    FIXME("stub %p\n", iface);
    return E_NOTIMPL;
}

static HRESULT WINAPI IOleClientSite_fnOnShowWindow(IOleClientSite *iface, BOOL fShow)
{
    IOleClientSiteImpl *This = impl_from_IOleClientSite(iface);
    if (!This->reOle)
        return CO_E_RELEASED;

    FIXME("stub %p\n", iface);
    return E_NOTIMPL;
}

static HRESULT WINAPI IOleClientSite_fnRequestNewObjectLayout(IOleClientSite *iface)
{
    IOleClientSiteImpl *This = impl_from_IOleClientSite(iface);
    if (!This->reOle)
        return CO_E_RELEASED;

    FIXME("stub %p\n", iface);
    return E_NOTIMPL;
}

static const IOleClientSiteVtbl ocst = {
    IOleClientSite_fnQueryInterface,
    IOleClientSite_fnAddRef,
    IOleClientSite_fnRelease,
    IOleClientSite_fnSaveObject,
    IOleClientSite_fnGetMoniker,
    IOleClientSite_fnGetContainer,
    IOleClientSite_fnShowObject,
    IOleClientSite_fnOnShowWindow,
    IOleClientSite_fnRequestNewObjectLayout
};

static IOleClientSiteImpl *
CreateOleClientSite(IRichEditOleImpl *reOle)
{
    IOleClientSiteImpl *clientSite = heap_alloc(sizeof *clientSite);
    if (!clientSite)
        return NULL;

    clientSite->IOleClientSite_iface.lpVtbl = &ocst;
    clientSite->ref = 1;
    clientSite->reOle = reOle;
    return clientSite;
}

static HRESULT WINAPI
IRichEditOle_fnGetClientSite(IRichEditOle *me,
               LPOLECLIENTSITE *lplpolesite)
{
    IRichEditOleImpl *This = impl_from_IRichEditOle(me);

    TRACE("%p,%p\n",This, lplpolesite);

    if(!lplpolesite)
        return E_INVALIDARG;
    *lplpolesite = &This->clientSite->IOleClientSite_iface;
    IOleClientSite_AddRef(*lplpolesite);
    return S_OK;
}

static HRESULT WINAPI
IRichEditOle_fnGetClipboardData(IRichEditOle *me, CHARRANGE *lpchrg,
               DWORD reco, LPDATAOBJECT *lplpdataobj)
{
    IRichEditOleImpl *This = impl_from_IRichEditOle(me);
    ME_Cursor start;
    int nChars;

    TRACE("(%p,%p,%d)\n",This, lpchrg, reco);
    if(!lplpdataobj)
        return E_INVALIDARG;
    if(!lpchrg) {
        int nFrom, nTo, nStartCur = ME_GetSelectionOfs(This->editor, &nFrom, &nTo);
        start = This->editor->pCursors[nStartCur];
        nChars = nTo - nFrom;
    } else {
        ME_CursorFromCharOfs(This->editor, lpchrg->cpMin, &start);
        nChars = lpchrg->cpMax - lpchrg->cpMin;
    }
    return ME_GetDataObject(This->editor, &start, nChars, lplpdataobj);
}

static LONG WINAPI IRichEditOle_fnGetLinkCount(IRichEditOle *me)
{
    IRichEditOleImpl *This = impl_from_IRichEditOle(me);
    FIXME("stub %p\n",This);
    return E_NOTIMPL;
}

static HRESULT WINAPI
IRichEditOle_fnGetObject(IRichEditOle *me, LONG iob,
               REOBJECT *lpreobject, DWORD dwFlags)
{
    IRichEditOleImpl *This = impl_from_IRichEditOle(me);
    FIXME("stub %p\n",This);
    return E_NOTIMPL;
}

static LONG WINAPI
IRichEditOle_fnGetObjectCount(IRichEditOle *me)
{
    IRichEditOleImpl *This = impl_from_IRichEditOle(me);
    FIXME("stub %p\n",This);
    return 0;
}

static HRESULT WINAPI
IRichEditOle_fnHandsOffStorage(IRichEditOle *me, LONG iob)
{
    IRichEditOleImpl *This = impl_from_IRichEditOle(me);
    FIXME("stub %p\n",This);
    return E_NOTIMPL;
}

static HRESULT WINAPI
IRichEditOle_fnImportDataObject(IRichEditOle *me, LPDATAOBJECT lpdataobj,
               CLIPFORMAT cf, HGLOBAL hMetaPict)
{
    IRichEditOleImpl *This = impl_from_IRichEditOle(me);
    FIXME("stub %p\n",This);
    return E_NOTIMPL;
}

static HRESULT WINAPI
IRichEditOle_fnInPlaceDeactivate(IRichEditOle *me)
{
    IRichEditOleImpl *This = impl_from_IRichEditOle(me);
    FIXME("stub %p\n",This);
    return E_NOTIMPL;
}

static HRESULT WINAPI
IRichEditOle_fnInsertObject(IRichEditOle *me, REOBJECT *reo)
{
    IRichEditOleImpl *This = impl_from_IRichEditOle(me);
    TRACE("(%p,%p)\n", This, reo);

    if (reo->cbStruct < sizeof(*reo)) return STG_E_INVALIDPARAMETER;

    ME_InsertOLEFromCursor(This->editor, reo, 0);
    ME_CommitUndo(This->editor);
    ME_UpdateRepaint(This->editor, FALSE);
    return S_OK;
}

static HRESULT WINAPI IRichEditOle_fnSaveCompleted(IRichEditOle *me, LONG iob,
               LPSTORAGE lpstg)
{
    IRichEditOleImpl *This = impl_from_IRichEditOle(me);
    FIXME("stub %p\n",This);
    return E_NOTIMPL;
}

static HRESULT WINAPI
IRichEditOle_fnSetDvaspect(IRichEditOle *me, LONG iob, DWORD dvaspect)
{
    IRichEditOleImpl *This = impl_from_IRichEditOle(me);
    FIXME("stub %p\n",This);
    return E_NOTIMPL;
}

static HRESULT WINAPI IRichEditOle_fnSetHostNames(IRichEditOle *me,
               LPCSTR lpstrContainerApp, LPCSTR lpstrContainerObj)
{
    IRichEditOleImpl *This = impl_from_IRichEditOle(me);
    FIXME("stub %p %s %s\n",This, lpstrContainerApp, lpstrContainerObj);
    return E_NOTIMPL;
}

static HRESULT WINAPI
IRichEditOle_fnSetLinkAvailable(IRichEditOle *me, LONG iob, BOOL fAvailable)
{
    IRichEditOleImpl *This = impl_from_IRichEditOle(me);
    FIXME("stub %p\n",This);
    return E_NOTIMPL;
}

static const IRichEditOleVtbl revt = {
    IRichEditOle_fnQueryInterface,
    IRichEditOle_fnAddRef,
    IRichEditOle_fnRelease,
    IRichEditOle_fnGetClientSite,
    IRichEditOle_fnGetObjectCount,
    IRichEditOle_fnGetLinkCount,
    IRichEditOle_fnGetObject,
    IRichEditOle_fnInsertObject,
    IRichEditOle_fnConvertObject,
    IRichEditOle_fnActivateAs,
    IRichEditOle_fnSetHostNames,
    IRichEditOle_fnSetLinkAvailable,
    IRichEditOle_fnSetDvaspect,
    IRichEditOle_fnHandsOffStorage,
    IRichEditOle_fnSaveCompleted,
    IRichEditOle_fnInPlaceDeactivate,
    IRichEditOle_fnContextSensitiveHelp,
    IRichEditOle_fnGetClipboardData,
    IRichEditOle_fnImportDataObject
};

/* ITextRange interface */
static inline ITextRangeImpl *impl_from_ITextRange(ITextRange *iface)
{
    return CONTAINING_RECORD(iface, ITextRangeImpl, ITextRange_iface);
}

static HRESULT WINAPI ITextRange_fnQueryInterface(ITextRange *me, REFIID riid, void **ppvObj)
{
    *ppvObj = NULL;
    if (IsEqualGUID(riid, &IID_IUnknown)
        || IsEqualGUID(riid, &IID_IDispatch)
        || IsEqualGUID(riid, &IID_ITextRange))
    {
        *ppvObj = me;
        ITextRange_AddRef(me);
        return S_OK;
    }

    return E_NOINTERFACE;
}

static ULONG WINAPI ITextRange_fnAddRef(ITextRange *me)
{
    ITextRangeImpl *This = impl_from_ITextRange(me);
    return InterlockedIncrement(&This->ref);
}

static ULONG WINAPI ITextRange_fnRelease(ITextRange *me)
{
    ITextRangeImpl *This = impl_from_ITextRange(me);
    ULONG ref = InterlockedDecrement(&This->ref);

    TRACE ("%p ref=%u\n", This, ref);
    if (ref == 0)
    {
        if (This->reOle)
        {
            CursorPair_Release(This->cursorPair);
            list_remove(&This->entry);
            This->reOle = NULL;
        }
        heap_free(This);
    }
    return ref;
}

static HRESULT WINAPI ITextRange_fnGetTypeInfoCount(ITextRange *me, UINT *pctinfo)
{
    ITextRangeImpl *This = impl_from_ITextRange(me);
    if (!This->reOle)
        return CO_E_RELEASED;

    FIXME("not implemented %p\n", This);
    return E_NOTIMPL;
}

static HRESULT WINAPI ITextRange_fnGetTypeInfo(ITextRange *me, UINT iTInfo, LCID lcid,
                                               ITypeInfo **ppTInfo)
{
    ITextRangeImpl *This = impl_from_ITextRange(me);
    if (!This->reOle)
        return CO_E_RELEASED;

    FIXME("not implemented %p\n", This);
    return E_NOTIMPL;
}

static HRESULT WINAPI ITextRange_fnGetIDsOfNames(ITextRange *me, REFIID riid, LPOLESTR *rgszNames,
                                                 UINT cNames, LCID lcid, DISPID *rgDispId)
{
    ITextRangeImpl *This = impl_from_ITextRange(me);
    if (!This->reOle)
        return CO_E_RELEASED;

    FIXME("not implemented %p\n", This);
    return E_NOTIMPL;
}

static HRESULT WINAPI ITextRange_fnInvoke(ITextRange *me, DISPID dispIdMember, REFIID riid,
                                          LCID lcid, WORD wFlags, DISPPARAMS *pDispParams,
                                          VARIANT *pVarResult, EXCEPINFO *pExcepInfo,
                                          UINT *puArgErr)
{
    ITextRangeImpl *This = impl_from_ITextRange(me);
    if (!This->reOle)
        return CO_E_RELEASED;

    FIXME("not implemented %p\n", This);
    return E_NOTIMPL;
}

static HRESULT WINAPI ITextRange_fnGetText(ITextRange *me, BSTR *pbstr)
{
    ITextRangeImpl *This = impl_from_ITextRange(me);
    if (!This->reOle)
        return CO_E_RELEASED;

    FIXME("not implemented %p\n", This);
    return E_NOTIMPL;
}

static HRESULT WINAPI ITextRange_fnSetText(ITextRange *me, BSTR bstr)
{
    ITextRangeImpl *This = impl_from_ITextRange(me);
    if (!This->reOle)
        return CO_E_RELEASED;

    FIXME("not implemented %p\n", This);
    return E_NOTIMPL;
}

static HRESULT range_GetChar(ME_TextEditor *editor, ME_Cursor *cursor, LONG *pch)
{
    WCHAR wch[2];

    ME_GetTextW(editor, wch, 1, cursor, 1, FALSE, cursor->pRun->next->type == diTextEnd);
    *pch = wch[0];

    return S_OK;
}

static HRESULT WINAPI ITextRange_fnGetChar(ITextRange *me, LONG *pch)
{
    ITextRangeImpl *This = impl_from_ITextRange(me);
    ME_Cursor cursor;

    if (!This->reOle)
        return CO_E_RELEASED;
    TRACE("%p\n", pch);
    if (!pch)
        return E_INVALIDARG;

    ME_CursorFromCharOfs(This->reOle->editor, This->cursorPair->start, &cursor);
    return range_GetChar(This->reOle->editor, &cursor, pch);
}

static HRESULT WINAPI ITextRange_fnSetChar(ITextRange *me, LONG ch)
{
    ITextRangeImpl *This = impl_from_ITextRange(me);
    if (!This->reOle)
        return CO_E_RELEASED;

    FIXME("not implemented %p\n", This);
    return E_NOTIMPL;
}

static HRESULT CreateITextRange(IRichEditOleImpl *reOle, LONG start, LONG end, ITextRange** ppRange);

static HRESULT WINAPI ITextRange_fnGetDuplicate(ITextRange *me, ITextRange **ppRange)
{
    ITextRangeImpl *This = impl_from_ITextRange(me);
    if (!This->reOle)
        return CO_E_RELEASED;

    TRACE("%p %p\n", This, ppRange);
    if (!ppRange)
        return E_INVALIDARG;

    return CreateITextRange(This->reOle, This->cursorPair->start, This->cursorPair->end, ppRange);
}

static HRESULT WINAPI ITextRange_fnGetFormattedText(ITextRange *me, ITextRange **ppRange)
{
    ITextRangeImpl *This = impl_from_ITextRange(me);
    if (!This->reOle)
        return CO_E_RELEASED;

    FIXME("not implemented %p\n", This);
    return E_NOTIMPL;
}

static HRESULT WINAPI ITextRange_fnSetFormattedText(ITextRange *me, ITextRange *pRange)
{
    ITextRangeImpl *This = impl_from_ITextRange(me);
    if (!This->reOle)
        return CO_E_RELEASED;

    FIXME("not implemented %p\n", This);
    return E_NOTIMPL;
}

static HRESULT WINAPI ITextRange_fnGetStart(ITextRange *me, LONG *pcpFirst)
{
    ITextRangeImpl *This = impl_from_ITextRange(me);
    if (!This->reOle)
        return CO_E_RELEASED;

    if (!pcpFirst)
        return E_INVALIDARG;
    *pcpFirst = This->cursorPair->start;
    TRACE("%d\n", *pcpFirst);
    return S_OK;
}

static HRESULT WINAPI ITextRange_fnSetStart(ITextRange *me, LONG cpFirst)
{
    ITextRangeImpl *This = impl_from_ITextRange(me);
    if (!This->reOle)
        return CO_E_RELEASED;

    FIXME("not implemented %p\n", This);
    return E_NOTIMPL;
}

static HRESULT WINAPI ITextRange_fnGetEnd(ITextRange *me, LONG *pcpLim)
{
    ITextRangeImpl *This = impl_from_ITextRange(me);
    if (!This->reOle)
        return CO_E_RELEASED;

    if (!pcpLim)
        return E_INVALIDARG;
    *pcpLim = This->cursorPair->end;
    TRACE("%d\n", *pcpLim);
    return S_OK;
}

static HRESULT WINAPI ITextRange_fnSetEnd(ITextRange *me, LONG cpLim)
{
    ITextRangeImpl *This = impl_from_ITextRange(me);
    if (!This->reOle)
        return CO_E_RELEASED;

    FIXME("not implemented %p\n", This);
    return E_NOTIMPL;
}

static HRESULT WINAPI ITextRange_fnGetFont(ITextRange *me, ITextFont **pFont)
{
    ITextRangeImpl *This = impl_from_ITextRange(me);
    if (!This->reOle)
        return CO_E_RELEASED;

    FIXME("not implemented %p\n", This);
    return E_NOTIMPL;
}

static HRESULT WINAPI ITextRange_fnSetFont(ITextRange *me, ITextFont *pFont)
{
    ITextRangeImpl *This = impl_from_ITextRange(me);
    if (!This->reOle)
        return CO_E_RELEASED;

    FIXME("not implemented %p\n", This);
    return E_NOTIMPL;
}

static HRESULT WINAPI ITextRange_fnGetPara(ITextRange *me, ITextPara **ppPara)
{
    ITextRangeImpl *This = impl_from_ITextRange(me);
    if (!This->reOle)
        return CO_E_RELEASED;

    FIXME("not implemented %p\n", This);
    return E_NOTIMPL;
}

static HRESULT WINAPI ITextRange_fnSetPara(ITextRange *me, ITextPara *pPara)
{
    ITextRangeImpl *This = impl_from_ITextRange(me);
    if (!This->reOle)
        return CO_E_RELEASED;

    FIXME("not implemented %p\n", This);
    return E_NOTIMPL;
}

static HRESULT WINAPI ITextRange_fnGetStoryLength(ITextRange *me, LONG *pcch)
{
    ITextRangeImpl *This = impl_from_ITextRange(me);
    if (!This->reOle)
        return CO_E_RELEASED;

    FIXME("not implemented %p\n", This);
    return E_NOTIMPL;
}

static HRESULT WINAPI ITextRange_fnGetStoryType(ITextRange *me, LONG *pValue)
{
    ITextRangeImpl *This = impl_from_ITextRange(me);
    if (!This->reOle)
        return CO_E_RELEASED;

    FIXME("not implemented %p\n", This);
    return E_NOTIMPL;
}

static HRESULT range_Collapse(LONG bStart, LONG *start, LONG *end)
{
  if (*end == *start)
      return S_FALSE;

  if (bStart == tomEnd || bStart == tomFalse)
      *start = *end;
  else
      *end = *start;
  return S_OK;
}

static HRESULT WINAPI ITextRange_fnCollapse(ITextRange *me, LONG bStart)
{
    ITextRangeImpl *This = impl_from_ITextRange(me);
    if (!This->reOle)
        return CO_E_RELEASED;

    return range_Collapse(bStart, &This->cursorPair->start, &This->cursorPair->end);
}

static HRESULT WINAPI ITextRange_fnExpand(ITextRange *me, LONG Unit, LONG *pDelta)
{
    ITextRangeImpl *This = impl_from_ITextRange(me);
    if (!This->reOle)
        return CO_E_RELEASED;

    FIXME("not implemented %p\n", This);
    return E_NOTIMPL;
}

static HRESULT WINAPI ITextRange_fnGetIndex(ITextRange *me, LONG Unit, LONG *pIndex)
{
    ITextRangeImpl *This = impl_from_ITextRange(me);
    if (!This->reOle)
        return CO_E_RELEASED;

    FIXME("not implemented %p\n", This);
    return E_NOTIMPL;
}

static HRESULT WINAPI ITextRange_fnSetIndex(ITextRange *me, LONG Unit, LONG Index,
                                            LONG Extend)
{
    ITextRangeImpl *This = impl_from_ITextRange(me);
    if (!This->reOle)
        return CO_E_RELEASED;

    FIXME("not implemented %p\n", This);
    return E_NOTIMPL;
}

static HRESULT WINAPI ITextRange_fnSetRange(ITextRange *me, LONG cpActive, LONG cpOther)
{
    ITextRangeImpl *This = impl_from_ITextRange(me);
    if (!This->reOle)
        return CO_E_RELEASED;

    FIXME("not implemented %p\n", This);
    return E_NOTIMPL;
}

static HRESULT WINAPI ITextRange_fnInRange(ITextRange *me, ITextRange *pRange, LONG *pb)
{
    ITextRangeImpl *This = impl_from_ITextRange(me);
    if (!This->reOle)
        return CO_E_RELEASED;

    FIXME("not implemented %p\n", This);
    return E_NOTIMPL;
}

static HRESULT WINAPI ITextRange_fnInStory(ITextRange *me, ITextRange *pRange, LONG *pb)
{
    ITextRangeImpl *This = impl_from_ITextRange(me);
    if (!This->reOle)
        return CO_E_RELEASED;

    FIXME("not implemented %p\n", This);
    return E_NOTIMPL;
}

static HRESULT WINAPI ITextRange_fnIsEqual(ITextRange *me, ITextRange *pRange, LONG *pb)
{
    ITextRangeImpl *This = impl_from_ITextRange(me);
    if (!This->reOle)
        return CO_E_RELEASED;

    FIXME("not implemented %p\n", This);
    return E_NOTIMPL;
}

static HRESULT WINAPI ITextRange_fnSelect(ITextRange *me)
{
    ITextRangeImpl *This = impl_from_ITextRange(me);
    if (!This->reOle)
        return CO_E_RELEASED;

    FIXME("not implemented %p\n", This);
    return E_NOTIMPL;
}

static HRESULT WINAPI ITextRange_fnStartOf(ITextRange *me, LONG Unit, LONG Extend,
                                           LONG *pDelta)
{
    ITextRangeImpl *This = impl_from_ITextRange(me);
    if (!This->reOle)
        return CO_E_RELEASED;

    FIXME("not implemented %p\n", This);
    return E_NOTIMPL;
}

static HRESULT WINAPI ITextRange_fnEndOf(ITextRange *me, LONG Unit, LONG Extend,
                                         LONG *pDelta)
{
    ITextRangeImpl *This = impl_from_ITextRange(me);
    if (!This->reOle)
        return CO_E_RELEASED;

    FIXME("not implemented %p\n", This);
    return E_NOTIMPL;
}

static HRESULT WINAPI ITextRange_fnMove(ITextRange *me, LONG Unit, LONG Count, LONG *pDelta)
{
    ITextRangeImpl *This = impl_from_ITextRange(me);
    if (!This->reOle)
        return CO_E_RELEASED;

    FIXME("not implemented %p\n", This);
    return E_NOTIMPL;
}

static HRESULT WINAPI ITextRange_fnMoveStart(ITextRange *me, LONG Unit, LONG Count,
                                             LONG *pDelta)
{
    ITextRangeImpl *This = impl_from_ITextRange(me);
    if (!This->reOle)
        return CO_E_RELEASED;

    FIXME("not implemented %p\n", This);
    return E_NOTIMPL;
}

static HRESULT WINAPI ITextRange_fnMoveEnd(ITextRange *me, LONG Unit, LONG Count,
                                           LONG *pDelta)
{
    ITextRangeImpl *This = impl_from_ITextRange(me);
    if (!This->reOle)
        return CO_E_RELEASED;

    FIXME("not implemented %p\n", This);
    return E_NOTIMPL;
}

static HRESULT WINAPI ITextRange_fnMoveWhile(ITextRange *me, VARIANT *Cset, LONG Count,
                                             LONG *pDelta)
{
    ITextRangeImpl *This = impl_from_ITextRange(me);
    if (!This->reOle)
        return CO_E_RELEASED;

    FIXME("not implemented %p\n", This);
    return E_NOTIMPL;
}

static HRESULT WINAPI ITextRange_fnMoveStartWhile(ITextRange *me, VARIANT *Cset, LONG Count,
                                                  LONG *pDelta)
{
    ITextRangeImpl *This = impl_from_ITextRange(me);
    if (!This->reOle)
        return CO_E_RELEASED;

    FIXME("not implemented %p\n", This);
    return E_NOTIMPL;
}

static HRESULT WINAPI ITextRange_fnMoveEndWhile(ITextRange *me, VARIANT *Cset, LONG Count,
                                                LONG *pDelta)
{
    ITextRangeImpl *This = impl_from_ITextRange(me);
    if (!This->reOle)
        return CO_E_RELEASED;

    FIXME("not implemented %p\n", This);
    return E_NOTIMPL;
}

static HRESULT WINAPI ITextRange_fnMoveUntil(ITextRange *me, VARIANT *Cset, LONG Count,
                                             LONG *pDelta)
{
    ITextRangeImpl *This = impl_from_ITextRange(me);
    if (!This->reOle)
        return CO_E_RELEASED;

    FIXME("not implemented %p\n", This);
    return E_NOTIMPL;
}

static HRESULT WINAPI ITextRange_fnMoveStartUntil(ITextRange *me, VARIANT *Cset, LONG Count,
                                                  LONG *pDelta)
{
    ITextRangeImpl *This = impl_from_ITextRange(me);
    if (!This->reOle)
        return CO_E_RELEASED;

    FIXME("not implemented %p\n", This);
    return E_NOTIMPL;
}

static HRESULT WINAPI ITextRange_fnMoveEndUntil(ITextRange *me, VARIANT *Cset, LONG Count,
                                                LONG *pDelta)
{
    ITextRangeImpl *This = impl_from_ITextRange(me);
    if (!This->reOle)
        return CO_E_RELEASED;

    FIXME("not implemented %p\n", This);
    return E_NOTIMPL;
}

static HRESULT WINAPI ITextRange_fnFindText(ITextRange *me, BSTR bstr, LONG cch, LONG Flags,
                                            LONG *pLength)
{
    ITextRangeImpl *This = impl_from_ITextRange(me);
    if (!This->reOle)
        return CO_E_RELEASED;

    FIXME("not implemented %p\n", This);
    return E_NOTIMPL;
}

static HRESULT WINAPI ITextRange_fnFindTextStart(ITextRange *me, BSTR bstr, LONG cch,
                                                 LONG Flags, LONG *pLength)
{
    ITextRangeImpl *This = impl_from_ITextRange(me);
    if (!This->reOle)
        return CO_E_RELEASED;

    FIXME("not implemented %p\n", This);
    return E_NOTIMPL;
}

static HRESULT WINAPI ITextRange_fnFindTextEnd(ITextRange *me, BSTR bstr, LONG cch,
                                               LONG Flags, LONG *pLength)
{
    ITextRangeImpl *This = impl_from_ITextRange(me);
    if (!This->reOle)
        return CO_E_RELEASED;

    FIXME("not implemented %p\n", This);
    return E_NOTIMPL;
}

static HRESULT WINAPI ITextRange_fnDelete(ITextRange *me, LONG Unit, LONG Count,
                                          LONG *pDelta)
{
    ITextRangeImpl *This = impl_from_ITextRange(me);
    if (!This->reOle)
        return CO_E_RELEASED;

    FIXME("not implemented %p\n", This);
    return E_NOTIMPL;
}

static HRESULT WINAPI ITextRange_fnCut(ITextRange *me, VARIANT *pVar)
{
    ITextRangeImpl *This = impl_from_ITextRange(me);
    if (!This->reOle)
        return CO_E_RELEASED;

    FIXME("not implemented %p\n", This);
    return E_NOTIMPL;
}

static HRESULT WINAPI ITextRange_fnCopy(ITextRange *me, VARIANT *pVar)
{
    ITextRangeImpl *This = impl_from_ITextRange(me);
    if (!This->reOle)
        return CO_E_RELEASED;

    FIXME("not implemented %p\n", This);
    return E_NOTIMPL;
}

static HRESULT WINAPI ITextRange_fnPaste(ITextRange *me, VARIANT *pVar, LONG Format)
{
    ITextRangeImpl *This = impl_from_ITextRange(me);
    if (!This->reOle)
        return CO_E_RELEASED;

    FIXME("not implemented %p\n", This);
    return E_NOTIMPL;
}

static HRESULT WINAPI ITextRange_fnCanPaste(ITextRange *me, VARIANT *pVar, LONG Format,
                                            LONG *pb)
{
    ITextRangeImpl *This = impl_from_ITextRange(me);
    if (!This->reOle)
        return CO_E_RELEASED;

    FIXME("not implemented %p\n", This);
    return E_NOTIMPL;
}

static HRESULT WINAPI ITextRange_fnCanEdit(ITextRange *me, LONG *pb)
{
    ITextRangeImpl *This = impl_from_ITextRange(me);
    if (!This->reOle)
        return CO_E_RELEASED;

    FIXME("not implemented %p\n", This);
    return E_NOTIMPL;
}

static HRESULT WINAPI ITextRange_fnChangeCase(ITextRange *me, LONG Type)
{
    ITextRangeImpl *This = impl_from_ITextRange(me);
    if (!This->reOle)
        return CO_E_RELEASED;

    FIXME("not implemented %p\n", This);
    return E_NOTIMPL;
}

static HRESULT WINAPI ITextRange_fnGetPoint(ITextRange *me, LONG Type, LONG *cx, LONG *cy)
{
    ITextRangeImpl *This = impl_from_ITextRange(me);
    if (!This->reOle)
        return CO_E_RELEASED;

    FIXME("not implemented %p\n", This);
    return E_NOTIMPL;
}

static HRESULT WINAPI ITextRange_fnSetPoint(ITextRange *me, LONG x, LONG y, LONG Type,
                                            LONG Extend)
{
    ITextRangeImpl *This = impl_from_ITextRange(me);
    if (!This->reOle)
        return CO_E_RELEASED;

    FIXME("not implemented %p\n", This);
    return E_NOTIMPL;
}

static HRESULT WINAPI ITextRange_fnScrollIntoView(ITextRange *me, LONG Value)
{
    ITextRangeImpl *This = impl_from_ITextRange(me);
    if (!This->reOle)
        return CO_E_RELEASED;

    FIXME("not implemented %p\n", This);
    return E_NOTIMPL;
}

static HRESULT WINAPI ITextRange_fnGetEmbeddedObject(ITextRange *me, IUnknown **ppv)
{
    ITextRangeImpl *This = impl_from_ITextRange(me);
    if (!This->reOle)
        return CO_E_RELEASED;

    FIXME("not implemented %p\n", This);
    return E_NOTIMPL;
}

static const ITextRangeVtbl trvt = {
    ITextRange_fnQueryInterface,
    ITextRange_fnAddRef,
    ITextRange_fnRelease,
    ITextRange_fnGetTypeInfoCount,
    ITextRange_fnGetTypeInfo,
    ITextRange_fnGetIDsOfNames,
    ITextRange_fnInvoke,
    ITextRange_fnGetText,
    ITextRange_fnSetText,
    ITextRange_fnGetChar,
    ITextRange_fnSetChar,
    ITextRange_fnGetDuplicate,
    ITextRange_fnGetFormattedText,
    ITextRange_fnSetFormattedText,
    ITextRange_fnGetStart,
    ITextRange_fnSetStart,
    ITextRange_fnGetEnd,
    ITextRange_fnSetEnd,
    ITextRange_fnGetFont,
    ITextRange_fnSetFont,
    ITextRange_fnGetPara,
    ITextRange_fnSetPara,
    ITextRange_fnGetStoryLength,
    ITextRange_fnGetStoryType,
    ITextRange_fnCollapse,
    ITextRange_fnExpand,
    ITextRange_fnGetIndex,
    ITextRange_fnSetIndex,
    ITextRange_fnSetRange,
    ITextRange_fnInRange,
    ITextRange_fnInStory,
    ITextRange_fnIsEqual,
    ITextRange_fnSelect,
    ITextRange_fnStartOf,
    ITextRange_fnEndOf,
    ITextRange_fnMove,
    ITextRange_fnMoveStart,
    ITextRange_fnMoveEnd,
    ITextRange_fnMoveWhile,
    ITextRange_fnMoveStartWhile,
    ITextRange_fnMoveEndWhile,
    ITextRange_fnMoveUntil,
    ITextRange_fnMoveStartUntil,
    ITextRange_fnMoveEndUntil,
    ITextRange_fnFindText,
    ITextRange_fnFindTextStart,
    ITextRange_fnFindTextEnd,
    ITextRange_fnDelete,
    ITextRange_fnCut,
    ITextRange_fnCopy,
    ITextRange_fnPaste,
    ITextRange_fnCanPaste,
    ITextRange_fnCanEdit,
    ITextRange_fnChangeCase,
    ITextRange_fnGetPoint,
    ITextRange_fnSetPoint,
    ITextRange_fnScrollIntoView,
    ITextRange_fnGetEmbeddedObject
};
/* ITextRange interface */

static HRESULT WINAPI
ITextDocument_fnQueryInterface(ITextDocument* me, REFIID riid,
    void** ppvObject)
{
    IRichEditOleImpl *This = impl_from_ITextDocument(me);
    return IRichEditOle_QueryInterface(&This->IRichEditOle_iface, riid, ppvObject);
}

static ULONG WINAPI
ITextDocument_fnAddRef(ITextDocument* me)
{
    IRichEditOleImpl *This = impl_from_ITextDocument(me);
    return IRichEditOle_AddRef(&This->IRichEditOle_iface);
}

static ULONG WINAPI
ITextDocument_fnRelease(ITextDocument* me)
{
    IRichEditOleImpl *This = impl_from_ITextDocument(me);
    return IRichEditOle_Release(&This->IRichEditOle_iface);
}

static HRESULT WINAPI
ITextDocument_fnGetTypeInfoCount(ITextDocument* me,
    UINT* pctinfo)
{
    IRichEditOleImpl *This = impl_from_ITextDocument(me);
    FIXME("stub %p\n",This);
    return E_NOTIMPL;
}

static HRESULT WINAPI
ITextDocument_fnGetTypeInfo(ITextDocument* me, UINT iTInfo, LCID lcid,
    ITypeInfo** ppTInfo)
{
    IRichEditOleImpl *This = impl_from_ITextDocument(me);
    FIXME("stub %p\n",This);
    return E_NOTIMPL;
}

static HRESULT WINAPI
ITextDocument_fnGetIDsOfNames(ITextDocument* me, REFIID riid,
    LPOLESTR* rgszNames, UINT cNames, LCID lcid, DISPID* rgDispId)
{
    IRichEditOleImpl *This = impl_from_ITextDocument(me);
    FIXME("stub %p\n",This);
    return E_NOTIMPL;
}

static HRESULT WINAPI
ITextDocument_fnInvoke(ITextDocument* me, DISPID dispIdMember,
    REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS* pDispParams,
    VARIANT* pVarResult, EXCEPINFO* pExcepInfo, UINT* puArgErr)
{
    IRichEditOleImpl *This = impl_from_ITextDocument(me);
    FIXME("stub %p\n",This);
    return E_NOTIMPL;
}

static HRESULT WINAPI
ITextDocument_fnGetName(ITextDocument* me, BSTR* pName)
{
    IRichEditOleImpl *This = impl_from_ITextDocument(me);
    FIXME("stub %p\n",This);
    return E_NOTIMPL;
}

static HRESULT WINAPI
ITextDocument_fnGetSelection(ITextDocument* me, ITextSelection** ppSel)
{
    IRichEditOleImpl *This = impl_from_ITextDocument(me);
    TRACE("(%p)\n", me);

    if(!ppSel)
      return E_INVALIDARG;
    *ppSel = &This->txtSel->ITextSelection_iface;
    ITextSelection_AddRef(*ppSel);
    return S_OK;
}

static HRESULT WINAPI
ITextDocument_fnGetStoryCount(ITextDocument* me, LONG* pCount)
{
    IRichEditOleImpl *This = impl_from_ITextDocument(me);
    FIXME("stub %p\n",This);
    return E_NOTIMPL;
}

static HRESULT WINAPI
ITextDocument_fnGetStoryRanges(ITextDocument* me,
    ITextStoryRanges** ppStories)
{
    IRichEditOleImpl *This = impl_from_ITextDocument(me);
    FIXME("stub %p\n",This);
    return E_NOTIMPL;
}

static HRESULT WINAPI
ITextDocument_fnGetSaved(ITextDocument* me, LONG* pValue)
{
    IRichEditOleImpl *This = impl_from_ITextDocument(me);
    FIXME("stub %p\n",This);
    return E_NOTIMPL;
}

static HRESULT WINAPI
ITextDocument_fnSetSaved(ITextDocument* me, LONG Value)
{
    IRichEditOleImpl *This = impl_from_ITextDocument(me);
    FIXME("stub %p\n",This);
    return E_NOTIMPL;
}

static HRESULT WINAPI
ITextDocument_fnGetDefaultTabStop(ITextDocument* me, float* pValue)
{
    IRichEditOleImpl *This = impl_from_ITextDocument(me);
    FIXME("stub %p\n",This);
    return E_NOTIMPL;
}

static HRESULT WINAPI
ITextDocument_fnSetDefaultTabStop(ITextDocument* me, float Value)
{
    IRichEditOleImpl *This = impl_from_ITextDocument(me);
    FIXME("stub %p\n",This);
    return E_NOTIMPL;
}

static HRESULT WINAPI
ITextDocument_fnNew(ITextDocument* me)
{
    IRichEditOleImpl *This = impl_from_ITextDocument(me);
    FIXME("stub %p\n",This);
    return E_NOTIMPL;
}

static HRESULT WINAPI
ITextDocument_fnOpen(ITextDocument* me, VARIANT* pVar, LONG Flags,
    LONG CodePage)
{
    IRichEditOleImpl *This = impl_from_ITextDocument(me);
    FIXME("stub %p\n",This);
    return E_NOTIMPL;
}

static HRESULT WINAPI
ITextDocument_fnSave(ITextDocument* me, VARIANT* pVar, LONG Flags,
    LONG CodePage)
{
    IRichEditOleImpl *This = impl_from_ITextDocument(me);
    FIXME("stub %p\n",This);
    return E_NOTIMPL;
}

static HRESULT WINAPI
ITextDocument_fnFreeze(ITextDocument* me, LONG* pCount)
{
    IRichEditOleImpl *This = impl_from_ITextDocument(me);
    FIXME("stub %p\n",This);
    return E_NOTIMPL;
}

static HRESULT WINAPI
ITextDocument_fnUnfreeze(ITextDocument* me, LONG* pCount)
{
    IRichEditOleImpl *This = impl_from_ITextDocument(me);
    FIXME("stub %p\n",This);
    return E_NOTIMPL;
}

static HRESULT WINAPI
ITextDocument_fnBeginEditCollection(ITextDocument* me)
{
    IRichEditOleImpl *This = impl_from_ITextDocument(me);
    FIXME("stub %p\n",This);
    return E_NOTIMPL;
}

static HRESULT WINAPI
ITextDocument_fnEndEditCollection(ITextDocument* me)
{
    IRichEditOleImpl *This = impl_from_ITextDocument(me);
    FIXME("stub %p\n",This);
    return E_NOTIMPL;
}

static HRESULT WINAPI
ITextDocument_fnUndo(ITextDocument* me, LONG Count, LONG* prop)
{
    IRichEditOleImpl *This = impl_from_ITextDocument(me);
    FIXME("stub %p\n",This);
    return E_NOTIMPL;
}

static HRESULT WINAPI
ITextDocument_fnRedo(ITextDocument* me, LONG Count, LONG* prop)
{
    IRichEditOleImpl *This = impl_from_ITextDocument(me);
    FIXME("stub %p\n",This);
    return E_NOTIMPL;
}

static HRESULT CreateITextRange(IRichEditOleImpl *reOle, LONG start, LONG end, ITextRange** ppRange)
{
    ITextRangeImpl *txtRge = heap_alloc(sizeof(ITextRangeImpl));

    if (!txtRge)
        return E_OUTOFMEMORY;
    txtRge->ITextRange_iface.lpVtbl = &trvt;
    txtRge->ref = 1;
    txtRge->reOle = reOle;
    txtRge->cursorPair = CreateCursorPair(start, end);
    if (!txtRge->cursorPair)
        return E_OUTOFMEMORY;
    list_add_head(&reOle->rangelist, &txtRge->entry);
    *ppRange = &txtRge->ITextRange_iface;
    return S_OK;
}

static HRESULT WINAPI
ITextDocument_fnRange(ITextDocument* me, LONG cp1, LONG cp2,
    ITextRange** ppRange)
{
    IRichEditOleImpl *This = impl_from_ITextDocument(me);
    const int len = ME_GetTextLength(This->editor) + 1;

    TRACE("%p %p %d %d\n", This, ppRange, cp1, cp2);
    if (!ppRange)
        return E_INVALIDARG;

    cp1 = max(cp1, 0);
    cp2 = max(cp2, 0);
    cp1 = min(cp1, len);
    cp2 = min(cp2, len);
    if (cp1 > cp2)
    {
        LONG tmp;
        tmp = cp1;
        cp1 = cp2;
        cp2 = tmp;
    }
    if (cp1 == len)
        cp1 = cp2 = len - 1;

    return CreateITextRange(This, cp1, cp2, ppRange);
}

static HRESULT WINAPI
ITextDocument_fnRangeFromPoint(ITextDocument* me, LONG x, LONG y,
    ITextRange** ppRange)
{
    IRichEditOleImpl *This = impl_from_ITextDocument(me);
    FIXME("stub %p\n",This);
    return E_NOTIMPL;
}

static const ITextDocumentVtbl tdvt = {
    ITextDocument_fnQueryInterface,
    ITextDocument_fnAddRef,
    ITextDocument_fnRelease,
    ITextDocument_fnGetTypeInfoCount,
    ITextDocument_fnGetTypeInfo,
    ITextDocument_fnGetIDsOfNames,
    ITextDocument_fnInvoke,
    ITextDocument_fnGetName,
    ITextDocument_fnGetSelection,
    ITextDocument_fnGetStoryCount,
    ITextDocument_fnGetStoryRanges,
    ITextDocument_fnGetSaved,
    ITextDocument_fnSetSaved,
    ITextDocument_fnGetDefaultTabStop,
    ITextDocument_fnSetDefaultTabStop,
    ITextDocument_fnNew,
    ITextDocument_fnOpen,
    ITextDocument_fnSave,
    ITextDocument_fnFreeze,
    ITextDocument_fnUnfreeze,
    ITextDocument_fnBeginEditCollection,
    ITextDocument_fnEndEditCollection,
    ITextDocument_fnUndo,
    ITextDocument_fnRedo,
    ITextDocument_fnRange,
    ITextDocument_fnRangeFromPoint
};

static inline ITextSelectionImpl *impl_from_ITextSelection(ITextSelection *iface)
{
    return CONTAINING_RECORD(iface, ITextSelectionImpl, ITextSelection_iface);
}

static HRESULT WINAPI ITextSelection_fnQueryInterface(
    ITextSelection *me,
    REFIID riid,
    void **ppvObj)
{
    *ppvObj = NULL;
    if (IsEqualGUID(riid, &IID_IUnknown)
        || IsEqualGUID(riid, &IID_IDispatch)
        || IsEqualGUID(riid, &IID_ITextRange)
        || IsEqualGUID(riid, &IID_ITextSelection))
    {
        *ppvObj = me;
        ITextSelection_AddRef(me);
        return S_OK;
    }

    return E_NOINTERFACE;
}

static ULONG WINAPI ITextSelection_fnAddRef(ITextSelection *me)
{
    ITextSelectionImpl *This = impl_from_ITextSelection(me);
    return InterlockedIncrement(&This->ref);
}

static ULONG WINAPI ITextSelection_fnRelease(ITextSelection *me)
{
    ITextSelectionImpl *This = impl_from_ITextSelection(me);
    ULONG ref = InterlockedDecrement(&This->ref);
    if (ref == 0)
        heap_free(This);
    return ref;
}

static HRESULT WINAPI ITextSelection_fnGetTypeInfoCount(ITextSelection *me, UINT *pctinfo)
{
    ITextSelectionImpl *This = impl_from_ITextSelection(me);
    if (!This->reOle)
        return CO_E_RELEASED;

    FIXME("not implemented\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI ITextSelection_fnGetTypeInfo(ITextSelection *me, UINT iTInfo, LCID lcid,
    ITypeInfo **ppTInfo)
{
    ITextSelectionImpl *This = impl_from_ITextSelection(me);
    if (!This->reOle)
        return CO_E_RELEASED;

    FIXME("not implemented\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI ITextSelection_fnGetIDsOfNames(ITextSelection *me, REFIID riid,
    LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId)
{
    ITextSelectionImpl *This = impl_from_ITextSelection(me);
    if (!This->reOle)
        return CO_E_RELEASED;

    FIXME("not implemented\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI ITextSelection_fnInvoke(
    ITextSelection *me,
    DISPID dispIdMember,
    REFIID riid,
    LCID lcid,
    WORD wFlags,
    DISPPARAMS *pDispParams,
    VARIANT *pVarResult,
    EXCEPINFO *pExcepInfo,
    UINT *puArgErr)
{
    FIXME("not implemented\n");
    return E_NOTIMPL;
}

/*** ITextRange methods ***/
static HRESULT WINAPI ITextSelection_fnGetText(ITextSelection *me, BSTR *pbstr)
{
    ITextSelectionImpl *This = impl_from_ITextSelection(me);
    ME_Cursor *start = NULL, *end = NULL;
    int nChars, endOfs;
    BOOL bEOP;

    if (!This->reOle)
        return CO_E_RELEASED;
    TRACE("%p\n", pbstr);
    if (!pbstr)
        return E_INVALIDARG;

    ME_GetSelection(This->reOle->editor, &start, &end);
    endOfs = ME_GetCursorOfs(end);
    nChars = endOfs - ME_GetCursorOfs(start);
    if (!nChars)
    {
        *pbstr = NULL;
        return S_OK;
    }

    *pbstr = SysAllocStringLen(NULL, nChars);
    if (!*pbstr)
        return E_OUTOFMEMORY;

    bEOP = (end->pRun->next->type == diTextEnd && endOfs > ME_GetTextLength(This->reOle->editor));
    ME_GetTextW(This->reOle->editor, *pbstr, nChars, start, nChars, FALSE, bEOP);
    TRACE("%s\n", wine_dbgstr_w(*pbstr));

    return S_OK;
}

static HRESULT WINAPI ITextSelection_fnSetText(ITextSelection *me, BSTR bstr)
{
    ITextSelectionImpl *This = impl_from_ITextSelection(me);
    if (!This->reOle)
        return CO_E_RELEASED;

    FIXME("not implemented\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI ITextSelection_fnGetChar(ITextSelection *me, LONG *pch)
{
    ITextSelectionImpl *This = impl_from_ITextSelection(me);
    ME_Cursor *start = NULL, *end = NULL;

    if (!This->reOle)
        return CO_E_RELEASED;
    TRACE("%p\n", pch);
    if (!pch)
        return E_INVALIDARG;

    ME_GetSelection(This->reOle->editor, &start, &end);
    return range_GetChar(This->reOle->editor, start, pch);
}

static HRESULT WINAPI ITextSelection_fnSetChar(ITextSelection *me, LONG ch)
{
    ITextSelectionImpl *This = impl_from_ITextSelection(me);
    if (!This->reOle)
        return CO_E_RELEASED;

    FIXME("not implemented\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI ITextSelection_fnGetDuplicate(ITextSelection *me, ITextRange **ppRange)
{
    ITextSelectionImpl *This = impl_from_ITextSelection(me);
    if (!This->reOle)
        return CO_E_RELEASED;

    FIXME("not implemented\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI ITextSelection_fnGetFormattedText(ITextSelection *me, ITextRange **ppRange)
{
    ITextSelectionImpl *This = impl_from_ITextSelection(me);
    if (!This->reOle)
        return CO_E_RELEASED;

    FIXME("not implemented\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI ITextSelection_fnSetFormattedText(ITextSelection *me, ITextRange *pRange)
{
    ITextSelectionImpl *This = impl_from_ITextSelection(me);
    if (!This->reOle)
        return CO_E_RELEASED;

    FIXME("not implemented\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI ITextSelection_fnGetStart(ITextSelection *me, LONG *pcpFirst)
{
    ITextSelectionImpl *This = impl_from_ITextSelection(me);
    LONG lim;
    if (!This->reOle)
        return CO_E_RELEASED;

    if (!pcpFirst)
        return E_INVALIDARG;
    ME_GetSelectionOfs(This->reOle->editor, pcpFirst, &lim);
    TRACE("%d\n", *pcpFirst);
    return S_OK;
}

static HRESULT WINAPI ITextSelection_fnSetStart(ITextSelection *me, LONG cpFirst)
{
    ITextSelectionImpl *This = impl_from_ITextSelection(me);
    if (!This->reOle)
        return CO_E_RELEASED;

    FIXME("not implemented\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI ITextSelection_fnGetEnd(ITextSelection *me, LONG *pcpLim)
{
    ITextSelectionImpl *This = impl_from_ITextSelection(me);
    LONG first;
    if (!This->reOle)
        return CO_E_RELEASED;

    if (!pcpLim)
        return E_INVALIDARG;
    ME_GetSelectionOfs(This->reOle->editor, &first, pcpLim);
    TRACE("%d\n", *pcpLim);
    return S_OK;
}

static HRESULT WINAPI ITextSelection_fnSetEnd(ITextSelection *me, LONG cpLim)
{
    ITextSelectionImpl *This = impl_from_ITextSelection(me);
    if (!This->reOle)
        return CO_E_RELEASED;

    FIXME("not implemented\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI ITextSelection_fnGetFont(ITextSelection *me, ITextFont **pFont)
{
    ITextSelectionImpl *This = impl_from_ITextSelection(me);
    if (!This->reOle)
        return CO_E_RELEASED;

    FIXME("not implemented\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI ITextSelection_fnSetFont(ITextSelection *me, ITextFont *pFont)
{
    ITextSelectionImpl *This = impl_from_ITextSelection(me);
    if (!This->reOle)
        return CO_E_RELEASED;

    FIXME("not implemented\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI ITextSelection_fnGetPara(ITextSelection *me, ITextPara **ppPara)
{
    ITextSelectionImpl *This = impl_from_ITextSelection(me);
    if (!This->reOle)
        return CO_E_RELEASED;

    FIXME("not implemented\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI ITextSelection_fnSetPara(ITextSelection *me, ITextPara *pPara)
{
    ITextSelectionImpl *This = impl_from_ITextSelection(me);
    if (!This->reOle)
        return CO_E_RELEASED;

    FIXME("not implemented\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI ITextSelection_fnGetStoryLength(ITextSelection *me, LONG *pcch)
{
    ITextSelectionImpl *This = impl_from_ITextSelection(me);
    if (!This->reOle)
        return CO_E_RELEASED;

    FIXME("not implemented\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI ITextSelection_fnGetStoryType(ITextSelection *me, LONG *pValue)
{
    ITextSelectionImpl *This = impl_from_ITextSelection(me);
    if (!This->reOle)
        return CO_E_RELEASED;

    FIXME("not implemented\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI ITextSelection_fnCollapse(ITextSelection *me, LONG bStart)
{
    ITextSelectionImpl *This = impl_from_ITextSelection(me);
    LONG start, end;
    HRESULT hres;
    if (!This->reOle)
        return CO_E_RELEASED;

    ME_GetSelectionOfs(This->reOle->editor, &start, &end);
    hres = range_Collapse(bStart, &start, &end);
    if (SUCCEEDED(hres))
        ME_SetSelection(This->reOle->editor, start, end);
    return hres;
}

static HRESULT WINAPI ITextSelection_fnExpand(ITextSelection *me, LONG Unit, LONG *pDelta)
{
    ITextSelectionImpl *This = impl_from_ITextSelection(me);
    if (!This->reOle)
        return CO_E_RELEASED;

    FIXME("not implemented\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI ITextSelection_fnGetIndex(ITextSelection *me, LONG Unit, LONG *pIndex)
{
    ITextSelectionImpl *This = impl_from_ITextSelection(me);
    if (!This->reOle)
        return CO_E_RELEASED;

    FIXME("not implemented\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI ITextSelection_fnSetIndex(ITextSelection *me, LONG Unit, LONG Index,
    LONG Extend)
{
    ITextSelectionImpl *This = impl_from_ITextSelection(me);
    if (!This->reOle)
        return CO_E_RELEASED;

    FIXME("not implemented\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI ITextSelection_fnSetRange(ITextSelection *me, LONG cpActive, LONG cpOther)
{
    ITextSelectionImpl *This = impl_from_ITextSelection(me);
    if (!This->reOle)
        return CO_E_RELEASED;

    FIXME("not implemented\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI ITextSelection_fnInRange(ITextSelection *me, ITextRange *pRange, LONG *pb)
{
    ITextSelectionImpl *This = impl_from_ITextSelection(me);
    if (!This->reOle)
        return CO_E_RELEASED;

    FIXME("not implemented\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI ITextSelection_fnInStory(ITextSelection *me, ITextRange *pRange, LONG *pb)
{
    ITextSelectionImpl *This = impl_from_ITextSelection(me);
    if (!This->reOle)
        return CO_E_RELEASED;

    FIXME("not implemented\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI ITextSelection_fnIsEqual(ITextSelection *me, ITextRange *pRange, LONG *pb)
{
    ITextSelectionImpl *This = impl_from_ITextSelection(me);
    if (!This->reOle)
        return CO_E_RELEASED;

    FIXME("not implemented\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI ITextSelection_fnSelect(ITextSelection *me)
{
    ITextSelectionImpl *This = impl_from_ITextSelection(me);
    if (!This->reOle)
        return CO_E_RELEASED;

    FIXME("not implemented\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI ITextSelection_fnStartOf(ITextSelection *me, LONG Unit, LONG Extend,
    LONG *pDelta)
{
    ITextSelectionImpl *This = impl_from_ITextSelection(me);
    if (!This->reOle)
        return CO_E_RELEASED;

    FIXME("not implemented\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI ITextSelection_fnEndOf(ITextSelection *me, LONG Unit, LONG Extend,
    LONG *pDelta)
{
    ITextSelectionImpl *This = impl_from_ITextSelection(me);
    if (!This->reOle)
        return CO_E_RELEASED;

    FIXME("not implemented\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI ITextSelection_fnMove(ITextSelection *me, LONG Unit, LONG Count, LONG *pDelta)
{
    ITextSelectionImpl *This = impl_from_ITextSelection(me);
    if (!This->reOle)
        return CO_E_RELEASED;

    FIXME("not implemented\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI ITextSelection_fnMoveStart(ITextSelection *me, LONG Unit, LONG Count,
    LONG *pDelta)
{
    ITextSelectionImpl *This = impl_from_ITextSelection(me);
    if (!This->reOle)
        return CO_E_RELEASED;

    FIXME("not implemented\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI ITextSelection_fnMoveEnd(ITextSelection *me, LONG Unit, LONG Count,
    LONG *pDelta)
{
    ITextSelectionImpl *This = impl_from_ITextSelection(me);
    if (!This->reOle)
        return CO_E_RELEASED;

    FIXME("not implemented\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI ITextSelection_fnMoveWhile(ITextSelection *me, VARIANT *Cset, LONG Count,
    LONG *pDelta)
{
    ITextSelectionImpl *This = impl_from_ITextSelection(me);
    if (!This->reOle)
        return CO_E_RELEASED;

    FIXME("not implemented\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI ITextSelection_fnMoveStartWhile(ITextSelection *me, VARIANT *Cset, LONG Count,
    LONG *pDelta)
{
    ITextSelectionImpl *This = impl_from_ITextSelection(me);
    if (!This->reOle)
        return CO_E_RELEASED;

    FIXME("not implemented\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI ITextSelection_fnMoveEndWhile(ITextSelection *me, VARIANT *Cset, LONG Count,
    LONG *pDelta)
{
    ITextSelectionImpl *This = impl_from_ITextSelection(me);
    if (!This->reOle)
        return CO_E_RELEASED;

    FIXME("not implemented\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI ITextSelection_fnMoveUntil(ITextSelection *me, VARIANT *Cset, LONG Count,
    LONG *pDelta)
{
    ITextSelectionImpl *This = impl_from_ITextSelection(me);
    if (!This->reOle)
        return CO_E_RELEASED;

    FIXME("not implemented\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI ITextSelection_fnMoveStartUntil(ITextSelection *me, VARIANT *Cset, LONG Count,
    LONG *pDelta)
{
    ITextSelectionImpl *This = impl_from_ITextSelection(me);
    if (!This->reOle)
        return CO_E_RELEASED;

    FIXME("not implemented\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI ITextSelection_fnMoveEndUntil(ITextSelection *me, VARIANT *Cset, LONG Count,
    LONG *pDelta)
{
    ITextSelectionImpl *This = impl_from_ITextSelection(me);
    if (!This->reOle)
        return CO_E_RELEASED;

    FIXME("not implemented\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI ITextSelection_fnFindText(ITextSelection *me, BSTR bstr, LONG cch, LONG Flags,
    LONG *pLength)
{
    ITextSelectionImpl *This = impl_from_ITextSelection(me);
    if (!This->reOle)
        return CO_E_RELEASED;

    FIXME("not implemented\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI ITextSelection_fnFindTextStart(ITextSelection *me, BSTR bstr, LONG cch,
    LONG Flags, LONG *pLength)
{
    ITextSelectionImpl *This = impl_from_ITextSelection(me);
    if (!This->reOle)
        return CO_E_RELEASED;

    FIXME("not implemented\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI ITextSelection_fnFindTextEnd(ITextSelection *me, BSTR bstr, LONG cch,
    LONG Flags, LONG *pLength)
{
    ITextSelectionImpl *This = impl_from_ITextSelection(me);
    if (!This->reOle)
        return CO_E_RELEASED;

    FIXME("not implemented\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI ITextSelection_fnDelete(ITextSelection *me, LONG Unit, LONG Count,
    LONG *pDelta)
{
    ITextSelectionImpl *This = impl_from_ITextSelection(me);
    if (!This->reOle)
        return CO_E_RELEASED;

    FIXME("not implemented\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI ITextSelection_fnCut(ITextSelection *me, VARIANT *pVar)
{
    ITextSelectionImpl *This = impl_from_ITextSelection(me);
    if (!This->reOle)
        return CO_E_RELEASED;

    FIXME("not implemented\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI ITextSelection_fnCopy(ITextSelection *me, VARIANT *pVar)
{
    ITextSelectionImpl *This = impl_from_ITextSelection(me);
    if (!This->reOle)
        return CO_E_RELEASED;

    FIXME("not implemented\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI ITextSelection_fnPaste(ITextSelection *me, VARIANT *pVar, LONG Format)
{
    ITextSelectionImpl *This = impl_from_ITextSelection(me);
    if (!This->reOle)
        return CO_E_RELEASED;

    FIXME("not implemented\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI ITextSelection_fnCanPaste(ITextSelection *me, VARIANT *pVar, LONG Format,
    LONG *pb)
{
    ITextSelectionImpl *This = impl_from_ITextSelection(me);
    if (!This->reOle)
        return CO_E_RELEASED;

    FIXME("not implemented\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI ITextSelection_fnCanEdit(ITextSelection *me, LONG *pb)
{
    ITextSelectionImpl *This = impl_from_ITextSelection(me);
    if (!This->reOle)
        return CO_E_RELEASED;

    FIXME("not implemented\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI ITextSelection_fnChangeCase(ITextSelection *me, LONG Type)
{
    ITextSelectionImpl *This = impl_from_ITextSelection(me);
    if (!This->reOle)
        return CO_E_RELEASED;

    FIXME("not implemented\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI ITextSelection_fnGetPoint(ITextSelection *me, LONG Type, LONG *cx, LONG *cy)
{
    ITextSelectionImpl *This = impl_from_ITextSelection(me);
    if (!This->reOle)
        return CO_E_RELEASED;

    FIXME("not implemented\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI ITextSelection_fnSetPoint(ITextSelection *me, LONG x, LONG y, LONG Type,
    LONG Extend)
{
    ITextSelectionImpl *This = impl_from_ITextSelection(me);
    if (!This->reOle)
        return CO_E_RELEASED;

    FIXME("not implemented\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI ITextSelection_fnScrollIntoView(ITextSelection *me, LONG Value)
{
    ITextSelectionImpl *This = impl_from_ITextSelection(me);
    if (!This->reOle)
        return CO_E_RELEASED;

    FIXME("not implemented\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI ITextSelection_fnGetEmbeddedObject(ITextSelection *me, IUnknown **ppv)
{
    ITextSelectionImpl *This = impl_from_ITextSelection(me);
    if (!This->reOle)
        return CO_E_RELEASED;

    FIXME("not implemented\n");
    return E_NOTIMPL;
}

/*** ITextSelection methods ***/
static HRESULT WINAPI ITextSelection_fnGetFlags(ITextSelection *me, LONG *pFlags)
{
    ITextSelectionImpl *This = impl_from_ITextSelection(me);
    if (!This->reOle)
        return CO_E_RELEASED;

    FIXME("not implemented\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI ITextSelection_fnSetFlags(ITextSelection *me, LONG Flags)
{
    ITextSelectionImpl *This = impl_from_ITextSelection(me);
    if (!This->reOle)
        return CO_E_RELEASED;

    FIXME("not implemented\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI ITextSelection_fnGetType(ITextSelection *me, LONG *pType)
{
    ITextSelectionImpl *This = impl_from_ITextSelection(me);
    if (!This->reOle)
        return CO_E_RELEASED;

    FIXME("not implemented\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI ITextSelection_fnMoveLeft(ITextSelection *me, LONG Unit, LONG Count,
    LONG Extend, LONG *pDelta)
{
    ITextSelectionImpl *This = impl_from_ITextSelection(me);
    if (!This->reOle)
        return CO_E_RELEASED;

    FIXME("not implemented\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI ITextSelection_fnMoveRight(ITextSelection *me, LONG Unit, LONG Count,
    LONG Extend, LONG *pDelta)
{
    ITextSelectionImpl *This = impl_from_ITextSelection(me);
    if (!This->reOle)
        return CO_E_RELEASED;

    FIXME("not implemented\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI ITextSelection_fnMoveUp(ITextSelection *me, LONG Unit, LONG Count,
    LONG Extend, LONG *pDelta)
{
    ITextSelectionImpl *This = impl_from_ITextSelection(me);
    if (!This->reOle)
        return CO_E_RELEASED;

    FIXME("not implemented\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI ITextSelection_fnMoveDown(ITextSelection *me, LONG Unit, LONG Count,
    LONG Extend, LONG *pDelta)
{
    ITextSelectionImpl *This = impl_from_ITextSelection(me);
    if (!This->reOle)
        return CO_E_RELEASED;

    FIXME("not implemented\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI ITextSelection_fnHomeKey(ITextSelection *me, LONG Unit, LONG Extend,
    LONG *pDelta)
{
    ITextSelectionImpl *This = impl_from_ITextSelection(me);
    if (!This->reOle)
        return CO_E_RELEASED;

    FIXME("not implemented\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI ITextSelection_fnEndKey(ITextSelection *me, LONG Unit, LONG Extend,
    LONG *pDelta)
{
    ITextSelectionImpl *This = impl_from_ITextSelection(me);
    if (!This->reOle)
        return CO_E_RELEASED;

    FIXME("not implemented\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI ITextSelection_fnTypeText(ITextSelection *me, BSTR bstr)
{
    ITextSelectionImpl *This = impl_from_ITextSelection(me);
    if (!This->reOle)
        return CO_E_RELEASED;

    FIXME("not implemented\n");
    return E_NOTIMPL;
}

static const ITextSelectionVtbl tsvt = {
    ITextSelection_fnQueryInterface,
    ITextSelection_fnAddRef,
    ITextSelection_fnRelease,
    ITextSelection_fnGetTypeInfoCount,
    ITextSelection_fnGetTypeInfo,
    ITextSelection_fnGetIDsOfNames,
    ITextSelection_fnInvoke,
    ITextSelection_fnGetText,
    ITextSelection_fnSetText,
    ITextSelection_fnGetChar,
    ITextSelection_fnSetChar,
    ITextSelection_fnGetDuplicate,
    ITextSelection_fnGetFormattedText,
    ITextSelection_fnSetFormattedText,
    ITextSelection_fnGetStart,
    ITextSelection_fnSetStart,
    ITextSelection_fnGetEnd,
    ITextSelection_fnSetEnd,
    ITextSelection_fnGetFont,
    ITextSelection_fnSetFont,
    ITextSelection_fnGetPara,
    ITextSelection_fnSetPara,
    ITextSelection_fnGetStoryLength,
    ITextSelection_fnGetStoryType,
    ITextSelection_fnCollapse,
    ITextSelection_fnExpand,
    ITextSelection_fnGetIndex,
    ITextSelection_fnSetIndex,
    ITextSelection_fnSetRange,
    ITextSelection_fnInRange,
    ITextSelection_fnInStory,
    ITextSelection_fnIsEqual,
    ITextSelection_fnSelect,
    ITextSelection_fnStartOf,
    ITextSelection_fnEndOf,
    ITextSelection_fnMove,
    ITextSelection_fnMoveStart,
    ITextSelection_fnMoveEnd,
    ITextSelection_fnMoveWhile,
    ITextSelection_fnMoveStartWhile,
    ITextSelection_fnMoveEndWhile,
    ITextSelection_fnMoveUntil,
    ITextSelection_fnMoveStartUntil,
    ITextSelection_fnMoveEndUntil,
    ITextSelection_fnFindText,
    ITextSelection_fnFindTextStart,
    ITextSelection_fnFindTextEnd,
    ITextSelection_fnDelete,
    ITextSelection_fnCut,
    ITextSelection_fnCopy,
    ITextSelection_fnPaste,
    ITextSelection_fnCanPaste,
    ITextSelection_fnCanEdit,
    ITextSelection_fnChangeCase,
    ITextSelection_fnGetPoint,
    ITextSelection_fnSetPoint,
    ITextSelection_fnScrollIntoView,
    ITextSelection_fnGetEmbeddedObject,
    ITextSelection_fnGetFlags,
    ITextSelection_fnSetFlags,
    ITextSelection_fnGetType,
    ITextSelection_fnMoveLeft,
    ITextSelection_fnMoveRight,
    ITextSelection_fnMoveUp,
    ITextSelection_fnMoveDown,
    ITextSelection_fnHomeKey,
    ITextSelection_fnEndKey,
    ITextSelection_fnTypeText
};

static ITextSelectionImpl *
CreateTextSelection(IRichEditOleImpl *reOle)
{
    ITextSelectionImpl *txtSel = heap_alloc(sizeof *txtSel);
    if (!txtSel)
        return NULL;

    txtSel->ITextSelection_iface.lpVtbl = &tsvt;
    txtSel->ref = 1;
    txtSel->reOle = reOle;
    return txtSel;
}

LRESULT CreateIRichEditOle(ME_TextEditor *editor, LPVOID *ppObj)
{
    IRichEditOleImpl *reo;

    reo = heap_alloc(sizeof(IRichEditOleImpl));
    if (!reo)
        return 0;

    reo->IRichEditOle_iface.lpVtbl = &revt;
    reo->ITextDocument_iface.lpVtbl = &tdvt;
    reo->ref = 1;
    reo->editor = editor;
    reo->txtSel = CreateTextSelection(reo);
    if (!reo->txtSel)
    {
        heap_free(reo);
        return 0;
    }
    reo->clientSite = CreateOleClientSite(reo);
    if (!reo->clientSite)
    {
        ITextSelection_Release(&reo->txtSel->ITextSelection_iface);
        heap_free(reo);
        return 0;
    }
    TRACE("Created %p\n",reo);
    *ppObj = reo;
    list_init(&reo->rangelist);

    return 1;
}

static void convert_sizel(const ME_Context *c, const SIZEL* szl, SIZE* sz)
{
  /* sizel is in .01 millimeters, sz in pixels */
  sz->cx = MulDiv(szl->cx, c->dpi.cx, 2540);
  sz->cy = MulDiv(szl->cy, c->dpi.cy, 2540);
}

/******************************************************************************
 * ME_GetOLEObjectSize
 *
 * Sets run extent for OLE objects.
 */
void ME_GetOLEObjectSize(const ME_Context *c, ME_Run *run, SIZE *pSize)
{
  IDataObject*  ido;
  FORMATETC     fmt;
  STGMEDIUM     stgm;
  DIBSECTION    dibsect;
  ENHMETAHEADER emh;

  assert(run->nFlags & MERF_GRAPHICS);
  assert(run->ole_obj);

  if (run->ole_obj->sizel.cx != 0 || run->ole_obj->sizel.cy != 0)
  {
    convert_sizel(c, &run->ole_obj->sizel, pSize);
    if (c->editor->nZoomNumerator != 0)
    {
      pSize->cx = MulDiv(pSize->cx, c->editor->nZoomNumerator, c->editor->nZoomDenominator);
      pSize->cy = MulDiv(pSize->cy, c->editor->nZoomNumerator, c->editor->nZoomDenominator);
    }
    return;
  }

  if (IOleObject_QueryInterface(run->ole_obj->poleobj, &IID_IDataObject, (void**)&ido) != S_OK)
  {
      FIXME("Query Interface IID_IDataObject failed!\n");
      pSize->cx = pSize->cy = 0;
      return;
  }
  fmt.cfFormat = CF_BITMAP;
  fmt.ptd = NULL;
  fmt.dwAspect = DVASPECT_CONTENT;
  fmt.lindex = -1;
  fmt.tymed = TYMED_GDI;
  if (IDataObject_GetData(ido, &fmt, &stgm) != S_OK)
  {
    fmt.cfFormat = CF_ENHMETAFILE;
    fmt.tymed = TYMED_ENHMF;
    if (IDataObject_GetData(ido, &fmt, &stgm) != S_OK)
    {
      FIXME("unsupported format\n");
      pSize->cx = pSize->cy = 0;
      IDataObject_Release(ido);
      return;
    }
  }

  switch (stgm.tymed)
  {
  case TYMED_GDI:
    GetObjectW(stgm.u.hBitmap, sizeof(dibsect), &dibsect);
    pSize->cx = dibsect.dsBm.bmWidth;
    pSize->cy = dibsect.dsBm.bmHeight;
    if (!stgm.pUnkForRelease) DeleteObject(stgm.u.hBitmap);
    break;
  case TYMED_ENHMF:
    GetEnhMetaFileHeader(stgm.u.hEnhMetaFile, sizeof(emh), &emh);
    pSize->cx = emh.rclBounds.right - emh.rclBounds.left;
    pSize->cy = emh.rclBounds.bottom - emh.rclBounds.top;
    if (!stgm.pUnkForRelease) DeleteEnhMetaFile(stgm.u.hEnhMetaFile);
    break;
  default:
    FIXME("Unsupported tymed %d\n", stgm.tymed);
    break;
  }
  IDataObject_Release(ido);
  if (c->editor->nZoomNumerator != 0)
  {
    pSize->cx = MulDiv(pSize->cx, c->editor->nZoomNumerator, c->editor->nZoomDenominator);
    pSize->cy = MulDiv(pSize->cy, c->editor->nZoomNumerator, c->editor->nZoomDenominator);
  }
}

void ME_DrawOLE(ME_Context *c, int x, int y, ME_Run *run,
                ME_Paragraph *para, BOOL selected)
{
  IDataObject*  ido;
  FORMATETC     fmt;
  STGMEDIUM     stgm;
  DIBSECTION    dibsect;
  ENHMETAHEADER emh;
  HDC           hMemDC;
  SIZE          sz;
  BOOL          has_size;

  assert(run->nFlags & MERF_GRAPHICS);
  assert(run->ole_obj);
  if (IOleObject_QueryInterface(run->ole_obj->poleobj, &IID_IDataObject, (void**)&ido) != S_OK)
  {
    FIXME("Couldn't get interface\n");
    return;
  }
  has_size = run->ole_obj->sizel.cx != 0 || run->ole_obj->sizel.cy != 0;
  fmt.cfFormat = CF_BITMAP;
  fmt.ptd = NULL;
  fmt.dwAspect = DVASPECT_CONTENT;
  fmt.lindex = -1;
  fmt.tymed = TYMED_GDI;
  if (IDataObject_GetData(ido, &fmt, &stgm) != S_OK)
  {
    fmt.cfFormat = CF_ENHMETAFILE;
    fmt.tymed = TYMED_ENHMF;
    if (IDataObject_GetData(ido, &fmt, &stgm) != S_OK)
    {
      FIXME("Couldn't get storage medium\n");
      IDataObject_Release(ido);
      return;
    }
  }
  switch (stgm.tymed)
  {
  case TYMED_GDI:
    GetObjectW(stgm.u.hBitmap, sizeof(dibsect), &dibsect);
    hMemDC = CreateCompatibleDC(c->hDC);
    SelectObject(hMemDC, stgm.u.hBitmap);
    if (has_size)
    {
      convert_sizel(c, &run->ole_obj->sizel, &sz);
    } else {
      sz.cx = MulDiv(dibsect.dsBm.bmWidth, c->dpi.cx, 96);
      sz.cy = MulDiv(dibsect.dsBm.bmHeight, c->dpi.cy, 96);
    }
    if (c->editor->nZoomNumerator != 0)
    {
      sz.cx = MulDiv(sz.cx, c->editor->nZoomNumerator, c->editor->nZoomDenominator);
      sz.cy = MulDiv(sz.cy, c->editor->nZoomNumerator, c->editor->nZoomDenominator);
    }
    if (sz.cx == dibsect.dsBm.bmWidth && sz.cy == dibsect.dsBm.bmHeight)
    {
      BitBlt(c->hDC, x, y - sz.cy,
             dibsect.dsBm.bmWidth, dibsect.dsBm.bmHeight,
             hMemDC, 0, 0, SRCCOPY);
    } else {
      StretchBlt(c->hDC, x, y - sz.cy, sz.cx, sz.cy,
                 hMemDC, 0, 0, dibsect.dsBm.bmWidth,
                 dibsect.dsBm.bmHeight, SRCCOPY);
    }
    DeleteDC(hMemDC);
    if (!stgm.pUnkForRelease) DeleteObject(stgm.u.hBitmap);
    break;
  case TYMED_ENHMF:
    GetEnhMetaFileHeader(stgm.u.hEnhMetaFile, sizeof(emh), &emh);
    if (has_size)
    {
      convert_sizel(c, &run->ole_obj->sizel, &sz);
    } else {
      sz.cy = MulDiv(emh.rclBounds.bottom - emh.rclBounds.top, c->dpi.cx, 96);
      sz.cx = MulDiv(emh.rclBounds.right - emh.rclBounds.left, c->dpi.cy, 96);
    }
    if (c->editor->nZoomNumerator != 0)
    {
      sz.cx = MulDiv(sz.cx, c->editor->nZoomNumerator, c->editor->nZoomDenominator);
      sz.cy = MulDiv(sz.cy, c->editor->nZoomNumerator, c->editor->nZoomDenominator);
    }

    {
      RECT    rc;

      rc.left = x;
      rc.top = y - sz.cy;
      rc.right = x + sz.cx;
      rc.bottom = y;
      PlayEnhMetaFile(c->hDC, stgm.u.hEnhMetaFile, &rc);
    }
    if (!stgm.pUnkForRelease) DeleteEnhMetaFile(stgm.u.hEnhMetaFile);
    break;
  default:
    FIXME("Unsupported tymed %d\n", stgm.tymed);
    selected = FALSE;
    break;
  }
  if (selected && !c->editor->bHideSelection)
    PatBlt(c->hDC, x, y - sz.cy, sz.cx, sz.cy, DSTINVERT);
  IDataObject_Release(ido);
}

void ME_DeleteReObject(REOBJECT* reo)
{
    if (reo->poleobj)   IOleObject_Release(reo->poleobj);
    if (reo->pstg)      IStorage_Release(reo->pstg);
    if (reo->polesite)  IOleClientSite_Release(reo->polesite);
    FREE_OBJ(reo);
}

void ME_CopyReObject(REOBJECT* dst, const REOBJECT* src)
{
    *dst = *src;

    if (dst->poleobj)   IOleObject_AddRef(dst->poleobj);
    if (dst->pstg)      IStorage_AddRef(dst->pstg);
    if (dst->polesite)  IOleClientSite_AddRef(dst->polesite);
}
