//	Copyright (C) 2011 by Artem A. Gevorkyan (gevorkyan.org)
//
//	Permission is hereby granted, free of charge, to any person obtaining a copy
//	of this software and associated documentation files (the "Software"), to deal
//	in the Software without restriction, including without limitation the rights
//	to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
//	copies of the Software, and to permit persons to whom the Software is
//	furnished to do so, subject to the following conditions:
//
//	The above copyright notice and this permission notice shall be included in
//	all copies or substantial portions of the Software.
//
//	THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
//	IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
//	FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
//	AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
//	LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
//	OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
//	THE SOFTWARE.

#pragma once

#include <atlbase.h>
#include <atlcom.h>
#include <memory>
#include <string>

#pragma warning(disable: 4278)
#pragma warning(disable: 4146)
	#import <MSADDNDR.DLL> raw_interfaces_only rename_namespace("msaddin")
	#import <dte80a.olb> rename_namespace("msaddin")
#pragma warning(default: 4146)
#pragma warning(default: 4278)

namespace ea
{
	template <class AppT, const CLSID *ClassID, int RegResID>
	class ATL_NO_VTABLE addin
		: public CComObjectRootEx<CComSingleThreadModel>,
			public CComCoClass<addin<AppT, ClassID, RegResID>, ClassID>,
			public IDispatchImpl<msaddin::IDTExtensibility2, &__uuidof(msaddin::IDTExtensibility2), &__uuidof(msaddin::__AddInDesignerObjects), 1, 0>
	{
		std::auto_ptr<AppT> _application;

		// IDTExtensibility2 methods
		STDMETHODIMP OnConnection(IDispatch *host, msaddin::ext_ConnectMode connectMode, IDispatch *instance, SAFEARRAY **custom);
		STDMETHODIMP OnDisconnection(msaddin::ext_DisconnectMode removeMode, SAFEARRAY **custom);
		STDMETHODIMP OnAddInsUpdate(SAFEARRAY **custom);
		STDMETHODIMP OnStartupComplete(SAFEARRAY **custom);
		STDMETHODIMP OnBeginShutdown(SAFEARRAY **custom);

	public:
		DECLARE_REGISTRY_RESOURCEID(RegResID)
		DECLARE_NOT_AGGREGATABLE(addin)

		BEGIN_COM_MAP(addin)
			COM_INTERFACE_ENTRY2(IDispatch, msaddin::IDTExtensibility2)
			COM_INTERFACE_ENTRY(msaddin::IDTExtensibility2)
		END_COM_MAP()
	};

	struct command
	{
		virtual ~command() {	}
		
		virtual std::wstring id() const = 0;
		virtual std::wstring caption() const = 0;
		virtual std::wstring description() const = 0;
		virtual void update_ui(msaddin::CommandPtr cmd, IDispatchPtr command_bars) const = 0;
		virtual bool query_status(msaddin::_DTEPtr dte, bool &enabled, std::wstring *caption, std::wstring *description) const = 0;
		virtual void execute(msaddin::_DTEPtr dte, VARIANT *input, VARIANT *output) const = 0;
	};


	template <class AppT, const CLSID *ClassID, int RegResID>
	inline STDMETHODIMP addin<AppT, ClassID, RegResID>::OnConnection(IDispatch *host, msaddin::ext_ConnectMode connectMode, IDispatch *instance, SAFEARRAY **custom)
	try
	{
		_application.reset(new AppT(IDispatchPtr(host, true)));
		return S_OK;
	}
	catch (...)
	{
		return E_FAIL;
	}

	template <class AppT, const CLSID *ClassID, int RegResID>
	inline STDMETHODIMP addin<AppT, ClassID, RegResID>::OnDisconnection(msaddin::ext_DisconnectMode /*removeMode*/, SAFEARRAY ** /*custom*/)
	{
		_application.reset();
		return S_OK;
	}

	template <class AppT, const CLSID *ClassID, int RegResID>
	inline STDMETHODIMP addin<AppT, ClassID, RegResID>::OnAddInsUpdate(SAFEARRAY ** /*custom*/)
	{
		return E_NOTIMPL;
	}

	template <class AppT, const CLSID *ClassID, int RegResID>
	inline STDMETHODIMP addin<AppT, ClassID, RegResID>::OnStartupComplete(SAFEARRAY ** /*custom*/)
	{
		return E_NOTIMPL;
	}

	template <class AppT, const CLSID *ClassID, int RegResID>
	inline STDMETHODIMP addin<AppT, ClassID, RegResID>::OnBeginShutdown(SAFEARRAY ** /*custom*/)
	{
		return E_NOTIMPL;
	}
}