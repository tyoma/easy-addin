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
#include <vector>

#pragma warning(disable: 4278)
	#import <msaddndr.dll> rename_namespace("msaddin") raw_interfaces_only
	#import <dte80a.olb> no_implementation
#pragma warning(default: 4278)

namespace std
{
	using tr1::shared_ptr;
}

namespace ea
{
	struct command;
	struct command_target;
	typedef std::shared_ptr<command> command_ptr;

	inline _bstr_t str2bstr(const std::wstring &value)
	{	return _bstr_t(value.c_str());	}

	template <class AppT, const CLSID *ClassID, int RegResID>
	class ATL_NO_VTABLE addin
		: public CComObjectRootEx<CComSingleThreadModel>,
			public CComCoClass<addin<AppT, ClassID, RegResID>, ClassID>,
			public IDispatchImpl<msaddin::IDTExtensibility2, &__uuidof(msaddin::IDTExtensibility2), &__uuidof(msaddin::__AddInDesignerObjects), 1, 0>,
			public IDispatchImpl<EnvDTE::IDTCommandTarget, &__uuidof(EnvDTE::IDTCommandTarget), &__uuidof(EnvDTE::__EnvDTE), 7, 0>
	{
		EnvDTE::_DTEPtr _dte;
		std::wstring _regid;
		std::auto_ptr<AppT> _application;
		std::vector<command_ptr> _commands;

		// IDTExtensibility2 methods
		STDMETHODIMP OnConnection(IDispatch *host, msaddin::ext_ConnectMode connectMode, IDispatch *instance, SAFEARRAY **custom);
		STDMETHODIMP OnDisconnection(msaddin::ext_DisconnectMode removeMode, SAFEARRAY **custom);
		STDMETHODIMP OnAddInsUpdate(SAFEARRAY **custom);
		STDMETHODIMP OnStartupComplete(SAFEARRAY **custom);
		STDMETHODIMP OnBeginShutdown(SAFEARRAY **custom);

		// IDTCommandTarget methods
		STDMETHODIMP raw_QueryStatus(BSTR id, EnvDTE::vsCommandStatusTextWanted needed_text, EnvDTE::vsCommandStatus *status, VARIANT *text);
		STDMETHODIMP raw_Exec(BSTR id, EnvDTE::vsCommandExecOption execute_option, VARIANT *input, VARIANT *output, VARIANT_BOOL *handled);

		void query_commands(void *application);
		void query_commands(command_target *application);

		void setup_ui(EnvDTE::_DTEPtr dte, EnvDTE::AddInPtr addin_instance);

		command_ptr find_command(std::wstring id) const;

	public:
		DECLARE_REGISTRY_RESOURCEID(RegResID)
		DECLARE_NOT_AGGREGATABLE(addin)

		BEGIN_COM_MAP(addin)
			COM_INTERFACE_ENTRY2(IDispatch, msaddin::IDTExtensibility2)
			COM_INTERFACE_ENTRY(msaddin::IDTExtensibility2)
			COM_INTERFACE_ENTRY(EnvDTE::IDTCommandTarget)
		END_COM_MAP()
	};

	struct command_target
	{
		virtual ~command_target() {	}

		virtual void get_commands(std::vector<command_ptr> &commands) const = 0;
	};

	struct command
	{
		virtual ~command() {	}
		
		virtual std::wstring id() const = 0;
		virtual std::wstring caption() const = 0;
		virtual std::wstring description() const = 0;
		virtual void update_ui(EnvDTE::CommandPtr cmd, IDispatchPtr command_bars) const = 0;
		virtual bool query_status(EnvDTE::_DTEPtr dte, bool &checked, std::wstring *caption, std::wstring *description) const = 0;
		virtual void execute(EnvDTE::_DTEPtr dte, VARIANT *input, VARIANT *output) const = 0;
	};

	
	template <class AppT, const CLSID *ClassID, int RegResID>
	inline STDMETHODIMP addin<AppT, ClassID, RegResID>::OnConnection(IDispatch *host, msaddin::ext_ConnectMode connectMode, IDispatch *instance, SAFEARRAY ** /*custom*/)
	try
	{
		DISPPARAMS dispparams = { 0 };
		_variant_t vregid;
		unsigned int arg_error;

		_dte = host;
		if (instance)
			if (S_OK == instance->Invoke(3 /*get_RegID*/, IID_NULL, 0, DISPATCH_PROPERTYGET, &dispparams, &vregid, NULL, &arg_error))
				_regid = _bstr_t(vregid);
			else
				return E_UNEXPECTED;
		_application.reset(new AppT(IDispatchPtr(host, true)));
		query_commands(_application.get());
		if (5 /*ext_cm_UISetup*/ == connectMode)
			setup_ui(_dte, instance);
		return S_OK;
	}
	catch (...)
	{
		return E_FAIL;
	}

	template <class AppT, const CLSID *ClassID, int RegResID>
	inline STDMETHODIMP addin<AppT, ClassID, RegResID>::OnDisconnection(msaddin::ext_DisconnectMode /*removeMode*/, SAFEARRAY ** /*custom*/)
	{
		_dte = 0;
		_commands.clear();
		_application.reset();
		return S_OK;
	}

	template <class AppT, const CLSID *ClassID, int RegResID>
	inline STDMETHODIMP addin<AppT, ClassID, RegResID>::OnAddInsUpdate(SAFEARRAY ** /*custom*/)
	{	return E_NOTIMPL;	}

	template <class AppT, const CLSID *ClassID, int RegResID>
	inline STDMETHODIMP addin<AppT, ClassID, RegResID>::OnStartupComplete(SAFEARRAY ** /*custom*/)
	{	return E_NOTIMPL;	}

	template <class AppT, const CLSID *ClassID, int RegResID>
	inline STDMETHODIMP addin<AppT, ClassID, RegResID>::OnBeginShutdown(SAFEARRAY ** /*custom*/)
	{	return E_NOTIMPL;	}

	template <class AppT, const CLSID *ClassID, int RegResID>
	inline STDMETHODIMP addin<AppT, ClassID, RegResID>::raw_QueryStatus(BSTR id, EnvDTE::vsCommandStatusTextWanted /*needed_text*/, EnvDTE::vsCommandStatus *status, VARIANT * /*text*/)
	{
		if (id == NULL)
			return E_INVALIDARG;
		if (command_ptr c = find_command(id))
		{
			bool checked;
			bool enabled = c->query_status(_dte, checked, 0, 0);

			*status = static_cast<EnvDTE::vsCommandStatus>(EnvDTE::vsCommandStatusSupported + (enabled ? EnvDTE::vsCommandStatusEnabled : 0) + (checked ? EnvDTE::vsCommandStatusLatched : 0));
			return S_OK;
		}
		return E_UNEXPECTED;
	}

	template <class AppT, const CLSID *ClassID, int RegResID>
	inline STDMETHODIMP addin<AppT, ClassID, RegResID>::raw_Exec(BSTR id, EnvDTE::vsCommandExecOption /*execute_option*/, VARIANT *input, VARIANT *output, VARIANT_BOOL *handled)
	{
		if (id == NULL)
			return E_INVALIDARG;
		if (command_ptr c = find_command(id))
		{
			c->execute(_dte, input, output);
			*handled = VARIANT_TRUE;
			return S_OK;
		}
		return E_UNEXPECTED;
	}

	template <class AppT, const CLSID *ClassID, int RegResID>
	inline void addin<AppT, ClassID, RegResID>::query_commands(void * /*application*/)
	{	}

	template <class AppT, const CLSID *ClassID, int RegResID>
	inline void addin<AppT, ClassID, RegResID>::query_commands(command_target *application)
	{
		_commands.clear();
		application->get_commands(_commands);
	}
	
	template <class AppT, const CLSID *ClassID, int RegResID>
	inline void addin<AppT, ClassID, RegResID>::setup_ui(EnvDTE::_DTEPtr dte, EnvDTE::AddInPtr addin_instance)
	{
		using namespace std;

		IDispatchPtr command_bars;
		EnvDTE::CommandsPtr dte_commands;
		vector<command_ptr> commands;

		dte->get_CommandBars(&command_bars);
		dte->get_Commands(&dte_commands);
		for (vector<command_ptr>::const_iterator i = _commands.begin(); i != _commands.end(); ++i)
		{
			EnvDTE::CommandPtr dte_command;

			dte_commands->raw_AddNamedCommand(addin_instance, str2bstr((*i)->id()), str2bstr((*i)->caption()), str2bstr((*i)->description()), VARIANT_TRUE, 0, NULL, 16, &dte_command);
			(*i)->update_ui(dte_command, command_bars);
		}
	}

	template <class AppT, const CLSID *ClassID, int RegResID>
	inline command_ptr addin<AppT, ClassID, RegResID>::find_command(std::wstring id) const
	{
		if (id.find(_regid + L".") == 0)
		{
			id.erase(0, _regid.size() + 1);
			for (vector<command_ptr>::const_iterator i = _commands.begin(); i != _commands.end(); ++i)
				if (id == (*i)->id())
					return *i;
		}
		return command_ptr();
	}
}
