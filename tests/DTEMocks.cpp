#include "DTEMocks.h"

using namespace Microsoft::VisualStudio::TestTools::UnitTesting;

namespace ea
{
	namespace tests
	{
		DTEMock::DTEMock()
			: command_bars(new mock_com_object<IDispatch>())
		{	}

		DTEMock::DTEMock(bool &released_flag)
			: base_type(released_flag), command_bars(new mock_com_object<IDispatch>())
		{	}

		STDMETHODIMP DTEMock::get_CommandBars(IDispatch **ppcbs)
		{	return command_bars->QueryInterface(ppcbs);	}

		STDMETHODIMP DTEMock::get_Commands(EnvDTE::Commands **ppCommands)
		{	return static_cast<IUnknown *>((new CommandsMock(commands_list)))->QueryInterface(ppCommands);	}


		CommandsMock::CommandsMock(std::vector<DTEMock::command> &commands_list)
			: _commands_list(&commands_list)
		{	}

		
		STDMETHODIMP CommandsMock::raw_AddNamedCommand(EnvDTE::AddIn *AddInInstance, BSTR Name, BSTR ButtonText, BSTR Tooltip, VARIANT_BOOL MSOButton, long /*Bitmap*/, SAFEARRAY **ContextUIGUIDs, long vsCommandDisabledFlagsValue, EnvDTE::Command **pVal)
		{
			Assert::IsTrue(::SysStringLen(Name) == wcslen(Name));
			Assert::IsTrue(::SysStringLen(ButtonText) == wcslen(ButtonText));
			Assert::IsTrue(::SysStringLen(Tooltip) == wcslen(Tooltip));
			Assert::IsTrue(VARIANT_TRUE == MSOButton);
			Assert::IsTrue(0 == ContextUIGUIDs);
			Assert::IsTrue(16 == vsCommandDisabledFlagsValue);
			
			EnvDTE::CommandPtr created_command(new CommandMock);

			DTEMock::command c = {	AddInInstance, Name, ButtonText, Tooltip, created_command	};

			_commands_list->push_back(c);
			created_command->QueryInterface(pVal);

			return S_OK;
		}


		CommandMock::CommandMock()
		{	}


		AddInMock::AddInMock()
		{	}

		AddInMock::AddInMock(bool &released_flag)
			: base_type(released_flag)
		{	}

		STDMETHODIMP AddInMock::get_ProgID(BSTR *lpbstr)
		{
			if (!lpbstr)
				return E_POINTER;
			*lpbstr = _bstr_t(prog_id.c_str()).copy();
			return S_OK;
		}
	}
}
