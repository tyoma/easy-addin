#include "DTEMocks.h"

using namespace Microsoft::VisualStudio::TestTools::UnitTesting;

namespace ea
{
	namespace tests
	{
		DTEMock::DTEMock()
		{	}

		DTEMock::DTEMock(bool &released_flag)
			: base_type(released_flag)
		{	}

		STDMETHODIMP DTEMock::get_Commands(msaddin::Commands **ppCommands)
		{	return static_cast<IUnknown *>((new CommandsMock(commands_list)))->QueryInterface(ppCommands);	}


		CommandsMock::CommandsMock(std::vector<DTEMock::command> &commands_list)
			: _commands_list(&commands_list)
		{	}

		STDMETHODIMP CommandsMock::raw_AddNamedCommand(msaddin::AddIn *AddInInstance, BSTR Name, BSTR ButtonText, BSTR Tooltip, VARIANT_BOOL MSOButton, long Bitmap, SAFEARRAY **ContextUIGUIDs, long vsCommandDisabledFlagsValue, msaddin::Command **pVal)
		{
			Assert::IsTrue(::SysStringLen(Name) == wcslen(Name));
			Assert::IsTrue(::SysStringLen(ButtonText) == wcslen(ButtonText));
			Assert::IsTrue(::SysStringLen(Tooltip) == wcslen(Tooltip));
			Assert::IsTrue(VARIANT_TRUE == MSOButton);
			Assert::IsTrue(0 == ContextUIGUIDs);
			Assert::IsTrue(16 == vsCommandDisabledFlagsValue);
			
			DTEMock::command c = {	AddInInstance, Name, ButtonText, Tooltip	};

			_commands_list->push_back(c);

			return S_OK;
		}


		AddInMock::AddInMock()
		{	}

		AddInMock::AddInMock(bool &released_flag)
			: base_type(released_flag)
		{	}

	}
}
