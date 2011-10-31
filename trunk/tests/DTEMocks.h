#pragma once

#include <string>
#include <vector>

#pragma warning(disable: 4278)
#pragma warning(disable: 4146)
	#import <dte80a.olb> rename_namespace("msaddin")
#pragma warning(default: 4146)
#pragma warning(default: 4278)

namespace ea
{
	namespace tests
	{
		template <typename InterfaceT>
		class mock_com_object : public InterfaceT
		{
			int _references;
			bool *_released_flag;

		public:
			STDMETHODIMP QueryInterface(REFIID riid, void **ppvObject)
			{
				if(riid != IID_IUnknown && riid != IID_IDispatch && riid != __uuidof(InterfaceT))
					return E_NOINTERFACE;
				*ppvObject =(IUnknown *)(InterfaceT *)this;
				AddRef();
				return S_OK;
			}

			STDMETHODIMP_(ULONG) AddRef()
			{	return ++_references;	}

			STDMETHODIMP_(ULONG) Release()
			{
				if(!--_references)
				{
					delete this;
					return 0;
				}
				return _references;
			}

			STDMETHODIMP GetIDsOfNames(REFIID /*riid*/, LPOLESTR * /*rgszNames*/, UINT /*cNames*/, LCID /*lcid*/, DISPID * /*rgDispId*/)
			{	return E_NOTIMPL;	}

			STDMETHODIMP GetTypeInfo(UINT /*iTInfo*/, LCID /*lcid*/, ITypeInfo ** /*ppTInfo*/)
			{	return E_NOTIMPL;	}

			STDMETHODIMP GetTypeInfoCount(UINT * /*pctinfo*/)
			{	return E_NOTIMPL;	}

			STDMETHODIMP Invoke(DISPID /*dispIdMember*/, REFIID /*riid*/, LCID /*lcid*/, WORD /*wFlags*/, DISPPARAMS * /*pDispParams*/, VARIANT * /*pVarResult*/, EXCEPINFO * /*pExcepInfo*/, UINT * /*puArgErr*/)
			{	return E_NOTIMPL;	}

		protected:
			typedef mock_com_object base_type;

		public:
			mock_com_object()
				: _released_flag(0), _references(0)
			{	}

			mock_com_object(bool &released_flag)
				: _released_flag(&released_flag), _references(0)
			{	}

			virtual ~mock_com_object()
			{
				if(_released_flag)
					*_released_flag = true;
			}
		};


		class DTEMock : public mock_com_object<msaddin::_DTE>
		{
			STDMETHODIMP get_Name(BSTR * /*lpbstrReturn*/)	{	return E_NOTIMPL;	}
			STDMETHODIMP get_FileName(BSTR * /*lpbstrReturn*/)	{	return E_NOTIMPL;	}
			STDMETHODIMP get_Version(BSTR * /*lpbstrReturn*/)	{	return E_NOTIMPL;	}
			STDMETHODIMP get_CommandBars(IDispatch **ppcbs);
			STDMETHODIMP get_Windows(msaddin::Windows ** /*ppwnsVBWindows*/)	{	return E_NOTIMPL;	}
			STDMETHODIMP get_Events(msaddin::Events ** /*ppevtEvents*/)	{	return E_NOTIMPL;	}
			STDMETHODIMP get_AddIns(msaddin::AddIns ** /*lpppAddIns*/)	{	return E_NOTIMPL;	}
			STDMETHODIMP get_MainWindow(msaddin::Window ** /*ppWin*/)	{	return E_NOTIMPL;	}
			STDMETHODIMP get_ActiveWindow(msaddin::Window ** /*ppwinActive*/)	{	return E_NOTIMPL;	}
			STDMETHODIMP raw_Quit()	{	return E_NOTIMPL;	}
			STDMETHODIMP get_DisplayMode(msaddin::vsDisplay * /*lpDispMode*/)	{	return E_NOTIMPL;	}
			STDMETHODIMP put_DisplayMode(msaddin::vsDisplay /*lpDispMode*/)	{	return E_NOTIMPL;	}
			STDMETHODIMP get_Solution(msaddin::_Solution ** /*ppSolution*/)	{	return E_NOTIMPL;	}
			STDMETHODIMP get_Commands(msaddin::Commands **ppCommands);
			STDMETHODIMP raw_GetObject(BSTR /*Name*/, IDispatch ** /*ppObject*/)	{	return E_NOTIMPL;	}
			STDMETHODIMP get_Properties(BSTR /*Category*/, BSTR /*Page*/, msaddin::Properties ** /*ppObject*/)	{	return E_NOTIMPL;	}
			STDMETHODIMP get_SelectedItems(msaddin::SelectedItems ** /*ppSelectedItems*/)	{	return E_NOTIMPL;	}
			STDMETHODIMP get_CommandLineArguments(BSTR * /*lpbstrReturn*/)	{	return E_NOTIMPL;	}
			STDMETHODIMP raw_OpenFile(BSTR /*ViewKind*/, BSTR /*FileName*/, msaddin::Window ** /*ppWin*/)	{	return E_NOTIMPL;	}
			STDMETHODIMP get_IsOpenFile(BSTR /*ViewKind*/, BSTR /*FileName*/, VARIANT_BOOL * /*lpfReturn*/)	{	return E_NOTIMPL;	}
			STDMETHODIMP get_DTE(msaddin::_DTE ** /*lppaReturn*/)	{	return E_NOTIMPL;	}
			STDMETHODIMP get_LocaleID(long * /*lpReturn*/)	{	return E_NOTIMPL;	}
			STDMETHODIMP get_WindowConfigurations(msaddin::WindowConfigurations ** /*WindowConfigurationsObject*/)	{	return E_NOTIMPL;	}
			STDMETHODIMP get_Documents(msaddin::Documents ** /*ppDocuments*/)	{	return E_NOTIMPL;	}
			STDMETHODIMP get_ActiveDocument(msaddin::Document ** /*ppDocument*/)	{	return E_NOTIMPL;	}
			STDMETHODIMP raw_ExecuteCommand(BSTR /*CommandName*/, BSTR /*CommandArgs*/)	{	return E_NOTIMPL;	}
			STDMETHODIMP get_Globals(msaddin::Globals ** /*ppGlobals*/)	{	return E_NOTIMPL;	}
			STDMETHODIMP get_StatusBar(msaddin::StatusBar ** /*ppStatusBar*/)	{	return E_NOTIMPL;	}
			STDMETHODIMP get_FullName(BSTR * /*lpbstrReturn*/)	{	return E_NOTIMPL;	}
			STDMETHODIMP get_UserControl(VARIANT_BOOL * /*UserControl*/)	{	return E_NOTIMPL;	}
			STDMETHODIMP put_UserControl(VARIANT_BOOL /*UserControl*/)	{	return E_NOTIMPL;	}
			STDMETHODIMP get_ObjectExtenders(msaddin::ObjectExtenders ** /*ppObjectExtenders*/)	{	return E_NOTIMPL;	}
			STDMETHODIMP get_Find(msaddin::Find ** /*ppFind*/)	{	return E_NOTIMPL;	}
			STDMETHODIMP get_Mode(msaddin::vsIDEMode * /*pVal*/)	{	return E_NOTIMPL;	}
			STDMETHODIMP raw_LaunchWizard(BSTR /*VSZFile*/, SAFEARRAY ** /*ContextParams*/, msaddin::wizardResult * /*pResult*/)	{	return E_NOTIMPL;	}
			STDMETHODIMP get_ItemOperations(msaddin::ItemOperations ** /*ppItemOperations*/)	{	return E_NOTIMPL;	}
			STDMETHODIMP get_UndoContext(msaddin::UndoContext ** /*ppUndoContext*/)	{	return E_NOTIMPL;	}
			STDMETHODIMP get_Macros(msaddin::Macros ** /*ppMacros*/)	{	return E_NOTIMPL;	}
			STDMETHODIMP get_ActiveSolutionProjects(VARIANT * /*pProjects*/)	{	return E_NOTIMPL;	}
			STDMETHODIMP get_MacrosIDE(msaddin::_DTE ** /*pDTE*/)	{	return E_NOTIMPL;	}
			STDMETHODIMP get_RegistryRoot(BSTR * /*pVal*/)	{	return E_NOTIMPL;	}
			STDMETHODIMP get_Application(msaddin::_DTE ** /*pVal*/)	{	return E_NOTIMPL;	}
			STDMETHODIMP get_ContextAttributes(msaddin::ContextAttributes ** /*ppVal*/)	{	return E_NOTIMPL;	}
			STDMETHODIMP get_SourceControl(msaddin::SourceControl ** /*ppVal*/)	{	return E_NOTIMPL;	}
			STDMETHODIMP get_SuppressUI(VARIANT_BOOL * /*pVal*/)	{	return E_NOTIMPL;	}
			STDMETHODIMP put_SuppressUI(VARIANT_BOOL /*pVal*/)	{	return E_NOTIMPL;	}
			STDMETHODIMP get_Debugger(msaddin::Debugger ** /*ppDebugger*/)	{	return E_NOTIMPL;	}
			STDMETHODIMP raw_SatelliteDllPath(BSTR /*Path*/, BSTR /*Name*/, BSTR * /*pFullPath*/)	{	return E_NOTIMPL;	}
			STDMETHODIMP get_Edition(BSTR * /*ProductEdition*/)	{	return E_NOTIMPL;	}

		public:
			DTEMock();
			DTEMock(bool &released_flag);

			struct command
			{
				msaddin::AddInPtr addin_instance;
				std::wstring id, caption, description;
				msaddin::CommandPtr created_command;
			};

			std::vector<command> commands_list;
			IDispatchPtr command_bars;
		};


		class CommandsMock : public mock_com_object<msaddin::Commands>
		{
			std::vector<DTEMock::command> *_commands_list;

			STDMETHODIMP get_DTE(msaddin::_DTE ** /*lppaReturn*/)	{	return E_NOTIMPL;	}
			STDMETHODIMP get_Parent(msaddin::_DTE ** /*lppaReturn*/)	{	return E_NOTIMPL;	}
			STDMETHODIMP raw_Add(BSTR /*Guid*/, long /*ID*/, VARIANT * /*Control*/)	{	return E_NOTIMPL;	}
			STDMETHODIMP raw_Raise(BSTR /*Guid*/, long /*ID*/, VARIANT * /*CustomIn*/, VARIANT * /*CustomOut*/)	{	return E_NOTIMPL;	}
			STDMETHODIMP raw_CommandInfo(IDispatch * /*CommandBarControl*/, BSTR * /*Guid*/, long * /*ID*/)	{	return E_NOTIMPL;	}
			STDMETHODIMP get_Count(long * /*lplReturn*/)	{	return E_NOTIMPL;	}
			STDMETHODIMP raw_Item(VARIANT /*index*/, long /*ID*/, msaddin::Command ** /*lppcReturn*/)	{	return E_NOTIMPL;	}
			STDMETHODIMP raw__NewEnum(IUnknown ** /*lppiuReturn*/)	{	return E_NOTIMPL;	}
			STDMETHODIMP raw_AddNamedCommand(msaddin::AddIn *AddInInstance, BSTR Name, BSTR ButtonText, BSTR Tooltip, VARIANT_BOOL MSOButton, long Bitmap, SAFEARRAY **ContextUIGUIDs, long vsCommandDisabledFlagsValue, msaddin::Command **pVal);
			STDMETHODIMP raw_AddCommandBar(BSTR /*Name*/, msaddin::vsCommandBarType /*Type*/, IDispatch * /*CommandBarParent*/, long /*Position*/, IDispatch ** /*pVal*/)	{	return E_NOTIMPL;	}
			STDMETHODIMP raw_RemoveCommandBar(IDispatch * /*CommandBar*/)	{	return E_NOTIMPL;	}

		public:
			CommandsMock(std::vector<DTEMock::command> &commands_list);
		};


		class CommandMock : public mock_com_object<msaddin::Command>
		{
			STDMETHODIMP get_Name(BSTR * /*lpbstr*/)	{	return E_NOTIMPL;	}
			STDMETHODIMP get_Collection(msaddin::Commands ** /*lppcReturn*/)	{	return E_NOTIMPL;	}
			STDMETHODIMP get_DTE(msaddin::_DTE ** /*lppaReturn*/)	{	return E_NOTIMPL;	}
			STDMETHODIMP get_Guid(BSTR * /*lpbstr*/)	{	return E_NOTIMPL;	}
			STDMETHODIMP get_ID(long * /*lReturn*/)	{	return E_NOTIMPL;	}
			STDMETHODIMP get_IsAvailable(VARIANT_BOOL * /*pAvail*/)	{	return E_NOTIMPL;	}
			STDMETHODIMP raw_AddControl(IDispatch * /*Owner*/, long /*Position*/, IDispatch ** /*pCommandBarControl*/)	{	return E_NOTIMPL;	}
			STDMETHODIMP raw_Delete()	{	return E_NOTIMPL;	}
			STDMETHODIMP get_Bindings(VARIANT * /*pVar*/)	{	return E_NOTIMPL;	}
			STDMETHODIMP put_Bindings(VARIANT /*pVar*/)	{	return E_NOTIMPL;	}
			STDMETHODIMP get_LocalizedName(BSTR * /*lpbstr*/)	{	return E_NOTIMPL;	}

		public:
			CommandMock();
		};


		class AddInMock : public mock_com_object<msaddin::AddIn>
		{
			STDMETHODIMP get_Description(BSTR * /*lpbstr*/)	{	return E_NOTIMPL;	}
			STDMETHODIMP put_Description(BSTR /*lpbstr*/)	{	return E_NOTIMPL;	}
			STDMETHODIMP get_Collection(msaddin::AddIns ** /*lppaddins*/)	{	return E_NOTIMPL;	}
			STDMETHODIMP get_ProgID(BSTR * /*lpbstr*/)	{	return E_NOTIMPL;	}
			STDMETHODIMP get_Guid(BSTR * /*lpbstr*/)	{	return E_NOTIMPL;	}
			STDMETHODIMP get_Connected(VARIANT_BOOL * /*lpfConnect*/)	{	return E_NOTIMPL;	}
			STDMETHODIMP put_Connected(VARIANT_BOOL /*lpfConnect*/)	{	return E_NOTIMPL;	}
			STDMETHODIMP get_Object(IDispatch ** /*lppdisp*/)	{	return E_NOTIMPL;	}
			STDMETHODIMP put_Object(IDispatch * /*lppdisp*/)	{	return E_NOTIMPL;	}
			STDMETHODIMP get_DTE(msaddin::_DTE ** /*lppaReturn*/)	{	return E_NOTIMPL;	}
			STDMETHODIMP get_Name(BSTR * /*lpbstr*/)	{	return E_NOTIMPL;	}
			STDMETHODIMP raw_Remove()	{	return E_NOTIMPL;	}
			STDMETHODIMP get_SatelliteDllPath(BSTR * /*pbstrPath*/)	{	return E_NOTIMPL;	}

		public:
			AddInMock();
			AddInMock(bool &released_flag);
		};
	}
}
