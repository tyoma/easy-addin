#pragma once

#include <string>
#include <vector>

#pragma warning(disable: 4278)
#pragma warning(disable: 4146)
	#import <dte80a.olb>
#pragma warning(default: 4146)
#pragma warning(default: 4278)

namespace ea
{
	namespace tests
	{
		class checked_dtor
		{
			bool *_released_flag;

		public:
			checked_dtor(bool &released_flag)
				: _released_flag(&released_flag)
			{	}

			checked_dtor()
				: _released_flag(0)
			{	}

			virtual ~checked_dtor()
			{
				if (_released_flag)
					*_released_flag = true;
			}
		};

		template <typename InterfaceT>
		class mock_com_object : public InterfaceT, public checked_dtor
		{
			int _references;

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
				: _references(0)
			{	}

			mock_com_object(bool &released_flag)
				: checked_dtor(released_flag), _references(0)
			{	}
		};


		class DTEMock : public mock_com_object<EnvDTE::_DTE>
		{
			STDMETHODIMP get_Name(BSTR * /*lpbstrReturn*/)	{	return E_NOTIMPL;	}
			STDMETHODIMP get_FileName(BSTR * /*lpbstrReturn*/)	{	return E_NOTIMPL;	}
			STDMETHODIMP get_Version(BSTR * /*lpbstrReturn*/)	{	return E_NOTIMPL;	}
			STDMETHODIMP get_CommandBars(IDispatch **ppcbs);
			STDMETHODIMP get_Windows(EnvDTE::Windows ** /*ppwnsVBWindows*/)	{	return E_NOTIMPL;	}
			STDMETHODIMP get_Events(EnvDTE::Events ** /*ppevtEvents*/)	{	return E_NOTIMPL;	}
			STDMETHODIMP get_AddIns(EnvDTE::AddIns ** /*lpppAddIns*/)	{	return E_NOTIMPL;	}
			STDMETHODIMP get_MainWindow(EnvDTE::Window ** /*ppWin*/)	{	return E_NOTIMPL;	}
			STDMETHODIMP get_ActiveWindow(EnvDTE::Window ** /*ppwinActive*/)	{	return E_NOTIMPL;	}
			STDMETHODIMP raw_Quit()	{	return E_NOTIMPL;	}
			STDMETHODIMP get_DisplayMode(EnvDTE::vsDisplay * /*lpDispMode*/)	{	return E_NOTIMPL;	}
			STDMETHODIMP put_DisplayMode(EnvDTE::vsDisplay /*lpDispMode*/)	{	return E_NOTIMPL;	}
			STDMETHODIMP get_Solution(EnvDTE::_Solution ** /*ppSolution*/)	{	return E_NOTIMPL;	}
			STDMETHODIMP get_Commands(EnvDTE::Commands **ppCommands);
			STDMETHODIMP raw_GetObject(BSTR /*Name*/, IDispatch ** /*ppObject*/)	{	return E_NOTIMPL;	}
			STDMETHODIMP get_Properties(BSTR /*Category*/, BSTR /*Page*/, EnvDTE::Properties ** /*ppObject*/)	{	return E_NOTIMPL;	}
			STDMETHODIMP get_SelectedItems(EnvDTE::SelectedItems ** /*ppSelectedItems*/)	{	return E_NOTIMPL;	}
			STDMETHODIMP get_CommandLineArguments(BSTR * /*lpbstrReturn*/)	{	return E_NOTIMPL;	}
			STDMETHODIMP raw_OpenFile(BSTR /*ViewKind*/, BSTR /*FileName*/, EnvDTE::Window ** /*ppWin*/)	{	return E_NOTIMPL;	}
			STDMETHODIMP get_IsOpenFile(BSTR /*ViewKind*/, BSTR /*FileName*/, VARIANT_BOOL * /*lpfReturn*/)	{	return E_NOTIMPL;	}
			STDMETHODIMP get_DTE(EnvDTE::_DTE ** /*lppaReturn*/)	{	return E_NOTIMPL;	}
			STDMETHODIMP get_LocaleID(long * /*lpReturn*/)	{	return E_NOTIMPL;	}
			STDMETHODIMP get_WindowConfigurations(EnvDTE::WindowConfigurations ** /*WindowConfigurationsObject*/)	{	return E_NOTIMPL;	}
			STDMETHODIMP get_Documents(EnvDTE::Documents ** /*ppDocuments*/)	{	return E_NOTIMPL;	}
			STDMETHODIMP get_ActiveDocument(EnvDTE::Document ** /*ppDocument*/)	{	return E_NOTIMPL;	}
			STDMETHODIMP raw_ExecuteCommand(BSTR /*CommandName*/, BSTR /*CommandArgs*/)	{	return E_NOTIMPL;	}
			STDMETHODIMP get_Globals(EnvDTE::Globals ** /*ppGlobals*/)	{	return E_NOTIMPL;	}
			STDMETHODIMP get_StatusBar(EnvDTE::StatusBar ** /*ppStatusBar*/)	{	return E_NOTIMPL;	}
			STDMETHODIMP get_FullName(BSTR * /*lpbstrReturn*/)	{	return E_NOTIMPL;	}
			STDMETHODIMP get_UserControl(VARIANT_BOOL * /*UserControl*/)	{	return E_NOTIMPL;	}
			STDMETHODIMP put_UserControl(VARIANT_BOOL /*UserControl*/)	{	return E_NOTIMPL;	}
			STDMETHODIMP get_ObjectExtenders(EnvDTE::ObjectExtenders ** /*ppObjectExtenders*/)	{	return E_NOTIMPL;	}
			STDMETHODIMP get_Find(EnvDTE::Find ** /*ppFind*/)	{	return E_NOTIMPL;	}
			STDMETHODIMP get_Mode(EnvDTE::vsIDEMode * /*pVal*/)	{	return E_NOTIMPL;	}
			STDMETHODIMP raw_LaunchWizard(BSTR /*VSZFile*/, SAFEARRAY ** /*ContextParams*/, EnvDTE::wizardResult * /*pResult*/)	{	return E_NOTIMPL;	}
			STDMETHODIMP get_ItemOperations(EnvDTE::ItemOperations ** /*ppItemOperations*/)	{	return E_NOTIMPL;	}
			STDMETHODIMP get_UndoContext(EnvDTE::UndoContext ** /*ppUndoContext*/)	{	return E_NOTIMPL;	}
			STDMETHODIMP get_Macros(EnvDTE::Macros ** /*ppMacros*/)	{	return E_NOTIMPL;	}
			STDMETHODIMP get_ActiveSolutionProjects(VARIANT * /*pProjects*/)	{	return E_NOTIMPL;	}
			STDMETHODIMP get_MacrosIDE(EnvDTE::_DTE ** /*pDTE*/)	{	return E_NOTIMPL;	}
			STDMETHODIMP get_RegistryRoot(BSTR * /*pVal*/)	{	return E_NOTIMPL;	}
			STDMETHODIMP get_Application(EnvDTE::_DTE ** /*pVal*/)	{	return E_NOTIMPL;	}
			STDMETHODIMP get_ContextAttributes(EnvDTE::ContextAttributes ** /*ppVal*/)	{	return E_NOTIMPL;	}
			STDMETHODIMP get_SourceControl(EnvDTE::SourceControl ** /*ppVal*/)	{	return E_NOTIMPL;	}
			STDMETHODIMP get_SuppressUI(VARIANT_BOOL * /*pVal*/)	{	return E_NOTIMPL;	}
			STDMETHODIMP put_SuppressUI(VARIANT_BOOL /*pVal*/)	{	return E_NOTIMPL;	}
			STDMETHODIMP get_Debugger(EnvDTE::Debugger ** /*ppDebugger*/)	{	return E_NOTIMPL;	}
			STDMETHODIMP raw_SatelliteDllPath(BSTR /*Path*/, BSTR /*Name*/, BSTR * /*pFullPath*/)	{	return E_NOTIMPL;	}
			STDMETHODIMP get_Edition(BSTR * /*ProductEdition*/)	{	return E_NOTIMPL;	}

		public:
			DTEMock();
			DTEMock(bool &released_flag);

			struct command
			{
				EnvDTE::AddInPtr addin_instance;
				std::wstring id, caption, description;
				EnvDTE::CommandPtr created_command;
			};

			std::vector<command> commands_list;
			IDispatchPtr command_bars;
		};


		class CommandsMock : public mock_com_object<EnvDTE::Commands>
		{
			std::vector<DTEMock::command> *_commands_list;

			STDMETHODIMP get_DTE(EnvDTE::_DTE ** /*lppaReturn*/)	{	return E_NOTIMPL;	}
			STDMETHODIMP get_Parent(EnvDTE::_DTE ** /*lppaReturn*/)	{	return E_NOTIMPL;	}
			STDMETHODIMP raw_Add(BSTR /*Guid*/, long /*ID*/, VARIANT * /*Control*/)	{	return E_NOTIMPL;	}
			STDMETHODIMP raw_Raise(BSTR /*Guid*/, long /*ID*/, VARIANT * /*CustomIn*/, VARIANT * /*CustomOut*/)	{	return E_NOTIMPL;	}
			STDMETHODIMP raw_CommandInfo(IDispatch * /*CommandBarControl*/, BSTR * /*Guid*/, long * /*ID*/)	{	return E_NOTIMPL;	}
			STDMETHODIMP get_Count(long * /*lplReturn*/)	{	return E_NOTIMPL;	}
			STDMETHODIMP raw_Item(VARIANT /*index*/, long /*ID*/, EnvDTE::Command ** /*lppcReturn*/)	{	return E_NOTIMPL;	}
			STDMETHODIMP raw__NewEnum(IUnknown ** /*lppiuReturn*/)	{	return E_NOTIMPL;	}
			STDMETHODIMP raw_AddNamedCommand(EnvDTE::AddIn *AddInInstance, BSTR Name, BSTR ButtonText, BSTR Tooltip, VARIANT_BOOL MSOButton, long Bitmap, SAFEARRAY **ContextUIGUIDs, long vsCommandDisabledFlagsValue, EnvDTE::Command **pVal);
			STDMETHODIMP raw_AddCommandBar(BSTR /*Name*/, EnvDTE::vsCommandBarType /*Type*/, IDispatch * /*CommandBarParent*/, long /*Position*/, IDispatch ** /*pVal*/)	{	return E_NOTIMPL;	}
			STDMETHODIMP raw_RemoveCommandBar(IDispatch * /*CommandBar*/)	{	return E_NOTIMPL;	}

		public:
			CommandsMock(std::vector<DTEMock::command> &commands_list);
		};


		class CommandMock : public mock_com_object<EnvDTE::Command>
		{
			STDMETHODIMP get_Name(BSTR * /*lpbstr*/)	{	return E_NOTIMPL;	}
			STDMETHODIMP get_Collection(EnvDTE::Commands ** /*lppcReturn*/)	{	return E_NOTIMPL;	}
			STDMETHODIMP get_DTE(EnvDTE::_DTE ** /*lppaReturn*/)	{	return E_NOTIMPL;	}
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


		class AddInMock : public mock_com_object<EnvDTE::AddIn>
		{
			STDMETHODIMP Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr);

			STDMETHODIMP get_Description(BSTR * /*lpbstr*/)	{	return E_NOTIMPL;	}
			STDMETHODIMP put_Description(BSTR /*lpbstr*/)	{	return E_NOTIMPL;	}
			STDMETHODIMP get_Collection(EnvDTE::AddIns ** /*lppaddins*/)	{	return E_NOTIMPL;	}
			STDMETHODIMP get_ProgID(BSTR * /*lpbstr*/)	{	return E_NOTIMPL;	}
			STDMETHODIMP get_Guid(BSTR * /*lpbstr*/)	{	return E_NOTIMPL;	}
			STDMETHODIMP get_Connected(VARIANT_BOOL * /*lpfConnect*/)	{	return E_NOTIMPL;	}
			STDMETHODIMP put_Connected(VARIANT_BOOL /*lpfConnect*/)	{	return E_NOTIMPL;	}
			STDMETHODIMP get_Object(IDispatch ** /*lppdisp*/)	{	return E_NOTIMPL;	}
			STDMETHODIMP put_Object(IDispatch * /*lppdisp*/)	{	return E_NOTIMPL;	}
			STDMETHODIMP get_DTE(EnvDTE::_DTE ** /*lppaReturn*/)	{	return E_NOTIMPL;	}
			STDMETHODIMP get_Name(BSTR * /*lpbstr*/)	{	return E_NOTIMPL;	}
			STDMETHODIMP raw_Remove()	{	return E_NOTIMPL;	}
			STDMETHODIMP get_SatelliteDllPath(BSTR * /*pbstrPath*/)	{	return E_NOTIMPL;	}

		public:
			AddInMock();
			AddInMock(bool &released_flag);

			std::wstring prog_id;
		};
	}
}
