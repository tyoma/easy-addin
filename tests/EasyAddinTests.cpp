#include <easy-addin/addin.h>

#include "DTEMocks.h"

#include <atlbase.h>
#include <atlcom.h>
#include <ut/assert.h>
#include <ut/test.h>

using namespace std;

class CEasyAddinTestsModule : public CAtlDllModuleT<CEasyAddinTestsModule> {	} _AtlModule;

namespace ea
{
	namespace tests
	{
		namespace
		{
			struct __declspec(uuid("8B9A13D1-1FCC-4DA6-BCC9-5882204833DC")) addin1_novscmdt
			{
				addin1_novscmdt(EnvDTE::_DTEPtr dte)
				{	++instance_count;	}

				~addin1_novscmdt()
				{	--instance_count;	}

				static int instance_count;
			};

			struct __declspec(uuid("B1F62666-3E61-4E61-A946-0E18D9D46C03")) addin2_novscmdt
			{
				addin2_novscmdt(EnvDTE::_DTEPtr dte)
				{	++instance_count;	}

				~addin2_novscmdt()
				{	--instance_count;	}

				static int instance_count;
			};

			struct __declspec(uuid("C5B0AE02-698E-4B11-AF14-462520BDA453")) addin_throwing_novscmdt
			{
				addin_throwing_novscmdt(EnvDTE::_DTEPtr dte)
				{	throw 0;	}
			};

			struct __declspec(uuid("26BCC010-DC3A-4973-9C78-9F2EA102A2A8")) addin_with_param_storing
			{
				addin_with_param_storing(EnvDTE::_DTEPtr dte)
				{	this->dte = dte;	}

				~addin_with_param_storing()
				{	dte = 0;	}

				static EnvDTE::_DTEPtr dte;
			};

			struct __declspec(uuid("9D26A098-715F-454A-8FD1-7D73ECEE3001")) addin_with_commandlist : command_target
			{
				addin_with_commandlist(EnvDTE::_DTEPtr /*dte*/)
				{	}

				void get_commands(vector< shared_ptr<command> > &commands) const
				{
					// Append instead of assign - commands must be clean
					commands.insert(commands.end(), addin_with_commandlist::commands.begin(), addin_with_commandlist::commands.end());
				}

				static vector< shared_ptr<command> > commands;
			};

			class mock_command_impl : public command, public checked_dtor
			{
				wstring _id, _caption, _description;

			public:
				mock_command_impl(const wstring &id, const wstring &caption, const wstring &description)
					: _id(id), _caption(caption), _description(description), throw_on_command(false)
				{	}

				mock_command_impl(const wstring &id, const wstring &caption, const wstring &description, bool &release_flag)
					: checked_dtor(release_flag), _id(id), _caption(caption), _description(description), throw_on_command(false)
				{	}

				virtual wstring id() const
				{	return _id;	}

				virtual wstring caption() const
				{	return _caption;	}

				virtual wstring description() const
				{	return _description;	}

				mutable vector< pair<EnvDTE::CommandPtr /*cmd*/, IDispatchPtr /*command_bars*/> > update_ui_log;

				virtual void update_ui(EnvDTE::CommandPtr cmd, IDispatchPtr command_bars) const
				{	update_ui_log.push_back(make_pair(cmd, command_bars));	}

				bool throw_on_command;

				bool command_enabled, command_checked;
				mutable vector<EnvDTE::_DTEPtr /*dte*/> query_status_log;

				virtual bool query_status(EnvDTE::_DTEPtr dte, bool &checked, wstring * /*caption*/, wstring * /*description*/) const
				{
					if (throw_on_command)
						throw 0;

					query_status_log.push_back(dte);
					checked = command_checked;
					return command_enabled;
				}

				typedef pair< EnvDTE::_DTEPtr /*dte*/, pair<VARIANT * /*input*/, VARIANT * /*output*/> > execute_log_entry;
				mutable vector<execute_log_entry> execute_log;

				virtual void execute(EnvDTE::_DTEPtr dte, VARIANT *input, VARIANT *output) const
				{
					if (throw_on_command)
						throw 0;

					execute_log.push_back(make_pair(dte, make_pair(input, output)));
				}
			};

			int addin1_novscmdt::instance_count = 0;
			int addin2_novscmdt::instance_count = 0;
			EnvDTE::_DTEPtr addin_with_param_storing::dte;
			vector< shared_ptr<command> > addin_with_commandlist::commands;

			typedef addin<addin1_novscmdt, &__uuidof(addin1_novscmdt), 123> Addin1Impl;
			typedef addin<addin2_novscmdt, &__uuidof(addin2_novscmdt), 234> Addin2Impl;
			typedef addin<addin_throwing_novscmdt, &__uuidof(addin_throwing_novscmdt), 234> AddinThrowingImpl;
			typedef addin<addin_with_param_storing, &__uuidof(addin_with_param_storing), 234> AddinParamStoringImpl;
			typedef addin<addin_with_commandlist, &__uuidof(addin_with_commandlist), 234> AddinWithCommandListImpl;
		}

		begin_test_suite( EasyAddinTests )
			init( Init )
			{
				addin1_novscmdt::instance_count = 0;
				addin2_novscmdt::instance_count = 0;
				addin_with_param_storing::dte = 0;
				addin_with_commandlist::commands.clear();
			}

			test( CreateAddinInstance )
			{
				// INIT
				CComPtr<IUnknown> a1;
				CComPtr<IDispatch> a2, a3;
				
				// ACT
				Addin1Impl::CreateInstance(&a1);
				Addin2Impl::CreateInstance(&a2);
				Addin2Impl::CreateInstance(&a3);

				// ASSERT
				assert_not_null(a1);
				assert_not_null(a2);
				assert_not_null(a3);
				assert_not_equal(a2, a3);
			}


			test( CreatingAddinInstanceDoesNotCreateApplicationInstance )
			{
				// INIT
				CComPtr<IUnknown> a1, a2;
				
				// ACT
				Addin1Impl::CreateInstance(&a1);
				Addin2Impl::CreateInstance(&a2);
				
				// ASSERT
				assert_equal(0, addin1_novscmdt::instance_count);
				assert_equal(0, addin2_novscmdt::instance_count);
			}


			test( ApplicationInstanceIsCreatedAtOnConnection )
			{
				// INIT
				CComPtr<msaddin::IDTExtensibility2> a1, a2, a3;
				
				// ACT
				Addin1Impl::CreateInstance(&a1);
				HRESULT hr1 = a1->OnConnection(0, msaddin::ext_cm_AfterStartup, 0, 0);
				
				// ASSERT
				assert_equal(1, addin1_novscmdt::instance_count);
				assert_equal(0, addin2_novscmdt::instance_count);
				assert_equal(S_OK, hr1);
				
				// ACT
				Addin2Impl::CreateInstance(&a2);
				Addin2Impl::CreateInstance(&a3);
				HRESULT hr2 = a2->OnConnection(0, msaddin::ext_cm_Startup, 0, 0);
				HRESULT hr3 = a3->OnConnection(0, msaddin::ext_cm_External, 0, 0);
				
				// ASSERT
				assert_equal(1, addin1_novscmdt::instance_count);
				assert_equal(2, addin2_novscmdt::instance_count);
				assert_equal(S_OK, hr2);
				assert_equal(S_OK, hr3);
			}


			test( ReturnEFailOnExceptionFromConstruction )
			{
				// INIT
				CComPtr<msaddin::IDTExtensibility2> a;
				
				// ACT
				AddinThrowingImpl::CreateInstance(&a);
				HRESULT hr = a->OnConnection(0, msaddin::ext_cm_AfterStartup, 0, 0);
				
				// ASSERT
				assert_equal(E_FAIL, hr);
			}


			test( ApplicationInstancesAreDestroyedAtAddinFinalRelease )
			{
				// INIT
				CComPtr<msaddin::IDTExtensibility2> a1, a2, a3;

				Addin1Impl::CreateInstance(&a1);
				Addin2Impl::CreateInstance(&a2);
				Addin2Impl::CreateInstance(&a3);
				a1->OnConnection(0, msaddin::ext_cm_AfterStartup, 0, 0);
				a2->OnConnection(0, msaddin::ext_cm_Startup, 0, 0);
				a3->OnConnection(0, msaddin::ext_cm_External, 0, 0);

				// ACT
				a1.Release();
				
				// ASSERT
				assert_equal(0, addin1_novscmdt::instance_count);
				assert_equal(2, addin2_novscmdt::instance_count);

				// ACT
				a2.Release();
				
				// ASSERT
				assert_equal(0, addin1_novscmdt::instance_count);
				assert_equal(1, addin2_novscmdt::instance_count);

				// ACT
				a3.Release();
				
				// ASSERT
				assert_equal(0, addin1_novscmdt::instance_count);
				assert_equal(0, addin2_novscmdt::instance_count);
			}


			test( ApplicationInstancesAreDestroyedAtDisconnection )
			{
				// INIT
				CComPtr<msaddin::IDTExtensibility2> a1, a2, a3;

				Addin1Impl::CreateInstance(&a1);
				Addin2Impl::CreateInstance(&a2);
				Addin2Impl::CreateInstance(&a3);
				a1->OnConnection(0, msaddin::ext_cm_AfterStartup, 0, 0);
				a2->OnConnection(0, msaddin::ext_cm_Startup, 0, 0);
				a3->OnConnection(0, msaddin::ext_cm_External, 0, 0);

				// ACT
				a1->OnDisconnection(msaddin::ext_dm_HostShutdown, 0);
				
				// ASSERT
				assert_equal(0, addin1_novscmdt::instance_count);
				assert_equal(2, addin2_novscmdt::instance_count);

				// ACT
				a2->OnDisconnection(msaddin::ext_dm_HostShutdown, 0);
				
				// ASSERT
				assert_equal(0, addin1_novscmdt::instance_count);
				assert_equal(1, addin2_novscmdt::instance_count);

				// ACT
				a3->OnDisconnection(msaddin::ext_dm_UserClosed, 0);
				
				// ASSERT
				assert_equal(0, addin1_novscmdt::instance_count);
				assert_equal(0, addin2_novscmdt::instance_count);
			}


			test( CommandInstancesAreDestroyedAtDisconnectionAndFinalRelease )
			{
				// INIT
				CComPtr<msaddin::IDTExtensibility2> a1, a2;
				CComPtr<DTEMock> dte(new DTEMock);
				CComPtr<AddInMock> addin_instance(new AddInMock);
				EnvDTE::IDTCommandTargetPtr cmd_target;
				bool c1_released = false, c2_released = false, c3_released = false;

				AddinWithCommandListImpl::CreateInstance(&a1);
				addin_with_commandlist::commands.clear();
				AddinWithCommandListImpl::CreateInstance(&a2);

				addin_instance->prog_id = L"TestAddIn.Connect";

				addin_with_commandlist::commands.push_back(shared_ptr<mock_command_impl>(new mock_command_impl(L"run", L"", L"", c1_released)));
				addin_with_commandlist::commands.push_back(shared_ptr<mock_command_impl>(new mock_command_impl(L"exit", L"", L"", c2_released)));
				a1->OnConnection(dte, msaddin::ext_cm_Startup, addin_instance, 0);
				addin_with_commandlist::commands.clear();
				addin_with_commandlist::commands.push_back(shared_ptr<mock_command_impl>(new mock_command_impl(L"stop", L"", L"", c3_released)));
				a2->OnConnection(dte, msaddin::ext_cm_Startup, addin_instance, 0);
				addin_with_commandlist::commands.clear();

				// ACT
				a1->OnDisconnection(msaddin::ext_dm_HostShutdown, 0);
				a2.Release();

				// ASSERT
				assert_is_true(c1_released);
				assert_is_true(c2_released);
				assert_is_true(c3_released);
			}


			test( CommandInstancesAreDestroyedAtNewConnection )
			{
				// INIT
				CComPtr<msaddin::IDTExtensibility2> a;
				CComPtr<DTEMock> dte(new DTEMock);
				CComPtr<AddInMock> addin_instance(new AddInMock);
				EnvDTE::IDTCommandTargetPtr cmd_target;
				bool c1_released = false, c2_released = false;

				AddinWithCommandListImpl::CreateInstance(&a);

				addin_instance->prog_id = L"TestAddIn.Connect";

				addin_with_commandlist::commands.push_back(shared_ptr<mock_command_impl>(new mock_command_impl(L"run", L"", L"", c1_released)));
				addin_with_commandlist::commands.push_back(shared_ptr<mock_command_impl>(new mock_command_impl(L"exit", L"", L"", c2_released)));
				a->OnConnection(dte, msaddin::ext_cm_Startup, addin_instance, 0);
				addin_with_commandlist::commands.clear();

				// ASSERT
				assert_is_false(c1_released);
				assert_is_false(c2_released);

				// ACT
				a->OnConnection(dte, msaddin::ext_cm_Startup, addin_instance, 0);

				// ASSERT
				assert_is_true(c1_released);
				assert_is_true(c2_released);
			}


			test( DisconnectionSucceeds )
			{
				// INIT
				CComPtr<msaddin::IDTExtensibility2> a;

				Addin1Impl::CreateInstance(&a);
				a->OnConnection(0, msaddin::ext_cm_AfterStartup, 0, 0);

				// ACT / ASSERT
				assert_equal(S_OK, a->OnDisconnection(msaddin::ext_dm_HostShutdown, 0));
			}


			test( DTEIsPassedToAppCtor )
			{
				// INIT
				CComPtr<msaddin::IDTExtensibility2> a;
				EnvDTE::_DTEPtr dte1(new DTEMock), dte2(new DTEMock);

				AddinParamStoringImpl::CreateInstance(&a);

				// ACT
				a->OnConnection(dte1, msaddin::ext_cm_AfterStartup, 0, 0);

				// ASSERT
				assert_equal(dte1, addin_with_param_storing::dte);

				// INIT
				a->OnDisconnection(msaddin::ext_dm_UserClosed, 0);

				// ACT
				a->OnConnection(dte2, msaddin::ext_cm_AfterStartup, 0, 0);

				// ASSERT
				assert_equal(dte2, addin_with_param_storing::dte);
			}


			test( DTEIsReleasedIfNotNecessary )
			{
				// INIT
				bool released = false;
				CComPtr<msaddin::IDTExtensibility2> a;
				EnvDTE::_DTEPtr dte(new DTEMock(released));

				Addin1Impl::CreateInstance(&a);

				// ACT
				a->OnConnection(dte, msaddin::ext_cm_AfterStartup, 0, 0);
				dte = 0;

				// ASSERT
				assert_is_false(released);

				// ACT
				a->OnDisconnection(msaddin::ext_dm_HostShutdown, 0);

				// ASSERT
				assert_is_true(released);
			}


			test( DoNotCreateCommandsForEmptyAddinCommandsList )
			{
				// INIT
				CComPtr<msaddin::IDTExtensibility2> a;
				CComPtr<DTEMock> dte(new DTEMock);

				AddinWithCommandListImpl::CreateInstance(&a);

				// ACT
				a->OnConnection(dte, (msaddin::ext_ConnectMode)5, 0, 0);

				// ASSERT
				assert_is_true(dte->commands_list.empty());
			}


			test( CommandsAreNotCreatedOnNonUISetupConnectMode )
			{
				// INIT
				CComPtr<msaddin::IDTExtensibility2> a;
				CComPtr<DTEMock> dte(new DTEMock);
				EnvDTE::AddInPtr addin_instance(new AddInMock);

				addin_with_commandlist::commands.push_back(shared_ptr<command>(new mock_command_impl(
					L"a", L"A {C8DE681B-F9CB-46FF-96E3-44C8DE97964E}", L"{D4FBBBDD-7730-4E77-BF8D-197920A39E0A}")));

				AddinWithCommandListImpl::CreateInstance(&a);

				// ACT
				a->OnConnection(dte, msaddin::ext_cm_AfterStartup, (EnvDTE::AddIn *)addin_instance, 0);
				a->OnConnection(dte, msaddin::ext_cm_Startup, (EnvDTE::AddIn *)addin_instance, 0);
				a->OnConnection(dte, msaddin::ext_cm_External, (EnvDTE::AddIn *)addin_instance, 0);
				a->OnConnection(dte, msaddin::ext_cm_CommandLine, (EnvDTE::AddIn *)addin_instance, 0);

				// ASSERT
				assert_is_true(dte->commands_list.empty());
			}


			test( CommandsAreCreatedOnUISetup1 )
			{
				// INIT
				CComPtr<msaddin::IDTExtensibility2> a;
				CComPtr<DTEMock> dte(new DTEMock);
				EnvDTE::AddInPtr addin_instance(new AddInMock);

				addin_with_commandlist::commands.push_back(shared_ptr<command>(new mock_command_impl(
					L"a", L"A {C8DE681B-F9CB-46FF-96E3-44C8DE97964E}", L"{D4FBBBDD-7730-4E77-BF8D-197920A39E0A}")));
				addin_with_commandlist::commands.push_back(shared_ptr<command>(new mock_command_impl(
					L"b", L"B {9633565D-55D0-4A17-8DAF-15604DDB3491}", L"{7FA3A73A-4D80-407B-8CA4-0C7904526197} {D4FBBBDD-7730-4E77-BF8D-197920A39E0A}")));
				addin_with_commandlist::commands.push_back(shared_ptr<command>(new mock_command_impl(
					L"c", L"C {7B718B04-91E7-4EA3-97AE-DDEE9F6E9817}", L"BF8D 197920A39E0A")));

				AddinWithCommandListImpl::CreateInstance(&a);

				// ACT / ASSERT
				assert_equal(S_OK, a->OnConnection(dte, (msaddin::ext_ConnectMode)5, (EnvDTE::AddIn *)addin_instance, 0));

				// ASSERT
				assert_equal(3u, dte->commands_list.size());

				assert_equal(IUnknownPtr(addin_instance), IUnknownPtr(dte->commands_list[0].addin_instance));
				assert_equal(L"a", dte->commands_list[0].id);
				assert_equal(L"A {C8DE681B-F9CB-46FF-96E3-44C8DE97964E}", dte->commands_list[0].caption);
				assert_equal(L"{D4FBBBDD-7730-4E77-BF8D-197920A39E0A}", dte->commands_list[0].description);
				
				assert_equal(IUnknownPtr(addin_instance), IUnknownPtr(dte->commands_list[1].addin_instance));
				assert_equal(L"b", dte->commands_list[1].id);
				assert_equal(L"B {9633565D-55D0-4A17-8DAF-15604DDB3491}", dte->commands_list[1].caption);
				assert_equal(L"{7FA3A73A-4D80-407B-8CA4-0C7904526197} {D4FBBBDD-7730-4E77-BF8D-197920A39E0A}", dte->commands_list[1].description);

				assert_equal(IUnknownPtr(addin_instance), dte->commands_list[2].addin_instance);
				assert_equal(L"c", dte->commands_list[2].id);
				assert_equal(L"C {7B718B04-91E7-4EA3-97AE-DDEE9F6E9817}", dte->commands_list[2].caption);
				assert_equal(L"BF8D 197920A39E0A", dte->commands_list[2].description);
			}


			test( CommandsAreCreatedOnUISetup2 )
			{
				// INIT
				bool addin_instance_released = false;
				EnvDTE::AddInPtr addin_instance(new AddInMock(addin_instance_released));
				CComPtr<msaddin::IDTExtensibility2> a;
				CComPtr<DTEMock> dte(new DTEMock);

				addin_with_commandlist::commands.push_back(shared_ptr<command>(new mock_command_impl(
					L"run", L"Run 96E3", L"D4FBBBDD-7730-4E77-BF8D")));
				addin_with_commandlist::commands.push_back(shared_ptr<command>(new mock_command_impl(
					L"stop", L"Stop DDEE9F6E9817", L"BF8D")));

				AddinWithCommandListImpl::CreateInstance(&a);

				// ACT
				assert_equal(S_OK, a->OnConnection(dte, (msaddin::ext_ConnectMode)5, (EnvDTE::AddIn *)addin_instance, 0));

				// ASSERT
				assert_equal(2u, dte->commands_list.size());

				assert_equal(IUnknownPtr(addin_instance), IUnknownPtr(dte->commands_list[0].addin_instance));
				assert_equal(L"run", dte->commands_list[0].id);
				assert_equal(L"Run 96E3", dte->commands_list[0].caption);
				assert_equal(L"D4FBBBDD-7730-4E77-BF8D", dte->commands_list[0].description);

				assert_equal(IUnknownPtr(addin_instance), IUnknownPtr(dte->commands_list[1].addin_instance));
				assert_equal(L"stop", dte->commands_list[1].id);
				assert_equal(L"Stop DDEE9F6E9817", dte->commands_list[1].caption);
				assert_equal(L"BF8D", dte->commands_list[1].description);

				assert_is_false(addin_instance_released);

				// ACT
				addin_instance = 0;
				dte->commands_list.clear();

				// ASSERT
				assert_is_true(addin_instance_released);
			}


			test( UICreationIsRequestedForEachCommandCreated )
			{
				// INIT
				bool addin_instance_released = false;
				EnvDTE::AddInPtr addin_instance(new AddInMock(addin_instance_released));
				CComPtr<msaddin::IDTExtensibility2> a;
				CComPtr<DTEMock> dte1(new DTEMock), dte2(new DTEMock);
				shared_ptr<mock_command_impl> c1(new mock_command_impl(L"run", L"Run 96E3", L"D4FBBBDD-7730-4E77-BF8D"));
				shared_ptr<mock_command_impl> c2(new mock_command_impl(L"stop", L"Stop DDEE9F6E9817", L"BF8D"));

				addin_with_commandlist::commands.push_back(c1);
				addin_with_commandlist::commands.push_back(c2);

				AddinWithCommandListImpl::CreateInstance(&a);

				// ACT
				a->OnConnection(dte1, (msaddin::ext_ConnectMode)5, (EnvDTE::AddIn *)addin_instance, 0);

				// ASSERT
				assert_equal(1u, c1->update_ui_log.size());
				assert_equal(1u, c2->update_ui_log.size());

				assert_is_true(dte1->commands_list[0].created_command);
				assert_equal(dte1->commands_list[0].created_command, c1->update_ui_log[0].first);
				assert_equal(dte1->command_bars, c1->update_ui_log[0].second);

				assert_equal(dte1->commands_list[1].created_command, c2->update_ui_log[0].first);
				assert_equal(dte1->command_bars, c2->update_ui_log[0].second);

				// ACT
				a->OnConnection(dte2, (msaddin::ext_ConnectMode)5, (EnvDTE::AddIn *)addin_instance, 0);

				// ASSERT
				assert_equal(2u, c1->update_ui_log.size());
				assert_equal(2u, c2->update_ui_log.size());

				assert_equal(dte2->commands_list[0].created_command, c1->update_ui_log[1].first);
				assert_equal(dte2->command_bars, c1->update_ui_log[1].second);

				assert_equal(dte2->commands_list[1].created_command, c2->update_ui_log[1].first);
				assert_equal(dte2->command_bars, c2->update_ui_log[1].second);
			}


			test( AddInImplImplementsCmdTarget )
			{
				// INIT
				CComPtr<msaddin::IDTExtensibility2> a;

				Addin1Impl::CreateInstance(&a);

				// ACT
				EnvDTE::IDTCommandTargetPtr cmd_target((IDispatch *)a);

				// ASSERT
				assert_is_true(cmd_target);
			}


			test( QueryingStatusOrExecFailsIfNotConnected )
			{
				// INIT
				EnvDTE::IDTCommandTargetPtr cmd_target;
				shared_ptr<mock_command_impl> c(new mock_command_impl(L"run", L"Run 96E3", L"D4FBBBDD-7730-4E77-BF8D"));
				EnvDTE::vsCommandStatus status;
				VARIANT_BOOL handled;

				addin_with_commandlist::commands.push_back(c);
				AddinWithCommandListImpl::CreateInstance(&cmd_target);

				// ACT / ASSERT
				assert_equal(E_UNEXPECTED, cmd_target->raw_QueryStatus(_bstr_t(L"testaddin.connect.run"), EnvDTE::vsCommandStatusTextWantedNone, &status, &vtMissing));
				assert_equal(E_UNEXPECTED, cmd_target->raw_Exec(_bstr_t(L"testaddin.connect.run"), EnvDTE::vsCommandExecOptionDoDefault, NULL, NULL, &handled));
			}


			test( QueryingStatusOrExecFailsWhenPrefixOrCommandAreInvalid )
			{
				// INIT
				CComPtr<msaddin::IDTExtensibility2> a;
				CComPtr<DTEMock> dte(new DTEMock);
				CComPtr<AddInMock> addin_instance(new AddInMock);
				EnvDTE::IDTCommandTargetPtr cmd_target;
				shared_ptr<mock_command_impl> c(new mock_command_impl(L"run", L"Run 96E3", L"D4FBBBDD-7730-4E77-BF8D"));
				EnvDTE::vsCommandStatus status;
				VARIANT_BOOL handled;

				addin_with_commandlist::commands.push_back(c);
				AddinWithCommandListImpl::CreateInstance(&a);
				cmd_target = (IDispatch *)a;

				addin_instance->prog_id = L"TestAddIn.Connect";

				a->OnConnection(dte, msaddin::ext_cm_Startup, addin_instance, 0);

				// ACT / ASSERT
				assert_equal(E_UNEXPECTED, cmd_target->raw_QueryStatus(_bstr_t(L"TestAddIn.connect.run"), EnvDTE::vsCommandStatusTextWantedNone, &status, &vtMissing));
				assert_equal(E_UNEXPECTED, cmd_target->raw_Exec(_bstr_t(L"TestAddIn.connect.run"), EnvDTE::vsCommandExecOptionDoDefault, NULL, NULL, &handled));
				assert_equal(E_UNEXPECTED, cmd_target->raw_QueryStatus(_bstr_t(L"TestAddIn2.Connect.run"), EnvDTE::vsCommandStatusTextWantedNone, &status, &vtMissing));
				assert_equal(E_UNEXPECTED, cmd_target->raw_Exec(_bstr_t(L"TestAddIn2.Connect.run"), EnvDTE::vsCommandExecOptionDoDefault, NULL, NULL, &handled));
				assert_equal(E_UNEXPECTED, cmd_target->raw_QueryStatus(_bstr_t(L"TestAddIn.connect.stop"), EnvDTE::vsCommandStatusTextWantedNone, &status, &vtMissing));
				assert_equal(E_UNEXPECTED, cmd_target->raw_Exec(_bstr_t(L"TestAddIn.Connect.stop"), EnvDTE::vsCommandExecOptionDoDefault, NULL, NULL, &handled));
				assert_equal(E_INVALIDARG, cmd_target->raw_QueryStatus(NULL, EnvDTE::vsCommandStatusTextWantedNone, &status, &vtMissing));
				assert_equal(E_INVALIDARG, cmd_target->raw_Exec(NULL, EnvDTE::vsCommandExecOptionDoDefault, NULL, NULL, &handled));

				// INIT
				addin_instance->prog_id = L"TestAddIn2.Connect";

				a->OnConnection(dte, msaddin::ext_cm_Startup, addin_instance, 0);

				// ACT / ASSERT
				assert_equal(E_UNEXPECTED, cmd_target->raw_QueryStatus(_bstr_t(L"TestAddIn2.connect.run"), EnvDTE::vsCommandStatusTextWantedNone, &status, &vtMissing));
				assert_equal(E_UNEXPECTED, cmd_target->raw_Exec(_bstr_t(L"TestAddIn2.connect.run"), EnvDTE::vsCommandExecOptionDoDefault, NULL, NULL, &handled));
				assert_equal(E_UNEXPECTED, cmd_target->raw_QueryStatus(_bstr_t(L"TestAddIn.Connect.run"), EnvDTE::vsCommandStatusTextWantedNone, &status, &vtMissing));
				assert_equal(E_UNEXPECTED, cmd_target->raw_Exec(_bstr_t(L"TestAddIn.Connect.run"), EnvDTE::vsCommandExecOptionDoDefault, NULL, NULL, &handled));
				assert_equal(E_UNEXPECTED, cmd_target->raw_QueryStatus(_bstr_t(L"TestAddIn2.Connect.stop"), EnvDTE::vsCommandStatusTextWantedNone, &status, &vtMissing));
				assert_equal(E_UNEXPECTED, cmd_target->raw_Exec(_bstr_t(L"TestAddIn2.Connect.stop"), EnvDTE::vsCommandExecOptionDoDefault, NULL, NULL, &handled));
				assert_equal(E_UNEXPECTED, cmd_target->raw_QueryStatus(_bstr_t(L"zTestAddIn2.Connect.run"), EnvDTE::vsCommandStatusTextWantedNone, &status, &vtMissing));
				assert_equal(E_UNEXPECTED, cmd_target->raw_Exec(_bstr_t(L"zTestAddIn2.Connect.run"), EnvDTE::vsCommandExecOptionDoDefault, NULL, NULL, &handled));
			}


			test( QueryingStatusOrExecSucceedsIfPrefixIsValid )
			{
				// INIT
				CComPtr<msaddin::IDTExtensibility2> a;
				CComPtr<DTEMock> dte(new DTEMock);
				CComPtr<AddInMock> addin_instance(new AddInMock);
				EnvDTE::IDTCommandTargetPtr cmd_target;
				shared_ptr<mock_command_impl> c1(new mock_command_impl(L"run", L"", L""));
				shared_ptr<mock_command_impl> c2(new mock_command_impl(L"exit", L"", L""));
				EnvDTE::vsCommandStatus status;
				VARIANT_BOOL handled;

				addin_with_commandlist::commands.push_back(c1);
				addin_with_commandlist::commands.push_back(c2);
				AddinWithCommandListImpl::CreateInstance(&a);
				cmd_target = (IDispatch *)a;

				addin_instance->prog_id = L"TestAddIn.Connect";

				a->OnConnection(dte, msaddin::ext_cm_Startup, addin_instance, 0);

				// ACT / ASSERT
				assert_equal(S_OK, cmd_target->raw_QueryStatus(_bstr_t(L"TestAddIn.Connect.run"), EnvDTE::vsCommandStatusTextWantedNone, &status, &vtMissing));
				assert_equal(S_OK, cmd_target->raw_Exec(_bstr_t(L"TestAddIn.Connect.run"), EnvDTE::vsCommandExecOptionDoDefault, NULL, NULL, &handled));
				assert_equal(S_OK, cmd_target->raw_QueryStatus(_bstr_t(L"TestAddIn.Connect.exit"), EnvDTE::vsCommandStatusTextWantedNone, &status, &vtMissing));
				assert_equal(S_OK, cmd_target->raw_Exec(_bstr_t(L"TestAddIn.Connect.exit"), EnvDTE::vsCommandExecOptionDoDefault, NULL, NULL, &handled));

				// INIT
				addin_instance->prog_id = L"TestAddIn2.Connect";

				a->OnConnection(dte, msaddin::ext_cm_Startup, addin_instance, 0);

				// ACT / ASSERT
				assert_equal(S_OK, cmd_target->raw_QueryStatus(_bstr_t(L"TestAddIn2.Connect.run"), EnvDTE::vsCommandStatusTextWantedNone, &status, &vtMissing));
				assert_equal(S_OK, cmd_target->raw_Exec(_bstr_t(L"TestAddIn2.Connect.run"), EnvDTE::vsCommandExecOptionDoDefault, NULL, NULL, &handled));
				assert_equal(S_OK, cmd_target->raw_QueryStatus(_bstr_t(L"TestAddIn2.Connect.exit"), EnvDTE::vsCommandStatusTextWantedNone, &status, &vtMissing));
				assert_equal(S_OK, cmd_target->raw_Exec(_bstr_t(L"TestAddIn2.Connect.exit"), EnvDTE::vsCommandExecOptionDoDefault, NULL, NULL, &handled));
			}


			test( ExecIsPRoperlyForwardedToCommand )
			{
				// INIT
				CComPtr<msaddin::IDTExtensibility2> a;
				CComPtr<DTEMock> dte(new DTEMock);
				CComPtr<AddInMock> addin_instance(new AddInMock);
				EnvDTE::IDTCommandTargetPtr cmd_target;
				shared_ptr<mock_command_impl> c1(new mock_command_impl(L"run", L"", L""));
				shared_ptr<mock_command_impl> c2(new mock_command_impl(L"exit", L"", L""));
				shared_ptr<mock_command_impl> c3(new mock_command_impl(L"stop", L"", L""));
				_variant_t dummies[10];
				VARIANT_BOOL handled = VARIANT_FALSE;

				addin_with_commandlist::commands.push_back(c1);
				addin_with_commandlist::commands.push_back(c2);
				addin_with_commandlist::commands.push_back(c3);
				AddinWithCommandListImpl::CreateInstance(&a);
				cmd_target = (IDispatch *)a;

				addin_instance->prog_id = L"MyAddIn.Connect";

				a->OnConnection(dte, msaddin::ext_cm_Startup, addin_instance, 0);

				// ACT
				cmd_target->raw_Exec(_bstr_t(L"MyAddIn.Connect.run"), EnvDTE::vsCommandExecOptionDoDefault, &dummies[0], &dummies[1], &handled);

				// ASSERT
				assert_equal(VARIANT_TRUE, handled);
				assert_equal(1u, c1->execute_log.size());
				assert_equal(0u, c2->execute_log.size());
				assert_equal(0u, c3->execute_log.size());
				assert_equal(dte, c1->execute_log[0].first);
				assert_equal(&dummies[0], c1->execute_log[0].second.first);
				assert_equal(&dummies[1], c1->execute_log[0].second.second);

				// ACT
				cmd_target->raw_Exec(_bstr_t(L"MyAddIn.Connect.run"), EnvDTE::vsCommandExecOptionDoDefault, &dummies[2], &dummies[3], &handled);
				cmd_target->raw_Exec(_bstr_t(L"MyAddIn.Connect.exit"), EnvDTE::vsCommandExecOptionDoDefault, &dummies[4], &dummies[5], &handled);

				// ASSERT
				assert_equal(2u, c1->execute_log.size());
				assert_equal(1u, c2->execute_log.size());
				assert_equal(0u, c3->execute_log.size());
				assert_equal(dte, c1->execute_log[1].first);
				assert_equal(&dummies[2], c1->execute_log[1].second.first);
				assert_equal(&dummies[3], c1->execute_log[1].second.second);
				assert_equal(dte, c2->execute_log[0].first);
				assert_equal(&dummies[4], c2->execute_log[0].second.first);
				assert_equal(&dummies[5], c2->execute_log[0].second.second);
			}


			test( QueryStatusIsPRoperlyForwardedToCommand )
			{
				// INIT
				CComPtr<msaddin::IDTExtensibility2> a;
				CComPtr<DTEMock> dte(new DTEMock);
				CComPtr<AddInMock> addin_instance(new AddInMock);
				EnvDTE::IDTCommandTargetPtr cmd_target;
				shared_ptr<mock_command_impl> c1(new mock_command_impl(L"run", L"", L""));
				shared_ptr<mock_command_impl> c2(new mock_command_impl(L"exit", L"", L""));
				shared_ptr<mock_command_impl> c3(new mock_command_impl(L"stop", L"", L""));
				shared_ptr<mock_command_impl> c4(new mock_command_impl(L"continue", L"", L""));
				EnvDTE::vsCommandStatus status[4];

				AddinWithCommandListImpl::CreateInstance(&a);
				cmd_target = (IDispatch *)a;

				addin_instance->prog_id = L"MyAddIn.Connect";

				addin_with_commandlist::commands.push_back(c1);
				addin_with_commandlist::commands.push_back(c2);
				addin_with_commandlist::commands.push_back(c3);
				addin_with_commandlist::commands.push_back(c4);
				a->OnConnection(dte, msaddin::ext_cm_Startup, addin_instance, 0);
				c1->command_enabled = true, c1->command_checked = true;
				c2->command_enabled = true, c2->command_checked = false;
				c3->command_enabled = false, c3->command_checked = true;
				c4->command_enabled = false, c4->command_checked = false;

				// ACT
				cmd_target->raw_QueryStatus(_bstr_t(L"MyAddIn.Connect.run"), EnvDTE::vsCommandStatusTextWantedNone, &status[0], &vtMissing);
				cmd_target->raw_QueryStatus(_bstr_t(L"MyAddIn.Connect.exit"), EnvDTE::vsCommandStatusTextWantedNone, &status[1], &vtMissing);
				cmd_target->raw_QueryStatus(_bstr_t(L"MyAddIn.Connect.stop"), EnvDTE::vsCommandStatusTextWantedNone, &status[2], &vtMissing);
				cmd_target->raw_QueryStatus(_bstr_t(L"MyAddIn.Connect.continue"), EnvDTE::vsCommandStatusTextWantedNone, &status[3], &vtMissing);

				// ASSERT
				assert_equal((EnvDTE::vsCommandStatusSupported | EnvDTE::vsCommandStatusEnabled | EnvDTE::vsCommandStatusLatched), status[0]);
				assert_equal((EnvDTE::vsCommandStatusSupported | EnvDTE::vsCommandStatusEnabled), status[1]);
				assert_equal((EnvDTE::vsCommandStatusSupported | EnvDTE::vsCommandStatusLatched), status[2]);
				assert_equal((EnvDTE::vsCommandStatusSupported), status[3]);
				assert_equal(1u, c1->query_status_log.size());
				assert_equal(dte, c1->query_status_log[0]);
			}


			test( ExecAndQueryStatusTranslateCPPExceptionsToEFail )
			{
				// INIT
				CComPtr<msaddin::IDTExtensibility2> a;
				CComPtr<DTEMock> dte(new DTEMock);
				CComPtr<AddInMock> addin_instance(new AddInMock);
				EnvDTE::IDTCommandTargetPtr cmd_target;
				shared_ptr<mock_command_impl> c(new mock_command_impl(L"run", L"", L""));
				EnvDTE::vsCommandStatus status;
				VARIANT_BOOL handled;

				AddinWithCommandListImpl::CreateInstance(&a);
				cmd_target = (IDispatch *)a;

				addin_instance->prog_id = L"MyAddIn.Connect";

				addin_with_commandlist::commands.push_back(c);
				a->OnConnection(dte, msaddin::ext_cm_Startup, addin_instance, 0);
				c->throw_on_command = true;

				// ACT / ASSERT
				assert_equal(E_FAIL, cmd_target->raw_QueryStatus(_bstr_t(L"MyAddIn.Connect.run"), EnvDTE::vsCommandStatusTextWantedNone, &status, &vtMissing));
				assert_equal(E_FAIL, cmd_target->raw_Exec(_bstr_t(L"MyAddIn.Connect.run"), EnvDTE::vsCommandExecOptionDoDefault, 0, 0, &handled));
			}
		end_test_suite
	}
}
