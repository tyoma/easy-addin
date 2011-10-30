#include <easy-addin/addin.h>

#include "DTEMocks.h"

#include <atlbase.h>
#include <atlcom.h>

using namespace System;
using namespace Microsoft::VisualStudio::TestTools::UnitTesting;

class CEasyAddinTestsModule : public CAtlDllModuleT<CEasyAddinTestsModule> {	} _AtlModule;

namespace ea
{
	namespace tests
	{
		namespace
		{
			struct __declspec(uuid("8B9A13D1-1FCC-4DA6-BCC9-5882204833DC")) addin1_novscmdt
			{
				addin1_novscmdt(msaddin::_DTEPtr dte)
				{	++instance_count;	}

				~addin1_novscmdt()
				{	--instance_count;	}

				static int instance_count;
			};

			struct __declspec(uuid("B1F62666-3E61-4E61-A946-0E18D9D46C03")) addin2_novscmdt
			{
				addin2_novscmdt(msaddin::_DTEPtr dte)
				{	++instance_count;	}

				~addin2_novscmdt()
				{	--instance_count;	}

				static int instance_count;
			};

			struct __declspec(uuid("C5B0AE02-698E-4B11-AF14-462520BDA453")) addin_throwing_novscmdt
			{
				addin_throwing_novscmdt(msaddin::_DTEPtr dte)
				{	throw 0;	}
			};

			struct __declspec(uuid("26BCC010-DC3A-4973-9C78-9F2EA102A2A8")) addin_with_param_storing
			{
				addin_with_param_storing(msaddin::_DTEPtr dte)
				{	this->dte = dte;	}

				~addin_with_param_storing()
				{	dte = 0;	}

				static msaddin::_DTEPtr dte;
			};

			int addin1_novscmdt::instance_count = 0;
			int addin2_novscmdt::instance_count = 0;
			msaddin::_DTEPtr addin_with_param_storing::dte;

			typedef addin<addin1_novscmdt, &__uuidof(addin1_novscmdt), 123> Addin1Impl;
			typedef addin<addin2_novscmdt, &__uuidof(addin2_novscmdt), 234> Addin2Impl;
			typedef addin<addin_throwing_novscmdt, &__uuidof(addin_throwing_novscmdt), 234> AddinThrowingImpl;
			typedef addin<addin_with_param_storing, &__uuidof(addin_with_param_storing), 234> AddinParamStoringImpl;
		}

		[TestClass]
		public ref class EasyAddinTests
		{
		public:
			[TestMethod]
			void CreateAddinInstance()
			{
				// INIT
				CComPtr<IUnknown> a1;
				CComPtr<IDispatch> a2, a3;
				
				// ACT
				Addin1Impl::CreateInstance(&a1);
				Addin2Impl::CreateInstance(&a2);
				Addin2Impl::CreateInstance(&a3);

				// ASSERT
				Assert::IsTrue(0 != a1);
				Assert::IsTrue(0 != a2);
				Assert::IsTrue(0 != a3);
				Assert::IsTrue(a2 != a3);
			}


			[TestMethod]
			void CreatingAddinInstanceDoesNotCreateApplicationInstance()
			{
				// INIT
				CComPtr<IUnknown> a1, a2;
				
				// ACT
				Addin1Impl::CreateInstance(&a1);
				Addin2Impl::CreateInstance(&a2);
				
				// ASSERT
				Assert::IsTrue(0 == addin1_novscmdt::instance_count);
				Assert::IsTrue(0 == addin2_novscmdt::instance_count);
			}


			[TestMethod]
			void ApplicationInstanceIsCreatedAtOnConnection()
			{
				// INIT
				CComPtr<msaddin::IDTExtensibility2> a1, a2, a3;
				
				// ACT
				Addin1Impl::CreateInstance(&a1);
				HRESULT hr1 = a1->OnConnection(0, msaddin::ext_cm_AfterStartup, 0, 0);
				
				// ASSERT
				Assert::IsTrue(1 == addin1_novscmdt::instance_count);
				Assert::IsTrue(0 == addin2_novscmdt::instance_count);
				Assert::IsTrue(S_OK == hr1);
				
				// ACT
				Addin2Impl::CreateInstance(&a2);
				Addin2Impl::CreateInstance(&a3);
				HRESULT hr2 = a2->OnConnection(0, msaddin::ext_cm_Startup, 0, 0);
				HRESULT hr3 = a3->OnConnection(0, msaddin::ext_cm_External, 0, 0);
				
				// ASSERT
				Assert::IsTrue(1 == addin1_novscmdt::instance_count);
				Assert::IsTrue(2 == addin2_novscmdt::instance_count);
				Assert::IsTrue(S_OK == hr2);
				Assert::IsTrue(S_OK == hr3);
			}


			[TestMethod]
			void ReturnEFailOnExceptionFromConstruction()
			{
				// INIT
				CComPtr<msaddin::IDTExtensibility2> a;
				
				// ACT
				AddinThrowingImpl::CreateInstance(&a);
				HRESULT hr = a->OnConnection(0, msaddin::ext_cm_AfterStartup, 0, 0);
				
				// ASSERT
				Assert::IsTrue(E_FAIL == hr);
			}


			[TestMethod]
			void ApplicationInstancesAreDestroyedAtAddinFinalRelease()
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
				Assert::IsTrue(0 == addin1_novscmdt::instance_count);
				Assert::IsTrue(2 == addin2_novscmdt::instance_count);

				// ACT
				a2.Release();
				
				// ASSERT
				Assert::IsTrue(0 == addin1_novscmdt::instance_count);
				Assert::IsTrue(1 == addin2_novscmdt::instance_count);

				// ACT
				a3.Release();
				
				// ASSERT
				Assert::IsTrue(0 == addin1_novscmdt::instance_count);
				Assert::IsTrue(0 == addin2_novscmdt::instance_count);
			}


			[TestMethod]
			void ApplicationInstancesAreDestroyedAtDisconnection()
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
				Assert::IsTrue(0 == addin1_novscmdt::instance_count);
				Assert::IsTrue(2 == addin2_novscmdt::instance_count);

				// ACT
				a2->OnDisconnection(msaddin::ext_dm_HostShutdown, 0);
				
				// ASSERT
				Assert::IsTrue(0 == addin1_novscmdt::instance_count);
				Assert::IsTrue(1 == addin2_novscmdt::instance_count);

				// ACT
				a3->OnDisconnection(msaddin::ext_dm_UserClosed, 0);
				
				// ASSERT
				Assert::IsTrue(0 == addin1_novscmdt::instance_count);
				Assert::IsTrue(0 == addin2_novscmdt::instance_count);
			}


			[TestMethod]
			void DisconnectionSucceeds()
			{
				// INIT
				CComPtr<msaddin::IDTExtensibility2> a;

				Addin1Impl::CreateInstance(&a);
				a->OnConnection(0, msaddin::ext_cm_AfterStartup, 0, 0);

				// ACT / ASSERT
				Assert::IsTrue(S_OK == a->OnDisconnection(msaddin::ext_dm_HostShutdown, 0));
			}


			[TestMethod]
			void DTEIsPassedToAppCtor()
			{
				// INIT
				CComPtr<msaddin::IDTExtensibility2> a;
				msaddin::_DTEPtr dte1(new DTEMock), dte2(new DTEMock);

				AddinParamStoringImpl::CreateInstance(&a);

				// ACT
				a->OnConnection(dte1, msaddin::ext_cm_AfterStartup, 0, 0);

				// ASSERT
				Assert::IsTrue(dte1 == addin_with_param_storing::dte);

				// INIT
				a->OnDisconnection(msaddin::ext_dm_UserClosed, 0);

				// ACT
				a->OnConnection(dte2, msaddin::ext_cm_AfterStartup, 0, 0);

				// ASSERT
				Assert::IsTrue(dte2 == addin_with_param_storing::dte);
			}


			[TestMethod]
			void DTEIsReleasedIfNotNecessary()
			{
				// INIT
				bool released = false;
				CComPtr<msaddin::IDTExtensibility2> a;
				msaddin::_DTEPtr dte(new DTEMock(released));

				Addin1Impl::CreateInstance(&a);

				// ACT
				a->OnConnection(dte, msaddin::ext_cm_AfterStartup, 0, 0);

				// ASSERT
				Assert::IsFalse(released);

				// ACT
				dte = 0;

				// ASSERT
				Assert::IsTrue(released);
			}
		};
	}
}
