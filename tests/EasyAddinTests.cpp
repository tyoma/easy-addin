#include <easy-addin/addin.h>

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
				addin1_novscmdt()
				{	++instance_count;	}

				~addin1_novscmdt()
				{	--instance_count;	}

				static int instance_count;
			};

			struct __declspec(uuid("B1F62666-3E61-4E61-A946-0E18D9D46C03")) addin2_novscmdt
			{
				addin2_novscmdt()
				{	++instance_count;	}

				~addin2_novscmdt()
				{	--instance_count;	}

				static int instance_count;
			};

			struct __declspec(uuid("C5B0AE02-698E-4B11-AF14-462520BDA453")) addin_throwing_novscmdt
			{
				addin_throwing_novscmdt()
				{	throw 0;	}
			};

			int addin1_novscmdt::instance_count = 0;
			int addin2_novscmdt::instance_count = 0;

			typedef addin<addin1_novscmdt, &__uuidof(addin1_novscmdt), 123> Addin1Impl;
			typedef addin<addin2_novscmdt, &__uuidof(addin2_novscmdt), 234> Addin2Impl;
			typedef addin<addin_throwing_novscmdt, &__uuidof(addin_throwing_novscmdt), 234> AddinThrowingImpl;
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
				HRESULT hr1 = a1->raw_OnConnection(0, msaddin::ext_cm_AfterStartup, 0, 0);
				
				// ASSERT
				Assert::IsTrue(1 == addin1_novscmdt::instance_count);
				Assert::IsTrue(0 == addin2_novscmdt::instance_count);
				Assert::IsTrue(S_OK == hr1);
				
				// ACT
				Addin2Impl::CreateInstance(&a2);
				Addin2Impl::CreateInstance(&a3);
				HRESULT hr2 = a2->raw_OnConnection(0, msaddin::ext_cm_Startup, 0, 0);
				HRESULT hr3 = a3->raw_OnConnection(0, msaddin::ext_cm_External, 0, 0);
				
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
				HRESULT hr = a->raw_OnConnection(0, msaddin::ext_cm_AfterStartup, 0, 0);
				
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
				a1->raw_OnConnection(0, msaddin::ext_cm_AfterStartup, 0, 0);
				a2->raw_OnConnection(0, msaddin::ext_cm_Startup, 0, 0);
				a3->raw_OnConnection(0, msaddin::ext_cm_External, 0, 0);

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
				a1->raw_OnConnection(0, msaddin::ext_cm_AfterStartup, 0, 0);
				a2->raw_OnConnection(0, msaddin::ext_cm_Startup, 0, 0);
				a3->raw_OnConnection(0, msaddin::ext_cm_External, 0, 0);

				// ACT
				a1->raw_OnDisconnection(msaddin::ext_dm_HostShutdown, 0);
				
				// ASSERT
				Assert::IsTrue(0 == addin1_novscmdt::instance_count);
				Assert::IsTrue(2 == addin2_novscmdt::instance_count);

				// ACT
				a2->raw_OnDisconnection(msaddin::ext_dm_HostShutdown, 0);
				
				// ASSERT
				Assert::IsTrue(0 == addin1_novscmdt::instance_count);
				Assert::IsTrue(1 == addin2_novscmdt::instance_count);

				// ACT
				a3->raw_OnDisconnection(msaddin::ext_dm_UserClosed, 0);
				
				// ASSERT
				Assert::IsTrue(0 == addin1_novscmdt::instance_count);
				Assert::IsTrue(0 == addin2_novscmdt::instance_count);
			}


			[TestMethod]
			void DisconnectionReturnsS_OK()
			{
				// INIT
				CComPtr<msaddin::IDTExtensibility2> a;

				Addin1Impl::CreateInstance(&a);
				a->raw_OnConnection(0, msaddin::ext_cm_AfterStartup, 0, 0);

				// ACT / ASSERT
				Assert::IsTrue(S_OK == a->raw_OnDisconnection(msaddin::ext_dm_HostShutdown, 0));
			}
		};
	}
}
