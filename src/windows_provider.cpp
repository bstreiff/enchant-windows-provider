// enchant_windows - an Enchant provider plugin that uses the Windows 8
//                   spell check API.
//
// Copyright (c) 2015 Brenda Streiff
//
// This library is free software; you can redistribute it and/or modify it
// under the terms of the GNU Lesser General Public License as published by
// the Free Software Foundation; either version 2.1 of the License, or (at
// your option) any later version.
//
// This library is distributed in the hope that it will be useful, but
// WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY
// or FITNESS FOR A PARTICULAR PURPOSE. See the GNU Lesser General Public
// License for more details.
//
// You should have received a copy of the GNU Lesser General Public License
// along with this library; if not, write to the Free Software Foundation,
// Inc., 59 Temple Place - Suite 330, Boston, MA 02110 - 1301, USA.

#include "enchant-provider.h"

#include <comdef.h>
#include <future>
#include <functional>
#include <memory>
#include <mutex>
#include <stdlib.h>
#include <spellcheck.h>
#include <thread>
#include <wtypes.h>
#include <wrl.h>

// ATL has a wider array of COM smart pointer classes, but isn't available
// in the Visual Studio Express editions. The WRL ComPtr works for us though.
using Microsoft::WRL::ComPtr;

ENCHANT_PLUGIN_DECLARE("windows")

// RAII class to wrap CoIninitalizeEx.
struct CoInitializer
{
	CoInitializer() :
		hr(CoInitializeEx(nullptr, COINIT_MULTITHREADED))
	{}
	~CoInitializer() _NOEXCEPT
	{
		if (SUCCEEDED(hr))
			CoUninitialize();
	}
	HRESULT hr;
};

// COM thread dispatcher. We're a DLL, and thus we're not allowed to call
// CoInitialize* on the application's thread. Larry Osterman has an article:
// http://blogs.msdn.com/b/larryosterman/archive/2004/05/12/130541.aspx
// So, punt all COM stuff to a worker thread under our control. This class
// provides a single-object queue for serializing methods on a worker thread.
// This could be replaced with a proper queue, but in practice this seems
// to work well enough.
class CoThreadDispatcher
{
public:
	CoThreadDispatcher() :
		running(false),
		dispatch_thread(std::thread(&CoThreadDispatcher::threadProc, this))
	{ }
	~CoThreadDispatcher()
	{
		dispatch([&]() { running = false; });
		dispatch_thread.join();
	}

	// Dispatch callable object 'f' on the COM worker thread. Blocks until
	// f returns.
	template<typename F>
	typename std::result_of<F()>::type dispatch(F&& f)
	{
		typedef typename std::result_of<F()>::type ResultType;

		// Package the callable object so we can get a future.
		std::packaged_task<ResultType(void)> task(std::forward<F>(f));
		auto result = task.get_future();

		{
			// Acquire the lock so we can queue the work.
			std::unique_lock<std::mutex> lock(processing_mutex);
			dispatched_function = [&task]() { task(); };

			// Tell the thread to go.
			dispatch_begin.notify_all();
		}

		// Wait for the future to have a result.
		result.wait();

		return result.get();
	}

private:
	void threadProc()
	{
		std::unique_lock<std::mutex> lock(processing_mutex);

		// Initialize COM in this thread.
		CoInitializer comInit;
		// We're good to go.
		running = true;

		while (running)
		{
			// Wait for work.
			dispatch_begin.wait(lock);
			// Do the work.
			dispatched_function();
		}
	}
	bool running;
	std::mutex processing_mutex;
	std::condition_variable dispatch_begin;
	std::function<void(void)> dispatched_function;
	std::thread dispatch_thread;
};

static std::mutex com_dispatcher_mutex;
static std::unique_ptr<CoThreadDispatcher> com_dispatcher;
static uint32_t com_dispatcher_refcount(0);

static void com_dispatcher_addref()
{
	std::lock_guard<std::mutex> lock(com_dispatcher_mutex);
	if (com_dispatcher_refcount == 0)
		com_dispatcher = std::make_unique<CoThreadDispatcher>();
	++com_dispatcher_refcount;
}

static void com_dispatcher_release()
{
	std::lock_guard<std::mutex> lock(com_dispatcher_mutex);
	if (com_dispatcher_refcount == 1)
		com_dispatcher.reset();
	--com_dispatcher_refcount;
}

// There is a MAX_WORD_LENGTH constant mentioned in the MSDN documentation
// for ISpellChecker::Add and AutoCorrect, but it's not actually in a header.
static const size_t kMaxWordLength = 128;
static const size_t kMaxUTF8WordLengthInBytes = kMaxWordLength*4;

struct ProviderUserData
{
	ComPtr<ISpellCheckerFactory> spellCheckerFactory;
};

struct DictUserData
{
	ComPtr<ISpellChecker> spellChecker;
};

static inline ProviderUserData* userdata(EnchantProvider* provider)
{
	return reinterpret_cast<ProviderUserData*>(provider->user_data);
}

static inline DictUserData* userdata(EnchantDict* dict)
{
	return reinterpret_cast<DictUserData*>(dict->user_data);
}

// Convert a UTF-8 string (from Enchant) to a new UTF-16 string (to pass into Windows API functions.)
static std::unique_ptr<wchar_t[]> copy_utf8_to_utf16(
	const char* u8str,
	size_t len)
{
	if (len > kMaxUTF8WordLengthInBytes)
		return nullptr;

	int requiredLengthInCharacters = MultiByteToWideChar(CP_UTF8, 0, u8str, static_cast<int>(len), nullptr, 0);
	auto newString = std::make_unique<wchar_t[]>(requiredLengthInCharacters+1);
	if (!newString)
		return nullptr;

	MultiByteToWideChar(CP_UTF8, 0, u8str, static_cast<int>(len), newString.get(), requiredLengthInCharacters);
	newString[requiredLengthInCharacters] = L'\0';
	return newString;
}

// Convert a UTF-16 (from Windows) to a new UTF-8 string (to give back to Enchant).
static std::unique_ptr<char[]> copy_utf16_to_utf8(
	const wchar_t* u16str,
	size_t len)
{
	if (len > kMaxWordLength)
		return nullptr;

	int requiredLengthInCharacters = WideCharToMultiByte(CP_UTF8, 0, u16str, static_cast<int>(len), nullptr, 0, nullptr, nullptr);
	auto newString = std::make_unique<char[]>(requiredLengthInCharacters+1);

	if (!newString)
		return nullptr;

	WideCharToMultiByte(CP_UTF8, 0, u16str, static_cast<int>(len), newString.get(), requiredLengthInCharacters, nullptr, nullptr);
	newString[requiredLengthInCharacters] = '\0';
	return newString;
}

// Convert an enumerator represented by an IEnumString into a null-terminated vector
// of null-terminated UTF-8 strings.
static void copy_string_list_from_enumerator(
	IEnumString* enumerator,
	char*** string_list,
	size_t* count)
{
	// Count the number of entries.
	size_t enumCount = 0;
	while (S_OK == enumerator->Skip(1))
		++enumCount;
	enumerator->Reset();

	auto list = std::make_unique<char*[]>(enumCount + 1);
	auto OleStringDeleter = [](LPOLESTR s) { CoTaskMemFree(s); };
	for (size_t i = 0; i < enumCount; ++i)
	{
		LPOLESTR nameRaw = nullptr;
		HRESULT hr = enumerator->Next(1, &nameRaw, nullptr);
		std::unique_ptr<OLECHAR, decltype(OleStringDeleter)> name(nameRaw, OleStringDeleter);

		if (FAILED(hr))
			return;

		list[i] = copy_utf16_to_utf8(name.get(), wcsnlen_s(name.get(), kMaxWordLength) + 1).release();
	}
	list[enumCount] = nullptr;

	*string_list = list.release();
	*count = enumCount;
}

// Enchant tags are of the form "en_US", Windows spellcheck languages are of the form "en-US".
static std::unique_ptr<wchar_t[]> copy_from_enchant_tag_to_windows_language(const char* tag)
{
	auto langTag = copy_utf8_to_utf16(tag, strnlen_s(tag, kMaxUTF8WordLengthInBytes)+1);
	if (!langTag)
		return nullptr;

	wchar_t* itr = langTag.get();
	while (*itr != '\0')
	{
		if (*itr == L'_')
			*itr = L'-';
		++itr;
	}

	return langTag;
}

// Returns 0 if word is correctly spelled, positive if not, negative if error.
static int windows_dict_check(
	EnchantDict* dict,
	const char *const word,
	size_t len)
{
	return com_dispatcher->dispatch([=]() -> int {
		auto utf16Word = copy_utf8_to_utf16(word, len);
		if (!utf16Word)
			return -1;

		ComPtr<IEnumSpellingError> errors;
		HRESULT hr = userdata(dict)->spellChecker->Check(utf16Word.get(), errors.GetAddressOf());
		if (FAILED(hr))
			return -1;

		// A correct 'test' returns an empty (not a null) enumeration.
		ComPtr<ISpellingError> error;
		hr = errors->Next(error.GetAddressOf());
		if (hr == S_OK)
		{
			// At least one error.
			return 1;
		}
		else
		{
			// No errors.
			return 0;
		}
	});
}

// Return a vector of strings that are suggestions for a word. Return null
// if no suggestions are available.
static char** windows_dict_suggest(
	EnchantDict* dict,
	const char *const word,
	size_t len,
	size_t* out_n_suggs)
{
	return com_dispatcher->dispatch([=]() -> char** {
		auto utf16Word = copy_utf8_to_utf16(word, len);
		if (!utf16Word)
			return nullptr;

		ComPtr<IEnumString> suggestionEnumerator;
		HRESULT hr = userdata(dict)->spellChecker->Suggest(utf16Word.get(), suggestionEnumerator.GetAddressOf());

		if (FAILED(hr))
			return nullptr;

		// If we returned S_FALSE, the word was spelled correctly and there are no suggestions.
		if (hr == S_FALSE)
			return nullptr;

		char** suggestions = nullptr;
		copy_string_list_from_enumerator(suggestionEnumerator.Get(), &suggestions, out_n_suggs);
		return suggestions;
	});
}

// Add a word to the user's personal dictionary.
static void windows_dict_add_to_personal(
	EnchantDict* dict,
	const char *const word,
	size_t len)
{
	com_dispatcher->dispatch([=]() -> void {
		auto utf16Word = copy_utf8_to_utf16(word, len);
		if (!utf16Word)
			return;

		HRESULT hr = userdata(dict)->spellChecker->Add(utf16Word.get());
		if (FAILED(hr))
			return;
	});
}

// Store a replacement for a particular spelling.
static void windows_dict_store_replacement(
	EnchantDict* dict,
	const char* const mis,
	size_t mis_len,
	const char* const cor,
	size_t cor_len)
{
	com_dispatcher->dispatch([=]() -> void {
		auto from = copy_utf8_to_utf16(mis, mis_len);
		if (!from)
			return;

		auto to = copy_utf8_to_utf16(cor, cor_len);
		if (!to)
			return;

		HRESULT hr = userdata(dict)->spellChecker->AutoCorrect(from.get(), to.get());
		if (FAILED(hr))
			return;
	});
}

// Add a word to the user's exclusion list.
static void windows_dict_add_to_exclude(
	EnchantDict* dict,
	const char* const word,
	size_t len)
{
	com_dispatcher->dispatch([=]() -> void {
		auto utf16Word = copy_utf8_to_utf16(word, len);
		if (!utf16Word)
			return;

		HRESULT hr = userdata(dict)->spellChecker->Ignore(utf16Word.get());
		if (FAILED(hr))
			return;
	});
}

// Request dictionary with language tag (such as 'en_US').
static EnchantDict* windows_provider_request_dict(
	EnchantProvider* provider,
	const char* const tag)
{
	return com_dispatcher->dispatch([=]() -> EnchantDict* {
		if (!userdata(provider)->spellCheckerFactory)
			return nullptr;

		auto wtag = copy_from_enchant_tag_to_windows_language(tag);
		if (!wtag)
			return nullptr;

		auto dict = std::make_unique<EnchantDict>();
		dict->check = windows_dict_check;
		dict->suggest = windows_dict_suggest;
		dict->add_to_personal = windows_dict_add_to_personal;
		dict->add_to_session = nullptr;
		dict->store_replacement = windows_dict_store_replacement;
		dict->add_to_exclude = windows_dict_add_to_exclude;

		auto dictdata = std::make_unique<DictUserData>();

		HRESULT hr = userdata(provider)->spellCheckerFactory->CreateSpellChecker(wtag.get(), dictdata->spellChecker.GetAddressOf());
		if (FAILED(hr))
			return nullptr;

		dict->user_data = dictdata.release();

		return dict.release();
	});
}

// Destroy an EnchantDict.
static void windows_provider_dispose_dict(
	EnchantProvider* provider,
	EnchantDict* dict)
{
	com_dispatcher->dispatch([=]() -> void {
		if (dict->user_data)
		{
			delete userdata(dict);
		}
		delete dict;
	});
}

// List all dictionary tags that are available from this provider.
static char** windows_provider_list_dicts(
	EnchantProvider* provider,
	size_t* out_n_dicts)
{
	return com_dispatcher->dispatch([=]() -> char** {
		if (!userdata(provider)->spellCheckerFactory)
			return nullptr;

		ComPtr<IEnumString> langEnumerator;
		HRESULT hr = userdata(provider)->spellCheckerFactory->get_SupportedLanguages(langEnumerator.GetAddressOf());
		if (FAILED(hr))
			return nullptr;

		char** langs = nullptr;
		copy_string_list_from_enumerator(langEnumerator.Get(), &langs, out_n_dicts);
		return langs;
	});
}

// Return whether or not a dictionary with a particular tag exists.
static int windows_provider_dictionary_exists(
	EnchantProvider* provider,
	const char* const tag)
{
	return com_dispatcher->dispatch([=]() -> int {
		if (!userdata(provider)->spellCheckerFactory)
			return -1;

		auto wtag = copy_from_enchant_tag_to_windows_language(tag);
		if (!wtag)
			return -1;

		BOOL isSupported = FALSE;
		userdata(provider)->spellCheckerFactory->IsSupported(wtag.get(), &isSupported);
		return (isSupported != FALSE);
	});
}

// Free a string list returned by windows_dict_suggest or windows_provider_list_dicts.
// The list and the items within were allocated with make_unique, so use the same
// deleter.
static void windows_provider_free_string_list(
	EnchantProvider* provider,
	char** str_list)
{
	com_dispatcher->dispatch([=]() -> void {
		if (str_list)
		{
			char** str = str_list;
			while (*str != nullptr)
			{
				std::default_delete<char[]>()(*str);
				++str;
			}
			std::default_delete<char*[]>()(str_list);
		}
	});
}

// Dispose a provider.
//
// Also decrements (and possibly destroys) the COM thread.
static void windows_provider_dispose(EnchantProvider* provider)
{
	com_dispatcher->dispatch([=]() -> void {
		if (provider->user_data)
		{
			ProviderUserData* providerdata = reinterpret_cast<ProviderUserData*>(provider->user_data);
			delete providerdata;
		}
		delete provider;
	});

	// One less provider using the dispatcher.
	com_dispatcher_release();
}


static const char* windows_provider_identify(EnchantProvider* provider) _NOEXCEPT
{
	return "windows";
}

static const char* windows_provider_describe(EnchantProvider* self) _NOEXCEPT
{
	return "Windows Provider";
}

#ifdef __cplusplus
extern "C" {
#endif

// Create a new provider. Can also create the COM thread.
__declspec(dllexport) EnchantProvider* init_enchant_provider() _NOEXCEPT
{
	// We're creating a dispatcher.
	com_dispatcher_addref();

	auto newProvider = com_dispatcher->dispatch([&]() -> EnchantProvider* {

		auto provider = std::make_unique<EnchantProvider>();
		provider->dispose = windows_provider_dispose;
		provider->request_dict = windows_provider_request_dict;
		provider->dispose_dict = windows_provider_dispose_dict;
		provider->dictionary_exists = windows_provider_dictionary_exists;
		provider->identify = windows_provider_identify;
		provider->describe = windows_provider_describe;
		provider->list_dicts = windows_provider_list_dicts;
		provider->free_string_list = windows_provider_free_string_list;

		auto userdata = std::make_unique<ProviderUserData>();

		HRESULT hr = CoCreateInstance(
			__uuidof(SpellCheckerFactory),
			nullptr,
			CLSCTX_INPROC_SERVER,
			__uuidof(ISpellCheckerFactory),
			reinterpret_cast<PVOID*>(userdata->spellCheckerFactory.GetAddressOf()));

		provider->user_data = userdata.release();

		return provider.release();
	});

	// We tried, but didn't get a new provider.
	if (!newProvider)
		com_dispatcher_release();

	return newProvider;
}

#ifdef __cplusplus
}
#endif

