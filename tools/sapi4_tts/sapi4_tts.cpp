// sapi4_tts.exe - speaks with old SAPI 4 voices (Microsoft Sam/Mike/Mary 4.0, L&H TruVoice - BonziBuddy's
// "Adult Male #2" - ...) into WAV files, for TF2Speech. 32-bit, because SAPI 4 engines only exist as 32-bit.
// MIT licence (part of TF2Speech). No Microsoft headers needed: the few SAPI 4 interfaces used are declared below
// with the IDs the SAPI 4 runtime registers in Windows.
//
// Runs as a helper: reads one command per line on stdin (UTF-8, fields separated by tabs), answers on stdout.
//   LIST                                          -> "VOICE\t<mode name>\t<speaker>\t<maker>\t<gender>" lines, then "END"
//   SPEAK\t<mode name>\t<pitch|->\t<speed|->\t<wav path>\t<text>   -> "OK" or "ERR <reason>"
//   QUIT
// Text is spoken as plain text (no SAPI 4 tags are interpreted).

#define WIN32_LEAN_AND_MEAN
#include <windows.h>
#include <objbase.h>
#include <stdio.h>
#include <string>
#include <vector>

typedef unsigned __int64 QWORD;

// ---------------------------------------------------------------- SAPI 4 declarations
#define TTSI_NAMELEN 262
#define LANG_LEN 64

struct LANGUAGEW4 { LANGID LanguageID; WCHAR szDialect[LANG_LEN]; };
struct TTSMODEINFOW4 {
    GUID gEngineID;
    WCHAR szMfgName[TTSI_NAMELEN];
    WCHAR szProductName[TTSI_NAMELEN];
    GUID gModeID;
    WCHAR szModeName[TTSI_NAMELEN];
    LANGUAGEW4 language;
    WCHAR szSpeaker[TTSI_NAMELEN];
    WCHAR szStyle[TTSI_NAMELEN];
    WORD wGender;
    WORD wAge;
    DWORD dwFeatures;
    DWORD dwInterfaces;
    DWORD dwEngineFeatures;
};
struct SDATA4 { PVOID pData; DWORD dwSize; };

static const CLSID CLSID_TTSEnumerator = {0xD67C0280, 0xC743, 0x11cd, {0x80, 0xE5, 0x00, 0xAA, 0x00, 0x3E, 0x4B, 0x50}};
static const CLSID CLSID_AudioDestFile = {0xD4623720, 0xE4B9, 0x11cf, {0x8D, 0x56, 0x00, 0xA0, 0xC9, 0x03, 0x4A, 0x7E}};
static const IID IID_ITTSEnumW = {0x6B837B20, 0x4A47, 0x101B, {0x93, 0x1A, 0x00, 0xAA, 0x00, 0x47, 0xBA, 0x4F}};
static const IID IID_ITTSCentralW = {0x28016060, 0x4A47, 0x101B, {0x93, 0x1A, 0x00, 0xAA, 0x00, 0x47, 0xBA, 0x4F}};
static const IID IID_ITTSAttributesW = {0x1287A280, 0x4A47, 0x101B, {0x93, 0x1A, 0x00, 0xAA, 0x00, 0x47, 0xBA, 0x4F}};
static const IID IID_ITTSNotifySinkW = {0xC0FA8F40, 0x4A46, 0x101B, {0x93, 0x1A, 0x00, 0xAA, 0x00, 0x47, 0xBA, 0x4F}};
static const IID IID_ITTSBufNotifySink = {0xe4963d40, 0xc743, 0x11cd, {0x80, 0xe5, 0x00, 0xaa, 0x00, 0x3e, 0x4b, 0x50}};
static const IID IID_IAudioFile = {0xFD7C2320, 0x3D6D, 0x11b9, {0xC0, 0x00, 0xFE, 0xD6, 0xCB, 0xA3, 0xB1, 0xA9}};

struct ITTSCentralW4;
struct ITTSEnumW4 : public IUnknown {
    virtual HRESULT STDMETHODCALLTYPE Next(ULONG celt, TTSMODEINFOW4 *rgelt, ULONG *pceltFetched) = 0;
    virtual HRESULT STDMETHODCALLTYPE Skip(ULONG celt) = 0;
    virtual HRESULT STDMETHODCALLTYPE Reset() = 0;
    virtual HRESULT STDMETHODCALLTYPE Clone(ITTSEnumW4 **ppenum) = 0;
    virtual HRESULT STDMETHODCALLTYPE Select(GUID gModeID, ITTSCentralW4 **ppCentral, IUnknown *pAudio) = 0;
};
struct ITTSCentralW4 : public IUnknown {
    virtual HRESULT STDMETHODCALLTYPE Inject(const WCHAR *pszTag) = 0;
    virtual HRESULT STDMETHODCALLTYPE ModeGet(TTSMODEINFOW4 *pInfo) = 0;
    virtual HRESULT STDMETHODCALLTYPE Phoneme(DWORD eCharset, DWORD dwFlags, SDATA4 dText, SDATA4 *pdPhoneme) = 0;
    virtual HRESULT STDMETHODCALLTYPE PosnGet(QWORD *pqTime) = 0;
    virtual HRESULT STDMETHODCALLTYPE TextData(DWORD eCharset, DWORD dwFlags, SDATA4 dText, PVOID pNotify, IID iidNotify) = 0;
    virtual HRESULT STDMETHODCALLTYPE ToFileTime(QWORD *pqTime, FILETIME *pft) = 0;
    virtual HRESULT STDMETHODCALLTYPE AudioPause() = 0;
    virtual HRESULT STDMETHODCALLTYPE AudioResume() = 0;
    virtual HRESULT STDMETHODCALLTYPE AudioReset() = 0;
    virtual HRESULT STDMETHODCALLTYPE Register(PVOID pNotify, IID iidNotify, DWORD *pdwKey) = 0;
    virtual HRESULT STDMETHODCALLTYPE UnRegister(DWORD dwKey) = 0;
};
struct ITTSAttributesW4 : public IUnknown {
    virtual HRESULT STDMETHODCALLTYPE PitchGet(WORD *pw) = 0;
    virtual HRESULT STDMETHODCALLTYPE PitchSet(WORD w) = 0;
    virtual HRESULT STDMETHODCALLTYPE RealTimeGet(DWORD *pdw) = 0;
    virtual HRESULT STDMETHODCALLTYPE RealTimeSet(DWORD dw) = 0;
    virtual HRESULT STDMETHODCALLTYPE SpeedGet(DWORD *pdw) = 0;
    virtual HRESULT STDMETHODCALLTYPE SpeedSet(DWORD dw) = 0;
    virtual HRESULT STDMETHODCALLTYPE VolumeGet(DWORD *pdw) = 0;
    virtual HRESULT STDMETHODCALLTYPE VolumeSet(DWORD dw) = 0;
};
struct IAudioFile4 : public IUnknown {
    virtual HRESULT STDMETHODCALLTYPE Register(IUnknown *pSink) = 0;
    virtual HRESULT STDMETHODCALLTYPE Set(LPCWSTR pszFile, DWORD dwFlags) = 0;
    virtual HRESULT STDMETHODCALLTYPE Add(LPCWSTR pszFile, DWORD dwFlags) = 0;
    virtual HRESULT STDMETHODCALLTYPE Flush() = 0;
    virtual HRESULT STDMETHODCALLTYPE RealTimeSet(WORD w) = 0;
    virtual HRESULT STDMETHODCALLTYPE RealTimeGet(WORD *pw) = 0;
};

// Told by the engine when it has finished writing the audio
struct NotifySink : public IUnknown {
    volatile LONG done = 0;
    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void **ppv) override {
        if (riid == IID_IUnknown || riid == IID_ITTSNotifySinkW) { *ppv = this; return S_OK; }
        *ppv = nullptr;
        return E_NOINTERFACE;
    }
    ULONG STDMETHODCALLTYPE AddRef() override { return 1; }   // lives for the whole program
    ULONG STDMETHODCALLTYPE Release() override { return 1; }
    virtual HRESULT STDMETHODCALLTYPE AttribChanged(DWORD) { return S_OK; }
    virtual HRESULT STDMETHODCALLTYPE AudioStart(QWORD) { return S_OK; }
    virtual HRESULT STDMETHODCALLTYPE AudioStop(QWORD) { InterlockedExchange(&done, 1); return S_OK; }
    virtual HRESULT STDMETHODCALLTYPE Visual(QWORD, WCHAR, WCHAR, DWORD, void *) { return S_OK; }
};

// ---------------------------------------------------------------- helpers
static std::wstring widen(const std::string &s) {
    if (s.empty()) return std::wstring();
    int n = MultiByteToWideChar(CP_UTF8, 0, s.data(), (int)s.size(), nullptr, 0);
    std::wstring w(n, L'\0');
    MultiByteToWideChar(CP_UTF8, 0, s.data(), (int)s.size(), &w[0], n);
    return w;
}
static std::string narrow(const WCHAR *w) {
    int n = WideCharToMultiByte(CP_UTF8, 0, w, -1, nullptr, 0, nullptr, nullptr);
    std::string s(n > 0 ? n - 1 : 0, '\0');
    if (n > 1) WideCharToMultiByte(CP_UTF8, 0, w, -1, &s[0], n, nullptr, nullptr);
    return s;
}
static std::string clean(std::string s) {  // nothing that would break the tab/line protocol
    for (auto &c : s) if (c == '\t' || c == '\r' || c == '\n') c = ' ';
    return s;
}
static std::vector<std::string> split_tabs(const std::string &line, size_t max_fields) {
    std::vector<std::string> out;
    size_t start = 0;
    while (out.size() + 1 < max_fields) {
        size_t tab = line.find('\t', start);
        if (tab == std::string::npos) break;
        out.push_back(line.substr(start, tab - start));
        start = tab + 1;
    }
    out.push_back(line.substr(start));
    return out;
}
static void reply(const std::string &s) { fputs(s.c_str(), stdout); fputc('\n', stdout); fflush(stdout); }

static std::vector<TTSMODEINFOW4> list_modes() {
    std::vector<TTSMODEINFOW4> modes;
    ITTSEnumW4 *e = nullptr;
    if (FAILED(CoCreateInstance(CLSID_TTSEnumerator, nullptr, CLSCTX_ALL, IID_ITTSEnumW, (void **)&e)) || !e) return modes;
    TTSMODEINFOW4 info;
    ULONG got = 0;
    while (e->Next(1, &info, &got) == S_OK && got == 1) modes.push_back(info);
    e->Release();
    return modes;
}

// ---------------------------------------------------------------- the selected voice
static NotifySink g_sink;
static ITTSCentralW4 *g_central = nullptr;
static ITTSAttributesW4 *g_attrs = nullptr;
static IAudioFile4 *g_file = nullptr;
static DWORD g_key = 0;
static std::wstring g_mode;
static WORD g_default_pitch = 0;
static DWORD g_default_speed = 0;

static void release_voice() {
    if (g_central && g_key) g_central->UnRegister(g_key);
    if (g_attrs) g_attrs->Release();
    if (g_central) g_central->Release();
    if (g_file) g_file->Release();
    g_attrs = nullptr; g_central = nullptr; g_file = nullptr; g_key = 0; g_mode.clear();
}

static std::string select_voice(const std::wstring &mode_name) {
    if (g_central && mode_name == g_mode) return "";
    release_voice();
    ITTSEnumW4 *e = nullptr;
    if (FAILED(CoCreateInstance(CLSID_TTSEnumerator, nullptr, CLSCTX_ALL, IID_ITTSEnumW, (void **)&e)) || !e)
        return "SAPI 4 isn't installed";
    TTSMODEINFOW4 info;
    ULONG got = 0;
    bool found = false;
    GUID mode{};
    while (e->Next(1, &info, &got) == S_OK && got == 1) {
        if (_wcsicmp(info.szModeName, mode_name.c_str()) == 0) { mode = info.gModeID; found = true; break; }
    }
    if (!found) { e->Release(); return "no such SAPI 4 voice"; }
    if (FAILED(CoCreateInstance(CLSID_AudioDestFile, nullptr, CLSCTX_ALL, IID_IAudioFile, (void **)&g_file)) || !g_file) {
        e->Release(); return "SAPI 4 audio-to-file component missing";
    }
    HRESULT hr = e->Select(mode, &g_central, g_file);
    e->Release();
    if (FAILED(hr) || !g_central) { release_voice(); return "the voice couldn't be started"; }
    g_file->RealTimeSet(0x100 * 32);  // write the file as fast as it can, not at speaking speed
    g_central->QueryInterface(IID_ITTSAttributesW, (void **)&g_attrs);
    if (g_attrs) { g_attrs->PitchGet(&g_default_pitch); g_attrs->SpeedGet(&g_default_speed); }
    g_central->Register(&g_sink, IID_ITTSNotifySinkW, &g_key);
    g_mode = mode_name;
    return "";
}

static std::string speak(const std::wstring &mode, const std::string &pitch, const std::string &speed,
                         const std::wstring &path, const std::wstring &text) {
    std::string err = select_voice(mode);
    if (!err.empty()) return err;
    if (g_attrs) {
        g_attrs->PitchSet(pitch == "-" ? g_default_pitch : (WORD)atoi(pitch.c_str()));
        g_attrs->SpeedSet(speed == "-" ? g_default_speed : (DWORD)atoi(speed.c_str()));
    }
    if (FAILED(g_file->Set(path.c_str(), 1))) return "couldn't create the WAV file";
    InterlockedExchange(&g_sink.done, 0);
    SDATA4 data;
    data.pData = (PVOID)text.c_str();
    data.dwSize = (DWORD)((text.size() + 1) * sizeof(WCHAR));
    if (FAILED(g_central->TextData(0 /*CHARSET_TEXT*/, 0 /*plain, no tags*/, data, nullptr, IID_ITTSBufNotifySink)))
        return "the voice refused the text";
    // SAPI 4 engines work through window messages on this thread, so keep them flowing until it says done
    ULONGLONG deadline = GetTickCount64() + 30000 + (ULONGLONG)text.size() * 200;
    while (!g_sink.done && GetTickCount64() < deadline) {
        MSG msg;
        while (PeekMessageW(&msg, nullptr, 0, 0, PM_REMOVE)) { TranslateMessage(&msg); DispatchMessageW(&msg); }
        if (!g_sink.done) MsgWaitForMultipleObjects(0, nullptr, FALSE, 10, QS_ALLINPUT);
    }
    g_file->Flush();
    // release the finished file so the TTS can read it
    g_file->Set(L"", 0);
    if (!g_sink.done) { g_central->AudioReset(); return "timed out"; }
    return "";
}

int main() {
    if (FAILED(CoInitialize(nullptr))) { reply("ERR COM"); return 1; }
    static char buf[1 << 16];
    while (fgets(buf, sizeof(buf), stdin)) {
        std::string line(buf);
        while (!line.empty() && (line.back() == '\n' || line.back() == '\r')) line.pop_back();
        if (line == "QUIT") break;
        if (line == "LIST") {
            for (auto &m : list_modes()) {
                reply("VOICE\t" + clean(narrow(m.szModeName)) + "\t" + clean(narrow(m.szSpeaker)) + "\t" +
                      clean(narrow(m.szMfgName)) + "\t" + std::to_string(m.wGender));
            }
            reply("END");
            continue;
        }
        if (line.rfind("SPEAK\t", 0) == 0) {
            auto f = split_tabs(line, 6);
            if (f.size() < 6) { reply("ERR bad command"); continue; }
            std::string err = speak(widen(f[1]), f[2], f[3], widen(f[4]), widen(f[5]));
            reply(err.empty() ? "OK" : "ERR " + err);
            continue;
        }
        reply("ERR unknown command");
    }
    release_voice();
    CoUninitialize();
    return 0;
}
