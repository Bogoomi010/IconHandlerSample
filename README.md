# 🧩 IconHandlerDll - Windows Shell Icon Handler

This DLL provides custom icons for specific file extensions via Windows Shell integration.

---

## 🔧 Register / Unregister the DLL

### ✅ Register the DLL
```
taskkill /f /im explorer.exe
regsvr32 x64\Release\IconHandlerDll.dll
start explorer.exe
```

### Unregister the DLL
```
regsvr32 /u C:\Users\admin\source\IconHandlerDll\x64\Release\IconHandlerDll.dll
taskkill /f /im explorer.exe
start explorer.exe
```

Re-register the DLL (when already registered)
If the DLL is already registered, follow these steps to unregister, rebuild, and register again:
```
regsvr32 /u x64\Release\IconHandlerDll.dll
taskkill /f /im explorer.exe
start explorer.exe

# Rebuild the project (in Visual Studio or via CLI)

taskkill /f /im explorer.exe
regsvr32 x64\Release\IconHandlerDll.dll
start explorer.exe
```
