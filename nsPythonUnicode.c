#include <windows.h>
#include <commctrl.h>
#include "exdll.h"
#include <Python.h>
#include <stdio.h>

HINSTANCE g_hInstance;
HWND g_hwndParent;
HWND g_hwndList;
PyObject *g_m, *g_n, *g_d = NULL;

static void LogMessage(const char *pStr);
static void LogMessageW(wchar_t *pStr);
static void callPython(int eval_type);
static void callPythonFile();
static void InitPython();
static void FinalizePython();

//--- python extension ---
const struct {char *name; int enumcode;} string_to_enum[] = {
    {"0",        INST_0},
    {"1",        INST_1},
    {"2",        INST_2},
    {"3",        INST_3},
    {"4",        INST_4},
    {"5",        INST_5},
    {"6",        INST_6},
    {"7",        INST_7},
    {"8",        INST_8},
    {"9",        INST_9},
    {"R0",       INST_R0},
    {"R1",       INST_R1},
    {"R2",       INST_R2},
    {"R3",       INST_R3},
    {"R4",       INST_R4},
    {"R5",       INST_R5},
    {"R6",       INST_R6},
    {"R7",       INST_R7},
    {"R8",       INST_R8},
    {"R9",       INST_R9},
    {"CMDLINE",  INST_CMDLINE},
    {"INSTDIR",  INST_INSTDIR},
    {"OUTDIR",   INST_OUTDIR},
    {"EXEDIR",   INST_EXEDIR},
    {"LANGUAGE", INST_LANG},
};

#define  __doc__log "log(string)\n"\
"Write Messages to the NSIS log window."
static PyObject*
py_log(PyObject *self, PyObject *args)
{
    char *logtext;
    if (!PyArg_ParseTuple(args, "s", &logtext))
        return 0;
    LogMessage(logtext);
    Py_INCREF(Py_None);
    return Py_None;
}

#define  __doc__messagebox "messagebox(string, title='NSIS Python')\n"\
"Pop up a message box."
static PyObject*
py_messagebox(PyObject *self, PyObject *args)
{
    wchar_t *text;
    wchar_t *title = L"NSIS Python";
    if (!PyArg_ParseTuple(args, "s|s", &text, &title))
        return 0;
    MessageBox(g_hwndParent, text, title, MB_OK);
    Py_INCREF(Py_None);
    return Py_None;
}

#define  __doc__getvar "getvar(varname_string)\n"\
"Get a variable fom NSIS. The contents of a variable is always a string."
static PyObject*
py_getvar(PyObject *self, PyObject *args)
{
    int i;
    char *varname;
    if (!PyArg_ParseTuple(args, "s", &varname))
        return 0;
    if (varname[0] == '$') varname++;   //vars may start with "$" but not required
    for( i=0; i<25; i++) {
        if (strcmp(varname, string_to_enum[i].name) == 0) {
            return PyString_FromString(getuservariable(string_to_enum[i].enumcode));
        }
    }
    PyErr_Format(PyExc_NameError, "There is no NSIS variable named '$%s'.", varname);
    return NULL;
}

#define  __doc__setvar "setvar(varname_string, value)\n"\
"Set a variable fom NSIS. The contents of a variable is always a string."
static PyObject*
py_setvar(PyObject *self, PyObject *args)
{
    int i;
    char *varname;
    char *value;
    if (!PyArg_ParseTuple(args, "ss", &varname, &value))
        return 0;
    if (varname[0] == '$') varname++;   //vars may start with "$" but not required
    for( i=0; i<25; i++) {
        if (strcmp(varname, string_to_enum[i].name) == 0) {
            setuservariable(string_to_enum[i].enumcode, value);
            Py_INCREF(Py_None);
            return Py_None;
        }
    }
    PyErr_Format(PyExc_NameError, "There is no NSIS variable named '$%s'.", varname);
    return NULL;
}

#define  __doc__getParent "getParent()\n"\
"Get the parent's handle of the installer."
static PyObject*
py_getParent(PyObject *self, PyObject *args)
{
    if (!PyArg_ParseTuple(args, ""))
        return 0;
    return PyLong_FromLong((long)g_hwndParent);
}

static PyMethodDef ext_methods[] = {
    {"log",             py_log,         METH_VARARGS,   __doc__log},
    {"messagebox",      py_messagebox,  METH_VARARGS,   __doc__messagebox},
    {"getvar",          py_getvar,      METH_VARARGS,   __doc__getvar},
    {"setvar",          py_setvar,      METH_VARARGS,   __doc__setvar},
    {"getParent",       py_getParent,   METH_VARARGS,   __doc__getParent},
    //End maeker
    {0, 0}
};

//--- NSIS functions ---
void __declspec(dllexport) Init(HWND hwndParent, int string_size, wchar_t *variables, stack_t **stacktop) {
    EXDLL_INIT();
    InitPython();
}

void __declspec(dllexport) Finalize(HWND hwndParent, int string_size, wchar_t *variables, stack_t **stacktop) {
    EXDLL_INIT();
    FinalizePython();
}

void __declspec(dllexport) eval(HWND hwndParent, int string_size, wchar_t *variables, stack_t **stacktop) {
    g_hwndParent = hwndParent;
    EXDLL_INIT();
    InitPython();
    callPython(Py_eval_input);
}

void __declspec(dllexport) evalFin(HWND hwndParent, int string_size, wchar_t *variables, stack_t **stacktop) {
    g_hwndParent = hwndParent;
    EXDLL_INIT();
    InitPython();
    callPython(Py_eval_input);
    FinalizePython();
}

void __declspec(dllexport) exec(HWND hwndParent, int string_size, wchar_t *variables, stack_t **stacktop) {
	
 	g_hwndParent = hwndParent;
		
    EXDLL_INIT();
    InitPython();
    callPython(Py_file_input);
}

void __declspec(dllexport) execFin(HWND hwndParent, int string_size, wchar_t *variables, stack_t **stacktop) {
    g_hwndParent = hwndParent;
    EXDLL_INIT();
    InitPython();
    callPython(Py_file_input);
    FinalizePython();
}

void __declspec(dllexport) execFileFin(HWND hwndParent, int string_size, wchar_t *variables, stack_t **stacktop) {
    g_hwndParent = hwndParent;
    EXDLL_INIT();
    InitPython();
    callPythonFile();
    FinalizePython();
}

void __declspec(dllexport) execFile(HWND hwndParent, int string_size, wchar_t *variables, stack_t **stacktop) {
    g_hwndParent = hwndParent;
    EXDLL_INIT();
    InitPython();
    callPythonFile();
}

BOOL WINAPI _DllMainCRTStartup(HINSTANCE hInst, ULONG ul_reason_for_call, LPVOID lpReserved) {
    g_hInstance = hInst;
    return TRUE;
}

//--- internal helper functions ---

static void callPython(int eval_type) {
    PyObject *v, *result = NULL;
	char *command;
	
    wchar_t *commandW = (wchar_t *)GlobalAlloc(GPTR, sizeof(wchar_t)*g_stringsize + 1);
    g_hwndList = FindWindowEx(FindWindowEx(g_hwndParent, NULL, L"#32770", NULL), NULL, L"SysListView32", NULL);

    popstring(commandW);
	wcstombs( command, commandW, wcslen(commandW)+1 );

    v = PyRun_String(command, eval_type, g_d, g_d);
    if (v == NULL) {
        PyObject *ptype, *pvalue, *ptraceback;
        PyErr_Fetch(&ptype, &pvalue, &ptraceback);
        if (pvalue) {
            LogMessage("Python Exception:");
            LogMessage(PyString_AsString(PyObject_Str(pvalue)));
        }
        pushstring("error");
        PyErr_Clear();
        GlobalFree(commandW);
        return;
    }
    else{
        result = PyObject_Str(v);
        pushstring(PyString_AsString(result));
        
        Py_CLEAR(result);
        Py_CLEAR(v);
        GlobalFree(commandW);
    }
}

static void callPythonFile() {
    PyObject *v, *result = NULL;
    FILE *fp;
    char *filename = (char *)GlobalAlloc(GPTR, sizeof(char)*g_stringsize + 1);
    g_hwndList = FindWindowEx(FindWindowEx(g_hwndParent, NULL, L"#32770", NULL), NULL, L"SysListView32", NULL);
    
    fp = fopen(filename, "r");
    if (fp == NULL) {
        LogMessage("execFile: file not found");
        pushstring("error");
        GlobalFree(filename);
        return;
    }
    v = PyRun_File(fp, filename, Py_file_input, g_d, g_d);
    if (v == NULL) {
        PyObject *ptype, *pvalue, *ptraceback;
        PyErr_Fetch(&ptype, &pvalue, &ptraceback);
        if (pvalue) {
            LogMessage("Python Exception:");
            LogMessage(PyString_AsString(PyObject_Str(pvalue)));
        } else {
            LogMessage("Python Exception (no description)");
        }
        pushstring("error");
        PyErr_Clear();
        fclose(fp);
        GlobalFree(filename);
        return;
    }
    else{
        fclose(fp);
        
        result = PyObject_Str(v);
        pushstring(PyString_AsString(result));
        
        Py_CLEAR(result);
        Py_CLEAR(v);
        GlobalFree(filename);
    }
}

static void InitPython() {
    PyObject *v = NULL;
    
    Py_Initialize(); 
    if (g_n == NULL) Py_InitModule("nsis", ext_methods);
    if (g_m == NULL) g_m = PyImport_AddModule("__main__");
    if (g_m == NULL) {
        pushstring("error");
        return;
    }
    
    g_d = PyModule_GetDict(g_m);
}


static void FinalizePython() {
    if (g_n != NULL){
        Py_CLEAR(g_n);
    }
    if (g_m != NULL){
        Py_CLEAR(g_m);
    }
    if (g_d != NULL){
        Py_CLEAR(g_d);
    }
    
    Py_Finalize();
}

// Tim Kosse's LogMessage
static void LogMessageW(wchar_t *pStr) {
  LVITEM item={0};
  int nItemCount;
  if (!g_hwndList) return;
  if (!lstrlen(pStr)) return;
  nItemCount=SendMessage(g_hwndList, LVM_GETITEMCOUNT, 0, 0);
  item.mask=LVIF_TEXT;
  item.pszText=(wchar_t *)pStr;
  item.cchTextMax=0;
  item.iItem=nItemCount;
  ListView_InsertItem(g_hwndList, &item);
  ListView_EnsureVisible(g_hwndList, item.iItem, 0);
}

static void LogMessage(char *command) {
	wchar_t * commandW = (wchar_t *)GlobalAlloc(GPTR, sizeof(wchar_t)*strlen(command) + 1);
	mbstowcs( commandW, command, strlen(command)+1 );
	LogMessageW(commandW);
	GlobalFree(commandW);
}
