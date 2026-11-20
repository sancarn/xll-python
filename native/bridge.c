#define PY_SSIZE_T_CLEAN
#include <Python.h>
#include <wchar.h>
#include <stdlib.h>

/* Excel owns the process lifetime. Never finalize CPython on XLL unload:
   extension modules and Python-created threads can still hold native state. */
static PyObject *dispatch;
#define API __declspec(dllexport)

static wchar_t *take_error(void) {
    PyObject *exc = PyErr_GetRaisedException();
    PyObject *value = exc ? PyObject_Str(exc) : NULL;
    wchar_t *msg = value ? PyUnicode_AsWideCharString(value, NULL) : NULL;
    wchar_t *copy = _wcsdup(msg ? msg : L"CPython call failed");
    PyMem_Free(msg);
    Py_XDECREF(value);
    Py_XDECREF(exc);
    PyErr_Clear();
    return copy;
}

API wchar_t *xp_start(const wchar_t *home, const wchar_t *exe, const wchar_t *support) {
    if (dispatch) return NULL;
    if (Py_IsInitialized()) return _wcsdup(L"Another add-in has initialized this CPython interpreter.");
    PyConfig config;
    PyConfig_InitIsolatedConfig(&config);
    config.install_signal_handlers = 0;
    config.parse_argv = 0;
    PyStatus status = PyConfig_SetString(&config, &config.home, home);
    if (!PyStatus_Exception(status)) status = PyConfig_SetString(&config, &config.program_name, exe);
    if (!PyStatus_Exception(status)) status = PyConfig_SetString(&config, &config.executable, exe);
    if (!PyStatus_Exception(status)) status = Py_InitializeFromConfig(&config);
    PyConfig_Clear(&config);
    if (PyStatus_Exception(status)) return _wcsdup(L"Could not initialize the packaged CPython runtime.");
    PyObject *dir = PyUnicode_FromWideChar(support, -1);
    PyObject *module = NULL;
    if (dir && PyList_Insert(PySys_GetObject("path"), 0, dir) == 0)
        module = PyImport_ImportModule("_xll_host");
    Py_XDECREF(dir);
    if (module) dispatch = PyObject_GetAttrString(module, "dispatch");
    Py_XDECREF(module);
    wchar_t *error = dispatch ? NULL : take_error();
    PyEval_SaveThread();
    return error;
}

API wchar_t *xp_call(const wchar_t *request, int *ok) {
    PyGILState_STATE gil = PyGILState_Ensure();
    PyObject *arg = PyUnicode_FromWideChar(request, -1);
    PyObject *result = arg ? PyObject_CallOneArg(dispatch, arg) : NULL;
    wchar_t *raw = result ? PyUnicode_AsWideCharString(result, NULL) : NULL;
    *ok = raw != NULL;
    wchar_t *reply = raw ? _wcsdup(raw) : take_error();
    PyMem_Free(raw);
    Py_XDECREF(result);
    Py_XDECREF(arg);
    PyGILState_Release(gil);
    return reply;
}

API void xp_free(void *p) { free(p); }

