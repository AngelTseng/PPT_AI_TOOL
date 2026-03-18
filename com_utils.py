from contextlib import contextmanager
import pythoncom


@contextmanager
def com_session():
    pythoncom.CoInitialize()
    try:
        yield
    finally:
        pythoncom.CoUninitialize()