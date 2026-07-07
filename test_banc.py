import sys
if 'modulo_convertidor' in sys.modules:
    del sys.modules['modulo_convertidor']

sys.path.insert(0, '.')
from modulo_convertidor import _parsear_pdf_bancolombia
import inspect

src = inspect.getsource(_parsear_pdf_bancolombia)
print(f"_ok en codigo: {'_ok' in src}")
print(f"Largo funcion: {len(src)}")

class FakeFile:
    def __init__(self, path):
        self.name = "test.pdf"
        with open(path, "rb") as f:
            self._data = f.read()
    def read(self):
        return self._data
    def seek(self, n):
        pass

ruta = r"C:\Users\ASUS\Downloads\ZIP_16600000605_000000901347233_20260707_07095223.pdf"
f = FakeFile(ruta)
datos = _parsear_pdf_bancolombia(f)
print(f"Registros: {len(datos['registros'])}")