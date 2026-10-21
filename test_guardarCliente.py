import os
import pytest
import openpyxl
from proyectoalgorit import inicializar_excel, obtener_siguiente_id

TEST_EXCEL = "test_inventario_completo.xlsx"

@pytest.fixture
def setup_excel():
    """Limpia y crea un archivo Excel de pruebas antes de cada función."""
    if os.path.exists(TEST_EXCEL):
        os.remove(TEST_EXCEL)
        
    inicializar_excel(TEST_EXCEL)
    yield TEST_EXCEL
    
    if os.path.exists(TEST_EXCEL):
        os.remove(TEST_EXCEL)


        # 6. Prueba para: guardar_cliente() -> Que guarde los datos de un cliente
def test_guardar_cliente(setup_excel):
    nombre, nit, direccion, email = "Carlos Pérez", "123456-7", "Ciudad", "carlos@mail.com"
    
    wb = openpyxl.load_workbook(setup_excel)
    ws = wb["clientes"]
    nuevo_id = obtener_siguiente_id(ws)
    ws.append([nuevo_id, nombre, nit, direccion, email])
    wb.save(setup_excel)
    
    wb_read = openpyxl.load_workbook(setup_excel)
    ws_read = wb_read["clientes"]
    filas = list(ws_read.iter_rows(values_only=True))
    assert filas[1] == (1, "Carlos Pérez", "123456-7", "Ciudad", "carlos@mail.com")