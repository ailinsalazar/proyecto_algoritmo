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


        # 2. Prueba para: obtener_siguiente_id(ws) -> Que genere el siguiente ID correctamente
def test_obtener_siguiente_id(setup_excel):
    wb = openpyxl.load_workbook(setup_excel)
    ws = wb["productos"]
    
    # Lista vacía debe dar ID 1
    assert obtener_siguiente_id(ws) == 1
    
    # Con registros debe calcular el consecutivo
    ws.append([1, "Lapicero", 2.50, 100])
    ws.append([2, "Cuaderno", 15.00, 50])
    wb.save(setup_excel)
    
    assert obtener_siguiente_id(ws) == 3