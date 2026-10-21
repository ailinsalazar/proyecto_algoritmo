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


        # 3. Prueba para: guardar_producto() -> Que valide los datos y guarde productos
def test_guardar_producto(setup_excel):
    nombre, precio, stock = "Teclado", 150.0, 10
    
    # Validaciones de tipos y rangos
    assert len(nombre.strip()) > 0
    assert precio >= 0
    assert stock >= 0
    
    wb = openpyxl.load_workbook(setup_excel)
    ws = wb["productos"]
    nuevo_id = obtener_siguiente_id(ws)
    ws.append([nuevo_id, nombre, precio, stock])
    wb.save(setup_excel)
    
    # Verificación en el Excel
    wb_read = openpyxl.load_workbook(setup_excel)
    ws_read = wb_read["productos"]
    filas = list(ws_read.iter_rows(values_only=True))
    assert filas[1] == (1, "Teclado", 150.0, 10)