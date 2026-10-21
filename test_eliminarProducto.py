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


        # 5. Prueba para: eliminar_producto() -> Que elimine el producto indicado
def test_eliminar_producto(setup_excel):
    wb = openpyxl.load_workbook(setup_excel)
    ws = wb["productos"]
    ws.append([1, "Monitor", 800.0, 5])
    ws.append([2, "Impresora", 600.0, 3])
    wb.save(setup_excel)
    
    # Eliminar producto ID 1
    prod_id = 1
    wb = openpyxl.load_workbook(setup_excel)
    ws = wb["productos"]
    for idx, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        if row[0] == prod_id:
            ws.delete_rows(idx)
            break
    wb.save(setup_excel)
    
    # Verificar que solo queda 1 registro de datos
    wb_read = openpyxl.load_workbook(setup_excel)
    ws_read = wb_read["productos"]
    filas = list(ws_read.iter_rows(values_only=True))
    assert len(filas) == 2
    assert filas[1][0] == 2  # Quedó la impresora con ID 2