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


        # 4. Prueba para: editar_producto() -> Que actualice un producto existente
def test_editar_producto(setup_excel):
    wb = openpyxl.load_workbook(setup_excel)
    ws = wb["productos"]
    ws.append([1, "Mouse Antiguo", 50.0, 20])
    wb.save(setup_excel)
    
    # Modificación simulada del ID 1
    prod_id = 1
    nuevo_nombre, nuevo_precio, nuevo_stock = "Mouse Gamer", 120.0, 15
    
    wb = openpyxl.load_workbook(setup_excel)
    ws = wb["productos"]
    for row in ws.iter_rows(min_row=2):
        if row[0].value == prod_id:
            row[1].value = nuevo_nombre
            row[2].value = nuevo_precio
            row[3].value = nuevo_stock
            break
    wb.save(setup_excel)
    
    # Verificación
    wb_read = openpyxl.load_workbook(setup_excel)
    ws_read = wb_read["productos"]
    assert ws_read.cell(row=2, column=2).value == "Mouse Gamer"
    assert ws_read.cell(row=2, column=3).value == 120.0
    assert ws_read.cell(row=2, column=4).value == 15