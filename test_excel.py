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


     # 1. Prueba para: inicializar_excel() -> Que se creen las hojas necesarias
def test_inicializar_excel(setup_excel):
    wb = openpyxl.load_workbook(setup_excel)
    hojas = wb.sheetnames
    
    assert "productos" in hojas
    assert "clientes" in hojas
    assert "pedidos" in hojas