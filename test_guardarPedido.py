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


        # 7. Prueba para: guardar_pedido() -> Que calcule el total y descuente existencias
def test_guardar_pedido(setup_excel):
    # Crear producto base
    wb = openpyxl.load_workbook(setup_excel)
    ws_prod = wb["productos"]
    ws_prod.append([1, "Silla", 100.0, 10])  # Stock inicial = 10
    wb.save(setup_excel)
    
    # Guardar pedido de 3 Sillas
    cant_comprada = 3
    wb = openpyxl.load_workbook(setup_excel)
    ws_prod = wb["productos"]
    
    # Descontar existencias y calcular total
    precio_unitario = float(ws_prod.cell(row=2, column=3).value)
    stock_actual = int(ws_prod.cell(row=2, column=4).value)
    
    ws_prod.cell(row=2, column=4, value=stock_actual - cant_comprada)
    
    ws_ped = wb["pedidos"]
    nuevo_id_ped = obtener_siguiente_id(ws_ped)
    total_pedido = precio_unitario * cant_comprada
    ws_ped.append([nuevo_id_ped, "Carlos", "Silla", cant_comprada, total_pedido])
    wb.save(setup_excel)
    
    # Verificaciones
    wb_read = openpyxl.load_workbook(setup_excel)
    assert wb_read["productos"].cell(row=2, column=4).value == 7     # 10 - 3 = 7
    assert wb_read["pedidos"].cell(row=2, column=5).value == 300.0  # 100 * 3 = 300